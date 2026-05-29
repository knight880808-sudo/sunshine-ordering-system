"""
浏览器 Cookie 持久化登录（7 天）
================================
使用 streamlit-cookies-controller；凭证为 HMAC 签名 JSON，不存密码。

须在页面组件挂载后调用 prepare_auth_from_cookie()（app.route 开头），
不能仅在 init_session 里读 Cookie（那时浏览器尚未回传）。
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import sys
import time
from typing import Any

import streamlit as st

COOKIE_NAME = "sunshine_auth_v1"
COOKIE_CTRL_KEY = "sunshine_cookies"
COOKIE_MAX_AGE_SEC = 7 * 24 * 3600
TOKEN_VERSION = 1

ROLE_BRANCH = "branch"
ROLE_WAREHOUSE = "warehouse"
ROLE_ADMIN = "admin"
VALID_ROLES = frozenset({ROLE_BRANCH, ROLE_WAREHOUSE, ROLE_ADMIN})


def _cookie_secret() -> bytes:
    s = os.getenv("SUNSHINE_AUTH_COOKIE_SECRET", "").strip()
    if not s:
        s = (
            os.getenv("SUNSHINE_ADMIN_PASSWORD", "sunshine")
            + "|sunshine-cookie-v1|"
            + os.getenv("SUNSHINE_WAREHOUSE_PASSWORD", "sunshine888")
        )
    return s.encode("utf-8")


def _get_controller():
    if "_sunshine_cookie_ctrl" not in st.session_state:
        from streamlit_cookies_controller import CookieController

        st.session_state._sunshine_cookie_ctrl = CookieController(key=COOKIE_CTRL_KEY)
    return st.session_state._sunshine_cookie_ctrl


def _is_logged_in_session() -> bool:
    role = st.session_state.get("role")
    if role not in VALID_ROLES:
        return False
    if role == ROLE_BRANCH:
        return bool(st.session_state.get("branch")) and (
            st.session_state.get("account_id") is not None
        )
    return True


def _sign_payload(payload: dict) -> str:
    body = base64.urlsafe_b64encode(
        json.dumps(payload, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    ).decode("ascii")
    sig = hmac.new(
        _cookie_secret(), body.encode("ascii"), hashlib.sha256
    ).hexdigest()
    return f"{body}.{sig}"


def _verify_token(token: str) -> dict | None:
    if not token or "." not in token:
        return None
    body, sig = token.rsplit(".", 1)
    expected = hmac.new(
        _cookie_secret(), body.encode("ascii"), hashlib.sha256
    ).hexdigest()
    if not hmac.compare_digest(expected, sig):
        return None
    try:
        raw = base64.urlsafe_b64decode(body.encode("ascii"))
        data = json.loads(raw.decode("utf-8"))
    except Exception:
        return None
    if not isinstance(data, dict):
        return None
    if int(data.get("v") or 0) != TOKEN_VERSION:
        return None
    if int(data.get("exp") or 0) < int(time.time()):
        return None
    role = (data.get("role") or "").strip()
    if role not in VALID_ROLES:
        return None
    return data


def _build_token(
    *,
    role: str,
    branch: str | None = None,
    account_id: int | None = None,
    account_username: str | None = None,
    account_perms: list[str] | None = None,
    lang: str | None = None,
    default_page: str | None = None,
) -> str:
    payload: dict[str, Any] = {
        "v": TOKEN_VERSION,
        "role": role,
        "exp": int(time.time()) + COOKIE_MAX_AGE_SEC,
    }
    if branch:
        payload["branch"] = branch
    if account_id is not None:
        payload["account_id"] = int(account_id)
    if account_username:
        payload["account_username"] = account_username
    if account_perms is not None:
        payload["account_perms"] = list(account_perms)
    if lang in ("en", "zh"):
        payload["lang"] = lang
    if default_page:
        payload["default_page"] = default_page
    return _sign_payload(payload)


def _app_runtime():
    return sys.modules.get("__main__")


def _query_page_from_url() -> str | None:
    """浏览器地址栏 ?p= 与 History API 同步的页面键。"""
    try:
        v = st.query_params.get("p")
    except Exception:
        return None
    if v is None:
        return None
    if isinstance(v, (list, tuple)):
        v = v[0] if v else None
    s = str(v).strip()
    return s or None


def _branch_allowed_keys(mod) -> frozenset[str]:
    keys = {code for code, _ in mod.BRANCH_PERM_ORDER if mod.has_branch_perm(code)}
    keys.add("order_done")
    return frozenset(keys)


def _resolve_restore_page(mod, role: str, cookie_page: str) -> str:
    """刷新时优先停留在地址栏当前页，其次 Cookie 里记录的上次页。"""
    url_p = _query_page_from_url()
    if role == mod.Role.ADMIN:
        allowed = mod.ADMIN_PAGE_KEYS
    elif role == mod.Role.WAREHOUSE:
        allowed = mod.WAREHOUSE_PAGE_KEYS
    else:
        allowed = _branch_allowed_keys(mod)

    if url_p and url_p in allowed:
        return url_p

    cp = (cookie_page or "").strip()
    if cp and cp in allowed:
        return cp

    if role == mod.Role.WAREHOUSE:
        return "pending"
    if role == mod.Role.ADMIN:
        return "dashboard"
    return mod.first_allowed_branch_page()


def _apply_restored_page(mod, page_key: str) -> None:
    """恢复登录后设置页面；不设置 _nav_from_sidebar，以便 render_* 仍读取 URL。"""
    st.session_state.page = page_key
    try:
        st.query_params[mod.URL_PAGE_QUERY_KEY] = page_key
    except Exception:
        pass


def _read_auth_token() -> str | None:
    ctrl = _get_controller()
    try:
        ctrl.refresh()
    except Exception:
        pass
    token = ctrl.get(COOKIE_NAME)
    if token:
        return str(token)
    try:
        all_c = ctrl.getAll() or {}
        if COOKIE_NAME in all_c and all_c[COOKIE_NAME]:
            return str(all_c[COOKIE_NAME])
    except Exception:
        pass
    pending = st.session_state.get("_sunshine_auth_cookie_pending")
    if pending:
        return str(pending)
    return None


def _restore_branch_session(data: dict) -> bool:
    mod = _app_runtime()
    if mod is None:
        return False
    account_id = data.get("account_id")
    branch = (data.get("branch") or "").strip()
    if account_id is None or not branch:
        return False
    try:
        with mod.db_conn() as conn:
            row = conn.execute(
                "SELECT * FROM user_accounts WHERE id = ?",
                (int(account_id),),
            ).fetchone()
    except Exception:
        return False
    if row is None or row["status"] != mod.ACCOUNT_STATUS_APPROVED:
        return False
    if (row["branch"] or "").strip() != branch:
        return False
    perms = mod._parse_permissions_json(row["permissions"])
    if not perms:
        return False
    st.session_state.role = mod.Role.BRANCH
    st.session_state.branch = branch
    st.session_state.account_id = int(row["id"])
    st.session_state.account_username = row["username"]
    st.session_state.account_perms = perms
    st.session_state.last_role = mod.Role.BRANCH
    st.session_state.last_branch = branch
    page = _resolve_restore_page(
        mod, mod.Role.BRANCH, (data.get("default_page") or "").strip()
    )
    _apply_restored_page(mod, page)
    return True


def _restore_staff_session(data: dict) -> bool:
    mod = _app_runtime()
    if mod is None:
        return False
    role = data["role"]
    if role not in (ROLE_WAREHOUSE, ROLE_ADMIN):
        return False
    st.session_state.role = role
    st.session_state.branch = None
    st.session_state.account_id = None
    st.session_state.account_username = None
    st.session_state.account_perms = []
    st.session_state.last_role = role
    page = _resolve_restore_page(mod, role, (data.get("default_page") or "").strip())
    _apply_restored_page(mod, page)
    return True


def _apply_lang_from_payload(data: dict) -> None:
    lang = data.get("lang")
    if lang in ("en", "zh"):
        st.session_state.lang = lang
    elif st.session_state.get("lang") is None:
        st.session_state.lang = "zh"


def bootstrap_from_cookie() -> bool:
    """从已 refresh 的 Cookie 缓存恢复 session。返回是否恢复成功。"""
    if _is_logged_in_session():
        return False

    token = _read_auth_token()
    if not token:
        return False

    data = _verify_token(token)
    if not data:
        clear_login_cookie()
        return False

    _apply_lang_from_payload(data)
    ok = (
        _restore_branch_session(data)
        if data["role"] == ROLE_BRANCH
        else _restore_staff_session(data)
    )
    if not ok:
        clear_login_cookie()
        return False

    st.session_state.pop("pending_role", None)
    st.session_state.pop("login_branch_context", None)
    st.session_state["_sunshine_restored_from_cookie"] = True
    return True


def prepare_auth_from_cookie() -> None:
    """
    在 route() 最前面调用：挂载 Cookie 组件 → refresh → 必要时 rerun 一次再恢复。
    解决 F5 刷新后仍跳登录页的问题。
    """
    if st.session_state.pop("_explicit_logout", False):
        _get_controller()
        if not st.session_state.get("_sunshine_cookie_bootstrapped"):
            try:
                _get_controller().refresh()
            except Exception:
                pass
            st.session_state._sunshine_cookie_bootstrapped = True
        st.session_state._sunshine_cookie_restore_attempted = True
        return

    if _is_logged_in_session():
        return

    _get_controller()

    if not st.session_state.get("_sunshine_cookie_bootstrapped"):
        try:
            _get_controller().refresh()
        except Exception:
            pass
        time.sleep(0.45)
        st.session_state._sunshine_cookie_bootstrapped = True
        st.rerun()

    if st.session_state.get("_sunshine_cookie_restore_attempted"):
        return

    st.session_state._sunshine_cookie_restore_attempted = True
    if bootstrap_from_cookie():
        st.rerun()


def touch_cookie_page(page_key: str) -> None:
    """侧栏换页时更新 Cookie 中的 default_page（静默，失败则忽略）。"""
    if not _is_logged_in_session() or not (page_key or "").strip():
        return
    try:
        role = st.session_state.get("role")
        persist_login_cookie(
            role=role,
            branch=st.session_state.get("branch"),
            account_id=st.session_state.get("account_id"),
            account_username=st.session_state.get("account_username"),
            account_perms=st.session_state.get("account_perms"),
            default_page=page_key,
        )
    except Exception:
        pass


def persist_login_cookie(
    *,
    role: str,
    branch: str | None = None,
    account_id: int | None = None,
    account_username: str | None = None,
    account_perms: list[str] | None = None,
    default_page: str | None = None,
) -> None:
    if role not in VALID_ROLES:
        return
    lang = st.session_state.get("lang")
    page = (default_page or st.session_state.get("page") or "").strip() or None
    token = _build_token(
        role=role,
        branch=branch,
        account_id=account_id,
        account_username=account_username,
        account_perms=account_perms,
        lang=lang if lang in ("en", "zh") else None,
        default_page=page,
    )
    try:
        ctrl = _get_controller()
        ctrl.set(
            COOKIE_NAME,
            token,
            max_age=float(COOKIE_MAX_AGE_SEC),
            path="/",
            same_site="lax",
        )
        st.session_state["_sunshine_auth_cookie_pending"] = token
        ctrl.refresh()
    except Exception:
        st.session_state["_sunshine_auth_cookie_pending"] = token


def clear_login_cookie() -> None:
    try:
        ctrl = _get_controller()
        ctrl.remove(COOKIE_NAME, path="/", same_site="lax")
        ctrl.refresh()
    except Exception:
        pass
    for k in (
        "_sunshine_auth_cookie_pending",
        "_sunshine_cookie_bootstrapped",
        "_sunshine_cookie_restore_attempted",
        "_sunshine_restored_from_cookie",
        COOKIE_CTRL_KEY,
        "_sunshine_cookie_ctrl",
    ):
        st.session_state.pop(k, None)


def init_auth_session() -> None:
    """兼容入口：实际恢复在 prepare_auth_from_cookie()（route 内）执行。"""
    return
