"""
供应链闭环 — 页面路由 + 管理员/仓库侧栏（二级折叠菜单）
================================================================
app.py 的 render_admin / render_warehouse 调用本模块。

页面键 st.session_state.page 与 app.py 内页面渲染函数保持不变。
"""
from __future__ import annotations

from typing import Any, Callable

import streamlit as st

URL_PAGE_QUERY_KEY = "p"

# 角色常量（与 app.Role 一致，避免 main ↔ app 循环导入）
ROLE_WAREHOUSE = "warehouse"
ROLE_ADMIN = "admin"

# 角色 → 允许的供应链页面
SUPPLY_CHAIN_PAGES_BY_ROLE: dict[str, frozenset[str]] = {
    ROLE_ADMIN: frozenset({
        "admin_suppliers",
        "supplier_order",
        "order_success",
        "verify_inbound",
        "verify_success",
    }),
    ROLE_WAREHOUSE: frozenset({
        "supplier_order",
        "order_success",
        "verify_inbound",
        "verify_success",
    }),
}

ADMIN_SUPPLY_PAGES = frozenset({
    "admin_suppliers",
    "supplier_order",
    "order_success",
    "verify_inbound",
    "verify_success",
})
WAREHOUSE_SUPPLY_PAGES = frozenset({
    "supplier_order",
    "order_success",
    "verify_inbound",
    "verify_success",
})

# 折叠分组 → 页面键（用于 expanded= 高亮当前组）
_WH_LOGISTICS_PAGES = frozenset({"pending", "short_in", "history"})
_WH_SUPPLY_PAGES = frozenset({
    "supplier_order", "order_success", "verify_inbound", "verify_success",
})
_WH_MORE_PAGES = frozenset({"inventory", "messages", "ai"})

_AD_BUSINESS_PAGES = frozenset({"dashboard", "all_orders", "expiry_dash"})
_AD_SUPPLY_PAGES = frozenset({
    "admin_suppliers", "supplier_order", "order_success",
    "verify_inbound", "verify_success",
})
_AD_MASTER_PAGES = frozenset({"images", "catalog", "product_master"})
_AD_SYSTEM_PAGES = frozenset({"accounts", "email", "export", "backup"})


def init_auth_session() -> None:
    """Cookie 恢复改在 route() 开头 prepare_auth_from_cookie()，此处保留空实现兼容。"""
    return


def set_page(page_key: str) -> None:
    st.session_state.page = page_key
    st.session_state._nav_from_sidebar = True
    try:
        st.query_params[URL_PAGE_QUERY_KEY] = page_key
    except Exception:
        pass
    try:
        import auth as sunshine_auth

        sunshine_auth.touch_cookie_page(page_key)
    except Exception:
        pass


def _sidebar_nav_button(page_key: str, label: str, *, key_suffix: str) -> None:
    """折叠面板内导航按钮。"""
    is_active = st.session_state.get("page") == page_key
    show_label = f"👉 {label}" if is_active else label
    if st.button(
        show_label,
        key=f"sb_{key_suffix}_{page_key}",
        use_container_width=True,
        type="primary" if is_active else "secondary",
    ):
        st.session_state.page = page_key
        st.session_state._nav_from_sidebar = True
        try:
            st.query_params[URL_PAGE_QUERY_KEY] = page_key
        except Exception:
            pass
        try:
            import auth as sunshine_auth

            sunshine_auth.touch_cookie_page(page_key)
        except Exception:
            pass
        st.session_state["confirming"] = False
        st.rerun()


def _render_expander_nav(
    title: str,
    items: list[tuple[str, str]],
    *,
    expanded: bool,
    key_suffix: str,
) -> list[tuple[str, str]]:
    with st.sidebar.expander(title, expanded=expanded):
        for page_key, label in items:
            _sidebar_nav_button(page_key, label, key_suffix=f"{key_suffix}_{page_key}")
    return items


def render_role_sidebar(role: str, *, messages_label: str) -> list[tuple[str, str]]:
    """
    按角色渲染侧栏折叠菜单，返回扁平 (page_key, label) 供当前页提示条使用。
    """
    page = st.session_state.get("page") or ""
    flat: list[tuple[str, str]] = []

    if role == ROLE_WAREHOUSE:
        flat.extend(
            _render_expander_nav(
                "📦 仓储与物流 / Logistics",
                [
                    ("pending", "📋 待发货单 / Pending"),
                    ("short_in", "🔔 缺货登记 / Shortages"),
                    ("history", "📜 历史记录 / History"),
                ],
                expanded=page in _WH_LOGISTICS_PAGES or page == "",
                key_suffix="wh_log",
            )
        )
        flat.extend(
            _render_expander_nav(
                "🏭 供应链管理 (闭环)",
                [
                    ("supplier_order", "🛍️ 采购下单 / Purchase"),
                    ("verify_inbound", "🔍 收货核对 / Receive"),
                ],
                expanded=page in _WH_SUPPLY_PAGES,
                key_suffix="wh_sc",
            )
        )
        flat.extend(
            _render_expander_nav(
                "📌 更多 / More",
                [
                    ("inventory", "📦 库存管理 / Inventory"),
                    ("messages", messages_label),
                    ("ai", "🤖 AI 助手 / AI"),
                ],
                expanded=page in _WH_MORE_PAGES,
                key_suffix="wh_more",
            )
        )
        return flat

    if role == ROLE_ADMIN:
        flat.extend(
            _render_expander_nav(
                "📊 经营大盘",
                [
                    ("dashboard", "📈 管理概览"),
                    ("all_orders", "📋 所有订单"),
                    ("expiry_dash", "📅 临期看板"),
                ],
                expanded=page in _AD_BUSINESS_PAGES or page == "",
                key_suffix="ad_a",
            )
        )
        flat.extend(
            _render_expander_nav(
                "🏭 供应链管理 (闭环)",
                [
                    ("admin_suppliers", "🏢 供应商管理"),
                    ("supplier_order", "🛍️ 供应商下单"),
                    ("verify_inbound", "🔍 收货核对"),
                ],
                expanded=page in _AD_SUPPLY_PAGES,
                key_suffix="ad_b",
            )
        )
        flat.extend(
            _render_expander_nav(
                "🧱 仓储主档",
                [
                    ("images", "🖼️ 图片管理"),
                    ("product_master", "📦 商品主档"),
                ],
                expanded=page in _AD_MASTER_PAGES,
                key_suffix="ad_c",
            )
        )
        flat.extend(
            _render_expander_nav(
                "⚙️ 系统维护",
                [
                    ("accounts", "👤 账号与权限"),
                    ("email", "📧 邮件设置"),
                    ("export", "📥 数据导出"),
                    ("backup", "💾 数据库备份"),
                ],
                expanded=page in _AD_SYSTEM_PAGES,
                key_suffix="ad_d",
            )
        )
        return flat

    return flat


def _bind_store(
    *,
    search_products: Callable[[str, int], Any],
    apply_inventory_receive: Callable[..., tuple[int, int]],
) -> None:
    from modules import supply_chain_store as sc

    sc.set_product_search_fn(search_products)
    sc.set_inventory_receive_fn(apply_inventory_receive)


def dispatch_supply_chain_page(
    page_key: str,
    *,
    render_page_heading: Callable[..., None],
    search_products: Callable[[str, int], Any],
    apply_inventory_receive: Callable[..., tuple[int, int]],
    audit_write: Callable[..., None] | None = None,
) -> bool:
    """若 page_key 属于供应链模块则渲染并返回 True。"""
    _bind_store(
        search_products=search_products,
        apply_inventory_receive=apply_inventory_receive,
    )

    from views.supply_chain import admin_suppliers, supplier_order, warehouse_verify

    pages: dict[str, Callable[[], None]] = {
        "admin_suppliers": lambda: admin_suppliers.render(render_page_heading),
        "supplier_order": lambda: supplier_order.render(
            render_page_heading, set_page, audit_write,
        ),
        "order_success": lambda: supplier_order.render(
            render_page_heading, set_page, audit_write,
        ),
        "verify_inbound": lambda: warehouse_verify.render(
            render_page_heading, set_page, audit_write,
        ),
        "verify_success": lambda: warehouse_verify.render(
            render_page_heading, set_page, audit_write,
        ),
    }
    fn = pages.get(page_key)
    if fn is None:
        return False
    fn()
    return True
