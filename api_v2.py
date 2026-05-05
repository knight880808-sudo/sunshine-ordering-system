"""
api_v2.py — Supermarket Ordering System: Comprehensive REST API
===============================================================
Covers ALL business domains:
  auth · products · orders · dispatch · receive
  shortages · inventory · prices · notifications
  reports/export · audit-log · backup · accounts · dashboard

Run:
    pip install fastapi uvicorn pandas openpyxl bcrypt python-multipart
    python -m uvicorn api_v2:app --host 127.0.0.1 --port 5056 --reload

Env vars (override defaults):
    SUNSHINE_WAREHOUSE_PASSWORD   default: sunshine888
    SUNSHINE_ADMIN_PASSWORD       default: sunshine
    API_V2_SECRET_KEY             JWT signing secret (change in production!)
    DB_PATH                       default: orders.db
    PRODUCTS_PATH                 default: products.xlsx
    PRODUCT_IMAGES_DIR            default: product_images
"""
from __future__ import annotations

import csv
import hashlib
import hmac
import io
import json
import os
import secrets
import shutil
import sqlite3
import uuid
from contextlib import contextmanager
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any
from urllib.parse import quote as urlquote

import pandas as pd
from fastapi import (
    Depends, FastAPI, File, Form, HTTPException, Query, Request,
    UploadFile, status,
)
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from pydantic import BaseModel, ConfigDict, Field

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    _OPENPYXL = True
except ImportError:
    _OPENPYXL = False

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
ROOT = Path(__file__).resolve().parent
DB_PATH          = Path(os.getenv("DB_PATH", str(ROOT / "orders.db")))
PRODUCTS_PATH    = Path(os.getenv("PRODUCTS_PATH", str(ROOT / "products.xlsx")))
IMAGES_DIR       = Path(os.getenv("PRODUCT_IMAGES_DIR", str(ROOT / "product_images")))
BACKUP_DIR       = ROOT / "backups"
IMAGE_EXTS       = (".jpg", ".jpeg", ".png", ".gif", ".webp")

WAREHOUSE_PW = os.getenv("SUNSHINE_WAREHOUSE_PASSWORD", "sunshine888")
ADMIN_PW     = os.getenv("SUNSHINE_ADMIN_PASSWORD", "sunshine")
SECRET_KEY   = os.getenv("API_V2_SECRET_KEY", "change-me-in-production-sunshine2025")
TOKEN_TTL_H  = 12

BRANCHES: list[str] = [
    "NAMBER ONE STORE",
    "SUNSHINE MARKET",
    "SUNSHINE PS",
    "SUNSHINE FU-SANMA",
    "SUNSHINE FU-NEMO",
    "CHEN STORE-SARAKATA",
    "CHEN STORE-CHAPI",
]

BRANCH_PERM_CODES = ("order", "my_orders", "my_short", "messages", "ai")


class Role:
    BRANCH    = "branch"
    WAREHOUSE = "warehouse"
    ADMIN     = "admin"


class OrderStatus:
    PENDING    = "Pending"
    DISPATCHED = "Dispatched"
    RECEIVED   = "Received"
    ALL        = (PENDING, DISPATCHED, RECEIVED)


class ShortageStatus:
    OPEN      = "Open"
    RESENDING = "Resending"
    OOS       = "Out of Stock"
    RESOLVED  = "Resolved"
    ALL       = (OPEN, RESENDING, OOS, RESOLVED)


# ---------------------------------------------------------------------------
# DB helpers
# ---------------------------------------------------------------------------
def _ensure_schema() -> None:
    try:
        from db_schema import migrate
        migrate(DB_PATH)
    except ImportError:
        pass  # db_schema.py optional at import time


@contextmanager
def db_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


# ---------------------------------------------------------------------------
# Simple HMAC token auth (stateless, no JWT dep required)
# ---------------------------------------------------------------------------
def _make_token(payload: dict) -> str:
    payload_json = json.dumps(payload, separators=(",", ":"))
    sig = hmac.new(SECRET_KEY.encode(), payload_json.encode(), hashlib.sha256).hexdigest()
    raw = json.dumps({"p": payload_json, "s": sig}, separators=(",", ":"))
    return raw.encode().hex()


def _verify_token(token: str) -> dict:
    try:
        raw = bytes.fromhex(token).decode()
        d = json.loads(raw)
        payload_json = d["p"]
        sig = d["s"]
        expected = hmac.new(SECRET_KEY.encode(), payload_json.encode(), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(sig, expected):
            raise ValueError("bad sig")
        payload = json.loads(payload_json)
        exp = payload.get("exp", 0)
        if exp and datetime.fromisoformat(exp) < datetime.now():
            raise ValueError("expired")
        return payload
    except Exception:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Invalid or expired token")


bearer = HTTPBearer(auto_error=False)


def get_current_user(
    cred: HTTPAuthorizationCredentials | None = Depends(bearer),
) -> dict:
    if not cred:
        raise HTTPException(status_code=401, detail="Not authenticated")
    return _verify_token(cred.credentials)


def require_role(*roles: str):
    def _dep(user: dict = Depends(get_current_user)) -> dict:
        if user.get("role") not in roles:
            raise HTTPException(status_code=403, detail="Forbidden")
        return user
    return _dep


# ---------------------------------------------------------------------------
# Password helpers (SHA-256 + salt; bcrypt optional)
# ---------------------------------------------------------------------------
def _hash_password(plain: str, salt: str | None = None) -> tuple[str, str]:
    if salt is None:
        salt = secrets.token_hex(16)
    h = hashlib.sha256((salt + plain).encode()).hexdigest()
    return h, salt


def _check_password(plain: str, hashed: str, salt: str) -> bool:
    h, _ = _hash_password(plain, salt)
    return hmac.compare_digest(h, hashed)


# ---------------------------------------------------------------------------
# Audit helper
# ---------------------------------------------------------------------------
def _audit(
    conn: sqlite3.Connection,
    event_type: str,
    *,
    role: str | None = None,
    account_id: int | None = None,
    username: str | None = None,
    branch: str | None = None,
    order_id: str | None = None,
    detail: str | None = None,
    ip: str | None = None,
) -> None:
    conn.execute(
        """INSERT INTO audit_log
           (event_type,role,account_id,username,branch,order_id,detail,ip_addr,created_at)
           VALUES (?,?,?,?,?,?,?,?,?)""",
        (event_type, role, account_id, username, branch, order_id, detail, ip, now_iso()),
    )


# ---------------------------------------------------------------------------
# Products (Excel loader, same logic as pos_api.py)
# ---------------------------------------------------------------------------
_prod_cache_mtime: float | None = None
_prod_cache_list:  list[dict] | None = None


def _load_products() -> list[dict]:
    global _prod_cache_mtime, _prod_cache_list
    mtime = PRODUCTS_PATH.stat().st_mtime if PRODUCTS_PATH.exists() else -1.0
    if _prod_cache_list is not None and _prod_cache_mtime == mtime:
        return _prod_cache_list

    if not PRODUCTS_PATH.exists():
        _prod_cache_list, _prod_cache_mtime = [], mtime
        return []

    df = pd.read_excel(
        PRODUCTS_PATH,
        dtype={"ItemCode": str, "Barcode": str, "Category": str, "Unit": str},
    )
    df = df.rename(columns=lambda c: str(c).replace("﻿", "").strip())
    col_lower = {str(c).lower(): c for c in df.columns}

    def _num(col: str) -> None:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    for c in ("ItemCode", "Barcode", "Name", "Unit", "Category"):
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].astype(str).replace({"nan": "", "None": ""}).fillna("")

    for c in ("ItemCode", "Barcode"):
        df[c] = df[c].str.replace(r"\.0$", "", regex=True)

    price_col = next(
        (col_lower[k] for k in ("price","价格","单价","售价","零售价") if k in col_lower), None
    )
    df["Price"] = pd.to_numeric(df[price_col] if price_col else 0, errors="coerce").fillna(0.0)

    for c in ("StockCartons", "StockPcs", "StockTotal"):
        _num(c) if c in df.columns else df.__setitem__(c, 0.0)

    with db_conn() as conn:
        try:
            inv_rows = conn.execute(
                "SELECT item_code,barcode,name,stock_cartons,stock_pcs FROM inventory"
            ).fetchall()
            by_ic = {str(r["item_code"]).strip().lower(): r for r in inv_rows if r["item_code"]}
            by_bc = {str(r["barcode"]).strip().lower(): r for r in inv_rows if r["barcode"]}
            for idx, row in df.iterrows():
                ic = str(row["ItemCode"]).strip().lower()
                bc = str(row["Barcode"]).strip().lower()
                hit = by_ic.get(ic) or by_bc.get(bc)
                if hit:
                    df.at[idx, "StockCartons"] = int(hit["stock_cartons"] or 0)
                    df.at[idx, "StockPcs"]     = int(hit["stock_pcs"] or 0)
        except Exception:
            pass

        try:
            pr_rows = conn.execute(
                "SELECT item_code,barcode,price FROM product_prices"
            ).fetchall()
            p_by_ic = {str(r["item_code"]).strip().lower(): float(r["price"]) for r in pr_rows if r["item_code"]}
            p_by_bc = {str(r["barcode"]).strip().lower():   float(r["price"]) for r in pr_rows if r["barcode"]}
            for idx, row in df.iterrows():
                ic = str(row["ItemCode"]).strip().lower()
                bc = str(row["Barcode"]).strip().lower()
                p  = p_by_ic.get(ic) or p_by_bc.get(bc)
                if p is not None:
                    df.at[idx, "Price"] = p
        except Exception:
            pass

    out: list[dict] = []
    for _, row in df.iterrows():
        name = str(row.get("Name", "")).strip()
        if not name:
            continue
        ic   = str(row.get("ItemCode", "")).strip()
        bc   = str(row.get("Barcode",  "")).strip()
        pid  = ic or bc or f"row-{len(out)}"
        ct   = int(pd.to_numeric(row.get("StockCartons", 0), errors="coerce") or 0)
        pc   = int(pd.to_numeric(row.get("StockPcs",     0), errors="coerce") or 0)
        tot  = int(pd.to_numeric(row.get("StockTotal",   0), errors="coerce") or 0) or ct + pc
        out.append({
            "id":       pid,
            "itemCode": ic,
            "barcode":  bc,
            "name":     name,
            "unit":     str(row.get("Unit", "")).strip(),
            "price":    float(row.get("Price", 0) or 0),
            "category": str(row.get("Category", "")).strip() or None,
            "stockCartons": max(0, ct),
            "stockPcs":     max(0, pc),
            "stockTotal":   max(0, tot),
            "imageUrl": f"/api/v2/product-image/{urlquote(ic or bc or pid, safe='')}",
        })

    _prod_cache_list, _prod_cache_mtime = out, mtime
    return out


def _invalidate_product_cache() -> None:
    global _prod_cache_mtime
    _prod_cache_mtime = None


def _find_image(key: str) -> Path | None:
    for ext in IMAGE_EXTS:
        for prefix in ("", "bc_"):
            p = IMAGES_DIR / f"{prefix}{key}{ext}"
            if p.is_file():
                return p
    return None


# ---------------------------------------------------------------------------
# Pydantic models
# ---------------------------------------------------------------------------
class LoginBody(BaseModel):
    role:     str
    password: str | None = None
    branch:   str | None = None
    username: str | None = None


class OrderLineIn(BaseModel):
    model_config = ConfigDict(extra="ignore")
    item_code:    str  = ""
    barcode:      str  = ""
    name:         str
    unit:         str  = ""
    price:        float = 0
    qty_cartons:  int  = Field(ge=0, default=0)
    qty_pcs:      int  = Field(ge=0, default=0)
    is_manual:    bool = False


class OrderSubmitBody(BaseModel):
    branch: str
    lines:  list[OrderLineIn]
    notes:  str | None = None


class DispatchLineIn(BaseModel):
    order_id:         str
    dispatch_cartons: int = Field(ge=0, default=0)
    dispatch_pcs:     int = Field(ge=0, default=0)


class DispatchBody(BaseModel):
    lines: list[DispatchLineIn]


class ReceiveLineIn(BaseModel):
    order_id:        str
    receive_cartons: int = Field(ge=0, default=0)
    receive_pcs:     int = Field(ge=0, default=0)
    short_cartons:   int = Field(ge=0, default=0)
    short_pcs:       int = Field(ge=0, default=0)
    note:            str | None = None


class ReceiveBody(BaseModel):
    lines: list[ReceiveLineIn]


class ShortageUpdateBody(BaseModel):
    model_config = ConfigDict(extra="ignore")
    shortage_status: str | None = None
    warehouse_note:  str | None = None
    branch_note:     str | None = None


class InventoryAdjustLine(BaseModel):
    item_code:     str = ""
    barcode:       str = ""
    name:          str = ""
    delta_cartons: int = 0
    delta_pcs:     int = 0
    note:          str | None = None


class InventoryAdjustBody(BaseModel):
    lines:    list[InventoryAdjustLine]
    operator: str | None = None


class PriceUpdateLine(BaseModel):
    item_code: str = ""
    barcode:   str = ""
    name:      str = ""
    price:     float = Field(ge=0)


class PriceUpdateBody(BaseModel):
    lines:    list[PriceUpdateLine]
    operator: str | None = None


class NotificationCreateBody(BaseModel):
    target_role:   str
    target_branch: str | None = None
    title:         str
    body:          str
    ref_order_id:  str | None = None


class AccountCreateBody(BaseModel):
    username:     str
    display_name: str | None = None
    phone:        str | None = None
    password:     str
    branch:       str
    permissions:  list[str] = []


class AccountUpdateBody(BaseModel):
    model_config = ConfigDict(extra="ignore")
    display_name:  str | None = None
    phone:         str | None = None
    password:      str | None = None
    status:        str | None = None
    permissions:   list[str] | None = None
    reject_reason: str | None = None


class SupplierOrderBody(BaseModel):
    title:         str
    body:          str
    product_lines: list[dict] = []


class StockArrivalBody(BaseModel):
    title:        str
    body:         str
    product_list: str = ""
    is_active:    bool = True


# ---------------------------------------------------------------------------
# FastAPI app
# ---------------------------------------------------------------------------
_ensure_schema()

app = FastAPI(
    title="Supermarket Ordering System API v2",
    version="2.0.0",
    description="Multi-branch ordering, warehouse dispatch, shortage management, inventory & price control.",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://127.0.0.1:8502", "http://localhost:8502",
                   "http://127.0.0.1:5173", "http://localhost:5173",
                   "http://127.0.0.1:8503", "http://localhost:8503"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ===========================================================================
# AUTH
# ===========================================================================
@app.get("/api/v2/branches")
def list_branches():
    return {"branches": BRANCHES}


@app.post("/api/v2/auth/login")
def login(body: LoginBody, request: Request):
    role = body.role
    ip   = request.client.host if request.client else None

    if role == Role.WAREHOUSE:
        if body.password != WAREHOUSE_PW:
            raise HTTPException(status_code=401, detail="Wrong password")
        exp = (datetime.now() + timedelta(hours=TOKEN_TTL_H)).isoformat(timespec="seconds")
        token = _make_token({"role": role, "exp": exp})
        with db_conn() as conn:
            _audit(conn, "login", role=role, detail="Shared password login (warehouse)", ip=ip)
        return {"token": token, "role": role, "branch": None}

    if role == Role.ADMIN:
        if body.password != ADMIN_PW:
            raise HTTPException(status_code=401, detail="Wrong password")
        exp = (datetime.now() + timedelta(hours=TOKEN_TTL_H)).isoformat(timespec="seconds")
        token = _make_token({"role": role, "exp": exp})
        with db_conn() as conn:
            _audit(conn, "login", role=role, detail="Shared password login (admin)", ip=ip)
        return {"token": token, "role": role, "branch": None}

    if role == Role.BRANCH:
        # branch can login by account username/password OR just select branch (no-password)
        if body.username and body.password:
            with db_conn() as conn:
                row = conn.execute(
                    "SELECT * FROM user_accounts WHERE username=?", (body.username,)
                ).fetchone()
            if not row:
                raise HTTPException(status_code=401, detail="Account not found")
            if row["status"] == "pending":
                raise HTTPException(status_code=403, detail="Account pending approval")
            if row["status"] == "rejected":
                raise HTTPException(status_code=403, detail="Account rejected")
            if not _check_password(body.password, row["password_hash"], row["salt"]):
                raise HTTPException(status_code=401, detail="Wrong password")
            if body.branch and row["branch"] != body.branch:
                raise HTTPException(status_code=403, detail="Account belongs to a different branch")
            branch = row["branch"]
            perms  = json.loads(row["permissions"] or "[]")
            exp    = (datetime.now() + timedelta(hours=TOKEN_TTL_H)).isoformat(timespec="seconds")
            token  = _make_token({
                "role": Role.BRANCH, "branch": branch,
                "account_id": row["id"], "username": row["username"],
                "permissions": perms, "exp": exp,
            })
            with db_conn() as conn:
                _audit(conn, "login", role=Role.BRANCH, account_id=row["id"],
                       username=row["username"], branch=branch,
                       detail=f"Account login · id={row['id']}", ip=ip)
            return {"token": token, "role": Role.BRANCH, "branch": branch,
                    "accountId": row["id"], "username": row["username"], "permissions": perms}

        # No-password branch select
        if not body.branch or body.branch not in BRANCHES:
            raise HTTPException(status_code=400, detail="Invalid branch")
        exp   = (datetime.now() + timedelta(hours=TOKEN_TTL_H)).isoformat(timespec="seconds")
        token = _make_token({"role": Role.BRANCH, "branch": body.branch,
                              "permissions": list(BRANCH_PERM_CODES), "exp": exp})
        return {"token": token, "role": Role.BRANCH, "branch": body.branch,
                "permissions": list(BRANCH_PERM_CODES)}

    raise HTTPException(status_code=400, detail="Unknown role")


# ===========================================================================
# ACCOUNTS (admin manages, branch registers)
# ===========================================================================
@app.post("/api/v2/accounts/register")
def register_account(body: AccountCreateBody):
    """Branch staff self-registration (status=pending)."""
    if body.branch not in BRANCHES:
        raise HTTPException(status_code=400, detail="Invalid branch")
    if len(body.username) < 3:
        raise HTTPException(status_code=400, detail="Username too short")
    pw_hash, salt = _hash_password(body.password)
    try:
        with db_conn() as conn:
            conn.execute(
                """INSERT INTO user_accounts
                   (username,display_name,phone,password_hash,salt,branch,status,permissions,created_at)
                   VALUES (?,?,?,?,?,?,'pending','[]',?)""",
                (body.username, body.display_name, body.phone,
                 pw_hash, salt, body.branch, now_iso()),
            )
    except sqlite3.IntegrityError:
        raise HTTPException(status_code=409, detail="Username already taken")
    return {"ok": True, "message": "Registration submitted, awaiting admin approval"}


@app.get("/api/v2/accounts")
def list_accounts(
    status_filter: str | None = Query(None, alias="status"),
    branch: str | None = None,
    user: dict = Depends(require_role(Role.ADMIN)),
):
    clauses, params = [], []
    if status_filter:
        clauses.append("status=?"); params.append(status_filter)
    if branch:
        clauses.append("branch=?"); params.append(branch)
    where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
    with db_conn() as conn:
        rows = conn.execute(f"SELECT * FROM user_accounts {where} ORDER BY created_at DESC", params).fetchall()
    return {"accounts": [dict(r) for r in rows]}


@app.post("/api/v2/accounts")
def create_account(
    body: AccountCreateBody,
    user: dict = Depends(require_role(Role.ADMIN)),
):
    """Admin direct account creation (status=approved)."""
    if body.branch not in BRANCHES:
        raise HTTPException(status_code=400, detail="Invalid branch")
    pw_hash, salt = _hash_password(body.password)
    try:
        with db_conn() as conn:
            cur = conn.execute(
                """INSERT INTO user_accounts
                   (username,display_name,phone,password_hash,salt,branch,status,permissions,
                    created_at,approved_at,approved_by)
                   VALUES (?,?,?,?,?,?,'approved',?,?,?,?)""",
                (body.username, body.display_name, body.phone, pw_hash, salt, body.branch,
                 json.dumps(body.permissions), now_iso(), now_iso(), user.get("role","admin")),
            )
            return {"ok": True, "id": cur.lastrowid}
    except sqlite3.IntegrityError:
        raise HTTPException(status_code=409, detail="Username already taken")


@app.put("/api/v2/accounts/{account_id}")
def update_account(
    account_id: int,
    body: AccountUpdateBody,
    user: dict = Depends(require_role(Role.ADMIN)),
):
    with db_conn() as conn:
        row = conn.execute("SELECT * FROM user_accounts WHERE id=?", (account_id,)).fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Account not found")
        sets, params = [], []
        if body.display_name is not None:
            sets.append("display_name=?"); params.append(body.display_name)
        if body.phone is not None:
            sets.append("phone=?"); params.append(body.phone)
        if body.password is not None:
            h, s = _hash_password(body.password)
            sets += ["password_hash=?", "salt=?"]; params += [h, s]
        if body.status is not None:
            sets.append("status=?"); params.append(body.status)
            if body.status == "approved":
                sets += ["approved_at=?", "approved_by=?"]; params += [now_iso(), user.get("username","admin")]
        if body.permissions is not None:
            sets.append("permissions=?"); params.append(json.dumps(body.permissions))
        if body.reject_reason is not None:
            sets.append("reject_reason=?"); params.append(body.reject_reason)
        if not sets:
            return {"ok": True}
        params.append(account_id)
        conn.execute(f"UPDATE user_accounts SET {','.join(sets)} WHERE id=?", params)
    return {"ok": True}


@app.delete("/api/v2/accounts/{account_id}")
def delete_account(account_id: int, user: dict = Depends(require_role(Role.ADMIN))):
    with db_conn() as conn:
        conn.execute("DELETE FROM user_accounts WHERE id=?", (account_id,))
    return {"ok": True}


# ===========================================================================
# PRODUCTS
# ===========================================================================
@app.get("/api/v2/products")
def api_products(
    q:     str = "",
    limit: int = Query(600, ge=1, le=5000),
    _user: dict = Depends(get_current_user),
):
    all_p = _load_products()
    if q.strip():
        ql = q.strip().lower()
        all_p = [p for p in all_p if
                 ql in p["name"].lower() or
                 ql in (p["itemCode"] or "").lower() or
                 ql in (p["barcode"] or "").lower()]
    return {"items": all_p[:limit], "total": len(all_p)}


@app.get("/api/v2/product-image/{item_key:path}")
def api_product_image(item_key: str):
    path = _find_image(item_key)
    if not path:
        raise HTTPException(status_code=404, detail="No image")
    return FileResponse(path)


@app.post("/api/v2/product-image/{item_key:path}")
async def upload_product_image(
    item_key: str,
    file: UploadFile = File(...),
    _user: dict = Depends(require_role(Role.ADMIN, Role.WAREHOUSE)),
):
    IMAGES_DIR.mkdir(parents=True, exist_ok=True)
    ext  = Path(file.filename or "").suffix.lower() or ".jpg"
    dest = IMAGES_DIR / f"{item_key}{ext}"
    data = await file.read()
    dest.write_bytes(data)
    _invalidate_product_cache()
    return {"ok": True, "path": str(dest)}


# ===========================================================================
# ORDERS — submit
# ===========================================================================
@app.post("/api/v2/orders")
def submit_order(body: OrderSubmitBody, request: Request, user: dict = Depends(get_current_user)):
    if not body.lines:
        raise HTTPException(status_code=400, detail="Empty order")
    if body.branch not in BRANCHES:
        raise HTTPException(status_code=400, detail="Invalid branch")

    ts         = datetime.now()
    date_str   = ts.strftime("%Y%m%d")
    group_id   = f"GRP-{date_str}-{uuid.uuid4().hex[:8].upper()}"
    account_id = user.get("account_id")
    ip         = request.client.host if request.client else None
    inserted   = []

    with db_conn() as conn:
        for ln in body.lines:
            oid = f"ORD-{date_str}-{uuid.uuid4().hex[:8].upper()}"
            conn.execute(
                """INSERT INTO orders
                   (id,order_group_id,branch,item_code,barcode,name,unit,price,
                    qty_cartons,qty_pcs,is_manual,status,order_date,account_id,notes)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (oid, group_id, body.branch, ln.item_code, ln.barcode, ln.name,
                 ln.unit, ln.price, ln.qty_cartons, ln.qty_pcs, int(ln.is_manual),
                 OrderStatus.PENDING, ts.isoformat(timespec="seconds"),
                 account_id, body.notes),
            )
            inserted.append(oid)

        _audit(conn, "order_submit", role=user.get("role"), account_id=account_id,
               username=user.get("username"), branch=body.branch,
               order_id=group_id, detail=f"{len(inserted)} lines", ip=ip)

        # Notify warehouse
        conn.execute(
            """INSERT INTO notifications
               (target_role,target_branch,title,body,ref_order_id,created_at)
               VALUES ('warehouse',NULL,?,?,?,?)""",
            (f"New order from {body.branch}",
             f"{len(inserted)} item(s) — group {group_id}", group_id, now_iso()),
        )

    return {"ok": True, "groupId": group_id, "orderIds": inserted}


# ===========================================================================
# ORDERS — query
# ===========================================================================
@app.get("/api/v2/orders")
def list_orders(
    branch:     str | None = None,
    status_f:   str | None = Query(None, alias="status"),
    q:          str | None = None,
    date_from:  str | None = None,
    date_to:    str | None = None,
    limit:      int = Query(500, ge=1, le=2000),
    offset:     int = Query(0, ge=0),
    user: dict = Depends(get_current_user),
):
    role = user.get("role")
    clauses, params = [], []

    if role == Role.BRANCH:
        clauses.append("branch=?"); params.append(user.get("branch"))
    elif branch:
        clauses.append("branch=?"); params.append(branch)

    if status_f:
        clauses.append("status=?"); params.append(status_f)
    if q:
        ql = f"%{q.strip()}%"
        clauses.append("(id LIKE ? OR name LIKE ? OR barcode LIKE ? OR item_code LIKE ?)")
        params.extend([ql, ql, ql, ql])
    if date_from:
        clauses.append("order_date >= ?"); params.append(date_from)
    if date_to:
        clauses.append("order_date <= ?"); params.append(date_to + "T23:59:59")

    where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
    with db_conn() as conn:
        total = conn.execute(
            f"SELECT COUNT(*) FROM orders {where}", params
        ).fetchone()[0]
        rows = conn.execute(
            f"SELECT * FROM orders {where} ORDER BY order_date DESC LIMIT ? OFFSET ?",
            params + [limit, offset],
        ).fetchall()
    return {"orders": [dict(r) for r in rows], "total": total}


@app.get("/api/v2/orders/{order_id}")
def get_order(order_id: str, user: dict = Depends(get_current_user)):
    with db_conn() as conn:
        row = conn.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="Order not found")
    if user.get("role") == Role.BRANCH and row["branch"] != user.get("branch"):
        raise HTTPException(status_code=403, detail="Forbidden")
    return dict(row)


# ===========================================================================
# DISPATCH (warehouse → dispatched)
# ===========================================================================
@app.post("/api/v2/orders/dispatch")
def dispatch_orders(
    body: DispatchBody,
    request: Request,
    user: dict = Depends(require_role(Role.WAREHOUSE, Role.ADMIN)),
):
    if not body.lines:
        raise HTTPException(status_code=400, detail="No lines")
    now = now_iso()
    operator = user.get("username") or user.get("role")
    ip = request.client.host if request.client else None

    with db_conn() as conn:
        for ln in body.lines:
            row = conn.execute("SELECT * FROM orders WHERE id=?", (ln.order_id,)).fetchone()
            if not row:
                continue
            if row["status"] != OrderStatus.PENDING:
                continue
            conn.execute(
                """UPDATE orders SET status=?,dispatch_cartons=?,dispatch_pcs=?,
                   dispatch_at=?,dispatch_by=? WHERE id=?""",
                (OrderStatus.DISPATCHED, ln.dispatch_cartons, ln.dispatch_pcs,
                 now, operator, ln.order_id),
            )
            # Deduct inventory
            _inv_adjust(conn, row["item_code"], row["barcode"], row["name"],
                        "DISPATCH", -ln.dispatch_cartons, -ln.dispatch_pcs,
                        ref_order_id=ln.order_id, operator=operator)

            # Notify branch
            conn.execute(
                """INSERT INTO notifications
                   (target_role,target_branch,title,body,ref_order_id,created_at)
                   VALUES ('branch',?,?,?,?,?)""",
                (row["branch"], f"Order dispatched: {ln.order_id}",
                 f"{row['name']} — {ln.dispatch_cartons} ctn / {ln.dispatch_pcs} pcs",
                 ln.order_id, now),
            )
            _audit(conn, "dispatch", role=user.get("role"), branch=row["branch"],
                   order_id=ln.order_id,
                   detail=f"ctn={ln.dispatch_cartons} pcs={ln.dispatch_pcs}", ip=ip)

    return {"ok": True}


# ===========================================================================
# RECEIVE (branch confirms receipt)
# ===========================================================================
@app.post("/api/v2/orders/receive")
def receive_orders(
    body: ReceiveBody,
    request: Request,
    user: dict = Depends(require_role(Role.BRANCH, Role.WAREHOUSE, Role.ADMIN)),
):
    if not body.lines:
        raise HTTPException(status_code=400, detail="No lines")
    now      = now_iso()
    operator = user.get("username") or user.get("branch") or user.get("role")
    ip       = request.client.host if request.client else None
    has_short = False

    with db_conn() as conn:
        for ln in body.lines:
            row = conn.execute("SELECT * FROM orders WHERE id=?", (ln.order_id,)).fetchone()
            if not row:
                continue
            if row["status"] == Role.BRANCH and row["branch"] != user.get("branch"):
                continue
            conn.execute(
                """UPDATE orders SET status=?,receive_cartons=?,receive_pcs=?,
                   receive_at=?,receive_by=? WHERE id=?""",
                (OrderStatus.RECEIVED, ln.receive_cartons, ln.receive_pcs,
                 now, operator, ln.order_id),
            )
            if ln.short_cartons > 0 or ln.short_pcs > 0:
                has_short = True
                ordered_ct = row["dispatch_cartons"] or row["qty_cartons"]
                ordered_pc = row["dispatch_pcs"]     or row["qty_pcs"]
                shortage_id = conn.execute(
                    """INSERT INTO shortages
                       (order_id,branch,item_code,barcode,name,
                        ordered_cartons,ordered_pcs,received_cartons,received_pcs,
                        short_cartons,short_pcs,status,branch_note,created_at,updated_at)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,'Open',?,?,?)""",
                    (ln.order_id, row["branch"], row["item_code"], row["barcode"],
                     row["name"], ordered_ct, ordered_pc,
                     ln.receive_cartons, ln.receive_pcs,
                     ln.short_cartons, ln.short_pcs, ln.note, now, now),
                ).lastrowid
                # Notify warehouse
                conn.execute(
                    """INSERT INTO notifications
                       (target_role,title,body,ref_order_id,ref_shortage_id,created_at)
                       VALUES ('warehouse',?,?,?,?,?)""",
                    (f"Shortage from {row['branch']}: {row['name']}",
                     f"Short {ln.short_cartons} ctn / {ln.short_pcs} pcs — order {ln.order_id}",
                     ln.order_id, shortage_id, now),
                )
            _audit(conn, "receive_confirm", role=user.get("role"),
                   account_id=user.get("account_id"), username=operator,
                   branch=row["branch"], order_id=ln.order_id,
                   detail=f"ctn={ln.receive_cartons} pcs={ln.receive_pcs} short={'yes' if ln.short_cartons or ln.short_pcs else 'no'}",
                   ip=ip)

    return {"ok": True, "hasShortage": has_short}


# ===========================================================================
# SHORTAGES
# ===========================================================================
@app.get("/api/v2/shortages")
def list_shortages(
    branch:    str | None = None,
    status_f:  str | None = Query(None, alias="status"),
    q:         str | None = None,
    limit:     int = Query(200, ge=1, le=1000),
    user: dict = Depends(get_current_user),
):
    role = user.get("role")
    clauses, params = [], []
    if role == Role.BRANCH:
        clauses.append("branch=?"); params.append(user.get("branch"))
    elif branch:
        clauses.append("branch=?"); params.append(branch)
    if status_f:
        clauses.append("status=?"); params.append(status_f)
    if q:
        ql = f"%{q}%"
        clauses.append("(name LIKE ? OR order_id LIKE ?)"); params.extend([ql, ql])
    where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
    with db_conn() as conn:
        rows = conn.execute(
            f"SELECT * FROM shortages {where} ORDER BY created_at DESC LIMIT ?",
            params + [limit],
        ).fetchall()
    return {"shortages": [dict(r) for r in rows]}


@app.put("/api/v2/shortages/{shortage_id}")
def update_shortage(
    shortage_id: int,
    body: ShortageUpdateBody,
    user: dict = Depends(get_current_user),
):
    with db_conn() as conn:
        row = conn.execute("SELECT * FROM shortages WHERE id=?", (shortage_id,)).fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Not found")
        if user.get("role") == Role.BRANCH and row["branch"] != user.get("branch"):
            raise HTTPException(status_code=403, detail="Forbidden")
        sets, params = ["updated_at=?"], [now_iso()]
        if body.shortage_status:
            sets.append("status=?"); params.append(body.shortage_status)
        if body.warehouse_note is not None:
            sets.append("warehouse_note=?"); params.append(body.warehouse_note)
        if body.branch_note is not None:
            sets.append("branch_note=?"); params.append(body.branch_note)
        params.append(shortage_id)
        conn.execute(f"UPDATE shortages SET {','.join(sets)} WHERE id=?", params)

        # Notify other party
        if body.shortage_status:
            if user.get("role") == Role.BRANCH:
                tgt, tgt_b = "warehouse", None
            else:
                tgt, tgt_b = "branch", row["branch"]
            conn.execute(
                """INSERT INTO notifications
                   (target_role,target_branch,title,body,ref_shortage_id,created_at)
                   VALUES (?,?,?,?,?,?)""",
                (tgt, tgt_b,
                 f"Shortage #{shortage_id} updated → {body.shortage_status}",
                 row["name"], shortage_id, now_iso()),
            )
    return {"ok": True}


# ===========================================================================
# INVENTORY
# ===========================================================================
def _inv_adjust(
    conn: sqlite3.Connection,
    item_code: str, barcode: str, name: str,
    txn_type: str,
    delta_ct: int, delta_pc: int,
    ref_order_id: str | None = None,
    operator: str | None = None,
    note: str | None = None,
) -> None:
    """Upsert inventory and write txn row."""
    row = conn.execute(
        """SELECT id,stock_cartons,stock_pcs FROM inventory
           WHERE item_code=? OR (item_code='' AND barcode=?)""",
        (item_code or "_none_", barcode or "_none_"),
    ).fetchone()
    if row:
        b_ct, b_pc = int(row["stock_cartons"]), int(row["stock_pcs"])
        a_ct = max(0, b_ct + delta_ct)
        a_pc = max(0, b_pc + delta_pc)
        conn.execute(
            "UPDATE inventory SET stock_cartons=?,stock_pcs=?,updated_at=? WHERE id=?",
            (a_ct, a_pc, now_iso(), row["id"]),
        )
    else:
        b_ct = b_pc = 0
        a_ct = max(0, delta_ct)
        a_pc = max(0, delta_pc)
        conn.execute(
            """INSERT OR IGNORE INTO inventory
               (item_code,barcode,name,unit,stock_cartons,stock_pcs,updated_at)
               VALUES (?,?,?,'',?,?,?)""",
            (item_code, barcode, name, a_ct, a_pc, now_iso()),
        )

    conn.execute(
        """INSERT INTO inventory_txn
           (item_code,barcode,name,txn_type,before_cartons,before_pcs,
            delta_cartons,delta_pcs,after_cartons,after_pcs,
            ref_order_id,operator,note,created_at)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (item_code, barcode, name, txn_type,
         b_ct, b_pc, delta_ct, delta_pc, a_ct, a_pc,
         ref_order_id, operator, note, now_iso()),
    )


@app.get("/api/v2/inventory")
def get_inventory(
    q: str | None = None,
    limit: int = Query(500, ge=1, le=5000),
    user: dict = Depends(require_role(Role.WAREHOUSE, Role.ADMIN)),
):
    clauses, params = [], []
    if q:
        ql = f"%{q}%"
        clauses.append("(name LIKE ? OR item_code LIKE ? OR barcode LIKE ?)")
        params.extend([ql, ql, ql])
    where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
    with db_conn() as conn:
        rows = conn.execute(
            f"SELECT * FROM inventory {where} ORDER BY name LIMIT ?",
            params + [limit],
        ).fetchall()
    return {"inventory": [dict(r) for r in rows]}


@app.post("/api/v2/inventory/adjust")
def adjust_inventory(
    body: InventoryAdjustBody,
    user: dict = Depends(require_role(Role.WAREHOUSE, Role.ADMIN)),
):
    operator = body.operator or user.get("username") or user.get("role")
    with db_conn() as conn:
        for ln in body.lines:
            _inv_adjust(conn, ln.item_code, ln.barcode, ln.name,
                        "ADJUST", ln.delta_cartons, ln.delta_pcs,
                        operator=operator, note=ln.note)
    _invalidate_product_cache()
    return {"ok": True, "updated": len(body.lines)}


@app.post("/api/v2/inventory/import")
async def import_inventory(
    mode: str = Form("append"),  # append | overwrite
    operator: str = Form(""),
    file: UploadFile = File(...),
    user: dict = Depends(require_role(Role.WAREHOUSE, Role.ADMIN)),
):
    """Upload Excel with ItemCode/Barcode/Name/StockCartons/StockPcs columns."""
    data = await file.read()
    try:
        df = pd.read_excel(io.BytesIO(data))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Cannot parse Excel: {e}")

    df.columns = [str(c).strip() for c in df.columns]
    col_lower   = {c.lower(): c for c in df.columns}

    def _col(*candidates):
        for c in candidates:
            if c.lower() in col_lower:
                return col_lower[c.lower()]
        return None

    ic_col  = _col("ItemCode", "item_code", "编号")
    bc_col  = _col("Barcode", "barcode", "条码")
    nm_col  = _col("Name", "name", "名称", "商品名")
    ct_col  = _col("StockCartons", "stock_cartons", "箱数", "库存箱数")
    pc_col  = _col("StockPcs",     "stock_pcs",     "个数", "库存个数")

    if not (ct_col or pc_col):
        raise HTTPException(status_code=400, detail="No stock columns found (StockCartons/StockPcs/箱数/个数)")

    updated = 0
    op      = operator or user.get("username") or user.get("role", "import")
    txn     = "IMPORT_OVR" if mode == "overwrite" else "IMPORT_DELTA"

    with db_conn() as conn:
        for _, row in df.iterrows():
            ic   = str(row[ic_col]).strip()  if ic_col  and not pd.isna(row[ic_col])  else ""
            bc   = str(row[bc_col]).strip()  if bc_col  and not pd.isna(row[bc_col])  else ""
            name = str(row[nm_col]).strip()  if nm_col  and not pd.isna(row[nm_col])  else ""
            if not (ic or bc or name):
                continue
            d_ct = int(pd.to_numeric(row[ct_col], errors="coerce") or 0) if ct_col else 0
            d_pc = int(pd.to_numeric(row[pc_col], errors="coerce") or 0) if pc_col else 0

            if mode == "overwrite":
                existing = conn.execute(
                    "SELECT stock_cartons,stock_pcs FROM inventory WHERE item_code=? OR barcode=?",
                    (ic or "_", bc or "_"),
                ).fetchone()
                b_ct = int(existing["stock_cartons"]) if existing else 0
                b_pc = int(existing["stock_pcs"])     if existing else 0
                delta_ct, delta_pc = d_ct - b_ct, d_pc - b_pc
            else:
                delta_ct, delta_pc = d_ct, d_pc

            _inv_adjust(conn, ic, bc, name or ic or bc, txn,
                        delta_ct, delta_pc, operator=op)
            updated += 1

    _invalidate_product_cache()
    return {"ok": True, "updated": updated}


@app.get("/api/v2/inventory/txn")
def inventory_txn_log(
    item_code: str | None = None,
    txn_type:  str | None = None,
    limit:     int = Query(200, ge=1, le=2000),
    user: dict = Depends(require_role(Role.WAREHOUSE, Role.ADMIN)),
):
    clauses, params = [], []
    if item_code:
        clauses.append("(item_code=? OR barcode=?)"); params.extend([item_code, item_code])
    if txn_type:
        clauses.append("txn_type=?"); params.append(txn_type)
    where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
    with db_conn() as conn:
        rows = conn.execute(
            f"SELECT * FROM inventory_txn {where} ORDER BY created_at DESC LIMIT ?",
            params + [limit],
        ).fetchall()
    return {"txns": [dict(r) for r in rows]}


# ===========================================================================
# PRICES
# ===========================================================================
@app.get("/api/v2/prices")
def list_prices(
    q:     str | None = None,
    limit: int = Query(500, ge=1, le=5000),
    user: dict = Depends(require_role(Role.ADMIN, Role.WAREHOUSE)),
):
    prods = _load_products()
    with db_conn() as conn:
        rows = conn.execute("SELECT * FROM product_prices").fetchall()
    override = {
        (str(r["item_code"]).strip(), str(r["barcode"]).strip()): dict(r) for r in rows
    }
    result = []
    for p in prods:
        ic  = str(p.get("itemCode","")).strip()
        bc  = str(p.get("barcode","")).strip()
        ov  = override.get((ic, bc)) or override.get((ic, "")) or override.get(("", bc))
        entry = {
            "itemCode":    ic,
            "barcode":     bc,
            "name":        p["name"],
            "archivePrice": p["price"],
            "currentPrice": ov["price"] if ov else p["price"],
            "overridden":   bool(ov),
            "updatedAt":    ov["updated_at"] if ov else None,
            "operator":     ov["operator"]   if ov else None,
        }
        if not q or q.lower() in p["name"].lower() or q.lower() in ic.lower():
            result.append(entry)
    return {"prices": result[:limit], "total": len(result)}


@app.post("/api/v2/prices")
def update_prices(
    body: PriceUpdateBody,
    user: dict = Depends(require_role(Role.ADMIN)),
):
    op  = body.operator or user.get("username") or "admin"
    now = now_iso()
    with db_conn() as conn:
        for ln in body.lines:
            conn.execute(
                """INSERT INTO product_prices (item_code,barcode,name,price,operator,updated_at)
                   VALUES (?,?,?,?,?,?)
                   ON CONFLICT DO UPDATE SET price=excluded.price,operator=excluded.operator,
                   updated_at=excluded.updated_at""",
                (ln.item_code, ln.barcode, ln.name, ln.price, op, now),
            )
    _invalidate_product_cache()
    return {"ok": True, "updated": len(body.lines)}


@app.post("/api/v2/prices/import")
async def import_prices(
    operator: str = Form(""),
    file: UploadFile = File(...),
    user: dict = Depends(require_role(Role.ADMIN)),
):
    data = await file.read()
    try:
        df = pd.read_excel(io.BytesIO(data))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Cannot parse Excel: {e}")

    df.columns = [str(c).strip() for c in df.columns]
    col_lower  = {c.lower(): c for c in df.columns}

    def _col(*cands):
        for c in cands:
            if c.lower() in col_lower: return col_lower[c.lower()]
        return None

    ic_col  = _col("ItemCode","item_code","编号")
    bc_col  = _col("Barcode","barcode","条码")
    nm_col  = _col("Name","name","名称")
    pr_col  = _col("Price","price","价格","单价","售价")
    if not pr_col:
        raise HTTPException(status_code=400, detail="No price column found")

    op  = operator or user.get("username") or "admin"
    now = now_iso()
    updated = 0
    with db_conn() as conn:
        for _, row in df.iterrows():
            price = pd.to_numeric(row[pr_col], errors="coerce")
            if pd.isna(price) or price <= 0:
                continue
            ic   = str(row[ic_col]).strip() if ic_col  else ""
            bc   = str(row[bc_col]).strip() if bc_col  else ""
            name = str(row[nm_col]).strip() if nm_col  else ""
            if not (ic or bc):
                continue
            conn.execute(
                """INSERT INTO product_prices (item_code,barcode,name,price,operator,updated_at)
                   VALUES (?,?,?,?,?,?)
                   ON CONFLICT DO UPDATE SET price=excluded.price,operator=excluded.operator,
                   updated_at=excluded.updated_at""",
                (ic, bc, name, float(price), op, now),
            )
            updated += 1
    _invalidate_product_cache()
    return {"ok": True, "updated": updated}


# ===========================================================================
# NOTIFICATIONS
# ===========================================================================
@app.get("/api/v2/notifications")
def list_notifications(
    unread_only: bool = False,
    limit:       int  = Query(50, ge=1, le=500),
    user: dict = Depends(get_current_user),
):
    role   = user.get("role")
    branch = user.get("branch")
    clauses = [f"(target_role=? OR target_role='all'"]
    params  = [role]
    if branch:
        clauses[0] += " OR target_branch=?"
        params.append(branch)
    clauses[0] += ")"
    if unread_only:
        clauses.append("is_read=0")
    where = "WHERE " + " AND ".join(clauses)
    with db_conn() as conn:
        rows = conn.execute(
            f"SELECT * FROM notifications {where} ORDER BY created_at DESC LIMIT ?",
            params + [limit],
        ).fetchall()
    return {"notifications": [dict(r) for r in rows]}


@app.post("/api/v2/notifications/{notif_id}/read")
def mark_notification_read(notif_id: int, user: dict = Depends(get_current_user)):
    with db_conn() as conn:
        conn.execute("UPDATE notifications SET is_read=1 WHERE id=?", (notif_id,))
    return {"ok": True}


@app.post("/api/v2/notifications/read-all")
def mark_all_read(user: dict = Depends(get_current_user)):
    role   = user.get("role")
    branch = user.get("branch")
    with db_conn() as conn:
        if branch:
            conn.execute(
                "UPDATE notifications SET is_read=1 WHERE target_role=? OR target_branch=?",
                (role, branch),
            )
        else:
            conn.execute("UPDATE notifications SET is_read=1 WHERE target_role=?", (role,))
    return {"ok": True}


@app.post("/api/v2/notifications")
def create_notification(
    body: NotificationCreateBody,
    user: dict = Depends(require_role(Role.ADMIN, Role.WAREHOUSE)),
):
    with db_conn() as conn:
        conn.execute(
            """INSERT INTO notifications
               (target_role,target_branch,title,body,ref_order_id,created_at)
               VALUES (?,?,?,?,?,?)""",
            (body.target_role, body.target_branch, body.title, body.body,
             body.ref_order_id, now_iso()),
        )
    return {"ok": True}


# ===========================================================================
# STOCK ARRIVALS
# ===========================================================================
@app.get("/api/v2/stock-arrivals")
def list_arrivals(active_only: bool = True, user: dict = Depends(get_current_user)):
    where = "WHERE is_active=1" if active_only else ""
    with db_conn() as conn:
        rows = conn.execute(
            f"SELECT * FROM stock_arrivals {where} ORDER BY created_at DESC LIMIT 20"
        ).fetchall()
    return {"arrivals": [dict(r) for r in rows]}


@app.post("/api/v2/stock-arrivals")
def create_arrival(
    body: StockArrivalBody,
    user: dict = Depends(require_role(Role.ADMIN, Role.WAREHOUSE)),
):
    with db_conn() as conn:
        cur = conn.execute(
            """INSERT INTO stock_arrivals (title,body,product_list,is_active,created_at,created_by)
               VALUES (?,?,?,?,?,?)""",
            (body.title, body.body, body.product_list, int(body.is_active),
             now_iso(), user.get("role")),
        )
        # Push as broadcast notification
        conn.execute(
            """INSERT INTO notifications (target_role,title,body,created_at) VALUES ('all',?,?,?)""",
            (f"[Arrival] {body.title}", body.body, now_iso()),
        )
    return {"ok": True, "id": cur.lastrowid}


@app.put("/api/v2/stock-arrivals/{arrival_id}")
def update_arrival(
    arrival_id: int,
    body: StockArrivalBody,
    user: dict = Depends(require_role(Role.ADMIN, Role.WAREHOUSE)),
):
    with db_conn() as conn:
        conn.execute(
            "UPDATE stock_arrivals SET title=?,body=?,product_list=?,is_active=? WHERE id=?",
            (body.title, body.body, body.product_list, int(body.is_active), arrival_id),
        )
    return {"ok": True}


# ===========================================================================
# SUPPLIER ORDERS
# ===========================================================================
@app.post("/api/v2/supplier-orders")
def create_supplier_order(
    body: SupplierOrderBody,
    user: dict = Depends(require_role(Role.WAREHOUSE, Role.ADMIN)),
):
    with db_conn() as conn:
        cur = conn.execute(
            """INSERT INTO supplier_orders (title,body,product_lines,sent_by,created_at)
               VALUES (?,?,?,?,?)""",
            (body.title, body.body, json.dumps(body.product_lines),
             user.get("role"), now_iso()),
        )
        conn.execute(
            """INSERT INTO notifications (target_role,title,body,created_at) VALUES ('admin',?,?,?)""",
            (f"Supplier order: {body.title}", body.body[:200], now_iso()),
        )
    return {"ok": True, "id": cur.lastrowid}


@app.get("/api/v2/supplier-orders")
def list_supplier_orders(
    limit: int = Query(50, ge=1, le=500),
    user: dict = Depends(require_role(Role.WAREHOUSE, Role.ADMIN)),
):
    with db_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM supplier_orders ORDER BY created_at DESC LIMIT ?", (limit,)
        ).fetchall()
    return {"orders": [dict(r) for r in rows]}


# ===========================================================================
# DASHBOARD (admin)
# ===========================================================================
@app.get("/api/v2/dashboard")
def dashboard(user: dict = Depends(require_role(Role.ADMIN, Role.WAREHOUSE))):
    with db_conn() as conn:
        def _q(sql, *p):
            return conn.execute(sql, p).fetchone()

        today = datetime.now().strftime("%Y-%m-%d")
        week_ago = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")

        pending_count   = _q("SELECT COUNT(*) FROM orders WHERE status='Pending'")[0]
        dispatch_count  = _q("SELECT COUNT(*) FROM orders WHERE status='Dispatched'")[0]
        received_today  = _q("SELECT COUNT(*) FROM orders WHERE status='Received' AND receive_at>=?", today)[0]
        open_shortages  = _q("SELECT COUNT(*) FROM shortages WHERE status NOT IN ('Resolved','Out of Stock')")[0]
        orders_7d       = _q("SELECT COUNT(*) FROM orders WHERE order_date>=?", week_ago)[0]
        branch_stats    = conn.execute(
            """SELECT branch,
                  SUM(CASE WHEN status='Pending'    THEN 1 ELSE 0 END) AS pending,
                  SUM(CASE WHEN status='Dispatched' THEN 1 ELSE 0 END) AS dispatched,
                  SUM(CASE WHEN status='Received'   THEN 1 ELSE 0 END) AS received
               FROM orders WHERE order_date>=? GROUP BY branch""",
            (week_ago,),
        ).fetchall()
        recent_orders = conn.execute(
            "SELECT id,branch,name,status,order_date FROM orders ORDER BY order_date DESC LIMIT 10"
        ).fetchall()
        recent_short  = conn.execute(
            "SELECT id,branch,name,status,created_at FROM shortages ORDER BY created_at DESC LIMIT 5"
        ).fetchall()

    return {
        "pendingCount":   pending_count,
        "dispatchCount":  dispatch_count,
        "receivedToday":  received_today,
        "openShortages":  open_shortages,
        "orders7Days":    orders_7d,
        "branchStats":    [dict(r) for r in branch_stats],
        "recentOrders":   [dict(r) for r in recent_orders],
        "recentShortages":[dict(r) for r in recent_short],
    }


# ===========================================================================
# AUDIT LOG
# ===========================================================================
@app.get("/api/v2/audit-log")
def get_audit_log(
    event_type: str | None = None,
    branch:     str | None = None,
    username:   str | None = None,
    order_id:   str | None = None,
    date_from:  str | None = None,
    date_to:    str | None = None,
    limit:      int = Query(200, ge=1, le=2000),
    export_csv: bool = False,
    user: dict = Depends(require_role(Role.ADMIN)),
):
    clauses, params = [], []
    if event_type:  clauses.append("event_type=?");           params.append(event_type)
    if branch:      clauses.append("branch=?");               params.append(branch)
    if username:    clauses.append("username LIKE ?");        params.append(f"%{username}%")
    if order_id:    clauses.append("order_id LIKE ?");        params.append(f"%{order_id}%")
    if date_from:   clauses.append("created_at>=?");          params.append(date_from)
    if date_to:     clauses.append("created_at<=?");          params.append(date_to + "T23:59:59")
    where = ("WHERE " + " AND ".join(clauses)) if clauses else ""

    with db_conn() as conn:
        rows = conn.execute(
            f"SELECT * FROM audit_log {where} ORDER BY created_at DESC LIMIT ?",
            params + [limit],
        ).fetchall()

    data = [dict(r) for r in rows]
    if export_csv:
        buf = io.StringIO()
        if data:
            w = csv.DictWriter(buf, fieldnames=data[0].keys())
            w.writeheader(); w.writerows(data)
        buf.seek(0)
        return StreamingResponse(
            iter([buf.getvalue()]),
            media_type="text/csv",
            headers={"Content-Disposition": "attachment; filename=audit_log.csv"},
        )
    return {"logs": data, "total": len(data)}


# ===========================================================================
# REPORTS / EXPORT
# ===========================================================================
@app.get("/api/v2/reports/picking-list")
def export_picking_list(
    branch:    str | None = None,
    date_from: str | None = None,
    date_to:   str | None = None,
    user: dict = Depends(require_role(Role.ADMIN, Role.WAREHOUSE)),
):
    """Export pending/dispatched orders as Excel picking list."""
    clauses = ["status IN ('Pending','Dispatched')"]
    params  = []
    if branch:    clauses.append("branch=?");     params.append(branch)
    if date_from: clauses.append("order_date>=?"); params.append(date_from)
    if date_to:   clauses.append("order_date<=?"); params.append(date_to + "T23:59:59")
    where = "WHERE " + " AND ".join(clauses)

    with db_conn() as conn:
        rows = conn.execute(
            f"SELECT * FROM orders {where} ORDER BY branch, order_date", params
        ).fetchall()

    df = pd.DataFrame([dict(r) for r in rows]) if rows else pd.DataFrame()
    buf = io.BytesIO()
    if not df.empty:
        cols = ["id","branch","name","item_code","barcode","unit",
                "qty_cartons","qty_pcs","price","status","order_date"]
        df = df[[c for c in cols if c in df.columns]]
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        (df if not df.empty else pd.DataFrame(columns=["No data"])).to_excel(
            writer, index=False, sheet_name="Picking List"
        )
    buf.seek(0)
    return StreamingResponse(
        iter([buf.read()]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=picking_list.xlsx"},
    )


@app.get("/api/v2/reports/shortages")
def export_shortages_report(
    branch:     str | None = None,
    status_f:   str | None = Query(None, alias="status"),
    export_fmt: str = "excel",
    user: dict = Depends(require_role(Role.ADMIN, Role.WAREHOUSE)),
):
    clauses, params = [], []
    if branch:   clauses.append("branch=?");  params.append(branch)
    if status_f: clauses.append("status=?");  params.append(status_f)
    where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
    with db_conn() as conn:
        rows = conn.execute(
            f"SELECT * FROM shortages {where} ORDER BY created_at DESC", params
        ).fetchall()

    df = pd.DataFrame([dict(r) for r in rows]) if rows else pd.DataFrame()
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        (df if not df.empty else pd.DataFrame(columns=["No data"])).to_excel(
            writer, index=False, sheet_name="Shortages"
        )
    buf.seek(0)
    return StreamingResponse(
        iter([buf.read()]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=shortages_report.xlsx"},
    )


@app.get("/api/v2/reports/reconciliation")
def export_reconciliation(
    branch:    str | None = None,
    date_from: str | None = None,
    date_to:   str | None = None,
    user: dict = Depends(require_role(Role.ADMIN)),
):
    """Full order reconciliation: ordered vs dispatched vs received."""
    clauses, params = [], []
    if branch:    clauses.append("branch=?");     params.append(branch)
    if date_from: clauses.append("order_date>=?"); params.append(date_from)
    if date_to:   clauses.append("order_date<=?"); params.append(date_to + "T23:59:59")
    where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
    with db_conn() as conn:
        rows = conn.execute(
            f"SELECT * FROM orders {where} ORDER BY branch, order_date", params
        ).fetchall()

    df = pd.DataFrame([dict(r) for r in rows]) if rows else pd.DataFrame()
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        (df if not df.empty else pd.DataFrame(columns=["No data"])).to_excel(
            writer, index=False, sheet_name="Reconciliation"
        )
    buf.seek(0)
    return StreamingResponse(
        iter([buf.read()]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=reconciliation.xlsx"},
    )


# ===========================================================================
# TEMPLATES (download)
# ===========================================================================
@app.get("/api/v2/templates/inventory")
def inventory_template(_user: dict = Depends(get_current_user)):
    buf = io.BytesIO()
    df  = pd.DataFrame(columns=["ItemCode","Barcode","Name","Unit","StockCartons","StockPcs"])
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Inventory")
    buf.seek(0)
    return StreamingResponse(
        iter([buf.read()]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=inventory_template.xlsx"},
    )


@app.get("/api/v2/templates/price")
def price_template(_user: dict = Depends(get_current_user)):
    buf = io.BytesIO()
    df  = pd.DataFrame(columns=["ItemCode","Barcode","Name","Price"])
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Prices")
    buf.seek(0)
    return StreamingResponse(
        iter([buf.read()]),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=price_template.xlsx"},
    )


# ===========================================================================
# BACKUP
# ===========================================================================
@app.get("/api/v2/backups")
def list_backups(user: dict = Depends(require_role(Role.ADMIN))):
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    files = sorted(BACKUP_DIR.glob("*.db"), reverse=True)[:30]
    return {"backups": [{"name": f.name, "sizeMb": round(f.stat().st_size / 1e6, 2),
                          "createdAt": datetime.fromtimestamp(f.stat().st_mtime).isoformat()}
                        for f in files]}


@app.post("/api/v2/backups")
def create_backup(user: dict = Depends(require_role(Role.ADMIN))):
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = BACKUP_DIR / f"orders_backup_{ts}.db"
    shutil.copy2(DB_PATH, dest)
    return {"ok": True, "file": dest.name, "sizeMb": round(dest.stat().st_size / 1e6, 2)}


@app.get("/api/v2/backups/{filename}")
def download_backup(filename: str, user: dict = Depends(require_role(Role.ADMIN))):
    path = BACKUP_DIR / filename
    if not path.exists() or not path.name.endswith(".db"):
        raise HTTPException(status_code=404, detail="Backup not found")
    return FileResponse(path, filename=filename,
                        media_type="application/octet-stream")


# ===========================================================================
# POS (kept compatible with pos_api.py, shared product/price/inventory source)
# ===========================================================================
class CartLineIn(BaseModel):
    model_config = ConfigDict(extra="ignore")
    id:              str
    productId:       str
    name:            str
    unitPriceCents:  int = Field(ge=0)
    quantity:        int = Field(ge=1)
    maxQty:          int = Field(ge=1)


class CheckoutBody(BaseModel):
    model_config = ConfigDict(extra="ignore")
    lines:               list[CartLineIn]
    subtotalCents:       int = Field(ge=0)
    discountCents:       int = Field(ge=0)
    totalCents:          int = Field(ge=0)
    discountCode:        str | None = None
    manualPriceOverride: bool = False
    payment:             str


@app.post("/api/v2/pos/checkout")
def pos_checkout(body: CheckoutBody):
    if not body.lines:
        raise HTTPException(status_code=400, detail="Empty cart")
    oid     = f"POS-{uuid.uuid4().hex[:12].upper()}"
    created = now_iso()
    with db_conn() as conn:
        conn.execute(
            """INSERT INTO pos_orders
               (id,created_at,subtotal_cents,discount_cents,total_cents,
                discount_code,manual_override,payment)
               VALUES (?,?,?,?,?,?,?,?)""",
            (oid, created, body.subtotalCents, body.discountCents, body.totalCents,
             body.discountCode, 1 if body.manualPriceOverride else 0, body.payment),
        )
        for ln in body.lines:
            conn.execute(
                """INSERT INTO pos_order_lines
                   (order_id,line_id,product_id,name,unit_price_cents,quantity,max_qty)
                   VALUES (?,?,?,?,?,?,?)""",
                (oid, ln.id, ln.productId, ln.name,
                 ln.unitPriceCents, ln.quantity, ln.maxQty),
            )
    return {"orderId": oid, "createdAt": created}


@app.get("/api/v2/pos/stats/summary")
def pos_stats():
    with db_conn() as conn:
        row = conn.execute(
            "SELECT COUNT(*) AS n, COALESCE(SUM(total_cents),0) AS rev FROM pos_orders"
        ).fetchone()
    n   = int(row["n"] or 0)
    rev = int(row["rev"] or 0)
    return {"orderCount": n, "revenueCents": rev,
            "averageTicketCents": round(rev / n) if n else 0}


# ===========================================================================
# HEALTH
# ===========================================================================
@app.get("/api/v2/health")
def health():
    return {
        "ok":           True,
        "db":           DB_PATH.name,
        "dbExists":     DB_PATH.exists(),
        "productsPath": PRODUCTS_PATH.name,
        "productsExists": PRODUCTS_PATH.exists(),
    }
