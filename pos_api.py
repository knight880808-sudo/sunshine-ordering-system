"""
POS REST API — 衔接 pos-web 前端与 Supermarket 商品表 / 数据库。

运行（默认端口 5055，供 Vite 代理 /api）:
    pip install fastapi uvicorn pandas openpyxl
    python -m uvicorn pos_api:app --host 127.0.0.1 --port 5055 --reload

商品来源: products.xlsx（与同目录 Streamlit app 一致）
库存/价格覆盖: SQLite inventory / product_prices（与 app.py 逻辑对齐）
订单写入: orders.db 内独立表 pos_orders / pos_order_lines
"""
from __future__ import annotations

import sqlite3
import uuid
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from urllib.parse import quote as urlquote

import pandas as pd
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel, ConfigDict, Field

# -----------------------------------------------------------------------------
# Paths（与 app.py 同级目录）
# -----------------------------------------------------------------------------
ROOT = Path(__file__).resolve().parent
DB_PATH = ROOT / "orders.db"
PRODUCTS_PATH = ROOT / "products.xlsx"
PRODUCT_IMAGES_DIR = ROOT / "product_images"
IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".gif", ".webp")

POS_SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS pos_orders (
    id               TEXT    PRIMARY KEY,
    created_at       TEXT    NOT NULL,
    subtotal_cents   INTEGER NOT NULL,
    discount_cents   INTEGER NOT NULL,
    total_cents      INTEGER NOT NULL,
    discount_code    TEXT,
    manual_override  INTEGER NOT NULL DEFAULT 0,
    payment          TEXT    NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_pos_orders_created ON pos_orders(created_at DESC);

CREATE TABLE IF NOT EXISTS pos_order_lines (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id         TEXT    NOT NULL,
    line_id          TEXT    NOT NULL,
    product_id       TEXT    NOT NULL,
    name             TEXT    NOT NULL,
    unit_price_cents INTEGER NOT NULL,
    quantity         INTEGER NOT NULL,
    max_qty          INTEGER NOT NULL,
    FOREIGN KEY (order_id) REFERENCES pos_orders(id)
);
CREATE INDEX IF NOT EXISTS idx_pos_lines_order ON pos_order_lines(order_id);
"""


@contextmanager
def db_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_pos_tables() -> None:
    with db_conn() as conn:
        conn.executescript(POS_SCHEMA_SQL)


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


# -----------------------------------------------------------------------------
# products.xlsx 解析（精简自 app.py）
# -----------------------------------------------------------------------------
def _parse_products_excel_after_read(df: pd.DataFrame) -> tuple[pd.DataFrame, bool]:
    for col in [
        "ItemCode", "Barcode", "Name", "Unit", "Price", "Category",
        "StockCartons", "StockPcs", "StockTotal",
    ]:
        if col not in df.columns:
            df[col] = ""
    for c in ["ItemCode", "Barcode", "Name", "Unit", "Category"]:
        df[c] = df[c].astype(str).fillna("").replace({"nan": "", "None": ""})
    for c in ["ItemCode", "Barcode"]:
        df[c] = df[c].str.replace(r"\.0$", "", regex=True)

    col_by_lower = {str(c).strip().lower(): c for c in df.columns}

    def _find_col(candidates: list[str]) -> str | None:
        for cand in candidates:
            hit = col_by_lower.get(cand.strip().lower())
            if hit:
                return hit
        return None

    price_src = _find_col([
        "Price", "价格", "单价", "售价", "零售价", "批发价", "进货价",
        "UnitPrice", "unit_price", "ListPrice", "MSRP",
        "单价(元)", "价格(元)", "售价(元)",
    ])
    if price_src:
        df["Price"] = pd.to_numeric(df[price_src], errors="coerce").fillna(0.0)
    else:
        df["Price"] = pd.to_numeric(df["Price"], errors="coerce").fillna(0.0)

    ct_col = _find_col([
        "StockCartons", "库存箱数", "库存(箱)", "库存箱", "CartonsStock", "Stock_Cartons",
    ])
    pc_col = _find_col([
        "StockPcs", "库存个数", "库存(个)", "库存个", "PcsStock", "Stock_Pcs",
    ])
    total_col = _find_col([
        "StockTotal", "Stock", "库存", "可用库存", "Available", "QtyAvailable",
    ])
    has_stock = False
    if ct_col:
        df["StockCartons"] = pd.to_numeric(df[ct_col], errors="coerce").fillna(0.0)
        has_stock = True
    elif "StockCartons" not in df.columns:
        df["StockCartons"] = 0.0
    if pc_col:
        df["StockPcs"] = pd.to_numeric(df[pc_col], errors="coerce").fillna(0.0)
        has_stock = True
    elif "StockPcs" not in df.columns:
        df["StockPcs"] = 0.0
    if total_col:
        df["StockTotal"] = pd.to_numeric(df[total_col], errors="coerce").fillna(0.0)
        has_stock = True
    elif "StockTotal" not in df.columns:
        df["StockTotal"] = 0.0

    return df, has_stock


def _overlay_inventory(conn: sqlite3.Connection, df: pd.DataFrame) -> pd.DataFrame:
    try:
        inv = conn.execute(
            "SELECT item_code, barcode, name, stock_cartons, stock_pcs FROM inventory"
        ).fetchall()
    except Exception:
        return df
    inv_by_ic = {str(r["item_code"]).strip().lower(): r for r in inv if (r["item_code"] or "").strip()}
    inv_by_bc = {str(r["barcode"]).strip().lower(): r for r in inv if (r["barcode"] or "").strip()}
    inv_by_nm = {str(r["name"]).strip().lower(): r for r in inv if (r["name"] or "").strip()}
    for idx, row in df.iterrows():
        ic = str(row.get("ItemCode", "") or "").strip().lower()
        bc = str(row.get("Barcode", "") or "").strip().lower()
        nm = str(row.get("Name", "") or "").strip().lower()
        hit = (inv_by_ic.get(ic) if ic else None) or (inv_by_bc.get(bc) if bc else None) or (
            inv_by_nm.get(nm) if nm else None
        )
        if hit:
            df.at[idx, "StockCartons"] = int(hit["stock_cartons"] or 0)
            df.at[idx, "StockPcs"] = int(hit["stock_pcs"] or 0)
    if inv:
        df["StockTotal"] = (
            pd.to_numeric(df["StockCartons"], errors="coerce").fillna(0)
            + pd.to_numeric(df["StockPcs"], errors="coerce").fillna(0)
        )
    return df


def _overlay_prices(conn: sqlite3.Connection, df: pd.DataFrame) -> pd.DataFrame:
    try:
        price_rows = conn.execute(
            "SELECT item_code, barcode, name, price FROM product_prices"
        ).fetchall()
    except Exception:
        return df
    by_ic = {str(r["item_code"]).strip().lower(): float(r["price"] or 0) for r in price_rows if (r["item_code"] or "").strip()}
    by_bc = {str(r["barcode"]).strip().lower(): float(r["price"] or 0) for r in price_rows if (r["barcode"] or "").strip()}
    by_nm = {str(r["name"]).strip().lower(): float(r["price"] or 0) for r in price_rows if (r["name"] or "").strip()}
    for idx, row in df.iterrows():
        ic = str(row.get("ItemCode", "") or "").strip().lower()
        bc = str(row.get("Barcode", "") or "").strip().lower()
        nm = str(row.get("Name", "") or "").strip().lower()
        hit = (by_ic.get(ic) if ic else None)
        if hit is None:
            hit = by_bc.get(bc) if bc else None
        if hit is None:
            hit = by_nm.get(nm) if nm else None
        if hit is not None:
            df.at[idx, "Price"] = float(hit)
    return df


_products_cache_mtime: float | None = None
_products_cache_list: list[dict] | None = None


def load_products_for_pos() -> list[dict]:
    """返回前端 Product[] 形态（整表；配合缓存避免重复读 Excel）。"""
    global _products_cache_mtime, _products_cache_list
    mtime = PRODUCTS_PATH.stat().st_mtime if PRODUCTS_PATH.exists() else -1.0
    if _products_cache_list is not None and _products_cache_mtime == mtime:
        return _products_cache_list

    if not PRODUCTS_PATH.exists():
        _products_cache_list = []
        _products_cache_mtime = mtime
        return []

    df = pd.read_excel(
        PRODUCTS_PATH,
        dtype={"ItemCode": str, "Barcode": str, "Category": str, "Unit": str},
    )
    df = df.rename(columns=lambda c: str(c).replace("\ufeff", "").strip())
    df, _ = _parse_products_excel_after_read(df)
    df.loc[df["Category"].str.strip() == "", "Category"] = "General"

    with db_conn() as conn:
        df = _overlay_inventory(conn, df)
        df = _overlay_prices(conn, df)

    out: list[dict] = []
    for _, row in df.iterrows():
        ic = str(row.get("ItemCode", "") or "").strip()
        bc = str(row.get("Barcode", "") or "").strip()
        name = str(row.get("Name", "") or "").strip()
        if not name:
            continue
        pid = ic or bc or f"row-{len(out)}"
        price = float(row.get("Price") or 0)
        price_cents = max(0, int(round(price * 100)))
        st_total = int(pd.to_numeric(row.get("StockTotal", 0), errors="coerce") or 0)
        if st_total <= 0:
            st_total = int(pd.to_numeric(row.get("StockCartons", 0), errors="coerce") or 0) + int(
                pd.to_numeric(row.get("StockPcs", 0), errors="coerce") or 0
            )
        sku = bc or ic or pid
        cat = str(row.get("Category", "") or "").strip() or None
        # 前端通过同源 /api 访问图片
        image_url = f"/api/product-image/{urlquote(ic or bc or pid, safe='')}"
        out.append({
            "id": pid,
            "name": name,
            "sku": sku,
            "priceCents": price_cents,
            "stock": max(0, st_total),
            "category": cat,
            "imageUrl": image_url,
        })
    _products_cache_list = out
    _products_cache_mtime = mtime
    return out


def find_image_file(item_key: str) -> Path | None:
    raw = (item_key or "").strip()
    if not raw:
        return None
    # 与 Streamlit 一致：优先 ItemCode / Barcode 文件名
    for ext in IMAGE_EXTS:
        p = PRODUCT_IMAGES_DIR / f"{raw}{ext}"
        if p.is_file():
            return p
        p2 = PRODUCT_IMAGES_DIR / f"bc_{raw}{ext}"
        if p2.is_file():
            return p2
    return None


# -----------------------------------------------------------------------------
# Pydantic
# -----------------------------------------------------------------------------
class CartLineIn(BaseModel):
    model_config = ConfigDict(extra="ignore")

    id: str
    productId: str
    name: str
    unitPriceCents: int = Field(ge=0)
    quantity: int = Field(ge=1)
    maxQty: int = Field(ge=1)


class CheckoutBody(BaseModel):
    model_config = ConfigDict(extra="ignore")

    lines: list[CartLineIn]
    subtotalCents: int = Field(ge=0)
    discountCents: int = Field(ge=0)
    totalCents: int = Field(ge=0)
    discountCode: str | None = None
    manualPriceOverride: bool = False
    payment: str


# -----------------------------------------------------------------------------
# FastAPI
# -----------------------------------------------------------------------------
init_pos_tables()

app = FastAPI(title="Supermarket POS API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://127.0.0.1:8503",
        "http://localhost:8503",
        "http://127.0.0.1:5173",
        "http://localhost:5173",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/api/health")
def health():
    return {"ok": True, "db": DB_PATH.name, "productsPath": PRODUCTS_PATH.name}


@app.get("/api/products")
def api_products(q: str = "", limit: int = 600):
    """默认最多返回 limit 条（避免超大表拖垮浏览器）；q 按名称 / SKU 模糊过滤。"""
    try:
        limit = max(1, min(limit, 5000))
        all_items = load_products_for_pos()
        items = all_items
        if (q or "").strip():
            ql = q.strip().lower()
            items = [
                p
                for p in all_items
                if ql in (p.get("name") or "").lower()
                or ql in (p.get("sku") or "").lower()
                or ql in (p.get("id") or "").lower()
            ]
        return JSONResponse({"items": items[:limit], "total": len(all_items)})
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e)) from e


@app.get("/api/product-image/{item_key:path}")
def api_product_image(item_key: str):
    path = find_image_file(item_key)
    if not path:
        raise HTTPException(status_code=404, detail="No image")
    return FileResponse(path)


@app.post("/api/pos/checkout")
def api_checkout(body: CheckoutBody):
    if not body.lines:
        raise HTTPException(status_code=400, detail="Empty cart")

    oid = f"POS-{uuid.uuid4().hex[:12].upper()}"
    created = now_iso()

    try:
        with db_conn() as conn:
            conn.execute(
                """
                INSERT INTO pos_orders (
                    id, created_at, subtotal_cents, discount_cents, total_cents,
                    discount_code, manual_override, payment
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    oid,
                    created,
                    body.subtotalCents,
                    body.discountCents,
                    body.totalCents,
                    body.discountCode,
                    1 if body.manualPriceOverride else 0,
                    body.payment,
                ),
            )
            for ln in body.lines:
                conn.execute(
                    """
                    INSERT INTO pos_order_lines (
                        order_id, line_id, product_id, name,
                        unit_price_cents, quantity, max_qty
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        oid,
                        ln.id,
                        ln.productId,
                        ln.name,
                        ln.unitPriceCents,
                        ln.quantity,
                        ln.maxQty,
                    ),
                )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e)) from e

    return {"orderId": oid, "createdAt": created}


@app.get("/api/pos/stats/summary")
def api_pos_stats():
    with db_conn() as conn:
        row = conn.execute(
            """
            SELECT
                COUNT(*) AS n,
                COALESCE(SUM(total_cents), 0) AS revenue
            FROM pos_orders
            """
        ).fetchone()
    n = int(row["n"] or 0)
    revenue = int(row["revenue"] or 0)
    avg = round(revenue / n) if n else 0
    return {
        "orderCount": n,
        "revenueCents": revenue,
        "averageTicketCents": avg,
    }


@app.get("/api/pos/orders/recent")
def api_pos_orders_recent(limit: int = 30):
    limit = max(1, min(limit, 200))
    with db_conn() as conn:
        rows = conn.execute(
            """
            SELECT id, created_at, total_cents, payment, discount_code
            FROM pos_orders
            ORDER BY created_at DESC
            LIMIT ?
            """,
            (limit,),
        ).fetchall()
    return {
        "orders": [
            {
                "id": r["id"],
                "createdAt": r["created_at"],
                "totalCents": r["total_cents"],
                "payment": r["payment"],
                "discountCode": r["discount_code"],
            }
            for r in rows
        ]
    }
