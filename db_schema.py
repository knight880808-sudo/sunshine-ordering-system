"""
db_schema.py — Supermarket Ordering System: SQLite Schema & Migration
=====================================================================
Usage:
    python db_schema.py                 # create / migrate orders.db in CWD
    python db_schema.py --path /x/y.db  # specify DB path

Safe to run repeatedly (all DDL uses IF NOT EXISTS / ADD COLUMN guards).
Does NOT touch existing data.  After running, all required tables and
indexes exist and any missing columns are appended.
"""
from __future__ import annotations

import argparse
import sqlite3
from pathlib import Path

# ---------------------------------------------------------------------------
# Master schema SQL
# ---------------------------------------------------------------------------
SCHEMA_SQL = """
-- =========================================================
-- ORDERS  (core ordering table — one row per order line)
-- =========================================================
CREATE TABLE IF NOT EXISTS orders (
    id               TEXT    PRIMARY KEY,        -- e.g. ORD-20240501-XXXXX
    order_group_id   TEXT,                       -- groups lines from same submit
    branch           TEXT    NOT NULL,
    item_code        TEXT    NOT NULL DEFAULT '',
    barcode          TEXT    NOT NULL DEFAULT '',
    name             TEXT    NOT NULL,
    unit             TEXT    NOT NULL DEFAULT '',
    price            REAL    NOT NULL DEFAULT 0,
    qty_cartons      INTEGER NOT NULL DEFAULT 0,
    qty_pcs          INTEGER NOT NULL DEFAULT 0,
    is_manual        INTEGER NOT NULL DEFAULT 0,  -- 1 = typed-in product
    status           TEXT    NOT NULL DEFAULT 'Pending',
    order_date       TEXT    NOT NULL,
    -- Dispatch
    dispatch_cartons INTEGER,
    dispatch_pcs     INTEGER,
    dispatch_at      TEXT,
    dispatch_by      TEXT,
    -- Receive
    receive_cartons  INTEGER,
    receive_pcs      INTEGER,
    receive_at       TEXT,
    receive_by       TEXT,
    -- Metadata
    account_id       INTEGER,
    notes            TEXT
);
CREATE INDEX IF NOT EXISTS idx_orders_branch   ON orders(branch);
CREATE INDEX IF NOT EXISTS idx_orders_status   ON orders(status);
CREATE INDEX IF NOT EXISTS idx_orders_date     ON orders(order_date DESC);
CREATE INDEX IF NOT EXISTS idx_orders_group    ON orders(order_group_id);
CREATE INDEX IF NOT EXISTS idx_orders_item     ON orders(item_code, barcode);

-- =========================================================
-- SHORTAGES  (缺货上报与闭环)
-- =========================================================
CREATE TABLE IF NOT EXISTS shortages (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id         TEXT    NOT NULL,
    branch           TEXT    NOT NULL,
    item_code        TEXT    NOT NULL DEFAULT '',
    barcode          TEXT    NOT NULL DEFAULT '',
    name             TEXT    NOT NULL,
    ordered_cartons  INTEGER NOT NULL DEFAULT 0,
    ordered_pcs      INTEGER NOT NULL DEFAULT 0,
    received_cartons INTEGER NOT NULL DEFAULT 0,
    received_pcs     INTEGER NOT NULL DEFAULT 0,
    short_cartons    INTEGER NOT NULL DEFAULT 0,
    short_pcs        INTEGER NOT NULL DEFAULT 0,
    status           TEXT    NOT NULL DEFAULT 'Open',  -- Open/Resending/Out of Stock/Resolved
    branch_note      TEXT,
    warehouse_note   TEXT,
    created_at       TEXT    NOT NULL,
    updated_at       TEXT    NOT NULL,
    FOREIGN KEY (order_id) REFERENCES orders(id)
);
CREATE INDEX IF NOT EXISTS idx_shortages_branch  ON shortages(branch);
CREATE INDEX IF NOT EXISTS idx_shortages_status  ON shortages(status);
CREATE INDEX IF NOT EXISTS idx_shortages_order   ON shortages(order_id);
CREATE INDEX IF NOT EXISTS idx_shortages_created ON shortages(created_at DESC);

-- =========================================================
-- INVENTORY  (权威库存，按商品 key 唯一)
-- =========================================================
CREATE TABLE IF NOT EXISTS inventory (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    item_code     TEXT    NOT NULL DEFAULT '',
    barcode       TEXT    NOT NULL DEFAULT '',
    name          TEXT    NOT NULL,
    unit          TEXT    NOT NULL DEFAULT '',
    stock_cartons INTEGER NOT NULL DEFAULT 0,
    stock_pcs     INTEGER NOT NULL DEFAULT 0,
    updated_at    TEXT    NOT NULL
);
-- item_code 与 barcode 至少有一个非空，且组合唯一
CREATE UNIQUE INDEX IF NOT EXISTS uq_inventory_key ON inventory(
    COALESCE(NULLIF(item_code,''), '_'),
    COALESCE(NULLIF(barcode,''), '_')
);
CREATE INDEX IF NOT EXISTS idx_inv_item_code ON inventory(item_code);
CREATE INDEX IF NOT EXISTS idx_inv_barcode   ON inventory(barcode);
CREATE INDEX IF NOT EXISTS idx_inv_name      ON inventory(name);

-- =========================================================
-- INVENTORY_TXN  (库存流水 — 可审计)
-- =========================================================
CREATE TABLE IF NOT EXISTS inventory_txn (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    item_code      TEXT    NOT NULL DEFAULT '',
    barcode        TEXT    NOT NULL DEFAULT '',
    name           TEXT    NOT NULL DEFAULT '',
    txn_type       TEXT    NOT NULL,   -- IN / OUT / ADJUST / DISPATCH / RECEIVE / IMPORT
    before_cartons INTEGER NOT NULL DEFAULT 0,
    before_pcs     INTEGER NOT NULL DEFAULT 0,
    delta_cartons  INTEGER NOT NULL DEFAULT 0,
    delta_pcs      INTEGER NOT NULL DEFAULT 0,
    after_cartons  INTEGER NOT NULL DEFAULT 0,
    after_pcs      INTEGER NOT NULL DEFAULT 0,
    ref_order_id   TEXT,
    operator       TEXT,
    note           TEXT,
    created_at     TEXT    NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_txn_item    ON inventory_txn(item_code, barcode);
CREATE INDEX IF NOT EXISTS idx_txn_created ON inventory_txn(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_txn_type    ON inventory_txn(txn_type);
CREATE INDEX IF NOT EXISTS idx_txn_order   ON inventory_txn(ref_order_id);

-- =========================================================
-- PRODUCT_PRICES  (价格覆盖，运行时生效价)
-- =========================================================
CREATE TABLE IF NOT EXISTS product_prices (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    item_code   TEXT    NOT NULL DEFAULT '',
    barcode     TEXT    NOT NULL DEFAULT '',
    name        TEXT    NOT NULL DEFAULT '',
    price       REAL    NOT NULL,
    operator    TEXT,
    updated_at  TEXT    NOT NULL
);
CREATE UNIQUE INDEX IF NOT EXISTS uq_prices_key ON product_prices(
    COALESCE(NULLIF(item_code,''), '_'),
    COALESCE(NULLIF(barcode,''), '_')
);
CREATE INDEX IF NOT EXISTS idx_prices_item ON product_prices(item_code);

-- =========================================================
-- USER_ACCOUNTS  (分店账号)
-- =========================================================
CREATE TABLE IF NOT EXISTS user_accounts (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    username      TEXT    NOT NULL UNIQUE,
    display_name  TEXT,
    phone         TEXT,
    password_hash TEXT    NOT NULL,
    salt          TEXT    NOT NULL,
    branch        TEXT    NOT NULL,
    status        TEXT    NOT NULL DEFAULT 'pending',  -- pending/approved/rejected
    permissions   TEXT    NOT NULL DEFAULT '[]',       -- JSON array of perm codes
    note          TEXT,
    reject_reason TEXT,
    created_at    TEXT    NOT NULL,
    approved_at   TEXT,
    approved_by   TEXT
);
CREATE INDEX IF NOT EXISTS idx_accounts_branch ON user_accounts(branch);
CREATE INDEX IF NOT EXISTS idx_accounts_status ON user_accounts(status);

-- =========================================================
-- NOTIFICATIONS  (消息中心)
-- =========================================================
CREATE TABLE IF NOT EXISTS notifications (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    target_role       TEXT    NOT NULL,   -- branch / warehouse / admin / all
    target_branch     TEXT,
    target_account_id INTEGER,
    title             TEXT    NOT NULL,
    body              TEXT    NOT NULL,
    ref_order_id      TEXT,
    ref_shortage_id   INTEGER,
    is_read           INTEGER NOT NULL DEFAULT 0,
    created_at        TEXT    NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_notif_role    ON notifications(target_role, is_read);
CREATE INDEX IF NOT EXISTS idx_notif_branch  ON notifications(target_branch, is_read);
CREATE INDEX IF NOT EXISTS idx_notif_created ON notifications(created_at DESC);

-- =========================================================
-- STOCK_ARRIVALS  (进货/到货通知公告)
-- =========================================================
CREATE TABLE IF NOT EXISTS stock_arrivals (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    title        TEXT    NOT NULL,
    body         TEXT    NOT NULL,
    product_list TEXT,   -- free-text, one line per product
    is_active    INTEGER NOT NULL DEFAULT 1,
    created_at   TEXT    NOT NULL,
    created_by   TEXT
);
CREATE INDEX IF NOT EXISTS idx_arrivals_active ON stock_arrivals(is_active);

-- =========================================================
-- AUDIT_LOG  (操作审计)
-- =========================================================
CREATE TABLE IF NOT EXISTS audit_log (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    event_type  TEXT    NOT NULL,  -- login/logout/order_submit/receive_confirm/...
    role        TEXT,
    account_id  INTEGER,
    username    TEXT,
    branch      TEXT,
    order_id    TEXT,
    detail      TEXT,
    ip_addr     TEXT,
    created_at  TEXT    NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_audit_created ON audit_log(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_audit_event   ON audit_log(event_type);
CREATE INDEX IF NOT EXISTS idx_audit_branch  ON audit_log(branch);
CREATE INDEX IF NOT EXISTS idx_audit_order   ON audit_log(order_id);

-- =========================================================
-- BRANCH_CART_DRAFT  (购物车草稿，分账号)
-- =========================================================
CREATE TABLE IF NOT EXISTS branch_cart_draft (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    account_id  INTEGER NOT NULL,
    branch      TEXT    NOT NULL,
    cart_json   TEXT    NOT NULL DEFAULT '[]',
    updated_at  TEXT    NOT NULL,
    UNIQUE(account_id, branch)
);

-- =========================================================
-- SUPPLIER_ORDERS  (仓库向供货商下单记录)
-- =========================================================
CREATE TABLE IF NOT EXISTS supplier_orders (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    title         TEXT    NOT NULL,
    body          TEXT    NOT NULL,
    product_lines TEXT,   -- JSON or free-text
    sent_by       TEXT,
    created_at    TEXT    NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_supplier_created ON supplier_orders(created_at DESC);

-- =========================================================
-- POS_ORDERS / POS_ORDER_LINES  (收银子系统，独立订单)
-- =========================================================
CREATE TABLE IF NOT EXISTS pos_orders (
    id              TEXT    PRIMARY KEY,
    created_at      TEXT    NOT NULL,
    subtotal_cents  INTEGER NOT NULL,
    discount_cents  INTEGER NOT NULL,
    total_cents     INTEGER NOT NULL,
    discount_code   TEXT,
    manual_override INTEGER NOT NULL DEFAULT 0,
    payment         TEXT    NOT NULL
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

# ---------------------------------------------------------------------------
# Columns to add when migrating an existing DB (idempotent ADD COLUMN)
# ---------------------------------------------------------------------------
MIGRATION_COLUMNS: list[tuple[str, str, str]] = [
    # (table, column, definition)
    ("orders", "order_group_id",   "TEXT"),
    ("orders", "account_id",       "INTEGER"),
    ("orders", "notes",            "TEXT"),
    ("orders", "dispatch_cartons", "INTEGER"),
    ("orders", "dispatch_pcs",     "INTEGER"),
    ("orders", "dispatch_at",      "TEXT"),
    ("orders", "dispatch_by",      "TEXT"),
    ("orders", "receive_cartons",  "INTEGER"),
    ("orders", "receive_pcs",      "INTEGER"),
    ("orders", "receive_at",       "TEXT"),
    ("orders", "receive_by",       "TEXT"),
    ("shortages", "barcode",            "TEXT NOT NULL DEFAULT ''"),
    ("shortages", "received_cartons",   "INTEGER NOT NULL DEFAULT 0"),
    ("shortages", "received_pcs",       "INTEGER NOT NULL DEFAULT 0"),
    ("shortages", "short_cartons",      "INTEGER NOT NULL DEFAULT 0"),
    ("shortages", "short_pcs",          "INTEGER NOT NULL DEFAULT 0"),
    ("shortages", "updated_at",         "TEXT NOT NULL DEFAULT ''"),
    ("shortages", "created_at",         "TEXT NOT NULL DEFAULT ''"),  # old table used reported_date
    ("inventory_txn", "ref_order_id",   "TEXT"),           # old table used order_id
    ("inventory_txn", "delta_cartons",  "INTEGER DEFAULT 0"),  # old table used change_cartons
    ("inventory_txn", "delta_pcs",      "INTEGER DEFAULT 0"),  # old table used change_pcs
    ("inventory", "unit",               "TEXT NOT NULL DEFAULT ''"),
    ("notifications", "target_account_id", "INTEGER"),
    ("notifications", "ref_shortage_id",   "INTEGER"),
    ("user_accounts", "display_name",   "TEXT"),
    ("user_accounts", "phone",          "TEXT"),
    ("user_accounts", "reject_reason",  "TEXT"),
    ("user_accounts", "approved_at",    "TEXT"),
    ("user_accounts", "approved_by",    "TEXT"),
]


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def migrate(db_path: str | Path = "orders.db") -> None:
    """Create all tables and append any missing columns.  Safe to call repeatedly.

    Order matters for existing DBs:
      1. CREATE TABLE IF NOT EXISTS  (new tables only)
      2. ADD COLUMN                  (extend existing tables)
      3. CREATE INDEX IF NOT EXISTS  (after columns exist)
    """
    db_path = Path(db_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    try:
        # Split schema into CREATE TABLE and CREATE INDEX blocks so we can
        # add missing columns before attempting to index them.
        table_stmts = []
        index_stmts = []
        for stmt in SCHEMA_SQL.split(";"):
            s = stmt.strip()
            if not s:
                continue
            if s.upper().startswith("CREATE INDEX") or s.upper().startswith("CREATE UNIQUE INDEX"):
                index_stmts.append(s)
            else:
                table_stmts.append(s)

        for stmt in table_stmts:
            try:
                conn.execute(stmt)
            except Exception as e:
                print(f"  ! Table stmt skipped: {e}")

        conn.commit()
        _add_missing_columns(conn)

        for stmt in index_stmts:
            try:
                conn.execute(stmt)
            except Exception as e:
                print(f"  ! Index stmt skipped: {e}")

        conn.commit()
        print(f"[db_schema] Migration complete: {db_path}")
    finally:
        conn.close()


def _add_missing_columns(conn: sqlite3.Connection) -> None:
    """Attempt ADD COLUMN for each migration entry; skip if already present."""
    existing: dict[str, set[str]] = {}
    for tbl, col, defn in MIGRATION_COLUMNS:
        if tbl not in existing:
            try:
                rows = conn.execute(f"PRAGMA table_info({tbl})").fetchall()
                existing[tbl] = {r[1].lower() for r in rows}
            except Exception:
                existing[tbl] = set()
        if col.lower() not in existing[tbl]:
            try:
                conn.execute(f"ALTER TABLE {tbl} ADD COLUMN {col} {defn}")
                existing[tbl].add(col.lower())
                print(f"  + Added column {tbl}.{col}")
            except Exception as exc:
                print(f"  ! Could not add {tbl}.{col}: {exc}")


def get_connection(db_path: str | Path = "orders.db") -> sqlite3.Connection:
    """Return a configured connection (row_factory set).  Caller must close."""
    conn = sqlite3.connect(str(db_path), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Create / migrate Supermarket DB")
    parser.add_argument("--path", default="orders.db", help="SQLite DB path")
    args = parser.parse_args()
    migrate(args.path)
