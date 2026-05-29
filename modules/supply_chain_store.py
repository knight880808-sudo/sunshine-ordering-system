"""
供应链闭环 — 内存表 + Session 持久化（Streamlit 会话内）
未来替换 SQL Server 时，按各函数内注释的 SQL 逐条迁移即可。
"""
from __future__ import annotations

import uuid
from copy import deepcopy
from datetime import datetime
from typing import Any, Callable

import pandas as pd

# 由 app.py 在启动时注入：商品搜索、库存写入
_product_search_fn: Callable[[str, int], pd.DataFrame] | None = None
_inventory_receive_fn: Callable[..., tuple[int, int]] | None = None

PO_STATUS_PENDING = "Pending"
PO_STATUS_RECEIVED = "Received"

_SESSION_KEY = "_supply_chain_store"


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _new_id(prefix: str) -> str:
    return f"{prefix}-{uuid.uuid4().hex[:10].upper()}"


def set_product_search_fn(fn: Callable[[str, int], pd.DataFrame]) -> None:
    global _product_search_fn
    _product_search_fn = fn


def set_inventory_receive_fn(fn: Callable[..., tuple[int, int]]) -> None:
    global _inventory_receive_fn
    _inventory_receive_fn = fn


def _empty_store() -> dict[str, Any]:
    return {
        "Suppliers": [],
        "Purchase_Orders": [],
        "Purchase_Order_Items": [],
        "Receipt_Discrepancies": [],
    }


def _get_store(session_state: Any) -> dict[str, Any]:
    if _SESSION_KEY not in session_state:
        session_state[_SESSION_KEY] = _empty_store()
    return session_state[_SESSION_KEY]


def search_products(query: str, limit: int = 200) -> pd.DataFrame:
    if _product_search_fn is None:
        return pd.DataFrame()
    return _product_search_fn(query, limit)


# ---------------------------------------------------------------------------
# Suppliers
# ---------------------------------------------------------------------------
def list_suppliers(session_state: Any, *, active_only: bool = True) -> list[dict]:
    # -- SQL Server (future):
    # SELECT supplier_id, name, contact, phone, email, address, notes, is_active
    # FROM Suppliers WHERE (@active_only = 0 OR is_active = 1) ORDER BY name;
    store = _get_store(session_state)
    rows = list(store["Suppliers"])
    if active_only:
        rows = [r for r in rows if r.get("is_active", True)]
    return sorted(rows, key=lambda r: (r.get("name") or "").lower())


def get_supplier(session_state: Any, supplier_id: str) -> dict | None:
    for s in _get_store(session_state)["Suppliers"]:
        if s.get("supplier_id") == supplier_id:
            return dict(s)
    return None


def add_supplier(session_state: Any, data: dict) -> dict:
    # INSERT INTO Suppliers (supplier_id, name, contact, phone, email, address, notes, is_active, created_at)
    # VALUES (@id, @name, @contact, @phone, @email, @address, @notes, 1, @created_at);
    store = _get_store(session_state)
    row = {
        "supplier_id": _new_id("SUP"),
        "name": (data.get("name") or "").strip(),
        "contact": (data.get("contact") or "").strip(),
        "phone": (data.get("phone") or "").strip(),
        "email": (data.get("email") or "").strip(),
        "address": (data.get("address") or "").strip(),
        "notes": (data.get("notes") or "").strip(),
        "is_active": True,
        "created_at": _now(),
    }
    store["Suppliers"].append(row)
    return row


def update_supplier(session_state: Any, supplier_id: str, data: dict) -> bool:
    # UPDATE Suppliers SET name=@name, contact=@contact, phone=@phone, email=@email,
    #   address=@address, notes=@notes, is_active=@active WHERE supplier_id=@id;
    store = _get_store(session_state)
    for s in store["Suppliers"]:
        if s.get("supplier_id") == supplier_id:
            for k in ("name", "contact", "phone", "email", "address", "notes"):
                if k in data:
                    s[k] = (data[k] or "").strip()
            if "is_active" in data:
                s["is_active"] = bool(data["is_active"])
            return True
    return False


def delete_supplier(session_state: Any, supplier_id: str) -> bool:
    # UPDATE Suppliers SET is_active = 0 WHERE supplier_id = @id;
    # 或 DELETE FROM Suppliers WHERE supplier_id = @id;
    return update_supplier(session_state, supplier_id, {"is_active": False})


# ---------------------------------------------------------------------------
# Purchase orders
# ---------------------------------------------------------------------------
def list_purchase_orders(
    session_state: Any, *, status: str | None = None
) -> list[dict]:
    # SELECT * FROM Purchase_Orders WHERE (@status IS NULL OR status = @status)
    # ORDER BY created_at DESC;
    store = _get_store(session_state)
    rows = list(store["Purchase_Orders"])
    if status:
        rows = [r for r in rows if r.get("status") == status]
    return sorted(rows, key=lambda r: r.get("created_at") or "", reverse=True)


def get_purchase_order(session_state: Any, po_id: str) -> dict | None:
    for po in _get_store(session_state)["Purchase_Orders"]:
        if po.get("po_id") == po_id:
            return dict(po)
    return None


def list_po_items(session_state: Any, po_id: str) -> list[dict]:
    # SELECT * FROM Purchase_Order_Items WHERE po_id = @po_id ORDER BY line_no;
    return [
        dict(x)
        for x in _get_store(session_state)["Purchase_Order_Items"]
        if x.get("po_id") == po_id
    ]


def create_purchase_order(
    session_state: Any,
    *,
    supplier_id: str,
    lines: list[dict],
    remarks: str = "",
    created_by: str = "",
) -> dict:
    """环节 B：生成采购单，状态 Pending。"""
    store = _get_store(session_state)
    sup = get_supplier(session_state, supplier_id)
    if not sup:
        raise ValueError("supplier_not_found")
    if not lines:
        raise ValueError("empty_lines")

    po_id = _new_id("PO")
    # INSERT INTO Purchase_Orders (po_id, supplier_id, supplier_name, status, remarks, created_by, created_at)
    # VALUES (@po_id, @supplier_id, @supplier_name, N'Pending', @remarks, @created_by, @created_at);
    po = {
        "po_id": po_id,
        "supplier_id": supplier_id,
        "supplier_name": sup.get("name") or "",
        "status": PO_STATUS_PENDING,
        "remarks": (remarks or "").strip(),
        "created_by": (created_by or "").strip(),
        "created_at": _now(),
        "received_at": "",
    }
    store["Purchase_Orders"].append(po)

    line_no = 0
    for ln in lines:
        line_no += 1
        item = {
            "line_id": _new_id("POL"),
            "po_id": po_id,
            "line_no": line_no,
            "item_code": str(ln.get("item_code") or "").strip(),
            "barcode": str(ln.get("barcode") or "").strip(),
            "name": str(ln.get("name") or "").strip(),
            "unit": str(ln.get("unit") or "").strip(),
            "qty_cartons_ordered": int(ln.get("qty_cartons") or 0),
            "qty_pcs_ordered": int(ln.get("qty_pcs") or 0),
        }
        # INSERT INTO Purchase_Order_Items (line_id, po_id, line_no, item_code, barcode, name, unit,
        #   qty_cartons_ordered, qty_pcs_ordered) VALUES (...);
        store["Purchase_Order_Items"].append(item)

    return po


def receive_purchase_order(
    session_state: Any,
    *,
    po_id: str,
    received_lines: list[dict],
    operator: str = "",
) -> tuple[bool, str, list[dict]]:
    """
    环节 C：核对收货 → status Received，累加库存，记录差异。
    received_lines: item_code/barcode/name/unit + qty_cartons + qty_pcs (实际收货)
    """
    store = _get_store(session_state)
    po = get_purchase_order(session_state, po_id)
    if not po:
        return False, "po_not_found", []
    if po.get("status") != PO_STATUS_PENDING:
        return False, "po_not_pending", []

    ordered = {(
        (x.get("item_code") or "").strip(),
        (x.get("barcode") or "").strip(),
        (x.get("name") or "").strip(),
    ): x for x in list_po_items(session_state, po_id)}

    recv_map: dict[tuple[str, str, str], dict] = {}
    for ln in received_lines:
        key = (
            str(ln.get("item_code") or "").strip(),
            str(ln.get("barcode") or "").strip(),
            str(ln.get("name") or "").strip(),
        )
        recv_map[key] = ln

    discrepancies: list[dict] = []

    for key, ord_row in ordered.items():
        recv = recv_map.get(key, {})
        o_ct = int(ord_row.get("qty_cartons_ordered") or 0)
        o_pc = int(ord_row.get("qty_pcs_ordered") or 0)
        r_ct = int(recv.get("qty_cartons") or 0)
        r_pc = int(recv.get("qty_pcs") or 0)

        if r_ct != o_ct or r_pc != o_pc:
            diff_type = "short" if (r_ct < o_ct or r_pc < o_pc) else "over"
            disc = {
                "discrepancy_id": _new_id("RD"),
                "po_id": po_id,
                "item_code": ord_row.get("item_code") or "",
                "barcode": ord_row.get("barcode") or "",
                "name": ord_row.get("name") or "",
                "ordered_cartons": o_ct,
                "ordered_pcs": o_pc,
                "received_cartons": r_ct,
                "received_pcs": r_pc,
                "diff_type": diff_type,
                "created_at": _now(),
            }
            # INSERT INTO Receipt_Discrepancies (...) VALUES (...);
            store["Receipt_Discrepancies"].append(disc)
            discrepancies.append(disc)

        if r_ct > 0 or r_pc > 0:
            if _inventory_receive_fn is not None:
                _inventory_receive_fn(
                    item_code=ord_row.get("item_code") or "",
                    barcode=ord_row.get("barcode") or "",
                    name=ord_row.get("name") or "-",
                    unit=ord_row.get("unit") or "",
                    change_ct=r_ct,
                    change_pc=r_pc,
                    order_id=po_id,
                    operator=operator or "warehouse_receive",
                )
            # 未来 SQL Server:
            # UPDATE Inventory SET stock_cartons = stock_cartons + @r_ct, stock_pcs = stock_pcs + @r_pc
            # WHERE item_key = @key;
            # INSERT INTO Inventory_Txn (txn_type, ...) VALUES (N'PURCHASE_RECEIVE', ...);

    for po_row in store["Purchase_Orders"]:
        if po_row.get("po_id") == po_id:
            # UPDATE Purchase_Orders SET status = N'Received', received_at = @ts WHERE po_id = @po_id;
            po_row["status"] = PO_STATUS_RECEIVED
            po_row["received_at"] = _now()
            break

    return True, "", discrepancies


def list_discrepancies(session_state: Any, po_id: str | None = None) -> list[dict]:
    rows = list(_get_store(session_state)["Receipt_Discrepancies"])
    if po_id:
        rows = [r for r in rows if r.get("po_id") == po_id]
    return sorted(rows, key=lambda r: r.get("created_at") or "", reverse=True)


def seed_demo_suppliers(session_state: Any) -> None:
    """首次进入供应链模块时可选演示数据。"""
    store = _get_store(session_state)
    if store["Suppliers"]:
        return
    add_supplier(session_state, {
        "name": "Demo Supplier A",
        "contact": "Alice",
        "phone": "+677-000-0001",
        "email": "supplier-a@example.com",
        "address": "Honiara",
        "notes": "Demo seed",
    })
