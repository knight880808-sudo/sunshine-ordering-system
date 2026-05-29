"""
供应链闭环 — 统一目录
  admin_suppliers.py  环节 A · 供应商管理
  supplier_order.py   环节 B · 采购下单
  warehouse_verify.py 环节 C · 收货核对
"""
from __future__ import annotations

from views.supply_chain import admin_suppliers, supplier_order, warehouse_verify

SUPPLY_CHAIN_PAGE_KEYS = frozenset({
    "admin_suppliers",
    "supplier_order",
    "order_success",
    "verify_inbound",
    "verify_success",
})


def section_title(zh: bool) -> str:
    return "📦 供应链管理" if zh else "📦 Supply chain"


def nav_items_for_role(role: str, zh: bool) -> list[tuple[str, str]]:
    """侧栏「供应链管理」分组内的三个入口（按角色）。"""
    indent = "　" if zh else "  "
    if role == "admin":
        return [
            ("admin_suppliers", f"{indent}· 供应商管理" if zh else f"{indent}· Suppliers"),
            ("supplier_order", f"{indent}· 采购下单" if zh else f"{indent}· Purchase order"),
            ("verify_inbound", f"{indent}· 收货核对" if zh else f"{indent}· Receive"),
        ]
    if role == "warehouse":
        return [
            ("supplier_order", f"{indent}· 采购下单" if zh else f"{indent}· Purchase order"),
            ("verify_inbound", f"{indent}· 收货核对" if zh else f"{indent}· Receive"),
        ]
    return []
