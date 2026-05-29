"""
环节 A — 供应商管理（Admin）
数据写入 Suppliers 表（内存 / 未来 SQL Server）。
"""
from __future__ import annotations

import streamlit as st

from modules import supply_chain_store as sc


def _zh() -> bool:
    return st.session_state.get("lang") == "zh"


def render(render_page_heading, t_logout=None) -> None:
    zh = _zh()
    render_page_heading(
        "🏭 供应商管理" if zh else "🏭 Supplier Management",
        "添加和维护供货商；「向供货商下单」页面的下拉框实时读取本列表。"
        if zh
        else "Maintain suppliers; the purchase order page reads this list live.",
    )
    sc.seed_demo_suppliers(st.session_state)

    with st.expander("➕ " + ("新增供应商" if zh else "Add supplier"), expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            name = st.text_input("名称 / Name *", key="sup_name")
            contact = st.text_input("联系人 / Contact", key="sup_contact")
            phone = st.text_input("电话 / Phone", key="sup_phone")
        with c2:
            email = st.text_input("邮箱 / Email", key="sup_email")
            address = st.text_input("地址 / Address", key="sup_addr")
            notes = st.text_area("备注 / Notes", key="sup_notes", height=68)
        if st.button("保存 / Save", type="primary", key="sup_add_btn"):
            if not (name or "").strip():
                st.error("请填写供应商名称" if zh else "Supplier name is required")
            else:
                row = sc.add_supplier(st.session_state, {
                    "name": name,
                    "contact": contact,
                    "phone": phone,
                    "email": email,
                    "address": address,
                    "notes": notes,
                })
                st.success(
                    f"已添加 / Added: {row['supplier_id']} — {row['name']}"
                )
                st.rerun()

    st.divider()
    st.subheader("📋 " + ("供应商列表" if zh else "Supplier list"))
    rows = sc.list_suppliers(st.session_state, active_only=False)
    if not rows:
        st.info("暂无供应商，请先添加。" if zh else "No suppliers yet.")
        return

    for s in rows:
        active = s.get("is_active", True)
        label = f"{s.get('name', '')}  (`{s.get('supplier_id', '')}`)"
        if not active:
            label += " — " + ("已停用" if zh else "inactive")
        with st.expander(label, expanded=False):
            st.markdown(
                f"**Contact:** {s.get('contact') or '—'} · "
                f"**Phone:** {s.get('phone') or '—'} · "
                f"**Email:** {s.get('email') or '—'}"
            )
            if s.get("address"):
                st.caption(s.get("address"))
            if s.get("notes"):
                st.caption(s.get("notes"))
            bc1, bc2 = st.columns(2)
            with bc1:
                if active and st.button(
                    "停用 / Deactivate",
                    key=f"sup_off_{s['supplier_id']}",
                ):
                    sc.delete_supplier(st.session_state, s["supplier_id"])
                    st.rerun()
            with bc2:
                if not active and st.button(
                    "启用 / Activate",
                    key=f"sup_on_{s['supplier_id']}",
                ):
                    sc.update_supplier(
                        st.session_state, s["supplier_id"], {"is_active": True}
                    )
                    st.rerun()
