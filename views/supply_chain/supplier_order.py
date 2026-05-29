"""
环节 B — 向供货商下单（Warehouse / Admin）
供应商下拉 ← Suppliers；提交 → Purchase_Orders + Items，状态 Pending。
"""
from __future__ import annotations

import math

import pandas as pd
import streamlit as st

from modules import supply_chain_store as sc


PAGE_SIZE = 5


def _zh() -> bool:
    return st.session_state.get("lang") == "zh"


def _merge_cart_line(cart: list[dict], line: dict) -> None:
    key = (
        (line.get("item_code") or "").strip(),
        (line.get("barcode") or "").strip(),
        (line.get("name") or "").strip(),
    )
    for x in cart:
        if (
            (x.get("item_code") or "").strip(),
            (x.get("barcode") or "").strip(),
            (x.get("name") or "").strip(),
        ) == key:
            x["qty_cartons"] = int(x.get("qty_cartons", 0) or 0) + int(
                line.get("qty_cartons", 0) or 0
            )
            x["qty_pcs"] = int(x.get("qty_pcs", 0) or 0) + int(
                line.get("qty_pcs", 0) or 0
            )
            return
    cart.append(line)


def render(render_page_heading, set_page, audit_write=None) -> None:
    zh = _zh()
    sc.seed_demo_suppliers(st.session_state)

    if st.session_state.get("page") == "order_success":
        render_page_heading(
            "✅ 采购单已创建" if zh else "✅ Purchase order created",
            None,
        )
        po_id = st.session_state.get("sc_last_po_id") or ""
        if po_id:
            st.success(f"PO: `{po_id}` · " + ("状态：待收货 Pending" if zh else "Status: Pending"))
        if st.button(
            "继续下单 / New order",
            type="primary",
            key="sc_po_success_back",
        ):
            st.session_state.pop("sc_last_po_id", None)
            st.session_state.sc_po_cart = []
            set_page("supplier_order")
            st.rerun()
        role = (st.session_state.get("role") or "").strip()
        if role == "warehouse":
            if st.button("去收货核对 / Receive goods", key="sc_go_verify"):
                set_page("verify_inbound")
                st.rerun()
        else:
            st.info(
                "收货核对在侧栏「供应链管理 → 收货核对」。"
                if zh
                else "Use Supply chain → Receive in the sidebar."
            )
        return

    if "sc_po_cart" not in st.session_state:
        st.session_state.sc_po_cart = []
    if "sc_po_search_q" not in st.session_state:
        st.session_state.sc_po_search_q = ""
    if "sc_po_search_page" not in st.session_state:
        st.session_state.sc_po_search_page = 0

    render_page_heading(
        "🏭 向供货商下单" if zh else "🏭 Supplier purchase order",
        "选择供应商、搜索商品（每页 5 个），提交后生成待收货采购单。"
        if zh
        else "Pick supplier, search products (5 per page), submit to create a Pending PO.",
    )

    suppliers = sc.list_suppliers(st.session_state, active_only=True)
    if not suppliers:
        st.warning(
            "请先在「供应链管理 → 供应商管理」添加供应商。"
            if zh
            else "Add suppliers under Supply chain → Suppliers first."
        )
        return

    sup_options = {s["supplier_id"]: s["name"] for s in suppliers}
    sup_id = st.selectbox(
        "供应商 / Supplier *",
        options=list(sup_options.keys()),
        format_func=lambda k: sup_options[k],
        key="sc_po_supplier",
    )

    remarks = st.text_area("备注 / Remarks", key="sc_po_remarks", height=70)

    st.divider()
    st.markdown("**" + ("商品搜索" if zh else "Product search") + "**")
    with st.form("sc_po_search_form", clear_on_submit=False):
        q1, q2 = st.columns([5, 1])
        with q1:
            q_raw = st.text_input(
                "关键词 / Keyword",
                key="sc_po_q_input",
                placeholder="名称 / 条码 / 编号",
            )
        with q2:
            st.markdown("<div style='padding-top:28px'></div>", unsafe_allow_html=True)
            do_search = st.form_submit_button("搜索 / Search", type="primary")
    if do_search:
        st.session_state.sc_po_search_q = (q_raw or "").strip()
        st.session_state.sc_po_search_page = 0

    q = str(st.session_state.get("sc_po_search_q") or "").strip()
    results = sc.search_products(q, limit=200) if len(q) >= 1 else pd.DataFrame()

    if not results.empty:
        total = len(results)
        pages = max(1, math.ceil(total / PAGE_SIZE))
        page = int(st.session_state.get("sc_po_search_page") or 0)
        page = max(0, min(page, pages - 1))
        st.session_state.sc_po_search_page = page

        pc1, pc2, pc3 = st.columns([1, 2, 1])
        with pc1:
            if st.button("◀ 上一页", disabled=page <= 0, key="sc_po_prev"):
                st.session_state.sc_po_search_page = page - 1
                st.rerun()
        with pc2:
            st.caption(f"{page + 1} / {pages} · {total} " + ("条" if zh else "items"))
        with pc3:
            if st.button("下一页 ▶", disabled=page >= pages - 1, key="sc_po_next"):
                st.session_state.sc_po_search_page = page + 1
                st.rerun()

        start = page * PAGE_SIZE
        page_df = results.iloc[start : start + PAGE_SIZE]
        labels = []
        for _, row in page_df.iterrows():
            labels.append(
                f"{str(row.get('Name', '') or '')[:40]} | "
                f"{row.get('ItemCode', '')} | {row.get('Barcode', '')}"
            )
        pick_idx = st.selectbox(
            "本页选择 / Pick from page",
            options=list(range(len(page_df))),
            format_func=lambda i: labels[i],
            key="sc_po_pick",
        )
        ac1, ac2, ac3 = st.columns([1, 1, 2])
        with ac1:
            add_ct = st.number_input("箱 / Ct", 0, step=1, key="sc_po_add_ct")
        with ac2:
            add_pc = st.number_input("件 / Pcs", 0, step=1, key="sc_po_add_pc")
        with ac3:
            if st.button("加入明细 / Add line", key="sc_po_add_line"):
                if add_ct <= 0 and add_pc <= 0:
                    st.warning("请填写数量" if zh else "Enter quantity")
                else:
                    row = page_df.iloc[int(pick_idx)]
                    _merge_cart_line(
                        st.session_state.sc_po_cart,
                        {
                            "item_code": str(row.get("ItemCode", "") or ""),
                            "barcode": str(row.get("Barcode", "") or ""),
                            "name": str(row.get("Name", "") or ""),
                            "unit": str(row.get("Unit", "") or ""),
                            "qty_cartons": int(add_ct),
                            "qty_pcs": int(add_pc),
                        },
                    )
                    st.rerun()
    elif q:
        st.info("无匹配商品" if zh else "No products found")

    cart: list[dict] = st.session_state.sc_po_cart
    st.divider()
    st.markdown(f"**{'采购明细' if zh else 'Order lines'}** · {len(cart)}")
    if cart:
        for i, ln in enumerate(list(cart)):
            cols = st.columns([5, 1])
            with cols[0]:
                st.markdown(
                    f"{ln.get('name', '')} · 📦 {ln.get('qty_cartons', 0)} · "
                    f"🔢 {ln.get('qty_pcs', 0)}"
                )
            with cols[1]:
                if st.button("移除", key=f"sc_po_rm_{i}"):
                    cart.pop(i)
                    st.rerun()
    else:
        st.caption("—")

    st.divider()
    if st.button("提交采购单 / Submit PO", type="primary", key="sc_po_submit"):
        if not cart:
            st.error("请至少添加一行商品" if zh else "Add at least one line")
            return
        role = st.session_state.get("role") or ""
        try:
            po = sc.create_purchase_order(
                st.session_state,
                supplier_id=sup_id,
                lines=cart,
                remarks=remarks,
                created_by=role,
            )
        except ValueError as e:
            st.error(str(e))
            return
        if audit_write:
            try:
                audit_write(
                    "supplier_order",
                    extra={"po_id": po["po_id"], "supplier": po["supplier_name"], "lines": len(cart)},
                )
            except Exception:
                pass
        st.session_state.sc_po_cart = []
        st.session_state.sc_po_search_q = ""
        st.session_state.sc_last_po_id = po["po_id"]
        set_page("order_success")
        st.rerun()
