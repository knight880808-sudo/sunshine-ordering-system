"""
环节 C — 收货核对（Warehouse / Admin）
仅 Pending 采购单；确认后 → Received，入库，差异 → Receipt_Discrepancies。
"""
from __future__ import annotations

import streamlit as st

from modules import supply_chain_store as sc


def _zh() -> bool:
    return st.session_state.get("lang") == "zh"


def render(render_page_heading, set_page, audit_write=None) -> None:
    zh = _zh()

    if st.session_state.get("page") == "verify_success":
        render_page_heading(
            "✅ 收货已入库" if zh else "✅ Receipt confirmed",
            None,
        )
        po_id = st.session_state.get("sc_verify_po_id") or ""
        n_disc = int(st.session_state.get("sc_verify_disc_count") or 0)
        st.success(f"PO `{po_id}` → Received")
        if n_disc > 0:
            st.warning(
                f"记录 {n_disc} 条数量差异"
                if zh
                else f"{n_disc} discrepancy record(s) logged"
            )
        if st.button("继续收货 / Next receipt", type="primary", key="sc_vfy_done"):
            st.session_state.pop("sc_verify_po_id", None)
            st.session_state.pop("sc_verify_disc_count", None)
            set_page("verify_inbound")
            st.rerun()
        return

    render_page_heading(
        "📥 收货核对" if zh else "📥 Inbound verification",
        "仅显示待收货 Pending 采购单；核对实收数量后确认入库。"
        if zh
        else "Pending POs only; confirm received qty to post inventory.",
    )

    pending = sc.list_purchase_orders(st.session_state, status=sc.PO_STATUS_PENDING)
    if not pending:
        st.info("当前没有待收货的采购单。" if zh else "No pending purchase orders.")
        return

    po_labels = {
        p["po_id"]: f"{p['po_id']} · {p.get('supplier_name', '')} · {p.get('created_at', '')}"
        for p in pending
    }
    po_id = st.selectbox(
        "采购单号 / PO",
        options=list(po_labels.keys()),
        format_func=lambda k: po_labels[k],
        key="sc_verify_po_pick",
    )
    po = sc.get_purchase_order(st.session_state, po_id)
    if not po:
        st.error("采购单不存在")
        return

    st.caption(
        f"供应商: **{po.get('supplier_name', '')}** · "
        f"备注: {po.get('remarks') or '—'}"
    )

    items = sc.list_po_items(st.session_state, po_id)
    if not items:
        st.warning("该采购单没有明细行")
        return

    if f"sc_recv_{po_id}" not in st.session_state:
        st.session_state[f"sc_recv_{po_id}"] = {
            (it.get("line_id") or ""): {
                "qty_cartons": int(it.get("qty_cartons_ordered") or 0),
                "qty_pcs": int(it.get("qty_pcs_ordered") or 0),
            }
            for it in items
        }

    recv_state: dict = st.session_state[f"sc_recv_{po_id}"]
    st.markdown("**" + ("核对实收数量" if zh else "Received quantities") + "**")

    received_lines: list[dict] = []
    for it in items:
        lid = it.get("line_id") or ""
        cur = recv_state.get(lid, {})
        st.markdown(f"**{it.get('name', '')}** `{it.get('item_code', '')}`")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.caption("订购箱" if zh else "Ordered ct")
            st.write(int(it.get("qty_cartons_ordered") or 0))
        with c2:
            st.caption("订购件" if zh else "Ordered pcs")
            st.write(int(it.get("qty_pcs_ordered") or 0))
        with c3:
            r_ct = st.number_input(
                "实收箱" if zh else "Recv ct",
                min_value=0,
                value=int(cur.get("qty_cartons", 0)),
                key=f"sc_rct_{po_id}_{lid}",
            )
        with c4:
            r_pc = st.number_input(
                "实收件" if zh else "Recv pcs",
                min_value=0,
                value=int(cur.get("qty_pcs", 0)),
                key=f"sc_rpc_{po_id}_{lid}",
            )
        recv_state[lid] = {"qty_cartons": int(r_ct), "qty_pcs": int(r_pc)}
        received_lines.append({
            "item_code": it.get("item_code") or "",
            "barcode": it.get("barcode") or "",
            "name": it.get("name") or "",
            "unit": it.get("unit") or "",
            "qty_cartons": int(r_ct),
            "qty_pcs": int(r_pc),
        })

    st.session_state[f"sc_recv_{po_id}"] = recv_state

    if st.button(
        "确认收货并入库 / Confirm receipt",
        type="primary",
        key="sc_verify_confirm",
    ):
        ok, err, discs = sc.receive_purchase_order(
            st.session_state,
            po_id=po_id,
            received_lines=received_lines,
            operator=st.session_state.get("role") or "warehouse",
        )
        if not ok:
            msg = {
                "po_not_found": "采购单不存在",
                "po_not_pending": "该单已不是待收货状态",
            }.get(err, err)
            st.error(msg)
            return
        if audit_write:
            try:
                audit_write(
                    "receive_confirm",
                    extra={"po_id": po_id, "discrepancies": len(discs)},
                )
            except Exception:
                pass
        st.session_state.pop(f"sc_recv_{po_id}", None)
        st.session_state.sc_verify_po_id = po_id
        st.session_state.sc_verify_disc_count = len(discs)
        set_page("verify_success")
        st.rerun()

    st.divider()
    hist = sc.list_discrepancies(st.session_state, po_id=po_id)
    if hist:
        st.subheader("⚠️ " + ("本单历史差异" if zh else "Discrepancies for this PO"))
        st.dataframe(hist, use_container_width=True, hide_index=True)
