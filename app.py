"""
Streamlit demo: upload Inventory + QR codes files, get a filled Bulk Edit master template.

Run with:
    streamlit run app.py
"""
import io
from pathlib import Path

import pandas as pd
import streamlit as st

from build_bulk_edit import (
    build_qr_lookup,
    collect_rows,
    write_output,
    DEFAULT_INVENTORY,
    DEFAULT_QRCODES,
)

st.set_page_config(page_title="Bulk Edit Builder", page_icon=None, layout="wide")

st.title("Bulk Edit Master Template Builder")
st.write(
    "Upload the **StockTake inventory** and **QR codes** files. "
    "The app filters inventory rows, looks up each location's QR code, "
    "and generates the Bulk Edit master template for download."
)

with st.expander("How the mapping works", expanded=False):
    st.markdown(
        """
        1. Filter StockTake rows where **Room = Inventory**.
        2. For each row, read **Deal Id** (column E) and **Location** (column C).
        3. XLOOKUP the Location in the QR codes' **Description** column and return the matching **UserQrCode**.
           Annotated locations like `R58 (AUH-D1-INVOICE-21-04)` are matched on the prefix `R58`.
        4. Rows with no Deal Id are skipped (Appraisal code would be blank).
        5. Output is sorted by Deal Id.
        """
    )

col1, col2 = st.columns(2)
with col1:
    inv_file = st.file_uploader("Inventory file (StockTake)", type=["xlsx"], key="inv")
with col2:
    qr_file = st.file_uploader("QR codes file", type=["xlsx"], key="qr")

st.caption("Master template is generated from scratch — no template upload needed.")

samples_available = Path(DEFAULT_INVENTORY).exists() and Path(DEFAULT_QRCODES).exists()
use_defaults = False
if samples_available:
    use_defaults = st.checkbox(
        "Use the sample inputs already in this folder (for demo)",
        value=not any([inv_file, qr_file]),
    )

st.divider()

if st.button("Generate filled master template", type="primary"):
    # Inputs: either uploads or sample fallback
    if use_defaults:
        inv_src = DEFAULT_INVENTORY
        qr_src = DEFAULT_QRCODES
        for p in (inv_src, qr_src):
            if not Path(p).exists():
                st.error(f"Sample file not found: {p}")
                st.stop()
    else:
        if not (inv_file and qr_file):
            st.error("Please upload both the Inventory and QR codes files.")
            st.stop()
        inv_src = io.BytesIO(inv_file.getvalue())
        qr_src = io.BytesIO(qr_file.getvalue())

    with st.status("Processing...", expanded=True) as status:
        st.write("Loading QR code descriptions...")
        qr_lookup = build_qr_lookup(qr_src)
        st.write(f"  {len(qr_lookup):,} unique QR descriptions loaded")

        st.write("Filtering inventory and matching QR codes...")
        rows, unmatched, skipped = collect_rows(inv_src, qr_lookup)
        matched = len(rows) - len(unmatched)
        st.write(f"  {len(rows):,} inventory rows | matched {matched:,} | unmatched {len(unmatched)} | skipped (No Deal Id) {skipped}")

        st.write("Building master template...")
        out_buf = io.BytesIO()
        write_output(out_buf, rows)
        out_buf.seek(0)
        status.update(label="Done", state="complete")

    # Summary metrics
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Rows written", f"{len(rows):,}")
    c2.metric("Matched QR", f"{matched:,}")
    c3.metric("Unmatched QR", f"{len(unmatched):,}")
    c4.metric("Skipped (No Deal Id)", f"{skipped:,}")

    # Preview
    st.subheader("Preview")
    preview_df = pd.DataFrame(
        [(d, q) for d, q, _l in rows[:100]],
        columns=["Appraisal code", "User Qr Code"],
    )
    st.dataframe(preview_df, use_container_width=True, hide_index=True)
    st.caption(f"Showing first {len(preview_df)} of {len(rows):,} rows.")

    # Unmatched warning
    if unmatched:
        st.warning(f"{len(unmatched)} row(s) have no QR code match.")
        st.dataframe(
            pd.DataFrame(unmatched, columns=["Deal Id", "Location"]),
            use_container_width=True,
            hide_index=True,
        )

    # Download button
    st.download_button(
        label="Download filled master template",
        data=out_buf.getvalue(),
        file_name="Bulk Edit - filled.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
