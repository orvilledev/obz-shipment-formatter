"""
OBZ Shipment Formatter - Streamlit app.

Upload a packing slip (Excel or PDF) and download a formatted
"OBZ Box Contents" workbook with the "Box Contents" and
"Weight and Dimensions" tabs.
"""

from __future__ import annotations

import re

import streamlit as st

from converter import (
    build_box_contents_workbook,
    parse_packing_slip,
    workbook_to_bytes,
)

st.set_page_config(page_title="OBZ Shipment Formatter", page_icon="📦", layout="wide")

st.title("📦 OBZ Shipment Formatter")
st.caption(
    "Convert a packing slip into the **OBZ Box Contents** workbook "
    "(Box Contents + Weight and Dimensions tabs)."
)

uploaded = st.file_uploader(
    "Upload the packing slip",
    type=["xlsx", "xlsm", "xls", "csv", "pdf"],
    accept_multiple_files=False,
    help="Supported: Excel (.xlsx, .xlsm, .xls), CSV, or PDF packing slip.",
)


def _default_output_name(input_name: str) -> str:
    """PLOB00123935PL.xlsx -> PLOB00123935 Box Contents.xlsx (best effort)."""
    stem = re.sub(r"\.[^.]+$", "", input_name, flags=re.IGNORECASE)
    stem = re.sub(r"(PackingSlip(By)?\s*)", "", stem, flags=re.IGNORECASE).strip()
    stem = re.sub(r"PL$", "", stem).strip()
    base = stem if stem else "OBZ"
    return f"{base} Box Contents.xlsx"


if uploaded is None:
    st.info("Upload a packing slip (Excel or PDF) to get started.")
    st.stop()

try:
    cartons = parse_packing_slip(uploaded.getvalue(), filename=uploaded.name)
except Exception as exc:  # noqa: BLE001
    st.error(f"Could not read that packing slip: {exc}")
    st.stop()

if not cartons:
    st.warning(
        "No cartons were found in this file. Make sure it is a packing slip "
        "with Carton sections, UPC/GTIN items, and PIN dimensions."
    )
    st.stop()

total_items = sum(c.total_qty for c in cartons)
total_lines = sum(len(c.items) for c in cartons)
missing_dims = sum(
    1 for c in cartons if c.length is None or c.width is None or c.height is None
)

c1, c2, c3 = st.columns(3)
c1.metric("Boxes / Cartons", len(cartons))
c2.metric("Total Qty", total_items)
c3.metric("Total weight", round(sum((c.weight or 0) for c in cartons), 2))

if missing_dims:
    st.warning(
        f"{missing_dims} carton(s) have no **PIN** dimensions in this file "
        "(Length / Width / Height will be left blank). "
        "PIN values look like `24x20x16` next to each Carton header."
    )

wb = build_box_contents_workbook(cartons)
output_bytes = workbook_to_bytes(wb)

st.success(
    f"Generated workbook with {total_items} units "
    f"({total_lines} line items) across {len(cartons)} boxes."
)

st.download_button(
    label="⬇️ Download OBZ Box Contents.xlsx",
    data=output_bytes,
    file_name=_default_output_name(uploaded.name),
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

st.subheader("Preview - Box Contents")
st.dataframe(
    [
        {"UPC": item.upc, "Qty": int(item.qty), "Box": int(c.number)}
        for c in cartons
        for item in c.items
    ],
    use_container_width=True,
    height=360,
)

st.subheader("Preview - Weight and Dimensions")
st.dataframe(
    [
        {
            "Box Number": int(c.number),
            "Weight": c.weight,
            "Length": c.length,
            "Width": c.width,
            "Height": c.height,
        }
        for c in cartons
    ],
    use_container_width=True,
)

st.caption(
    "UPC is the 12-digit code (GTIN-14 with leading digits removed). "
    "Qty comes from the packing slip Qty column. "
    "Length / Width / Height come only from each carton's **PIN:** field "
    "on the uploaded packing slip — not from static defaults."
)
