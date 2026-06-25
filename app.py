"""
OBZ Shipment Formatter - Streamlit app.

Upload a packing slip (e.g. PLOB00123935PL.xlsx) and download a formatted
"OBZ Box Contents" workbook with the "Box Contents" and "Weight and Dimensions"
tabs, matching the reference template.
"""

from __future__ import annotations

import re

import streamlit as st

from converter import (
    DEFAULT_HEIGHT,
    DEFAULT_LENGTH,
    DEFAULT_WIDTH,
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

with st.sidebar:
    st.header("Fallback box dimensions")
    st.write(
        "Dimensions are read from each carton's **PIN:** field (e.g. `24x20x16`). "
        "These values are only used for cartons where that field is missing."
    )
    length = st.number_input("Length", min_value=0, value=DEFAULT_LENGTH, step=1)
    width = st.number_input("Width", min_value=0, value=DEFAULT_WIDTH, step=1)
    height = st.number_input("Height", min_value=0, value=DEFAULT_HEIGHT, step=1)

uploaded = st.file_uploader(
    "Upload the packing slip",
    type=["xlsx", "xlsm", "xls", "csv"],
    accept_multiple_files=False,
    help="Supported formats: Excel (.xlsx, .xlsm, .xls) and CSV (.csv).",
)


def _default_output_name(input_name: str) -> str:
    """PLOB00123935PL.xlsx -> PLOB00123935 Box Contents.xlsx (best effort)."""
    stem = re.sub(r"\.[^.]+$", "", input_name, flags=re.IGNORECASE)
    stem = re.sub(r"PL$", "", stem).strip()
    base = stem if stem else "OBZ"
    return f"{base} Box Contents.xlsx"


if uploaded is None:
    st.info("Upload a packing slip workbook to get started.")
    st.stop()

try:
    file_bytes = uploaded.getvalue()
    cartons = parse_packing_slip(file_bytes, filename=uploaded.name)
except Exception as exc:  # noqa: BLE001
    st.error(f"Could not read that file: {exc}")
    st.stop()

if not cartons:
    st.warning(
        "No cartons were found in this workbook. Make sure it is a packing slip "
        "with 'Carton' header rows and a 'UPC/GTIN' column."
    )
    st.stop()

total_items = sum(len(c.upcs) for c in cartons)

c1, c2, c3 = st.columns(3)
c1.metric("Boxes / Cartons", len(cartons))
c2.metric("Total items (Qty)", total_items)
total_weight = sum((c.weight or 0) for c in cartons)
c3.metric("Total weight", round(total_weight, 2))

# Build the workbook.
wb = build_box_contents_workbook(cartons, int(length), int(width), int(height))
output_bytes = workbook_to_bytes(wb)

st.success(
    f"Generated workbook with {total_items} items across {len(cartons)} boxes."
)

st.download_button(
    label="⬇️ Download OBZ Box Contents.xlsx",
    data=output_bytes,
    file_name=_default_output_name(uploaded.name),
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

# --- Previews -----------------------------------------------------------------
st.subheader("Preview - Box Contents")
box_rows = [
    {"UPC": upc, "Qty": 1, "Box": int(c.number)}
    for c in cartons
    for upc in c.upcs
]
st.dataframe(box_rows, use_container_width=True, height=360)

st.subheader("Preview - Weight and Dimensions")
wd_rows = [
    {
        "Box Number": int(c.number),
        "Weight": c.weight,
        "Length": c.length if c.length is not None else int(length),
        "Width": c.width if c.width is not None else int(width),
        "Height": c.height if c.height is not None else int(height),
    }
    for c in cartons
]
st.dataframe(wd_rows, use_container_width=True)

st.caption(
    "UPC is the 12-digit code (GTIN-14 with the leading digits removed). "
    "Qty, Box, Weight and dimensions are written as real numbers, not text."
)
