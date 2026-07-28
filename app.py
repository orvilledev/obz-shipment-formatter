"""
OBZ Shipment Formatter - Streamlit app.

Upload a packing slip (and optionally a DIMS file for accurate box dimensions)
and download a formatted "OBZ Box Contents" workbook with the
"Box Contents" and "Weight and Dimensions" tabs.
"""

from __future__ import annotations

import re

import streamlit as st

from converter import (
    DEFAULT_HEIGHT,
    DEFAULT_LENGTH,
    DEFAULT_WIDTH,
    apply_dims_to_cartons,
    build_box_contents_workbook,
    parse_dims_file,
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
        "Used only when a carton has no dimensions from a **DIMS** file "
        "or from the packing slip **PIN:** field."
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

dims_uploaded = st.file_uploader(
    "Upload the DIMS file (optional, for accurate L × W × H)",
    type=["xlsx", "xlsm", "xls", "csv"],
    accept_multiple_files=False,
    help=(
        "Companion 'DIMS …' workbook with Carton / Length / Width / Height columns. "
        "Required for accurate dimensions when the packing slip PIN field is empty."
    ),
)


def _default_output_name(input_name: str) -> str:
    """PLOB00123935PL.xlsx -> PLOB00123935 Box Contents.xlsx (best effort)."""
    stem = re.sub(r"\.[^.]+$", "", input_name, flags=re.IGNORECASE)
    stem = re.sub(r"(PackingSlip(By)?\s*)", "", stem, flags=re.IGNORECASE).strip()
    stem = re.sub(r"PL$", "", stem).strip()
    base = stem if stem else "OBZ"
    return f"{base} Box Contents.xlsx"


if uploaded is None:
    st.info("Upload a packing slip workbook to get started. Add the matching DIMS file for accurate box dimensions.")
    st.stop()

try:
    file_bytes = uploaded.getvalue()
    cartons = parse_packing_slip(file_bytes, filename=uploaded.name)
except Exception as exc:  # noqa: BLE001
    st.error(f"Could not read that packing slip: {exc}")
    st.stop()

if not cartons:
    st.warning(
        "No cartons were found in this workbook. Make sure it is a packing slip "
        "with 'Carton' header rows and a 'UPC/GTIN' column."
    )
    st.stop()

dims_matched = 0
if dims_uploaded is not None:
    try:
        dims = parse_dims_file(dims_uploaded.getvalue(), filename=dims_uploaded.name)
        dims_matched = apply_dims_to_cartons(cartons, dims)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not read that DIMS file: {exc}")
        st.stop()

total_items = sum(c.total_qty for c in cartons)
total_lines = sum(len(c.items) for c in cartons)
missing_dims = sum(
    1 for c in cartons if c.length is None or c.width is None or c.height is None
)

c1, c2, c3, c4 = st.columns(4)
c1.metric("Boxes / Cartons", len(cartons))
c2.metric("Total Qty", total_items)
total_weight = sum((c.weight or 0) for c in cartons)
c3.metric("Total weight", round(total_weight, 2))
c4.metric("DIMS matched", dims_matched if dims_uploaded else "—")

if dims_uploaded is None and missing_dims:
    st.warning(
        f"{missing_dims} carton(s) have no PIN dimensions. "
        "Upload the matching **DIMS** file for accurate Length / Width / Height "
        "(otherwise sidebar fallbacks are used)."
    )
elif dims_uploaded is not None and dims_matched < len(cartons):
    st.warning(
        f"DIMS matched {dims_matched} of {len(cartons)} cartons. "
        "Unmatched cartons will use PIN values or sidebar fallbacks."
    )
elif dims_uploaded is not None:
    st.success(f"Applied dimensions from DIMS to all {dims_matched} cartons.")

# Build the workbook.
wb = build_box_contents_workbook(cartons, int(length), int(width), int(height))
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

# --- Previews -----------------------------------------------------------------
st.subheader("Preview - Box Contents")
box_rows = [
    {"UPC": item.upc, "Qty": int(item.qty), "Box": int(c.number)}
    for c in cartons
    for item in c.items
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
        "Source": (
            "DIMS/PIN"
            if c.length is not None and c.width is not None and c.height is not None
            else "Fallback"
        ),
    }
    for c in cartons
]
st.dataframe(wd_rows, use_container_width=True)

st.caption(
    "UPC is the 12-digit code (GTIN-14 with the leading digits removed). "
    "Qty is read from the packing slip Qty column. "
    "Length / Width / Height prefer the DIMS file, then PIN, then sidebar fallbacks. "
    "Numbers are written as real numbers, not text."
)
