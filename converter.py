"""
OBZ Shipment Formatter - core conversion logic.

Reads a "Packing Slip" workbook (e.g. PLOB00123935PL.xlsx) and produces an
"OBZ Box Contents" workbook with two sheets:

    1. "Box Contents"          -> UPC | Qty | Box
    2. "Weight and Dimensions" -> Box Number | Weight | Length | Width | Height

The packing slip is laid out as a series of "Carton" blocks. Each block has a
header row that looks like:

    B: "Carton"   F: <carton number>   J: <tracking>   O: "Weight:"   T: <weight>   AA: "SSCC:"

followed (two rows down) by a sub-header row:

    B: "PO"   I: "Pick No."   M: "Item Number"   W: "UPC/GTIN"

and then one row per item, where column W holds the 14-digit GTIN.

The output strips the GTIN-14 down to the 12-digit UPC (the trailing 12 digits),
emits one line per scanned item with Qty = 1, and records the carton weight on
the second sheet. Box dimensions are not present in the packing slip, so they
default to standard values (configurable in the app).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from io import BytesIO

import openpyxl
from openpyxl.styles import Font

# Column indices (1-based) inside the packing slip ----------------------------
COL_LABEL = 2     # B - "Carton" / "PO"
COL_CARTON_NO = 6  # F - carton number on the header row
COL_WEIGHT = 20   # T - carton weight on the header row
COL_UPC = 23      # W - UPC/GTIN on the item rows
COL_PIN = 46      # AT - "PIN:" box dimensions (e.g. "24x20x16") on the header row

DEFAULT_LENGTH = 24
DEFAULT_WIDTH = 20
DEFAULT_HEIGHT = 16


@dataclass
class Carton:
    number: int
    weight: float | int | None = None
    length: float | int | None = None
    width: float | int | None = None
    height: float | int | None = None
    upcs: list[str] = field(default_factory=list)


def _clean_digits(value) -> str:
    """Return only the digit characters from a cell value."""
    if value is None:
        return ""
    text = str(value).strip()
    return "".join(ch for ch in text if ch.isdigit())


def gtin_to_upc(value) -> str:
    """Convert a GTIN-14 (e.g. '00840127879390') to a 12-digit UPC.

    The UPC-A is the trailing 12 digits of the GTIN, which also correctly
    preserves any leading zero that is genuinely part of the UPC.
    """
    digits = _clean_digits(value)
    if len(digits) > 12:
        return digits[-12:]
    return digits


def parse_dimensions(value):
    """Parse a 'PIN:' dimension string like '24x20x16' into (L, W, H).

    Returns a tuple of numbers (int/float). Missing/unparseable parts come back
    as None so the caller can fall back to defaults.
    """
    if value is None:
        return (None, None, None)
    parts = re.split(r"[xX\u00d7*]", str(value).strip())
    nums = [_as_number(p) for p in parts if p.strip() != ""]
    nums = (nums + [None, None, None])[:3]
    return (nums[0], nums[1], nums[2])


def _as_number(value):
    """Coerce a weight-like value to int or float (never a string)."""
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return value
    text = str(value).strip().replace(",", "")
    try:
        num = float(text)
    except ValueError:
        return None
    return int(num) if num.is_integer() else num


def parse_packing_slip(source) -> list[Carton]:
    """Parse the packing slip workbook into an ordered list of Cartons.

    `source` may be a path or a file-like object (BytesIO from an upload).
    """
    wb = openpyxl.load_workbook(source, data_only=True)
    ws = wb.active

    cartons: list[Carton] = []
    current: Carton | None = None

    for row in range(1, ws.max_row + 1):
        label = ws.cell(row=row, column=COL_LABEL).value
        label_str = str(label).strip().lower() if label is not None else ""

        # New carton block.
        if label_str == "carton":
            carton_no = _as_number(ws.cell(row=row, column=COL_CARTON_NO).value)
            weight = _as_number(ws.cell(row=row, column=COL_WEIGHT).value)
            length, width, height = parse_dimensions(
                ws.cell(row=row, column=COL_PIN).value
            )
            current = Carton(
                number=int(carton_no) if carton_no is not None else len(cartons) + 1,
                weight=weight,
                length=length,
                width=width,
                height=height,
            )
            cartons.append(current)
            continue

        # Item row: a UPC/GTIN that is a real code (not the "UPC/GTIN" header).
        raw_upc = ws.cell(row=row, column=COL_UPC).value
        digits = _clean_digits(raw_upc)
        if current is not None and len(digits) >= 12:
            current.upcs.append(gtin_to_upc(raw_upc))

    return cartons


def build_box_contents_workbook(
    cartons: list[Carton],
    length: int = DEFAULT_LENGTH,
    width: int = DEFAULT_WIDTH,
    height: int = DEFAULT_HEIGHT,
) -> openpyxl.Workbook:
    """Create the OBZ Box Contents workbook matching the reference template."""
    wb = openpyxl.Workbook()

    # --- Sheet 1: Box Contents ------------------------------------------------
    ws1 = wb.active
    ws1.title = "Box Contents"

    bold = Font(name="Calibri", size=11, bold=True)
    normal = Font(name="Calibri", size=11, bold=False)

    headers = [("A", "UPC", "0.00"), ("B", "Qty", "0"), ("C", "Box", "General")]
    for col, text, fmt in headers:
        cell = ws1[f"{col}1"]
        cell.value = text
        cell.font = bold
        cell.number_format = fmt

    row = 2
    for carton in cartons:
        for upc in carton.upcs:
            a = ws1.cell(row=row, column=1, value=upc)   # UPC kept as exact code
            a.number_format = "0.00"
            a.font = normal

            b = ws1.cell(row=row, column=2, value=1)     # Qty (real integer)
            b.number_format = "0"
            b.font = normal

            c = ws1.cell(row=row, column=3, value=int(carton.number))  # Box (real int)
            c.number_format = "General"
            c.font = normal
            row += 1

    ws1.column_dimensions["A"].width = 13.11
    ws1.column_dimensions["B"].width = 8.89

    # --- Sheet 2: Weight and Dimensions --------------------------------------
    ws2 = wb.create_sheet("Weight and Dimensions")
    dim_headers = ["Box Number", "Weight", "Length", "Width", "Height"]
    for idx, text in enumerate(dim_headers, start=1):
        cell = ws2.cell(row=1, column=idx, value=text)
        cell.font = bold

    for i, carton in enumerate(cartons, start=2):
        ws2.cell(row=i, column=1, value=int(carton.number)).font = normal
        w = _as_number(carton.weight)
        ws2.cell(row=i, column=2, value=w).font = normal
        # Prefer the dimensions parsed from the carton's "PIN:" field; fall back
        # to the provided defaults when the slip omits them.
        L = carton.length if carton.length is not None else length
        W = carton.width if carton.width is not None else width
        H = carton.height if carton.height is not None else height
        ws2.cell(row=i, column=3, value=_as_number(L)).font = normal
        ws2.cell(row=i, column=4, value=_as_number(W)).font = normal
        ws2.cell(row=i, column=5, value=_as_number(H)).font = normal

    ws2.column_dimensions["A"].width = 11.44
    ws2.column_dimensions["C"].width = 8.78

    return wb


def convert(
    source,
    length: int = DEFAULT_LENGTH,
    width: int = DEFAULT_WIDTH,
    height: int = DEFAULT_HEIGHT,
) -> tuple[openpyxl.Workbook, list[Carton]]:
    cartons = parse_packing_slip(source)
    wb = build_box_contents_workbook(cartons, length, width, height)
    return wb, cartons


def workbook_to_bytes(wb: openpyxl.Workbook) -> bytes:
    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
