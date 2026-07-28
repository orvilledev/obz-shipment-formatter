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
emits one line per packing-slip item with the real Qty from column AW, and records
the carton weight on the second sheet. Box dimensions come from, in order:
  1. an optional companion DIMS file (Length / Width / Height per carton),
  2. each carton's PIN field on the packing slip,
  3. configurable sidebar defaults.
"""

from __future__ import annotations

import csv
import re
from dataclasses import dataclass, field
from io import BytesIO, StringIO
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".csv"}

# Column indices (1-based) inside the packing slip ----------------------------
COL_LABEL = 2     # B - "Carton" / "PO"
COL_CARTON_NO = 6  # F - carton number on the header row
COL_LICENSE = 10  # J - license plate / tracking on the header row
COL_WEIGHT = 20   # T - carton weight on the header row
COL_UPC = 23      # W - UPC/GTIN on the item rows
COL_QTY = 49      # AW - Qty on the item rows
COL_PIN = 46      # AT - "PIN:" box dimensions (e.g. "24x20x16") on the header row

DEFAULT_LENGTH = 24
DEFAULT_WIDTH = 20
DEFAULT_HEIGHT = 16


@dataclass
class LineItem:
    upc: str
    qty: int


@dataclass
class Carton:
    number: int
    weight: float | int | None = None
    length: float | int | None = None
    width: float | int | None = None
    height: float | int | None = None
    license_plate: str | None = None
    items: list[LineItem] = field(default_factory=list)

    @property
    def upcs(self) -> list[str]:
        """Back-compat helper: list of UPC codes only."""
        return [item.upc for item in self.items]

    @property
    def total_qty(self) -> int:
        return sum(item.qty for item in self.items)


@dataclass
class DimsRecord:
    """One carton row from a DIMS workbook."""

    carton: int | None = None
    license_plate: str | None = None
    length: float | int | None = None
    width: float | int | None = None
    height: float | int | None = None
    weight: float | int | None = None
    units: int | None = None


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


@dataclass
class CellGrid:
    """Row/column grid with 1-based indexing, like an Excel worksheet."""

    _rows: list[list]

    @property
    def max_row(self) -> int:
        return len(self._rows)

    @property
    def max_column(self) -> int:
        return max((len(row) for row in self._rows), default=0)

    def cell_value(self, row: int, column: int):
        if row < 1 or column < 1:
            return None
        row_values = self._rows[row - 1]
        if column > len(row_values):
            return None
        value = row_values[column - 1]
        if value == "":
            return None
        return value


def _normalize_extension(filename: str | None) -> str:
    if not filename:
        return ".xlsx"
    return Path(filename).suffix.lower()


def _read_bytes(source) -> bytes:
    if isinstance(source, (bytes, bytearray)):
        return bytes(source)
    if hasattr(source, "read"):
        data = source.read()
        if hasattr(source, "seek"):
            source.seek(0)
        return data
    return Path(source).read_bytes()


def _grid_from_openpyxl(ws) -> CellGrid:
    rows: list[list] = []
    for row in ws.iter_rows(values_only=True):
        rows.append(list(row))
    return CellGrid(rows)


def _grid_from_xls(data: bytes) -> CellGrid:
    import xlrd

    book = xlrd.open_workbook(file_contents=data)
    sheet = book.sheet_by_index(0)
    rows: list[list] = []
    for r in range(sheet.nrows):
        row_values = []
        for c in range(sheet.ncols):
            cell = sheet.cell(r, c)
            if cell.ctype == xlrd.XL_CELL_EMPTY:
                row_values.append(None)
            elif cell.ctype == xlrd.XL_CELL_NUMBER:
                value = cell.value
                row_values.append(int(value) if value == int(value) else value)
            else:
                row_values.append(cell.value)
        rows.append(row_values)
    return CellGrid(rows)


def _grid_from_csv(data: bytes) -> CellGrid:
    text = None
    for encoding in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            text = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = data.decode("utf-8", errors="replace")

    rows: list[list] = []
    for row in csv.reader(StringIO(text)):
        rows.append(row)
    return CellGrid(rows)


def load_cell_grid(source, filename: str | None = None) -> CellGrid:
    """Load a packing slip from Excel (.xlsx/.xlsm/.xls) or CSV into a grid."""
    ext = _normalize_extension(filename)
    if ext not in SUPPORTED_EXTENSIONS:
        supported = ", ".join(sorted(SUPPORTED_EXTENSIONS))
        raise ValueError(f"Unsupported file type '{ext}'. Use one of: {supported}")

    data = _read_bytes(source)

    if ext == ".csv":
        return _grid_from_csv(data)
    if ext == ".xls":
        return _grid_from_xls(data)

    wb = openpyxl.load_workbook(BytesIO(data), data_only=True)
    return _grid_from_openpyxl(wb.active)


def _scan_carton_fields(grid: CellGrid, row: int):
    """Read carton number, license plate, weight, and PIN dimensions from a header row."""
    carton_no = _as_number(grid.cell_value(row, COL_CARTON_NO))
    license_plate = grid.cell_value(row, COL_LICENSE)
    if license_plate is not None:
        license_plate = str(license_plate).strip() or None
    weight = _as_number(grid.cell_value(row, COL_WEIGHT))
    length, width, height = parse_dimensions(grid.cell_value(row, COL_PIN))

    row_values = [
        grid.cell_value(row, col)
        for col in range(1, grid.max_column + 1)
    ]

    if length is None:
        for value in row_values:
            if value and re.search(r"\d+[xX\u00d7*]\d+", str(value)):
                length, width, height = parse_dimensions(value)
                break

    if weight is None:
        for idx, value in enumerate(row_values):
            if value and "weight" in str(value).lower():
                for later in row_values[idx + 1 :]:
                    parsed = _as_number(later)
                    if parsed is not None:
                        weight = parsed
                        break
                break

    if carton_no is None:
        seen_carton = False
        for value in row_values:
            if value is None:
                continue
            text = str(value).strip().lower()
            if text == "carton":
                seen_carton = True
                continue
            if seen_carton:
                parsed = _as_number(value)
                if parsed is not None:
                    carton_no = parsed
                    break

    if license_plate is None:
        # License plate is typically a long digit string on the carton header.
        for value in row_values:
            if value is None:
                continue
            text = str(value).strip()
            digits = _clean_digits(text)
            if len(digits) >= 8 and digits == text:
                license_plate = text
                break

    return carton_no, license_plate, weight, length, width, height


def _detect_header_column(grid: CellGrid, row: int, *names: str) -> int | None:
    """Find a column whose header cell contains any of the given names."""
    targets = [n.lower() for n in names]
    for col in range(1, grid.max_column + 1):
        value = grid.cell_value(row, col)
        if value is None:
            continue
        text = str(value).strip().lower()
        if any(t == text or t in text for t in targets):
            return col
    return None


def parse_packing_slip_grid(grid: CellGrid) -> list[Carton]:
    """Parse a cell grid into an ordered list of Cartons."""
    cartons: list[Carton] = []
    current: Carton | None = None
    col_upc = COL_UPC
    col_qty = COL_QTY

    for row in range(1, grid.max_row + 1):
        label = grid.cell_value(row, COL_LABEL)
        label_str = str(label).strip().lower() if label is not None else ""

        if label_str == "po":
            detected_upc = _detect_header_column(grid, row, "upc", "upc/gtin", "gtin")
            if detected_upc is not None:
                col_upc = detected_upc
            detected_qty = _detect_header_column(grid, row, "qty", "quantity")
            if detected_qty is not None:
                col_qty = detected_qty
            continue

        if label_str == "carton":
            carton_no, license_plate, weight, length, width, height = _scan_carton_fields(
                grid, row
            )
            current = Carton(
                number=int(carton_no) if carton_no is not None else len(cartons) + 1,
                weight=weight,
                length=length,
                width=width,
                height=height,
                license_plate=license_plate,
            )
            cartons.append(current)
            continue

        raw_upc = grid.cell_value(row, col_upc)
        digits = _clean_digits(raw_upc)
        if current is not None and len(digits) >= 12:
            qty_val = _as_number(grid.cell_value(row, col_qty))
            qty = int(qty_val) if qty_val is not None and qty_val > 0 else 1
            current.items.append(LineItem(upc=gtin_to_upc(raw_upc), qty=qty))

    return cartons


def parse_packing_slip(source, filename: str | None = None) -> list[Carton]:
    """Parse the packing slip workbook into an ordered list of Cartons.

    `source` may be a path, bytes, or file-like object (e.g. Streamlit upload).
    `filename` is used to detect the file type (.xlsx, .xls, .xlsm, .csv).
    """
    grid = load_cell_grid(source, filename)
    return parse_packing_slip_grid(grid)


def _find_header_row(grid: CellGrid, required: set[str], scan_rows: int = 15) -> tuple[int, dict[str, int]] | None:
    """Locate a header row containing the required column names.

    Returns (row_number, {normalized_name: column_index}) or None.
    """
    required_norm = {name.lower() for name in required}
    for row in range(1, min(grid.max_row, scan_rows) + 1):
        mapping: dict[str, int] = {}
        for col in range(1, grid.max_column + 1):
            value = grid.cell_value(row, col)
            if value is None:
                continue
            key = str(value).strip().lower()
            if key:
                mapping[key] = col
        if required_norm.issubset(mapping.keys()):
            return row, mapping
    return None


def parse_dims_file(source, filename: str | None = None) -> list[DimsRecord]:
    """Parse a DIMS workbook (Length / Width / Height / Weight per carton).

    Expected columns (header row, typically row 3):
        Carton | License Plate | Length | Width | Height | Weight | Units | ...
    """
    grid = load_cell_grid(source, filename)
    found = _find_header_row(
        grid, {"carton", "length", "width", "height"}
    )
    if found is None:
        raise ValueError(
            "Could not find a DIMS header row with Carton / Length / Width / Height columns."
        )

    header_row, cols = found
    col_carton = cols["carton"]
    col_length = cols["length"]
    col_width = cols["width"]
    col_height = cols["height"]
    col_license = cols.get("license plate")
    col_weight = cols.get("weight")
    col_units = cols.get("units")

    records: list[DimsRecord] = []
    for row in range(header_row + 1, grid.max_row + 1):
        carton_no = _as_number(grid.cell_value(row, col_carton))
        length = _as_number(grid.cell_value(row, col_length))
        width = _as_number(grid.cell_value(row, col_width))
        height = _as_number(grid.cell_value(row, col_height))

        # Skip total / footer rows that lack carton + dimensions.
        if carton_no is None and length is None:
            continue
        if carton_no is None:
            continue

        license_plate = None
        if col_license is not None:
            raw = grid.cell_value(row, col_license)
            if raw is not None:
                license_plate = str(raw).strip() or None

        weight = _as_number(grid.cell_value(row, col_weight)) if col_weight else None
        units_val = _as_number(grid.cell_value(row, col_units)) if col_units else None
        units = int(units_val) if units_val is not None else None

        records.append(
            DimsRecord(
                carton=int(carton_no),
                license_plate=license_plate,
                length=length,
                width=width,
                height=height,
                weight=weight,
                units=units,
            )
        )

    return records


def apply_dims_to_cartons(cartons: list[Carton], dims: list[DimsRecord]) -> int:
    """Merge DIMS Length/Width/Height (and Weight if present) onto packing-slip cartons.

    Match priority: license plate, then carton number.
    Returns the number of cartons updated.
    """
    by_license = {
        d.license_plate: d
        for d in dims
        if d.license_plate
    }
    by_number = {d.carton: d for d in dims if d.carton is not None}

    updated = 0
    for carton in cartons:
        match = None
        if carton.license_plate and carton.license_plate in by_license:
            match = by_license[carton.license_plate]
        elif carton.number in by_number:
            match = by_number[carton.number]

        if match is None:
            continue

        if match.length is not None:
            carton.length = match.length
        if match.width is not None:
            carton.width = match.width
        if match.height is not None:
            carton.height = match.height
        # Prefer DIMS weight when present (authoritative for the carton list).
        if match.weight is not None:
            carton.weight = match.weight
        updated += 1

    return updated


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
        for item in carton.items:
            a = ws1.cell(row=row, column=1, value=item.upc)  # UPC kept as exact code
            a.number_format = "0.00"
            a.font = normal

            b = ws1.cell(row=row, column=2, value=int(item.qty))  # Qty (real integer)
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
    filename: str | None = None,
    dims_source=None,
    dims_filename: str | None = None,
    length: int = DEFAULT_LENGTH,
    width: int = DEFAULT_WIDTH,
    height: int = DEFAULT_HEIGHT,
) -> tuple[openpyxl.Workbook, list[Carton]]:
    cartons = parse_packing_slip(source, filename=filename)
    if dims_source is not None:
        dims = parse_dims_file(dims_source, filename=dims_filename)
        apply_dims_to_cartons(cartons, dims)
    wb = build_box_contents_workbook(cartons, length, width, height)
    return wb, cartons


def workbook_to_bytes(wb: openpyxl.Workbook) -> bytes:
    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
