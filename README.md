# OBZ Shipment Formatter

A small Streamlit app that converts a **packing slip** file
(Excel `.xlsx` / `.xlsm` / `.xls` or `.csv`, e.g. `PLOB00123935PL.xlsx`) into the **OBZ Box Contents** workbook with two
tabs, matching the reference template exactly.

## What it does

Input: a packing slip laid out as a series of `Carton` blocks. Each block has a
header row (`Carton | <no.> | <tracking> | Weight: | <weight> | SSCC:`) followed
by item rows that contain a `UPC/GTIN` column.

Output: an `.xlsx` with two sheets.

| Sheet | Columns |
| --- | --- |
| **Box Contents** | `UPC` &nbsp;\|&nbsp; `Qty` &nbsp;\|&nbsp; `Box` |
| **Weight and Dimensions** | `Box Number` \| `Weight` \| `Length` \| `Width` \| `Height` |

Key transformations:

- **UPC** – the 14-digit GTIN (`00840127879390`) is reduced to the 12-digit
  UPC-A (`840127879390`) by keeping the trailing 12 digits.
- **Qty** – read from the packing slip **Qty** column (not assumed to be 1).
  When a line has Qty `2`, the output row keeps Qty as `2`.
- **Box** – the carton number the item belongs to.
- **Weight** – taken from each carton header (overridden by DIMS when provided).
- **Length / Width / Height** – prefer the companion **DIMS** file
  (`Carton | Length | Width | Height`), then each carton's **`PIN:`** field
  (e.g. `24x20x16`), then sidebar fallbacks. Many FBA packing slips leave PIN
  empty, so uploading the matching DIMS file is required for accurate dimensions.

All of `Qty`, `Box`, `Weight`, and the dimensions are written as **real numbers**
(int/float), never as text. The `UPC` is kept as the exact code string so no
leading digits are lost, matching the reference file.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then upload a packing slip (`.xlsx`, `.xlsm`, `.xls`, or `.csv`) — and preferably
the matching **DIMS** file for accurate box dimensions — and download the generated
`... Box Contents.xlsx`.

## Files

- `app.py` – Streamlit UI.
- `converter.py` – parsing + workbook generation (the reusable core logic).
- `requirements.txt` – pinned dependencies.
