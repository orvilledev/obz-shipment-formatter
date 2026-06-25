# OBZ Shipment Formatter

A small Streamlit app that converts a **packing slip** workbook
(e.g. `PLOB00123935PL.xlsx`) into the **OBZ Box Contents** workbook with two
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
- **Qty** – one line per scanned item, quantity `1`.
- **Box** – the carton number the item belongs to.
- **Weight** – taken from each carton header.
- **Length / Width / Height** – read from each carton's **`PIN:`** field
  (e.g. `24x20x16` → L=24, W=20, H=16). The sidebar values are only used as a
  fallback for cartons that have no `PIN:` field.

All of `Qty`, `Box`, `Weight`, and the dimensions are written as **real numbers**
(int/float), never as text. The `UPC` is kept as the exact code string so no
leading digits are lost, matching the reference file.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then upload a packing slip `.xlsx` and download the generated
`... Box Contents.xlsx`.

## Files

- `app.py` – Streamlit UI.
- `converter.py` – parsing + workbook generation (the reusable core logic).
- `requirements.txt` – pinned dependencies.
