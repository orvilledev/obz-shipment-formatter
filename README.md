# OBZ Shipment Formatter

A small Streamlit app that converts a **packing slip** file
(Excel `.xlsx` / `.xlsm` / `.xls`, `.csv`, or `.pdf`) into the **OBZ Box Contents**
workbook with two tabs, matching the reference template.

## What it does

Input: a packing slip laid out as a series of `Carton` blocks. Each block has a
header row (`Carton | <no.> | <tracking> | Weight: | <weight> | PIN: <LxWxH>`)
followed by item rows that contain `UPC/GTIN` and `Qty`.

Output: an `.xlsx` with two sheets.

| Sheet | Columns |
| --- | --- |
| **Box Contents** | `UPC` &nbsp;\|&nbsp; `Qty` &nbsp;\|&nbsp; `Box` |
| **Weight and Dimensions** | `Box Number` \| `Weight` \| `Length` \| `Width` \| `Height` |

Key transformations:

- **UPC** – the 14-digit GTIN (`00840127879390`) is reduced to the 12-digit
  UPC-A (`840127879390`) by keeping the trailing 12 digits.
- **Qty** – read from the packing slip **Qty** column (not assumed to be 1).
- **Box** – the carton number the item belongs to.
- **Weight** – taken from each carton header.
- **Length / Width / Height** – read from each carton's **`PIN:`** field on the
  uploaded packing slip (e.g. `24x20x16`). Missing PIN values are left blank —
  they are never filled with static defaults.

All of `Qty`, `Box`, `Weight`, and the dimensions are written as **real numbers**
(int/float), never as text. The `UPC` is kept as the exact code string so no
leading digits are lost.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then upload a packing slip (`.xlsx`, `.xlsm`, `.xls`, `.csv`, or `.pdf`) and
download the generated `... Box Contents.xlsx`.

## Files

- `app.py` – Streamlit UI.
- `converter.py` – parsing + workbook generation (the reusable core logic).
- `requirements.txt` – dependencies.
