# Trax CDC Export Report extractor

Scanned shipment PDF (SAD + Commercial Invoice + Packing List + Info page) -> one new row in `CDC_Export_Report_2026.xlsx`.

## Setup (Windows, once)
1. Install Python 3.10+.
2. Install Tesseract OCR (UB-Mannheim installer, default folder is auto-detected).
3. `pip install -r requirements.txt`

## Run
- Window:  `python gui.py`  -> Browse PDF, Browse Excel, **Extract**, check values, **Write to Excel**
- Command: `python extractor.py shipment.pdf --xlsx CDC_Export_Report_2026.xlsx`
  - `--dry-run` check only, `--extras` also fill E and J, `--force` allow duplicate INV.NO, `--out new.xlsx` don't touch original, `--debug` save OCR text to `ocr_debug/`

Close the Excel file before writing. A backup copy goes to `backup/` each time.

## Red-mark -> Excel column
| # | Field | Col | Best source |
|---|---|---|---|
| 1 | INV.NO | C | Info page, Packing list, Invoice |
| 2 | Customer | D | Consignee -> `CUSTOMER_MAP` |
| 3 | Customs declaration no | F | Info page, SAD |
| 4 | Date | G | Info page, SAD |
| 5 | HTS code | I | Packing list, SAD |
| 6 | Description | M | Packing list, Invoice |
| 7 | Qty | N | Packing list, Invoice, SAD |
| 8 | Total price USD | P | Invoice, SAD |
| 9 | Destination | R | Invoice "TO UK" -> `COUNTRY_MAP`, SAD |
| 10 | Rate KHR | S | SAD box 23 |
| 11 / 12 | NW / GW | U / V | Packing list, Invoice |
| 13 | Kind of fabric | X | Invoice, Packing list |

Column T = `=P*S` (template formula). Other columns (E, H, J, K, AB ...) are left blank unless `--extras`.

## How it stays accurate
- Red pen marks are erased from the image before OCR.
- Each value is read from 2-3 documents and majority-voted; disagreements print as `[WARN]`.
- Checks: qty x unit price = amount, NW <= GW, HTS starts 61, invoice vs packing list transport mode.
- `[ERROR]` = nothing is written. Duplicate INV.NO is refused.

## Edit rules
`CUSTOMER_MAP`, `COUNTRY_MAP`, `COLUMNS` at the top of `extractor.py`.
One PDF = one shipment (one Excel row). Other invoice layouts may need a regex tweak in `extract()`.
