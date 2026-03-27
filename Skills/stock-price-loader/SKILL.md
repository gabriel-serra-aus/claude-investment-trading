---
name: stock-price-loader
description: >
  Loads historical stock price data from a multi-tab Excel file into a Stock Prices table
  in a portfolio workbook. Use this skill whenever the user wants to import, load, sync, or
  update stock prices from a spreadsheet where each tab is a ticker symbol. Triggers on
  phrases like "load stock prices", "import price history", "add prices to the table",
  "update stock prices from file", or any request to move ticker/date/price data from one
  Excel file into a portfolio or prices table. Also triggers when the user points at a file
  with multiple ticker tabs and asks to populate a target table.
---

# Stock Price Loader

Loads historical price data from a multi-tab Excel source file into a target Stock Prices
table inside a portfolio workbook. Each tab in the source file represents one ticker, with
Date and price columns. All entries are written into a flat target table with columns:
Date, Ticker, USD Price, and AUD Price.

---

## Expected file structure

### Source file (e.g. `USD Stock Prices History.xlsx`)
- Each **sheet name** = ticker symbol (e.g. TSLA, AAPL, MSFT)
- Each sheet has a **header row** (row 1) followed by data rows
- **Column A** = Date (datetime or date)
- **Column B** = Close / Price (numeric, USD)

### Target file (e.g. `Portfolio.xlsx`)
- Contains a sheet named **"Stock Prices"** (or similar — confirm with user if unclear)
- That sheet has an Excel Table with columns in this order:
  1. **Date**
  2. **Ticker**
  3. **USD Price**
  4. **AUD Price** (formula: USD Price ÷ AUD→USD rate from a Rates table)

---

## Workflow

### Step 1 — Identify files

Confirm source and target paths from the user's message or the current folder.
If ambiguous, ask before proceeding.

### Step 2 — Read source file

Use the bundled script `scripts/load_stock_prices.py` as the foundation — adapt as needed
for the specific file paths and table names in this session.

Loop through every sheet in the source workbook:
- Skip the header row (row 1)
- For each data row: extract date (strip time component if datetime), price
- Collect as: `(date, ticker=sheet_name, usd_price)`
- Skip rows where date or price is None

### Step 3 — Write to target table

Open the target workbook. Find the Stock Prices sheet and its Excel Table.

- Clear any existing data rows (preserve header row 1)
- Sort all collected rows by (ticker, date) for clean ordering
- Write each row: Date → col 1, Ticker → col 2, USD Price → col 3
- Add AUD Price formula in col 4:
  `=IFERROR(C{row}/INDEX(Rates!$B:$B,MATCH(A{row},Rates!$A:$A,0)),"")`
  This divides USD Price by the AUD→USD rate from the Rates sheet for that date.
  If no Rates sheet exists, leave AUD Price blank.

### Step 4 — Update table reference

Update the table's `ref` to cover all rows including the new data:
`A1:D{last_row}`

Keep table style as `TableStyleLight2` with row stripes.

### Step 5 — Apply number formats only

**CRITICAL: Do NOT touch cell background colour, font, or font size.**
Only set `cell.number_format`:
- Date column: `mm-dd-yy`
- USD Price and AUD Price columns: `'_-"$"* #,##0.00_-;\\-"$"* #,##0.00_-;_-"$"* "-"??_-;_-@_-'`
- Ticker column: `General`

### Step 6 — Recalculate and verify

Run the recalc script to resolve formulas and check for errors:
```bash
python /sessions/zen-busy-keller/mnt/.skills/skills/xlsx/scripts/recalc.py <target_file> 120
```

If errors are found in `error_summary`, diagnose and fix before saving the final file.

### Step 7 — Save and report

Copy the finished file to the workspace folder so the user can open it.
Report: tickers loaded, rows per ticker, total rows, any errors encountered.

---

## Number format standards (from project memory)

These apply across all Excel work in this workspace:

| Column type     | Format |
|-----------------|--------|
| Date            | `mm-dd-yy` |
| Currency/Price  | `_-"$"* #,##0.00_-;\\-"$"* #,##0.00_-;_-"$"* "-"??_-;_-@_-` |
| Exchange rate   | `0.00000` |
| Quantity        | `#,##0.0000` |

---

## Edge cases

- **Ticker tab with no data**: skip it, log a warning
- **Missing Rates sheet**: skip AUD Price formula, leave column blank
- **Duplicate rows** (same date + ticker): write both — don't deduplicate unless asked
- **Non-standard column order in source**: check header row to find Date and Close/Price columns by name, not position
- **Table not found**: tell the user the table name you're looking for and ask for clarification
