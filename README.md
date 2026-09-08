# xPence — Personal Expense Report Generator

xPence turns a messy bank/credit-card export into a clean, interactive HTML spending report — no server, no account, no data leaving your machine. Point it at a CSV/Excel file, get back a single self-contained `.html` file you can open in any browser, share, or archive.

![xPence report preview](https://img.shields.io/badge/status-personal--finance--tool-blue)

## Why this exists

Bank statements are inconsistent, unlabeled, and painful to categorize by hand. xPence:

- **Auto-detects** date, description, debit, credit, and account-type columns regardless of header names or column order.
- **Classifies every transaction** into a spending category using keyword matching plus your own learned corrections — no manual tagging required after the first pass.
- **Learns as you go**: reclassify something once in the report and xPence remembers it for every future run.
- **Never phones home**: everything runs locally with Python + pandas. Your transaction data never leaves your computer.

## What's in the report

- **Review** — flags credit-side transactions that look like disguised expenses (e.g. a "refund" that's actually a purchase) so you can confirm or correct them.
- **All Spending** — a category breakdown bar chart, a spending-over-time chart, and a full sortable/searchable transaction table.
- **Monthly** — a month picker with a category pivot table, a spending-distribution pie chart, and that month's transactions.
- **Subscriptions** — automatically detected recurring charges.
- **Customize** — add your own categories and keyword rules.

All charts and totals update live as you reclassify transactions — nothing requires regenerating the report.

## Project files

| File | Purpose |
|---|---|
| `xpence_analyzer.py` | Core report generator. Reads your transaction file and produces the HTML report. |
| `xpence_gui.py` | Desktop data-scrubbing tool with a friendly UI: normalizes raw bank exports (TD, CIBC, Wealthsimple, and other banks) into the format xPence expects, stripping PII in the process. |
| `merchant_categories.json` | Master reference list of category keywords and community-known merchants. Safe to share/download — contains no personal data. |
| `user_overrides.json` | Your personal corrections, custom categories, and subscription preferences. Created automatically on first run. Never share this file — keep it local. |

## Getting started

### 1. Install dependencies

**Python is required** to run xPence — both `xpence_analyzer.py` (report generator) and `xpence_gui.py` (data-scrubbing tool) are Python scripts. Nothing in this project runs without a local Python installation.

| Dependency | Required for | Notes |
|---|---|---|
| **Python 3.10+** | Everything | [Download Python](https://www.python.org/downloads/) if you don't already have it. Check your version with `python3 --version`. |
| `pandas` | Everything | Core data-handling library. |
| `openpyxl` | Reading/writing `.xlsx` files | Needed for modern Excel files. |
| `xlrd` | Reading legacy `.xls` files | Only needed if your bank exports the old `.xls` format. |
| `odfpy` | Reading `.ods` files | Only needed for OpenDocument spreadsheets. |
| Tkinter | `xpence_gui.py` only | Powers the desktop scrubbing-tool UI. Bundled with most Python installs; on some Linux distros install it separately (e.g. `sudo apt install python3-tk`). |

Install the Python package dependencies with:

```bash
pip install pandas openpyxl xlrd odfpy
```

### 2. Get your transaction data into the right shape

If your data already comes as a CSV/Excel file with clear column headers (Date, Description, Debit, Credit, etc.), you can skip straight to step 3 — xPence will auto-detect the columns.

If you're starting from a raw bank export (e.g. a TD or CIBC download), run the scrubbing tool first:

```bash
python3 xpence_gui.py
```

This opens a desktop app where you can select one or more raw export files. It will:
- Strip personally identifiable information (card numbers, account numbers).
- Normalize dates, descriptions, and amounts into a consistent format.
- Identify the account type (Credit, Chequing, Savings) where the source data allows it — otherwise it safely assumes **Credit**, since that's the most common source for xPence users.
- Write out a clean `.xlsx` file ready for the analyzer.

### 3. Generate your report

```bash
python3 xpence_analyzer.py --file your_transactions.xlsx --output my_report.html
```

Then open `my_report.html` in any browser.

### 4. (Optional) Scope the report to one account type

If your data has multiple account types mixed together, you can generate a report for just one:

```bash
python3 xpence_analyzer.py --file your_transactions.xlsx --account-type Chequing
```

The report also includes an **account filter dropdown** on the "Spending by Category" chart, so you can flip between "All accounts" and individual account types without regenerating anything.

## Command-line options

| Flag | Description |
|---|---|
| `--file`, `-f` | Path to your CSV/Excel file. If omitted, xPence looks for a file starting with `xPence` in the current directory. |
| `--output`, `-o` | Output HTML path (default: `xpence_report.html`). |
| `--csv`, `-c` | Also write a category-by-month summary CSV. |
| `--merchants`, `-m` | Path to `merchant_categories.json` (default: beside the script). |
| `--overrides`, `-u` | Path to your `user_overrides.json` (default: beside the script, created automatically). |
| `--account-type`, `-a` | Only include transactions on this account type (e.g. `Credit`, `Chequing`, `Savings`). Only takes effect if your data has an identifiable account-type column. |
| `--xpr`, `-x` | Path to a previously-saved `.xpr` report snapshot — carries forward any reclassifications you made in the browser. |

## How transactions get classified

Priority order, highest to lowest:

1. Your explicit overrides (set by reclassifying a transaction in the report)
2. Your personally-learned merchant mappings
3. The shared community merchant list (`merchant_categories.json`)
4. Keyword matching against category definitions
5. Your custom categories (defined in the Customize tab)
6. `Other / Uncategorised`

## How account type is determined

1. If your source data has a column that looks like "Account Type", "Card Type", etc., that value is used directly.
2. If not, but the data was produced by `xpence_gui.py`, the scrubber makes its best guess from the filename (e.g. a file with "chequing" in the name) — otherwise it defaults to **Credit**, since that's the most common statement type.
3. If no account-type information exists anywhere, every transaction in the report is assumed to be on a **Credit** account.

A report only remembers a specific account-type selection (via `--account-type`) if the source data genuinely contained account-type information — an assumed default is never saved as if it were a verified fact, since the same dates and merchants could plausibly belong to a different account.

## Handling of credits and payments

Some export formats use a single signed "amount" column instead of separate debit/credit columns. In that case, negative amounts are automatically treated as **credits/payments** (e.g. a card payment or refund) rather than being discarded — they show up correctly in the "Credits / Payments" total and the Review tab.

## Privacy

- `xpence_gui.py` strips card numbers, account numbers, and similar PII columns before writing the scrubbed file.
- `user_overrides.json` contains your personal spending categorizations — it is created locally and is never uploaded anywhere. Add it to your `.gitignore` if you keep this project in version control.
- `merchant_categories.json` contains only generic merchant-name → category mappings and is safe to share or contribute back to.


## Contributing

Found a merchant that's miscategorized, or want to add support for another bank's export format? Pull requests to `merchant_categories.json` (community data, no PII) or the bank parsers in `xpence_gui.py` are welcome.
