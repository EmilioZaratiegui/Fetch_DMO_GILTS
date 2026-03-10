# Fetch DMO Gilts

Bulk-downloads the UK Debt Management Office (DMO) "Gilts in Issue" report (D1A) for every working day from a given start date to today, and consolidates all daily snapshots into a single CSV file.

---

## Prerequisites

**Python 3.10 or later**
Download from [python.org](https://www.python.org/downloads/) if not already installed.

**xlrd** (Python package)
```bash
pip install xlrd
```

**curl**
Used to download files from the DMO website.
- macOS and Linux: built-in
- Windows 10/11: built-in (verify by running `curl --version` in a terminal)
- Older Windows: download from [curl.se](https://curl.se/windows/)

---

## Setup

Clone or download this repository, then navigate to the `codes` folder:

```bash
cd path/to/Fetch_DMO_GILTS/codes
```

---

## Usage

Run with defaults (fetches from 2019-01-02 to today):

```bash
python fetch_dmo_gilts_bulk.py
```

The script will save:
- Downloaded `.xls` files → `inputs/`
- Consolidated CSV → `output/dmo_gilts_consolidated.csv`

### Options

| Argument | Default | Description |
|---|---|---|
| `--start` | `2019-01-02` | First date to fetch (YYYY-MM-DD) |
| `--output` | `../output/dmo_gilts_consolidated.csv` | Path for the consolidated CSV |
| `--xls-dir` | `../inputs` | Directory to cache downloaded XLS files |
| `--delay` | `1.5` | Seconds to wait between requests |

Example with custom arguments:

```bash
python fetch_dmo_gilts_bulk.py --start 2015-01-02 --delay 2.0
```

---

## Output

The consolidated CSV contains one row per gilt per report date, with the following columns:

| Column | Description |
|---|---|
| `report_date` | ISO date of the DMO report (YYYY-MM-DD) |
| `gilt_type` | Conventional / Index-linked (3-month) / Index-linked (8-month) |
| `maturity_bucket` | Ultra-Short / Short / Medium / Long (Conventional only) |
| `name` | Gilt name |
| `isin` | ISIN code (always starts with GB) |
| `redemption_date` | Maturity date |
| `first_issue_date` | Date of first issuance |
| `dividend_dates` | Coupon payment dates |
| `ex_dividend_date` | Ex-dividend date |
| `amount_in_issue_mn` | Nominal amount in issue (£ million) |
| `base_rpi` | Base RPI (Index-linked only) |
| `amount_incl_uplift_mn` | Amount including RPI uplift (Index-linked only) |

---

## Pre-built dataset

A consolidated CSV covering all working days from 2019-01-02 is included in `output/dmo_gilts_consolidated.csv`. If you just want the data, you can use it directly without running the script.

To bring it up to date, simply run the script — it will detect the existing file and only fetch dates that are missing.

---

## Resuming interrupted runs

The script is designed to be safely interrupted and restarted. Already-downloaded XLS files are cached in `inputs/xls_cache/` and skipped on re-runs. Dates already present in the output CSV are also skipped automatically.

---

## Folder structure

```
Fetch_DMO_GILTS/
├── README.md
├── CLAUDE.md          ← project notes for AI context
├── codes/
│   └── fetch_dmo_gilts_bulk.py
├── inputs/            ← downloaded XLS files
├── output/            ← consolidated CSV output
└── temp/              ← scratch files (safe to delete)
```
