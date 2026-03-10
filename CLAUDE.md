# Fetch DMO Gilts — Project Notes

## Purpose

This project bulk-downloads the UK Debt Management Office (DMO) "Gilts in Issue" report (report code **D1A**) for every UK working day from a given start date to today, then consolidates all daily snapshots into a single CSV file.

## Folder Structure

```
Fetch_DMO_GILTS/
├── CLAUDE.md          ← this file
├── codes/             ← Python scripts
│   └── fetch_dmo_gilts_bulk.py
├── inputs/            ← downloaded XLS files (gitignored)
├── output/            ← consolidated CSV output (dmo_gilts_consolidated.csv)
└── temp/              ← intermediate/scratch files (safe to delete)
```

## Main Script

**`codes/fetch_dmo_gilts_bulk.py`**

### What it does
1. Iterates over every UK working day (Mon–Fri) between `--start` and today.
2. For each date, downloads the D1A report as an `.xls` file from the DMO website using `curl`.
3. Parses the `.xls` directly using `xlrd` (no LibreOffice required).
4. Extracts gilt records across three sections: Conventional, Index-linked (3-month), and Index-linked (8-month).
5. Appends all records to a consolidated CSV.

The run is **resumable**: already-downloaded XLS files are cached and skipped, and dates already present in the output CSV are skipped too.

### Usage

```bash
pip install xlrd

python codes/fetch_dmo_gilts_bulk.py \
    --start 2019-01-02 \
    --output output/dmo_gilts_consolidated.csv \
    --xls-dir inputs \
    --delay 1.5
```

### Arguments

| Argument    | Default                        | Description                              |
|-------------|--------------------------------|------------------------------------------|
| `--start`   | `2019-01-02`                   | First date to fetch (YYYY-MM-DD)         |
| `--output`  | `dmo_gilts_consolidated.csv`   | Path for the consolidated CSV            |
| `--xls-dir` | `./xls_cache`                  | Directory to cache downloaded XLS files  |
| `--delay`   | `1.5`                          | Seconds to wait between HTTP requests    |

### Output columns

| Column                 | Description                                              |
|------------------------|----------------------------------------------------------|
| `report_date`          | ISO date of the DMO report (YYYY-MM-DD)                  |
| `gilt_type`            | Conventional / Index-linked (3-month) / Index-linked (8-month) |
| `maturity_bucket`      | Ultra-Short / Short / Medium / Long (Conventional only)  |
| `name`                 | Gilt name                                                |
| `isin`                 | ISIN code (always starts with GB)                        |
| `redemption_date`      | Maturity date                                            |
| `first_issue_date`     | Date of first issuance                                   |
| `dividend_dates`       | Coupon payment dates                                     |
| `ex_dividend_date`     | Ex-dividend date                                         |
| `amount_in_issue_mn`   | Nominal amount in issue (£ million)                      |
| `base_rpi`             | Base RPI (Index-linked only)                             |
| `amount_incl_uplift_mn`| Amount including RPI uplift (Index-linked only)          |

## Dependencies

- `xlrd` — reads `.xls` binary files natively in Python. Install with `pip install xlrd`.
- `curl` — used for HTTP downloads (available on macOS/Linux by default).

## Notes

- The DMO source URL is: `https://www.dmo.gov.uk/umbraco/surface/DataExport/GetDataExport?reportCode=D1A&exportFormatValue=xls&parameters=%26COBDate%3D{DD}%2F{MM}%2F{YYYY}`
- The D1A report is only published on UK working days. Weekends and UK bank holidays will return no data or a very small file (automatically skipped).
- The default start date is `2019-01-02` but the DMO has data going back further if a longer history is needed.
- XLS files are cached in `temp/xls_cache` so the script can be interrupted and resumed without re-downloading.
- The output CSV is opened in append mode; the script deduplicates by date on resume.
