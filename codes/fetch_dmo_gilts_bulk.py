"""
fetch_dmo_gilts_bulk.py
=======================
Downloads the DMO "Gilts in Issue" (D1A) Excel report for every UK working day
from START_DATE to today, then consolidates all data into a single CSV.

Usage:
    python fetch_dmo_gilts_bulk.py [--start YYYY-MM-DD] [--output gilts.csv]
                                   [--xls-dir ./xls_cache] [--delay 1.5]

Arguments:
    --start     First date to fetch (default: 2019-01-02)
    --output    Path for the consolidated CSV (default: ../output/dmo_gilts_consolidated.csv)
    --xls-dir   Directory to store/cache downloaded XLS files (default: ../inputs)
    --delay     Seconds to wait between requests (default: 1.5)

Resumable: already-downloaded XLS files are skipped on re-runs.
Requires:  xlrd (pip install xlrd)

File structure parsed
---------------------
Row 0  : "Data Date: DD-Mon-YYYY, ..."  ← report date
Rows 1–7: metadata / blanks
Row 8  : Conventional Gilts column headers (first cell = "Conventional Gilts")
Rows 9+: Conventional gilt rows; sub-bucket label rows (Ultra-Short/Short/Medium/Long)
         have blank ISIN → skipped as data, stored as current bucket context
Blank separator rows separate the three sections:
  - Conventional Gilts
  - Index-linked Gilts (3-month Indexation Lag)
  - Index-linked Gilts (8-month Indexation Lag)
Each index-linked section header row begins with "Index-linked Gilts"
Footer notes rows begin with "Note:" or "Page"
"""

import argparse
import csv
import io
import os
import subprocess
import time
from datetime import date, datetime, timedelta

import xlrd


# ── Constants ────────────────────────────────────────────────────────────────

DEFAULT_START   = date(2019, 1, 2)
DEFAULT_OUTPUT  = "../output/dmo_gilts_consolidated.csv"
DEFAULT_XLS_DIR = "../inputs"
DEFAULT_DELAY   = 1.5   # seconds between HTTP requests

DMO_URL_TEMPLATE = (
    "https://www.dmo.gov.uk/umbraco/surface/DataExport/GetDataExport"
    "?reportCode=D1A&exportFormatValue=xls"
    "&parameters=%26COBDate%3D{dd}%2F{mm}%2F{yyyy}"
)

OUTPUT_COLUMNS = [
    "report_date",
    "gilt_type",           # Conventional | Index-linked (3-month) | Index-linked (8-month)
    "maturity_bucket",     # Ultra-Short | Short | Medium | Long  (Conventional only)
    "name",
    "isin",
    "redemption_date",
    "first_issue_date",
    "dividend_dates",
    "ex_dividend_date",
    "amount_in_issue_mn",  # £ million nominal
    "base_rpi",            # Index-linked only
    "amount_incl_uplift_mn",  # Index-linked only
]

# Sub-bucket labels that appear as row separators in the Conventional section
BUCKET_LABELS = {"ultra-short", "short", "medium", "long"}

# ── Date helpers ─────────────────────────────────────────────────────────────

def last_uk_working_day(d: date) -> date:
    """Roll back to the most recent Mon–Fri."""
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    return d


def working_days_between(start: date, end: date):
    """Yield every Mon–Fri from start to end (inclusive)."""
    current = start
    while current <= end:
        if current.weekday() < 5:
            yield current
        current += timedelta(days=1)


# ── Download ─────────────────────────────────────────────────────────────────

def download_xls(report_date: date, xls_dir: str) -> str | None:
    """
    Download the XLS for report_date into xls_dir using curl.
    Returns the local file path, or None on failure.
    Skips download if the file already exists (resumable).
    """
    filename = f"dmo_gilts_{report_date.strftime('%Y%m%d')}.xls"
    filepath = os.path.join(xls_dir, filename)

    if os.path.exists(filepath) and os.path.getsize(filepath) > 0:
        return filepath   # already cached

    url = DMO_URL_TEMPLATE.format(
        dd=report_date.strftime("%d"),
        mm=report_date.strftime("%m"),
        yyyy=report_date.strftime("%Y"),
    )

    try:
        result = subprocess.run(
            [
                "curl", "--silent", "--show-error",
                "--location",
                "--max-time", "30",
                "--header", "Referer: https://www.dmo.gov.uk/",
                "--output", filepath,
                url,
            ],
            capture_output=True,
            text=True,
            timeout=60,
        )
    except (subprocess.TimeoutExpired, FileNotFoundError) as exc:
        print(f"  [WARN] curl failed: {exc}")
        return None

    if result.returncode != 0:
        print(f"  [WARN] curl error: {result.stderr.strip()}")
        return None

    if not os.path.exists(filepath) or os.path.getsize(filepath) < 100:
        print(f"  [WARN] Downloaded file too small — skipping")
        if os.path.exists(filepath):
            os.remove(filepath)
        return None

    return filepath

# ── XLS → rows conversion ────────────────────────────────────────────────────

def xls_to_csv_text(xls_path: str) -> str | None:
    """
    Read an .xls file using xlrd and return its first sheet as CSV text.
    Handles date cells natively — no LibreOffice dependency required.
    """
    try:
        wb = xlrd.open_workbook(xls_path)
    except Exception as exc:
        print(f"  [ERROR] xlrd failed to open {xls_path}: {exc}")
        return None

    ws = wb.sheet_by_index(0)
    output = io.StringIO()
    writer = csv.writer(output)

    for row_idx in range(ws.nrows):
        row = []
        for col_idx in range(ws.ncols):
            cell = ws.cell(row_idx, col_idx)
            if cell.ctype == xlrd.XL_CELL_DATE:
                try:
                    dt = xlrd.xldate_as_datetime(cell.value, wb.datemode)
                    row.append(dt.strftime("%d-%b-%Y"))
                except Exception:
                    row.append(str(cell.value))
            elif cell.ctype == xlrd.XL_CELL_NUMBER:
                val = cell.value
                # Emit integers without a trailing .0
                row.append(str(int(val)) if val == int(val) else str(val))
            else:
                row.append(str(cell.value))
        writer.writerow(row)

    return output.getvalue()


# ── Parsing ───────────────────────────────────────────────────────────────────

def _clean(val: str) -> str:
    """Strip whitespace and newlines from a cell value."""
    return val.strip().replace("\n", " ").replace("\r", "")


def _is_blank_row(row: list[str]) -> bool:
    return all(c.strip() == "" for c in row)


def _is_section_header(row: list[str]) -> bool:
    """True if this row is a section header (Conventional Gilts / Index-linked Gilts ...)"""
    first = _clean(row[0]).lower() if row else ""
    return first.startswith("conventional gilts") or first.startswith("index-linked gilts")


def _section_type_from_header(row: list[str]) -> str:
    first = _clean(row[0]).lower()
    if first.startswith("conventional"):
        return "Conventional"
    if "8-month" in first:
        return "Index-linked (8-month)"
    return "Index-linked (3-month)"


def parse_csv_text(csv_text: str, report_date: date) -> list[dict]:
    """
    Parse the CSV text into a list of row dicts.
    """
    reader = csv.reader(io.StringIO(csv_text))
    rows = list(reader)

    records = []
    current_section = None     # Conventional | Index-linked (3-month) | Index-linked (8-month)
    current_bucket  = ""       # Ultra-Short | Short | Medium | Long
    in_data         = False    # True once we've passed the first section header

    for row in rows:
        # Pad to at least 9 columns
        row = row + [""] * max(0, 9 - len(row))

        first = _clean(row[0]).lower()

        # Skip blank rows
        if _is_blank_row(row):
            continue

        # Skip footer rows
        if first.startswith("note:") or first.startswith("page"):
            continue

        # Section header row → update section context, reset bucket
        if _is_section_header(row):
            current_section = _section_type_from_header(row)
            current_bucket  = ""
            in_data         = True
            continue

        # Skip rows before we've hit any section header (title, total, etc.)
        if not in_data:
            continue

        # Skip column-header rows (second cell is "ISIN Code")
        if _clean(row[1]).lower() == "isin code":
            continue

        # Sub-bucket label row: col 0 is a bucket name, col 1 is blank
        if _clean(row[1]) == "" and _clean(row[0]).lower() in BUCKET_LABELS:
            current_bucket = _clean(row[0])
            continue

        # Data row: must have a non-empty ISIN (col 1)
        isin = _clean(row[1])
        if not isin or not isin.startswith("GB"):
            continue

        record = {
            "report_date":          report_date.isoformat(),
            "gilt_type":            current_section or "",
            "maturity_bucket":      current_bucket if current_section == "Conventional" else "",
            "name":                 _clean(row[0]),
            "isin":                 isin,
            "redemption_date":      _clean(row[2]),
            "first_issue_date":     _clean(row[3]),
            "dividend_dates":       _clean(row[4]),
            "ex_dividend_date":     _clean(row[5]),
            "amount_in_issue_mn":   _clean(row[6]),
            "base_rpi":             _clean(row[7]) if current_section != "Conventional" else "",
            "amount_incl_uplift_mn":_clean(row[8]) if current_section != "Conventional" else "",
        }
        records.append(record)

    return records


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Bulk-download DMO Gilts in Issue (D1A) and consolidate into CSV."
    )
    parser.add_argument("--start",  default=DEFAULT_START.isoformat(),
                        help=f"Start date YYYY-MM-DD (default: {DEFAULT_START})")
    parser.add_argument("--output", default=DEFAULT_OUTPUT,
                        help=f"Output CSV path (default: {DEFAULT_OUTPUT})")
    parser.add_argument("--xls-dir", default=DEFAULT_XLS_DIR,
                        help=f"Directory for cached XLS files (default: {DEFAULT_XLS_DIR})")
    parser.add_argument("--delay",  type=float, default=DEFAULT_DELAY,
                        help=f"Seconds between HTTP requests (default: {DEFAULT_DELAY})")
    args = parser.parse_args()

    start_date = datetime.strptime(args.start, "%Y-%m-%d").date()
    end_date   = last_uk_working_day(date.today())
    xls_dir    = args.xls_dir
    output     = args.output

    os.makedirs(xls_dir, exist_ok=True)

    all_dates = list(working_days_between(start_date, end_date))
    total     = len(all_dates)
    print(f"Date range : {start_date} → {end_date}  ({total} working days)")
    print(f"XLS cache  : {xls_dir}")
    print(f"Output     : {output}")
    print(f"Delay      : {args.delay}s between requests\n")

    # Open output CSV (append mode so we can resume partial runs)
    # Determine if file already has data so we can skip the header
    file_exists   = os.path.exists(output) and os.path.getsize(output) > 0
    written_dates = set()

    if file_exists:
        print(f"Output file exists — loading already-processed dates for dedup...")
        with open(output, "r", newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                written_dates.add(row.get("report_date", ""))
        print(f"  {len(written_dates)} dates already in output file\n")

    downloaded = 0
    parsed     = 0
    skipped    = 0
    errors     = 0

    with open(output, "a", newline="", encoding="utf-8") as out_file:
        writer = csv.DictWriter(out_file, fieldnames=OUTPUT_COLUMNS)
        if not file_exists:
            writer.writeheader()

        for i, d in enumerate(all_dates, 1):
            date_iso = d.isoformat()
            progress = f"[{i:4d}/{total}]"

            # Skip dates already in the output CSV
            if date_iso in written_dates:
                skipped += 1
                if i % 50 == 0:
                    print(f"{progress} ... (skipping already-processed dates)")
                continue

            print(f"{progress} {date_iso}", end="  ", flush=True)

            # 1. Download XLS
            xls_path = download_xls(d, xls_dir)
            if xls_path is None:
                print("→ download failed")
                errors += 1
                time.sleep(args.delay)
                continue
            downloaded += 1

            # 2. Convert to CSV text
            csv_text = xls_to_csv_text(xls_path)
            if csv_text is None:
                print("→ conversion failed")
                errors += 1
                continue

            # 3. Parse
            records = parse_csv_text(csv_text, d)
            if not records:
                print("→ no records parsed")
                errors += 1
                continue

            # 4. Write to output
            for rec in records:
                writer.writerow(rec)
            out_file.flush()
            parsed += 1
            print(f"→ {len(records)} rows")

            time.sleep(args.delay)

    print("\n" + "=" * 60)
    print(f"Done.")
    print(f"  Dates processed  : {parsed}")
    print(f"  Dates skipped    : {skipped}  (already in output)")
    print(f"  Errors/no-data   : {errors}")
    print(f"  Output file      : {output}")


if __name__ == "__main__":
    main()
