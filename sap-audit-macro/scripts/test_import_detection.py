"""Replicate modImport's header/column detection and run it over the real
request files.

The VBA cannot be executed here, so this mirrors it line for line. When the
two disagree, this file is wrong -- but a layout the detector cannot read is
found here rather than on someone's laptop halfway through an audit.
"""
import sys
import glob
import os

import openpyxl

SKIP_SHEET_PREFIX = "DS_INTERNAL_"

CAPTIONS_DATE = "date|payment date|posting date|value date|document date"
CAPTIONS_AMOUNT = "amount|amount in local currency|amount in lc|payment amount|value"
CAPTIONS_PARTY = (
    "name of party|name of the party|name 1|party|supplier|vendor|beneficiary|payee"
)
CAPTIONS_REFERENCE = (
    "transaction description|payment reference|description|narrative|"
    "customer reference|bank reference|reference"
)
CAPTIONS_COMMENT = "ey comments|ey comment|comments|comment|query|request"


def squeeze(text):
    """modUtil.Squeeze: collapse runs of whitespace, trim."""
    return " ".join(str(text).split())


def cell(rows, r, c):
    """1-based, out of range -> ''."""
    if r < 1 or r > len(rows):
        return ""
    row = rows[r - 1]
    if c < 1 or c > len(row):
        return ""
    v = row[c - 1]
    return "" if v is None else str(v).strip()


def match_column(rows, row, synonyms, ncols, search_above=False):
    for wanted in synonyms.split("|"):
        for col in range(1, min(40, ncols) + 1):
            if squeeze(cell(rows, row, col)).lower() == wanted:
                return col
            if search_above and squeeze(cell(rows, row - 1, col)).lower() == wanted:
                return col
    return 0


def find_columns(rows, ncols):
    for row in range(1, min(20, len(rows)) + 1):
        date_col = match_column(rows, row, CAPTIONS_DATE, ncols)
        amount_col = match_column(rows, row, CAPTIONS_AMOUNT, ncols)
        if date_col and amount_col:
            return {
                "header": row,
                "date": date_col,
                "amount": amount_col,
                "party": match_column(rows, row, CAPTIONS_PARTY, ncols, True),
                "reference": match_column(rows, row, CAPTIONS_REFERENCE, ncols, True),
                "comment": match_column(rows, row, CAPTIONS_COMMENT, ncols, True),
            }
    return None


def parse_day_first(text):
    """modImport.ParseDayFirst: dd/mm/yyyy, dd.mm.yyyy, dd-mm-yyyy."""
    import datetime

    cleaned = str(text).strip().replace(".", "/").replace("-", "/")
    if len(cleaned) < 6:
        return None
    parts = cleaned.split("/")
    if len(parts) != 3:
        return None
    try:
        d, m, y = (int(float(p)) for p in parts)
    except ValueError:
        return None
    if y < 100:
        y += 2000
    if not (1 <= d <= 31 and 1 <= m <= 12 and 1990 <= y <= 2100):
        return None
    try:
        return datetime.date(y, m, d)
    except ValueError:
        return None


def row_date(rows, row, col):
    """modImport.RowDate: a real date cell, else a day-first text date."""
    import datetime

    raw = rows[row - 1][col - 1] if col and row <= len(rows) else None
    if isinstance(raw, (datetime.datetime, datetime.date)):
        return raw.date() if isinstance(raw, datetime.datetime) else raw
    if raw is None:
        return None
    return parse_day_first(raw)


def row_amount(rows, row, col):
    """modImport.RowAmount: a number, else parsed from text."""
    raw = rows[row - 1][col - 1] if col and row <= len(rows) else None
    if isinstance(raw, (int, float)):
        return float(raw)
    if raw is None:
        return 0.0
    txt = str(raw).strip().replace(" ", "")
    negative = txt.endswith("-") or txt.startswith("-")
    txt = txt.strip("-")
    if txt.rfind(",") > txt.rfind("."):
        txt = txt.replace(".", "").replace(",", ".")
    else:
        txt = txt.replace(",", "")
    try:
        v = float(txt)
    except ValueError:
        return 0.0
    return -v if negative else v


def is_sample_row(rows, row, m):
    if row_date(rows, row, m["date"]) is None:
        return False
    return abs(row_amount(rows, row, m["amount"])) > 0.005


def letter(col):
    if col <= 0:
        return "(none)"
    s = ""
    while col:
        col, rem = divmod(col - 1, 26)
        s = chr(65 + rem) + s
    return s


def main():
    args = sys.argv[1:]
    if not args:
        print("usage: test_import_detection.py <xlsx> [<xlsx> ...]")
        return 1

    paths = []
    for a in args:
        paths.extend(sorted(glob.glob(a)) or [a])

    total_files = total_rows = 0
    problems = []

    for path in paths:
        wb = openpyxl.load_workbook(path, data_only=True)
        print("=" * 100)
        print(os.path.basename(path))
        total_files += 1
        file_rows = 0

        for ws in wb.worksheets:
            if ws.title.upper().startswith(SKIP_SHEET_PREFIX):
                continue
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) < 2:
                continue
            ncols = max((len(r) for r in rows), default=0)

            m = find_columns(rows, ncols)
            if not m:
                print(f"  '{ws.title}': NO HEADER FOUND")
                problems.append(f"{os.path.basename(path)} :: {ws.title}: no header")
                continue

            n = sum(1 for r in range(m["header"] + 1, len(rows) + 1)
                    if is_sample_row(rows, r, m))
            file_rows += n
            print(f"  '{ws.title}': header row {m['header']}, {n} samples  "
                  f"date={letter(m['date'])} amount={letter(m['amount'])} "
                  f"party={letter(m['party'])} ref={letter(m['reference'])} "
                  f"comment={letter(m['comment'])}")

            shown = 0
            for r in range(m["header"] + 1, len(rows) + 1):
                if shown >= 2:
                    break
                if is_sample_row(rows, r, m):
                    shown += 1
                    d = row_date(rows, r, m["date"])
                    amt = abs(row_amount(rows, r, m["amount"]))
                    party = cell(rows, r, m["party"])
                    print(f"       {d.strftime('%d/%m/%Y')}  {amt:>14,.2f}  {party[:30]}")

            if n == 0:
                problems.append(f"{os.path.basename(path)} :: {ws.title}: 0 rows")
            if m["party"] == 0:
                problems.append(f"{os.path.basename(path)} :: {ws.title}: no party column")

        total_rows += file_rows
        print(f"  -> {file_rows} samples")

    print("=" * 100)
    print(f"{total_files} file(s), {total_rows} samples")
    if problems:
        print("\nPROBLEMS:")
        for p in problems:
            print("  " + p)
        return 1
    print("\nevery sheet resolved a header, a party column and at least one row")
    return 0


if __name__ == "__main__":
    sys.exit(main())
