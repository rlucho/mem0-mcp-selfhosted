"""Mirror of modExportRead.StatementLineFacts, against a real FEBAN export.

A sample that stops at NO CLEARING or NO VENDOR PAYMENTS is the one an auditor
asks 'confirm the nature of the transaction' about, and the answer is already in
the statement: the note to payee, the business partner, the posting text. It was
never hidden, only buried -- the export is the whole month, so the row is one in
two thousand.

This checks the two things that can quietly go wrong:

  * the column lookup is EXACT, not 'contains'. 'Document Number' and
    'Subledger Doc.Number' sit next to each other, and a loose match takes
    whichever comes first -- which would hunt the wrong document all day and
    silently find nothing.
  * the document match ignores leading zeros. FEBAN writes '0000805039' in one
    list and '805039' in another.

Usage: test_statement_facts.py <feban-export.xlsx> [<document number> ...]
"""

import sys
from pathlib import Path

import openpyxl

WANTED = (
    "Business partner",
    "Customer",
    "Payment Notes",
    "Posting text",
    "External transaction",
    "Bank reference",
)


def digits_only(text):
    """DigitsOnly: significant digits, so 0000805039 == 805039."""
    kept = "".join(c for c in str(text) if c.isdigit())
    return kept.lstrip("0") or ("0" if kept else "")


def column_headed(rows, caption):
    """ColumnHeaded: exact match on row 1, never 'contains'."""
    for index, value in enumerate(rows[0], start=1):
        if str(value or "").strip().lower() == caption.lower():
            return index
    return 0


def statement_row(rows, document_number, amount=None):
    """StatementRow: by document number, else by amount."""
    wanted = digits_only(document_number)
    if wanted:
        col = column_headed(rows, "Document Number")
        if col:
            for r in range(1, len(rows)):
                if digits_only(rows[r][col - 1]) == wanted:
                    return r + 1
    if amount is None:
        return 0
    col = column_headed(rows, "Amount")
    if not col:
        return 0
    for r in range(1, len(rows)):
        try:
            here = abs(float(str(rows[r][col - 1]).replace(",", "")))
        except (TypeError, ValueError):
            continue
        if abs(here - abs(amount)) < 0.005:
            return r + 1
    return 0


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2

    path = Path(sys.argv[1])
    book = openpyxl.load_workbook(path, data_only=True, read_only=True)
    rows = [list(r) for r in book.worksheets[0].iter_rows(values_only=True)]
    print(f"{path.name}: {len(rows)} rows x {len(rows[0])} columns\n")

    print("column lookup (exact):")
    problems = []
    for caption in ("Document Number", "Subledger Doc.Number", *WANTED):
        col = column_headed(rows, caption)
        print(f"  {caption:<22} -> column {col or '(absent)'}")
    doc_col = column_headed(rows, "Document Number")
    sub_col = column_headed(rows, "Subledger Doc.Number")
    if doc_col and sub_col and doc_col == sub_col:
        problems.append("Document Number and Subledger Doc.Number resolved to one column")

    wanted_docs = sys.argv[2:]
    if not wanted_docs:
        # No documents named, so take a few real ones off the sheet itself and
        # prove they are found again -- including with leading zeros bolted on,
        # which is how the FI document reaches this code.
        wanted_docs = []
        for r in range(1, min(len(rows), 400)):
            value = rows[r][doc_col - 1] if doc_col else None
            if value and digits_only(value):
                wanted_docs.append(str(value))
            if len(wanted_docs) == 3:
                break

    print("\nrow lookup:")
    for document in wanted_docs:
        padded = "0000" + digits_only(document)
        hit = statement_row(rows, document)
        hit_padded = statement_row(rows, padded)
        ok = hit and hit == hit_padded
        print(f"  {document:<14} -> row {hit or '(not found)'}"
              f"   zero-padded '{padded}' -> row {hit_padded or '(not found)'}"
              f"   {'ok' if ok else 'MISMATCH'}")
        if not ok:
            problems.append(f"{document}: padded and unpadded disagree")
        if hit:
            facts = []
            for caption in WANTED:
                col = column_headed(rows, caption)
                if col:
                    value = str(rows[hit - 1][col - 1] or "").strip()
                    if value:
                        facts.append(f"      {caption}: {value[:70]}")
            if facts:
                print("\n".join(facts))
            else:
                print("      (every explaining field blank on this row)")

    print()
    if problems:
        for problem in problems:
            print(f"FAIL {problem}")
        return 1
    print("column lookup exact, document match ignores leading zeros")
    return 0


if __name__ == "__main__":
    sys.exit(main())
