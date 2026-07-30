#!/usr/bin/env python3
"""Tests for the list-parsing logic that lives in vba/modListFile.bas.

The VBA cannot be executed here, so the three routines most likely to get a
sample wrong -- SAP amount parsing, cleared-items column detection, and
confirming-party name matching -- are mirrored in Python and tested against
synthetic SAP list exports in both the European and US number formats.

Keep this in step with the VBA by hand. If you change ParseSapAmount,
LargestRow or NamesMatch in modListFile.bas / modUtil.bas, change the mirror
below and re-run:

    python3 scripts/test_list_parser.py
"""
import re

def parse_sap_amount(text):
    c = text.strip()
    if not c: return 0.0
    neg = False
    if c.endswith("-"): neg, c = True, c[:-1].strip()
    elif c.startswith("-"): neg, c = True, c[1:].strip()
    c = c.replace(" ", "").replace("\xa0", "")
    ld, lc = c.rfind("."), c.rfind(",")
    if lc > ld: c = c.replace(".", "").replace(",", ".")
    else: c = c.replace(",", "")
    try: v = float(c)
    except ValueError: return 0.0
    return -v if neg else v

def normalise(t): return re.sub(r'[^A-Z0-9]', '', t.upper())

AMOUNT_CAPS = "Amount in local currency|Amount in LC|Amount|LC amount|Amnt in loc.cur.|DMBTR|WRBTR".split("|")
SUPPLIER_CAPS = "Name|Name 1|Name of vendor|Vendor name|Supplier|Account name|NAME1|Text".split("|")
DOC_CAPS = "Document Number|DocumentNo|Doc. Number|Document no.|BELNR|Invoice reference".split("|")

def is_data_line(t):
    stripped = re.sub(r'[-|\s_]', '', t)
    return bool(stripped) and len(t.strip()) >= 3

def match_column(fields, caps):
    for want in (normalise(c) for c in caps if c.strip()):
        for i, f in enumerate(fields):
            if normalise(f) and normalise(f) == want: return i
    return -1

def largest_row(text):
    lines = text.replace("\r\n","\n").replace("\r","\n").split("\n")
    pipes = sum(l.count("|") for l in lines[:60]); tabs = sum(l.count("\t") for l in lines[:60])
    delim = "|" if pipes >= tabs and pipes > 0 else "\t"
    hdr, ac, sc, dc = -1, -1, -1, -1
    for i, l in enumerate(lines[:80]):
        if delim in l:
            f = l.split(delim)
            a, s = match_column(f, AMOUNT_CAPS), match_column(f, SUPPLIER_CAPS)
            if a >= 0 and s >= 0:
                hdr, ac, sc, dc = i, a, s, match_column(f, DOC_CAPS); break
    if hdr < 0: return None
    best, out, considered = 0.0, None, 0
    for l in lines[hdr+1:]:
        if not is_data_line(l): continue
        f = l.split(delim)
        if len(f) <= ac: continue
        v = abs(parse_sap_amount(f[ac]))
        if v > best:
            best = v
            out = dict(amount=v,
                       supplier=re.sub(r'\s+',' ',f[sc].strip()) if sc < len(f) else "",
                       doc=re.sub(r'\s+',' ',f[dc].strip()) if 0 <= dc < len(f) else "")
        considered += 1
    if out: out["rows"] = considered
    return out

def names_match(a, b):
    if not b.strip(): return False
    na = normalise(a)
    return all(normalise(t) in na for t in re.sub(r'\s+',' ',b.upper()).split() if t)

# --- Test 1: pipe-delimited SAP "unconverted" export, European numbers ---
t1 = """--------------------------------------------------------------------------
| Document Number | Name                    | Amnt in loc.cur. | Clearing  |
--------------------------------------------------------------------------
| 1900012345      | MONDI BIRMINGHAM LIMITED|     1.234.567,89 | 2000099   |
| 1900012346      | DS Smith Paper Limited  |     5.988.033,60 | 2000099   |
| 1900012347      | Smith Partnership       |       849.771,13 | 2000099   |
--------------------------------------------------------------------------
"""
# --- Test 2: tab-delimited, US numbers, trailing-minus credit ---
t2 = "Document no.\tVendor name\tAmount\n" \
     "5100001\tSANTANDER SCF\t4,483,676.08-\n" \
     "5100002\tDS SMITH LTD\t2,450,000.00\n"
# --- Test 3: no recognisable header ---
t3 = "| 111 | ACME LTD | 100,00 |\n| 222 | BETA GMBH | 250,00 |\n"

for name, txt, expect_supplier, expect_amount in [
    ("pipe / European",  t1, "DS Smith Paper Limited", 5988033.60),
    ("tab / US, credit", t2, "SANTANDER SCF",          4483676.08),
]:
    r = largest_row(txt)
    ok = r and r["supplier"] == expect_supplier and abs(r["amount"] - expect_amount) < 0.005
    print(f"{'PASS' if ok else 'FAIL'}  {name:18} -> {r}")

print(f"{'PASS' if largest_row(t3) is None else 'FAIL'}  no header          -> falls through to the guess path (returns None here)")

print("\nConfirming-party matching:")
for supplier, expected in [("SANTANDER SCF", True), ("SCF Santander", True), ("Santander  SCF", True),
                           ("SANTANDER SCF LONDON", True), ("DS SMITH LTD", False),
                           ("DS Smith Paper Limited", False), ("SANTANDER UK PLC", False)]:
    got = names_match(supplier, "SANTANDER SCF")
    print(f"  {'PASS' if got == expected else 'FAIL'}  {supplier!r:26} -> {got}")

print("\nAmount parsing:")
for text, expected in [("8.072.447,42", 8072447.42), ("8,072,447.42", 8072447.42),
                       ("5.988.033,60-", -5988033.60), ("2450000", 2450000.0),
                       ("  1 234,56  ", 1234.56), ("", 0.0), ("abc", 0.0), ("-2.450.000,00", -2450000.0)]:
    got = parse_sap_amount(text)
    print(f"  {'PASS' if abs(got-expected) < 0.005 else 'FAIL'}  {text!r:18} -> {got}")
