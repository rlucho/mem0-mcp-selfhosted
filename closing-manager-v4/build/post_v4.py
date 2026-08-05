"""V4-CIO post-pass applied to the generated Closing / Printing module text.

  Printing : ten unbounded Shell.Application "wait for the PDF" loops -> one
             bounded, responsive helper.  This is the loop that froze Excel and
             ended in "-2147417848 (80010108) Method 'NameSpace' ... failed".
  Closing  : locale-independent amount conversion (Run-time error 13),
             a status-bar breadcrumb at each stage, and one plain-language
             failure handler for the whole close.
"""
import re

def _ind(line):
    return line[:len(line) - len(line.lstrip())]

# --------------------------------------------------------------- Printing.bas
LABELS = {
    212:  'ZGLRME (errors only)',
    315:  'ZGLRME (full report)',
    401:  'report group EIS4',
    506:  'report group GIS4',
    601:  'ZGE132 (before posting)',
    660:  'ZGE132 (after posting)',
    1046: 'ZGR215 (document list)',
    1105: 'ZGR215 (documents)',
    1182: 'report group GTB1',
    1311: 'ZGE1174',
}

def printing(text):
    L = text.split('\n')
    sites = [i for i, l in enumerate(L) if 'Do Until File <> ""' in l]
    assert len(sites) == 10, 'expected 10 print-wait loops, found %d' % len(sites)
    assert sorted(LABELS) == sorted(i + 1 for i in sites), sorted(i + 1 for i in sites)
    for st in reversed(sites):
        ind = _ind(L[st])
        end = next(j for j in range(st + 1, len(L))
                   if L[j].strip() == 'Loop' and _ind(L[j]) == ind)
        L[st:end + 1] = [
            ind + "'V4-CIO FIX: bounded, responsive wait for the printed PDF. The old loop",
            ind + "'rebuilt a Shell.Application on every pass with no pause, so an empty",
            ind + "'temp folder became a tight spin: Excel showed \"Not Responding\" and the",
            ind + "'shell object eventually dropped out with error 80010108.",
            ind + 'File = CM_WaitForPrint(FTemp, fso, "%s")' % LABELS[st + 1],
        ]
    return instrument_reads(amounts('\n'.join(L), 6, 2), 6, 6)

# ---------------------------------------------------------------- Closing.bas
STEPS = [
    (93,  'connecting to your SAP session'),
    (110, 'reading the company code and period'),
    (168, "clearing the previous run's data"),
    (196, 'creating the working folders'),
    (277, 'reading the settings from the config sheet'),
    (289, 'selecting the closing variant for this company code'),
    (303, 'checking the company code and period are open in SAP'),
    (336, 'asking which reports to print'),
    (365, 'checking ZGLRME and AA02 for differences'),
    (468, 'printing ZGLRME'),
    (476, 'printing report group EIS4'),
    (481, 'printing report group GIS4'),
    (486, 'printing and checking ZGE132'),
    (560, 'posting the ZGE132 entries in SAP'),
    (571, 'checking ZGE132 after posting'),
    (661, 'reading the document numbers from SM35'),
    (689, 'printing the posted documents in ZGR215'),
    (700, 'printing report group EIS4 again (after posting)'),
    (735, 'printing report group GIS4 again (after posting)'),
    (770, 'running ZGLGWUL'),
    (808, 'checking the result of ZGLGWUL'),
    (820, 'printing report group GTB1'),
    (861, 'printing ZGE1174'),
    (868, 'checking ZGLRME again'),
    (921, 'printing the final ZGLRME'),
    (926, 'merging all the PDFs into the final report pack'),
    (931, 'checking for any remaining errors'),
]
EXITS_PLAIN  = [104, 330, 381, 444, 552, 650, 728, 763, 800]
EXITS_INLINE = [21, 811]
AMOUNT_LINE  = 1288
RUN_END      = 988

def closing(text):
    L = text.split('\n')
    edits = []

    i = AMOUNT_LINE - 1
    assert 'Right(ArrZGL(i, 5), 1) = "-"' in L[i], L[i]
    ind = _ind(L[i])
    edits.append((i, [
        ind + "'V4-CIO FIX: SAP writes amounts in the SAP user's decimal notation, which",
        ind + "'need not match this PC's Windows regional settings; the old implicit",
        ind + "'conversion then raised \"Run-time error 13: Type mismatch\". CM_Amount reads",
        ind + "'both conventions and the trailing minus, and explains itself if it cannot.",
        ind + 'ArrZGL(i, 5) = CM_Amount(ArrZGL(i, 5), i, "reading the amounts from the ZGLRME extract")',
    ]))

    for ln, txt in STEPS:
        i = ln - 1
        edits.append((i, [_ind(L[i]) + 'CM_Note "%s"' % txt, L[i]]))

    for ln in EXITS_PLAIN:
        i = ln - 1
        assert L[i].strip() == 'Exit Sub', (ln, L[i])
        edits.append((i, [_ind(L[i]) + 'CM_Done', L[i]]))

    for ln in EXITS_INLINE:
        i = ln - 1
        m = re.match(r'^If (.*) Then Exit Sub$', L[i].strip())
        assert m, (ln, L[i])
        ind = _ind(L[i])
        edits.append((i, [ind + 'If ' + m.group(1) + ' Then',
                          ind + '    CM_Done',
                          ind + '    Exit Sub',
                          ind + 'End If']))

    i = RUN_END - 1
    assert L[i].strip() == 'End Sub', L[i]
    edits.append((i, [
        '',
        "'V4-CIO: one handler for the whole close. RunClosing turns error handling on",
        "'nowhere else, so anything that fails here - or in any routine it calls -",
        "'lands below and is explained in plain language instead of showing a bare",
        "'\"Run-time error 13\" dialog with no idea what the macro was doing.",
        'CM_Done',
        'Exit Sub',
        '',
        'CM_Fail:',
        '    CM_Explain Err.Number, Err.Description',
        '    CM_Done',
        '',
        L[i],
    ]))

    for i, repl in sorted(edits, key=lambda e: -e[0]):
        L[i:i + 1] = repl

    i = L.index('Sub RunClosing()')
    L[i + 1:i + 1] = [
        '',
        "'V4-CIO: show progress on the status bar so a slow run cannot be mistaken",
        "'for a frozen one, and explain any failure in plain language.",
        'On Error GoTo CM_Fail',
        'CM_Begin %d' % len(STEPS),
    ]
    text = amounts('\n'.join(L), 0, 1)
    assert text.count(ZGLRME_OLD) == 1
    text = text.replace(ZGLRME_OLD, ZGLRME_NEW)
    return review_fixes_closing(instrument_reads(text, 3, 3))


# ------------------------------------------------------- amounts (both modules)
# Every place a SAP text export is turned into a number was locale-dependent:
# Round("1.234,56", 2) and CDbl("1.234,56") both go through VBA's implicit
# conversion, which reads the PC's Windows regional format. When the SAP user's
# decimal notation differs, the whole column raises "Run-time error 13: Type
# mismatch". CM_Amount reads either convention plus the trailing minus, and
# stops with an explained message if the value genuinely is not a number.
#
# Each If/Else pair collapses to one line because CM_Amount handles the sign
# itself - which is exactly equivalent, including a leading "-" that the old
# Else branch also passed straight through.

def _block(cond, neg, pos, tail):
    return ('            If Right(%s, 1) = "-" Then\n'
            '                %s\n'
            '            Else\n'
            '                %s\n'
            '            End If' % (cond, neg, pos)), tail

AMOUNT_BLOCKS = [
    # (exact block, replacement expression, plain-language "what it was doing")
    ('            If Right(arr(LBound(arr) + 1), 1) = "-" Then\n'
     '                AmL = -Round(Left(arr(LBound(arr) + 1), Len(arr(LBound(arr) + 1)) - 1), 2)\n'
     '            Else\n'
     '                AmL = Round(arr(LBound(arr) + 1), 2)\n'
     '            End If',
     'AmL = Round(CM_AmountReq(arr(LBound(arr) + 1), 0, _\n'
     '                                  "reading the local-currency total from the ZGE132 extract"), 2)'),

    ('            If Right(arr(UBound(arr)), 1) = "-" Then\n'
     '                Am = -Round(Left(arr(UBound(arr)), Len(arr(UBound(arr))) - 1), 2)\n'
     '            Else\n'
     '                Am = Round(arr(UBound(arr)), 2)\n'
     '            End If',
     'Am = Round(CM_AmountReq(arr(UBound(arr)), 0, _\n'
     '                                 "reading a profit-centre amount from the ZGE132 extract"), 2)'),

    ('            If Right(arr(UBound(arr) - 3), 1) = "-" Then\n'
     '                AmG = -Round(Left(arr(UBound(arr) - 3), Len(arr(UBound(arr) - 3)) - 1), 2)\n'
     '            Else\n'
     '                AmG = Round(arr(UBound(arr) - 3), 2)\n'
     '            End If',
     'AmG = Round(CM_AmountReq(arr(UBound(arr) - 3), 0, _\n'
     '                                  "reading the group-currency total from the ZGE132 extract"), 2)'),

    ('            If Right(arr(UBound(arr) - 2), 1) = "-" Then\n'
     '                Am = -Round(Left(arr(UBound(arr) - 2), Len(arr(UBound(arr) - 2)) - 1), 2)\n'
     '            Else\n'
     '                Am = Round(arr(UBound(arr) - 2), 2)\n'
     '            End If',
     'Am = Round(CM_AmountReq(arr(UBound(arr) - 2), 0, _\n'
     '                                 "reading a profit-centre amount from ZGE132 (group currency)"), 2)'),

    ('            If Right(arr(UBound(arr) - 1), 1) = "-" Then\n'
     '                Am = -CDbl(Left(arr(UBound(arr) - 1), Len(arr(UBound(arr) - 1)) - 1))\n'
     '            Else\n'
     '                Am = CDbl(arr(UBound(arr) - 1))\n'
     '            End If',
     'Am = CM_AmountReq(arr(UBound(arr) - 1), 0, _\n'
     '                           "reading the account 44400200 balance from the GTB1 report")'),

    ('            If Right(arr(UBound(arr) - 1), 1) = "-" Then\n'
     '                AmT = -CDbl(Left(arr(UBound(arr) - 1), Len(arr(UBound(arr) - 1)) - 1))\n'
     '            Else\n'
     '                AmT = CDbl(arr(UBound(arr) - 1))\n'
     '            End If',
     'AmT = CM_AmountReq(arr(UBound(arr) - 1), 0, _\n'
     '                            "reading the total from the GTB1 report")'),
]

AMOUNT_LINES = [
    ('Print_EIS4 = Round(Replace(Replace(arr(UBound(arr, 1) - 1), "(", "-"), ")", ""), 2)',
     'Print_EIS4 = Round(CM_AmountReq(Replace(Replace(arr(UBound(arr, 1) - 1), "(", "-"), ")", ""), 0, _\n'
     '                   "reading the total from the EIS4 report group"), 2)'),
    ('Print_GIS4 = Round(Replace(Replace(arr(UBound(arr, 1) - 1), "(", "-"), ")", ""), 2)',
     'Print_GIS4 = Round(CM_AmountReq(Replace(Replace(arr(UBound(arr, 1) - 1), "(", "-"), ")", ""), 0, _\n'
     '                   "reading the total from the GIS4 report group"), 2)'),
    ('        Am = Round(Replace(Replace(arr(UBound(arr, 1) - 1), "(", "-"), ")", ""), 2)',
     '        Am = Round(CM_AmountReq(Replace(Replace(arr(UBound(arr, 1) - 1), "(", "-"), ")", ""), 0, _\n'
     '                   "reading an amount from the AA02 report group"), 2)'),
]

NOTE = ("'V4-CIO FIX: locale-proof amount. Round()/CDbl() on a SAP string use this\n"
        "'PC's Windows regional format, so a mismatch with the SAP user's decimal\n"
        "'notation raised \"Run-time error 13: Type mismatch\" on the whole column.")

def amounts(text, expect_blocks, expect_lines):
    n_b = n_l = 0
    for old, new in AMOUNT_BLOCKS:
        c = text.count(old)
        if c == 0:
            continue
        assert c == 1, 'amount block matched %d times: %r' % (c, old[:60])
        ind = '            '
        repl = '\n'.join(ind + l for l in NOTE.split('\n')) + '\n' + ind + new
        text = text.replace(old, repl)
        n_b += 1
    for old, new in AMOUNT_LINES:
        c = text.count(old)
        if c == 0:
            continue
        assert c == 1, 'amount line matched %d times: %r' % (c, old[:60])
        ind = old[:len(old) - len(old.lstrip())]
        repl = '\n'.join(ind + l for l in NOTE.split('\n')) + '\n' + new
        text = text.replace(old, repl)
        n_l += 1
    assert n_b == expect_blocks, 'expected %d amount blocks, patched %d' % (expect_blocks, n_b)
    assert n_l == expect_lines, 'expected %d amount lines, patched %d' % (expect_lines, n_l)
    return text


# ----------------------------------------------------------------- Postings.bas
# The module that actually posts, and then verifies the posting. It carried five
# more copies of the same unbounded print-wait loop as Printing.bas -- including
# one inside Check_ZGE132AP, which runs immediately AFTER the entries have gone
# into SAP. A freeze there is the worst case: the postings are real, but the run
# never reaches the verification that would tell the operator so.
# It also carried eleven more locale-dependent amount conversions.

POSTINGS_LABELS = [
    'ZGE132 re-check (local currency)',
    'ZGE132 re-check (group currency)',
    'ZGLGWUL',
    'ZGE132 after posting',
    'ZGR215 (batch-input documents)',
]

POSTINGS_BLOCKS = [
    # Check_ZGE132AP, local currency
    ('            If Right(arr(LBound(arr) + 1), 1) = "-" Then\n'
     '                AmL = -Round(Left(arr(LBound(arr) + 1), Len(arr(LBound(arr) + 1)) - 1), 2)\n'
     '            Else\n'
     '                AmL = Round(arr(LBound(arr) + 1), 2)\n'
     '            End If',
     'AmL = Round(CM_AmountReq(arr(LBound(arr) + 1), 0, _\n'
     '                                     "re-reading the local-currency total from ZGE132 after posting"), 2)'),

    ('            If Right(arr(UBound(arr)), 1) = "-" Then\n'
     '                Am = Am + -Round(Left(arr(UBound(arr)), Len(arr(UBound(arr))) - 1), 2)\n'
     '            Else\n'
     '                Am = Am + Round(arr(UBound(arr)), 2)\n'
     '            End If',
     'Am = Am + Round(CM_AmountReq(arr(UBound(arr)), 0, _\n'
     '                                         "re-reading a profit-centre amount from ZGE132 after posting"), 2)'),

    ('            If Right(arr(UBound(arr) - 3), 1) = "-" Then\n'
     '                AmG = -Round(Left(arr(UBound(arr) - 3), Len(arr(UBound(arr) - 3)) - 1), 2)\n'
     '            Else\n'
     '                AmG = Round(arr(UBound(arr) - 3), 2)\n'
     '            End If',
     'AmG = Round(CM_AmountReq(arr(UBound(arr) - 3), 0, _\n'
     '                                     "re-reading the group-currency total from ZGE132 after posting"), 2)'),

    ('            If Right(arr(UBound(arr) - 2), 1) = "-" Then\n'
     '                Am = Am + -Round(Left(arr(UBound(arr) - 2), Len(arr(UBound(arr) - 2)) - 1), 2)\n'
     '            Else\n'
     '                Am = Am + Round(arr(UBound(arr) - 2), 2)\n'
     '            End If',
     'Am = Am + Round(CM_AmountReq(arr(UBound(arr) - 2), 0, _\n'
     '                                         "re-reading a profit-centre amount from ZGE132 (group currency)"), 2)'),
]

# Post_ZGLGWUL. Two defects in five lines: the locale-dependent Round, and
# Left(arr(U), Len(arr(U - 1))) -- the length of a DIFFERENT array element, so a
# negative amount was truncated to whatever the neighbouring column happened to
# be wide. The "|" (empty column) case is preserved exactly.
ZGLGWUL_OLD = ('                If Right(arr(UBound(arr)), 1) = "-" Then\n'
               '                    Am = -Round(Left(arr(UBound(arr)), Len(arr(UBound(arr) - 1))), 2)\n'
               '                ElseIf arr(UBound(arr)) = "|" Then\n'
               '                    Am = 0\n'
               '                Else\n'
               '                    Am = Round(arr(UBound(arr)), 2)\n'
               '                End If')
ZGLGWUL_NEW = ("                'V4-CIO FIX: locale-proof amount, and the trailing minus is now stripped\n"
               "                'correctly. The old negative branch read Left(arr(U), Len(arr(U - 1))) -\n"
               "                'the length of the PREVIOUS column - so a negative amount was cut to the\n"
               "                'wrong number of characters whenever the two columns differed in width.\n"
               '                If arr(UBound(arr)) = "|" Then\n'
               '                    Am = 0\n'
               '                Else\n'
               '                    Am = Round(CM_AmountReq(arr(UBound(arr)), 0, _\n'
               '                               "reading account 44400200 from the ZGLGWUL extract"), 2)\n'
               '                End If')

ZGE132AG_OLD = '            ElseIf Round(Trim(getLineData(line, "Amount", 1)), 2) <> 0 Then'
ZGE132AG_NEW = ("            'V4-CIO FIX: locale-proof amount (was Round on a SAP string -> error 13)\n"
                '            ElseIf Round(CM_AmountReq(Trim(getLineData(line, "Amount", 1)), 0, _\n'
                '                         "checking whether ZGE132 still has an amount to post"), 2) <> 0 Then')


def postings(text):
    L = text.split('\n')
    sites = [i for i, l in enumerate(L) if 'Do Until File <> ""' in l]
    assert len(sites) == len(POSTINGS_LABELS), \
        'expected %d print-wait loops in Postings, found %d' % (len(POSTINGS_LABELS), len(sites))
    for k in range(len(sites) - 1, -1, -1):
        st = sites[k]
        ind = _ind(L[st])
        end = next(j for j in range(st + 1, len(L))
                   if L[j].strip() == 'Loop' and _ind(L[j]) == ind)
        L[st:end + 1] = [
            ind + "'V4-CIO FIX: bounded, responsive wait for the printed file. The old loop",
            ind + "'rebuilt a Shell.Application on every pass with no pause - a tight spin that",
            ind + "'froze Excel and ended in error 80010108. In this module that freeze could",
            ind + "'strand a run with entries already posted but not yet verified.",
            ind + 'File = CM_WaitForPrint(FTemp, fso, "%s")' % POSTINGS_LABELS[k],
        ]
    text = '\n'.join(L)

    n = 0
    for old, new in POSTINGS_BLOCKS:
        c = text.count(old)
        assert c == 1, 'Postings amount block matched %d times: %r' % (c, old[:60])
        ind = '            '
        text = text.replace(old, '\n'.join(ind + l for l in NOTE.split('\n')) + '\n' + ind + new)
        n += 1
    assert text.count(ZGLGWUL_OLD) == 1, 'ZGLGWUL block not found'
    text = text.replace(ZGLGWUL_OLD, ZGLGWUL_NEW)
    assert text.count(ZGE132AG_OLD) == 1, 'ZGE132AG line not found'
    text = text.replace(ZGE132AG_OLD, ZGE132AG_NEW)
    assert n == 4
    return review_fixes_postings(instrument_reads(text, 4, 4))


# --------------------------------------------------- extract-read instrumentation
# So a bad amount can name the file, the line number and the line itself. Every
# "strim.LoadFromFile (FPath & "x.txt")" is followed by CM_Source, and every
# "line = strim.ReadText(-2)" by CM_Reading. Purely additive - no existing
# statement is touched.
import re as _re

_LOAD = _re.compile(r'^(\s*)strim\.LoadFromFile \(FPath & "([^"]+)"\)\s*$')
_READ = _re.compile(r'^(\s*)line = strim\.ReadText\(-2\)\s*$')

def instrument_reads(text, expect_loads, expect_reads):
    L = text.split('\n')
    out, nl, nr = [], 0, 0
    for line in L:
        out.append(line)
        m = _LOAD.match(line)
        if m:
            out.append('%sCM_Source "%s"' % (m.group(1), m.group(2)))
            nl += 1
            continue
        m = _READ.match(line)
        if m:
            out.append('%sCM_Reading line' % m.group(1))
            nr += 1
    assert nl == expect_loads, 'expected %d LoadFromFile, instrumented %d' % (expect_loads, nl)
    assert nr == expect_reads, 'expected %d ReadText, instrumented %d' % (expect_reads, nr)
    return '\n'.join(out)


# The ZGLRME amount comes from a worksheet, not a text file: name the sheet, and
# report the real sheet row (the array starts at A2, so array row i is sheet row
# i + 1). Reporting only - the loop still indexes with i.
ZGLRME_OLD = 'ArrZGL(i, 5) = CM_Amount(ArrZGL(i, 5), i, "reading the amounts from the ZGLRME extract")'
ZGLRME_NEW = ('CM_Source ""\n'
              '    ArrZGL(i, 5) = CM_Amount(ArrZGL(i, 5), i + 1, "reading the amounts from the ZGLRME extract")')


# ------------------------------------------------------- review fixes (originals)
# Three defects that were in the original macro, found by an independent read of
# the shipping code. All three are hangs or silent-wrong-number paths, which is
# the class this whole build exists to remove.

# (a) Range(row, col) where Cells(row, col) was meant. Worksheet.Range takes an
#     A1 string or a Range, so a Long gives run-time error 1004 EVERY time this
#     branch runs -- and it only runs after ZGE132/ZGLGWUL have already posted.
#     Every neighbouring statement, and the identical logic in CheckZGLRME, uses
#     .Cells. It is a typo, and the branch has never worked.
GTB1_BOLD = [
    ('                    Sheets("Errors").Range(EmptRow, 1).Font.Bold = True',
     '                    Sheets("Errors").Cells(EmptRow, 1).Font.Bold = True'),
    ('                    Sheets("Errors").Range(EmptRow1, 1).Font.Bold = True',
     '                    Sheets("Errors").Cells(EmptRow1, 1).Font.Bold = True'),
]

# (b) Post_ZGLGWUL arms On Error Resume Next and only disarms it inside the Else
#     branch. On the other branch the handler stays armed for the rest of the
#     procedure, so a later failure -- including the explained one CM_AmountReq
#     raises -- is swallowed and AA16 is written from a stale Am.
ZGLGWUL_ERR_OLD = (
    '    If Err.Number <> 0 Then\n'
    '        ErrTxt = .findById("wnd[1]/usr/txtMESSTXT1").Text')
ZGLGWUL_ERR_NEW = (
    '    If Err.Number <> 0 Then\n'
    "        'V4-CIO FIX: disarm here too. Without this the handler stayed armed\n"
    "        'for the rest of the procedure and swallowed everything after it,\n"
    "        'including the explained amount error - AA16 was then written from\n"
    "        'whatever Am happened to hold.\n"
    '        On Error Resume Next\n'
    '        ErrTxt = .findById("wnd[1]/usr/txtMESSTXT1").Text\n'
    '        Err.Clear\n'
    '        On Error GoTo 0')


def review_fixes_closing(text):
    for old, new in GTB1_BOLD:
        assert text.count(old) == 1, 'GTB1 bold line not found: %r' % old[:60]
        text = text.replace(old, new)
    return text


def review_fixes_postings(text):
    assert text.count(ZGLGWUL_ERR_OLD) == 1, 'Post_ZGLGWUL error branch not found'
    return text.replace(ZGLGWUL_ERR_OLD, ZGLGWUL_ERR_NEW)


# (c) Two loops in GlobalModule that cannot terminate on an empty or pipe-less
#     SAP export. CreateArray is called on t001.txt, zgxmit.txt, t001z.txt,
#     skb1.txt, closestatus.txt, zglrme.txt and zge132gwul.txt, so an export that
#     comes back empty hangs Excel with no message and no way out - exactly the
#     symptom this build set out to eliminate. Neither guard changes behaviour on
#     data the loops already handled.
IMPORT_OLD = "\n".join([
    'Do',
    '    Data = strix.ReadText(-2)',
    '    If VBA.Left(VBA.Trim(Data), 1) = "|" Then',
    '        For i = 2 To Len(Trim(Data))',
    '            If Mid(Trim(Data), i, 1) = "|" Then',
    '                FirstColumn = Trim(Mid(Trim(Data), 2, i - 2))',
    '                Exit For',
    '            End If',
    '        Next i',
    '        Exit Do',
    '    End If',
    'Loop',
])

IMPORT_NEW = "\n".join([
    "'V4-CIO FIX: the only way out of this loop was finding a line that starts",
    "'with \"|\". An empty or pipe-less export spun here for ever, hanging Excel",
    "'with no message. It now stops at end of file and says which file.",
    'Dim cmFound As Boolean',
    'Do Until strix.EOS',
    '    Data = strix.ReadText(-2)',
    '    If VBA.Left(VBA.Trim(Data), 1) = "|" Then',
    '        cmFound = True',
    '        For i = 2 To Len(Trim(Data))',
    '            If Mid(Trim(Data), i, 1) = "|" Then',
    '                FirstColumn = Trim(Mid(Trim(Data), 2, i - 2))',
    '                Exit For',
    '            End If',
    '        Next i',
    '        Exit Do',
    '    End If',
    'Loop',
    'If Not cmFound Then',
    '    strix.Close',
    '    Set strix = Nothing',
    '    CM_Step = "reading the SAP export " & nazwa',
    '    Err.Raise vbObjectError + 515, "ClosingManager", "EMPTY" & Chr$(1) & nazwa',
    'End If',
])

PROPER_OLD = "\n".join([
    '    For a = nStart To VBA.Len(Data)',
    '        If VBA.Mid(Data, a, 1) = "|" Then',
    '            nStart = a + 1',
    '            nEnd = nStart',
    '            Exit For',
    '        End If',
    '    Next',
])

PROPER_NEW = "\n".join([
    '    cmLast = nStart',
    '    For a = nStart To VBA.Len(Data)',
    '        If VBA.Mid(Data, a, 1) = "|" Then',
    '            nStart = a + 1',
    '            nEnd = nStart',
    '            Exit For',
    '        End If',
    '    Next',
    "    'V4-CIO FIX: nStart only moves when a \"|\" is found. With none left it",
    "    'stayed put and this loop ran for ever, while ReDim Preserve grew the",
    "    'array on every pass - a hang that ended in out-of-memory, if at all.",
    '    If nStart = cmLast Then Exit Do',
])


def review_fixes_globalmodule(text):
    assert text.count(IMPORT_OLD) == 1, 'ImportTitle loop not found'
    text = text.replace(IMPORT_OLD, IMPORT_NEW)
    assert text.count(PROPER_OLD) == 1, 'ProperArray loop not found'
    text = text.replace(PROPER_OLD, PROPER_NEW)
    old = 'Dim a, b, n\n'
    assert text.count(old) == 1, 'ProperArray Dim line not found'
    return text.replace(old, 'Dim a, b, n\nDim cmLast As Long\n')


# Admin.UpdateData also calls CreateArray, so it can now raise the explained
# empty-export error. Without a handler the operator gets a raw VBA dialog.
#
# Unlike RunClosing, UpdateData already contains four short "On Error Resume
# Next / On Error GoTo 0" windows. "On Error GoTo 0" DISABLES handling, so an
# outer handler armed at the top would be switched off by the first one. Each
# window is three lines - arm, one guarded statement, disarm - so the disarm is
# rewritten to restore the outer handler, which is what it always meant.
def admin(text):
    L = text.split('\n')
    i = L.index('Sub UpdateData()')
    end = next(j for j in range(i + 1, len(L)) if L[j].strip() == 'End Sub')

    n_restore = 0
    for j in range(i, end):
        if L[j].strip() == 'On Error GoTo 0':
            L[j] = L[j].replace('On Error GoTo 0', 'On Error GoTo CM_Fail')
            n_restore += 1
    assert n_restore == 4, 'expected 4 On Error GoTo 0 in UpdateData, found %d' % n_restore

    n_exit = 0
    for j in range(end - 1, i, -1):
        if L[j].strip() == 'Exit Sub':
            ind = L[j][:len(L[j]) - len(L[j].lstrip())]
            L[j:j] = [ind + 'CM_Done']
            n_exit += 1
    assert n_exit == 1, 'expected 1 Exit Sub in UpdateData, found %d' % n_exit
    end += n_exit

    L[end:end] = [
        '',
        "'V4-CIO: the same plain-language failure reporting as the close.",
        'CM_Done',
        'Exit Sub',
        '',
        'CM_Fail:',
        '    CM_Explain Err.Number, Err.Description',
        '    CM_Done',
        '',
    ]
    L[i + 1:i + 1] = ['', 'On Error GoTo CM_Fail', 'CM_Begin 0',
                      'CM_Note "refreshing the reference data from SAP"']
    return '\n'.join(L)
