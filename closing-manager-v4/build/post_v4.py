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
    return amounts('\n'.join(L), 6, 2)

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
    return amounts('\n'.join(L), 0, 1)


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
    return text
