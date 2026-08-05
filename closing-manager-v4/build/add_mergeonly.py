#!/usr/bin/env python3
"""Make the merge test standalone: it no longer depends on PDFCreator having
produced two PDFs first. Two built-in one-page PDFs are written directly to
disk, so GiosPSMC.exe can be proved on its own."""
import io, os

SRC = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'build_v4')
P   = os.path.join(SRC, 'mCloseEnv_V4.bas')
s   = io.open(P, encoding='utf-8', newline='').read()

# ------------------------------------------------- 1. three-way consent dialog
old = '''    ' 8) Print + merge rehearsal (optional - it really prints, so it asks) -----
    If pdfOK And winPrn <> "" And Not IsUrlPath(ThisWorkbook.Path) Then
        If MsgBox("Test printing and merging for real?" & vbCrLf & vbCrLf & _
                  "This prints two small test pages to PDFCreator, then merges them" & vbCrLf & _
                  "with GiosPSMC.exe - the exact chain that produces the report pack." & vbCrLf & _
                  "It is the only way to prove PDFCreator is saving automatically" & vbCrLf & _
                  "into " & CM_BASE_DRIVE & "pdf\\temp and that the merger works." & vbCrLf & vbCrLf & _
                  "It writes only test files under " & CM_BASE_DRIVE & "pdf\\ and deletes them" & vbCrLf & _
                  "afterwards. If SAP is available the pages come from SE16/T001 -" & vbCrLf & _
                  "a display only, which creates a temporary print job and no data." & vbCrLf & _
                  "Allow up to two minutes.", _
                  vbYesNo + vbQuestion, "Closing Manager - Preflight") = vbYes Then
            msg = msg & CM_PrintMergeTest(sapSess, prtBlocked)
            If prtBlocked Then okAll = False
        Else
            msg = msg & "[~] Print/merge not tested (skipped)." & vbCrLf
        End If
    End If
'''
assert s.count(old) == 1
new = '''    ' 8) Print + merge rehearsal (optional - it really prints, so it asks) -----
    If Not IsUrlPath(ThisWorkbook.Path) Then
        ans = MsgBox("Test printing and merging for real?" & vbCrLf & vbCrLf & _
              "YES     print two test pages, then merge them." & vbCrLf & _
              "        Proves the whole chain: SAP -> PDFCreator -> " & CM_BASE_DRIVE & "pdf\\temp" & vbCrLf & _
              "        -> GiosPSMC.exe. Allow up to two minutes." & vbCrLf & vbCrLf & _
              "NO      test the PDF merger only." & vbCrLf & _
              "        Uses two built-in test pages, so nothing is printed and" & vbCrLf & _
              "        PDFCreator is not involved. Takes a few seconds." & vbCrLf & vbCrLf & _
              "CANCEL  skip this check." & vbCrLf & vbCrLf & _
              "Either test writes only under " & CM_BASE_DRIVE & "pdf\\ and clears up after" & vbCrLf & _
              "itself. Printing uses SE16/T001 - a display, so no data is created" & vbCrLf & _
              "or changed in SAP.", _
              vbYesNoCancel + vbQuestion, "Closing Manager - Preflight")

        If ans = vbCancel Then
            msg = msg & "[~] Print/merge not tested (skipped)." & vbCrLf
        Else
            msg = msg & CM_PrintMergeTest(sapSess, prtBlocked, (ans = vbYes))
            If prtBlocked Then okAll = False
        End If
    End If
'''
s = s.replace(old, new)

old = '''    Dim pdfOK As Boolean, prtBlocked As Boolean
'''
assert s.count(old) == 1
s = s.replace(old, '''    Dim pdfOK As Boolean, prtBlocked As Boolean, ans As VbMsgBoxResult
''')

# ------------------------------------------------------ 2. rewrite the test
start = s.index('Public Function CM_PrintMergeTest(')
end   = s.index("'--- print SE16/T001 - a display, so nothing is created or changed -----------")
new_fn = r'''Public Function CM_PrintMergeTest(ByVal sess As Object, ByRef blocked As Boolean, _
                                  ByVal doPrint As Boolean) As String
    Dim fso As Object, out As String, route As String, why As String, src As String
    Dim tmp As String, work As String, exePath As String, merged As String
    Dim p As String, n As Long, i As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Call EnsureFolders

    tmp     = CM_BASE_DRIVE & "pdf\temp\"
    work    = CM_BASE_DRIVE & "pdf\_preflight\"
    exePath = CM_BASE_DRIVE & "pdf\merger\GiosPSMC.exe"
    merged  = work & "merged.pdf"

    On Error Resume Next
    fso.DeleteFolder work, True
    Err.Clear
    On Error GoTo 0
    fso.CreateFolder work

    If Not doPrint Then GoTo DoMerge

    'the close picks up whatever is sitting in \temp, so leftovers are a fault
    'in their own right - and they would make the print test meaningless
    n = CM_CountFiles(fso, tmp)
    If n > 0 Then
        If MsgBox(n & " leftover file(s) are sitting in" & vbCrLf & tmp & vbCrLf & vbCrLf & _
                  "The close treats whatever is in that folder as its freshly printed" & vbCrLf & _
                  "report, so leftovers can end up inside the report pack." & vbCrLf & vbCrLf & _
                  "Delete them now?", vbYesNo + vbExclamation, _
                  "Closing Manager - Preflight") = vbYes Then
            On Error Resume Next
            fso.DeleteFile tmp & "*.*", True
            Err.Clear
            On Error GoTo 0
            out = out & "[OK] Cleared " & n & " leftover file(s) from \pdf\temp." & vbCrLf
        Else
            blocked = True
            out = out & "[X] " & n & " leftover file(s) in " & tmp & _
                        " - printing not tested." & vbCrLf
            GoTo DoMerge
        End If
    End If

    Call SetPDFCreator          'same launch the close does

    'two test pages, printed one at a time so each can be identified
    For i = 1 To 2
        CM_Step = "preflight: printing test page " & i & " of 2"
        CM_Paint ""

        why = ""
        route = "SAP"
        If sess Is Nothing Then
            route = "Excel"
        Else
            why = CM_SapTestPrint(sess)
            If why <> "" Then route = "Excel"
        End If
        If route = "Excel" Then why = CM_ExcelTestPrint()

        If why <> "" Then
            blocked = True
            out = out & "[X] Could not send test page " & i & " to PDFCreator - " & why & "." & vbCrLf
            GoTo DoMerge
        End If

        p = CM_WaitForFile(fso, tmp, 90, "test page " & i)
        If p = "" Then
            blocked = True
            out = out & "[X] Test page " & i & " was printed from " & route & _
                        " but no PDF appeared in" & vbCrLf & _
                        "       " & tmp & " within 90 seconds." & vbCrLf & _
                        "       PDFCreator is not saving automatically to that folder," & vbCrLf & _
                        "       or it is waiting on a dialog. This is what makes the" & vbCrLf & _
                        "       close appear to freeze." & vbCrLf
            GoTo DoMerge
        End If

        On Error Resume Next
        fso.MoveFile p, work & "t" & i & ".pdf"
        If Err.Number <> 0 Then
            Err.Clear: On Error GoTo 0
            blocked = True
            out = out & "[X] The test PDF was created but could not be moved out of \pdf\temp." & vbCrLf
            GoTo DoMerge
        End If
        On Error GoTo 0

        out = out & "[OK] Test page " & i & ": printed from " & route & _
                    " and saved as PDF (" & CM_KB(fso, work & "t" & i & ".pdf") & " KB)." & vbCrLf
    Next i

DoMerge:
    'The merger is tested whether or not the printing worked. If we did not get
    'two printed PDFs, two built-in one-page PDFs are written straight to disk,
    'so GiosPSMC.exe is proved on its own and a broken PDFCreator cannot hide a
    'broken merger (or the other way round).
    src = "the 2 printed test pages"
    If Not (fso.FileExists(work & "t1.pdf") And fso.FileExists(work & "t2.pdf")) Then
        On Error Resume Next
        fso.DeleteFile work & "*.*", True
        Err.Clear
        On Error GoTo 0
        If CM_SeedPdf(work & "t1.pdf", "PDF merge test page 1") And _
           CM_SeedPdf(work & "t2.pdf", "PDF merge test page 2") Then
            src = "2 built-in test pages"
        Else
            blocked = True
            out = out & "[X] Could not write the built-in test pages to " & work & _
                        " - merge not tested." & vbCrLf
            GoTo CleanUp
        End If
    End If

    If Not fso.FileExists(exePath) Then
        blocked = True
        out = out & "[X] PDF merger not present at " & exePath & " - merge not tested." & vbCrLf
        GoTo CleanUp
    End If

    CM_Step = "preflight: merging " & src
    CM_Paint ""
    On Error Resume Next
    Err.Clear
    CreateObject("WScript.Shell").Run "%COMSPEC% /c " & exePath & " " & _
        Chr$(34) & work & "t1.pdf" & Chr$(34) & " " & _
        Chr$(34) & work & "t2.pdf" & Chr$(34) & " output " & merged, 0, False
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        blocked = True
        out = out & "[X] The PDF merger would not start: " & exePath & vbCrLf
        GoTo CleanUp
    End If
    On Error GoTo 0

    If Not CM_WaitForPath(fso, merged, 60, "the merged test file") Then
        blocked = True
        out = out & "[X] The merger ran but produced no merged file within 60 seconds." & vbCrLf & _
                    "       " & exePath & vbCrLf & _
                    "       Check the file is a real GiosPSMC.exe and is not blocked by" & vbCrLf & _
                    "       Windows (Properties > Unblock) or by antivirus." & vbCrLf
    ElseIf CM_KB(fso, merged) < 0 Then
        blocked = True
        out = out & "[X] The merger produced an unreadable file." & vbCrLf
    ElseIf Not CM_LooksLikePdf(merged) Then
        blocked = True
        out = out & "[X] The merger produced a file that is not a PDF." & vbCrLf
    Else
        out = out & "[OK] PDF merge works: " & src & " -> 1 file (" & _
                    CM_KB(fso, merged) & " KB)." & vbCrLf
    End If

CleanUp:
    On Error Resume Next
    fso.DeleteFolder work, True
    If doPrint Then fso.DeleteFile tmp & "*.*", True
    Err.Clear
    On Error GoTo 0
    CM_Done
    CM_PrintMergeTest = out
End Function


'--- a real PDF starts with %PDF- and ends with %%EOF ------------------------
Private Function CM_LooksLikePdf(ByVal path As String) As Boolean
    Dim f As Integer, head As String * 5, sz As Long, tail As String * 6
    On Error Resume Next
    f = FreeFile
    Open path For Binary Access Read As #f
    sz = LOF(f)
    If sz > 32 Then
        Get #f, 1, head
        Get #f, sz - 5, tail
    End If
    Close #f
    If Err.Number = 0 Then
        CM_LooksLikePdf = (head = "%PDF-") And (InStr(tail, "%%EOF") > 0)
    End If
    Err.Clear
    On Error GoTo 0
End Function


'============================================================================
' A BUILT-IN ONE-PAGE PDF
'----------------------------------------------------------------------------
' So the merger can be tested on its own, with no printer involved. This writes
' a minimal but fully valid PDF 1.4 - catalog, page tree, one A4 page, one
' Helvetica text block - with a correct cross-reference table, because a merger
' will refuse a file whose xref offsets are wrong.
'
' Everything written is 7-bit ASCII and the line ending is LF, so the byte
' length of the string equals its character length and the offsets below are
' exact on any Windows regional setting.
'============================================================================
Private Function CM_SeedPdf(ByVal path As String, ByVal caption As String) As Boolean
    Dim NL As String, body As String, out As String
    Dim o(1 To 5) As String, offs(1 To 5) As Long
    Dim i As Long, xrefAt As Long, f As Integer
    Dim b() As Byte

    NL = Chr$(10)
    body = "BT" & NL & _
           "/F1 16 Tf" & NL & _
           "60 760 Td" & NL & _
           "(Closing Manager - preflight test) Tj" & NL & _
           "0 -28 Td" & NL & _
           "(" & caption & ") Tj" & NL & _
           "ET" & NL

    o(1) = "<< /Type /Catalog /Pages 2 0 R >>"
    o(2) = "<< /Type /Pages /Kids [3 0 R] /Count 1 >>"
    o(3) = "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] " & _
           "/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>"
    o(4) = "<< /Length " & Len(body) & " >>" & NL & "stream" & NL & body & "endstream"
    o(5) = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"

    out = "%PDF-1.4" & NL
    For i = 1 To 5
        offs(i) = Len(out)
        out = out & i & " 0 obj" & NL & o(i) & NL & "endobj" & NL
    Next i

    xrefAt = Len(out)
    out = out & "xref" & NL & "0 6" & NL & "0000000000 65535 f " & NL
    For i = 1 To 5
        out = out & Format$(offs(i), "0000000000") & " 00000 n " & NL
    Next i
    out = out & "trailer" & NL & _
                "<< /Size 6 /Root 1 0 R >>" & NL & _
                "startxref" & NL & xrefAt & NL & "%%EOF" & NL

    On Error Resume Next
    If Dir(path) <> "" Then Kill path
    Err.Clear
    b = StrConv(out, vbFromUnicode)
    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, 1, b
    Close #f
    If Err.Number = 0 Then CM_SeedPdf = True
    Err.Clear
    On Error GoTo 0
End Function


'''
s = s[:start] + new_fn + s[end:]

io.open(P, 'w', encoding='utf-8', newline='').write(s)
print('merge test is now standalone (built-in seed PDFs, 3-way consent)')
