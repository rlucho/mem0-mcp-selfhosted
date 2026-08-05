#!/usr/bin/env python3
"""Preflight: end-to-end PDFCreator + merger rehearsal."""
import io, os

D   = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(os.path.dirname(D), 'build_v4')
P   = os.path.join(SRC, 'mCloseEnv_V4.bas')
env = io.open(P, encoding='utf-8', newline='').read()

# ---------------------------------------------------------------- 1. new check
# "PDFCreator installed" now also proves the Windows printer exists and is
# named exactly 'PDFCreator' - the name SAP is told to use.
old = '''    ' 4) PDFCreator ------------------------------------------------------------
    If fso.FileExists("C:\\Program Files\\PDFCreator\\PDFCreator.exe") Or _
       fso.FileExists("C:\\Program Files (x86)\\PDFCreator\\PDFCreator.exe") Then
        msg = msg & "[OK] PDFCreator is installed." & vbCrLf
    Else
        msg = msg & "[X] PDFCreator.exe not found (printer must be named 'PDFCreator')." & vbCrLf
        okAll = False
    End If
'''
assert env.count(old) == 1
new = '''    ' 4) PDFCreator ------------------------------------------------------------
    pdfOK = fso.FileExists("C:\\Program Files\\PDFCreator\\PDFCreator.exe") Or _
            fso.FileExists("C:\\Program Files (x86)\\PDFCreator\\PDFCreator.exe")
    If pdfOK Then
        msg = msg & "[OK] PDFCreator is installed." & vbCrLf
    Else
        msg = msg & "[X] PDFCreator.exe not found (printer must be named 'PDFCreator')." & vbCrLf
        okAll = False
    End If

    ' 4b) the Windows printer itself - SAP is told to print to the name below,
    '     so a renamed printer fails the close even with PDFCreator installed.
    winPrn = CM_PdfCreatorPrinter()
    If winPrn <> "" Then
        msg = msg & "[OK] Windows printer found: " & winPrn & "." & vbCrLf
    Else
        msg = msg & "[X] No Windows printer whose name contains 'PDFCreator'." & vbCrLf
        okAll = False
    End If
'''
env = env.replace(old, new)

old = ('    Dim sapSess As Object, sapBlocked As Boolean, runData As Boolean\n'
       '    Dim cc As String, dst As String\n')
assert env.count(old) == 1
env = env.replace(old,
       '    Dim sapSess As Object, sapBlocked As Boolean, runData As Boolean\n'
       '    Dim cc As String, dst As String, winPrn As String\n'
       '    Dim pdfOK As Boolean, prtBlocked As Boolean\n')

# ------------------------------------------------- 2. offer the print rehearsal
old = '''    msg = "CLOSING MANAGER  -  PREFLIGHT CHECK" & vbCrLf & _'''
assert env.count(old) == 1
new = '''    ' 8) Print + merge rehearsal (optional - it really prints, so it asks) -----
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

    msg = "CLOSING MANAGER  -  PREFLIGHT CHECK" & vbCrLf & _'''
env = env.replace(old, new)

# --------------------------------------------------------- 3. the test itself
env += r'''

'============================================================================
' PRINT + MERGE REHEARSAL
'----------------------------------------------------------------------------
' The close breaks most often at the printing step, and the old code could not
' tell the difference between "PDFCreator is slow" and "PDFCreator is never
' going to produce anything" - it simply waited for ever. Every other preflight
' check is a look; this one is a rehearsal, because the only way to know the
' chain works is to run it:
'
'     SAP  ->  front-end printer LOCLX  ->  Windows printer "PDFCreator"
'          ->  auto-saved PDF in C:\pdf\temp  ->  GiosPSMC.exe merge
'
' It writes only into C:\pdf\ and removes what it wrote. In SAP it displays
' SE16/T001 (company-code names) and prints that list: a display, so no
' business data is created or changed. Printing does raise a temporary SAP
' print job, which is why PreflightCheck asks first. If SAP is not available,
' or its print dialog does not appear, it falls back to printing a page from
' Excel and says so - that still proves PDFCreator and the merger.
'============================================================================
Public Function CM_PrintMergeTest(ByVal sess As Object, ByRef blocked As Boolean) As String
    Dim fso As Object, out As String, route As String, why As String
    Dim tmp As String, work As String, exePath As String, merged As String
    Dim p As String, n As Long, i As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Call EnsureFolders

    tmp     = CM_BASE_DRIVE & "pdf\temp\"
    work    = CM_BASE_DRIVE & "pdf\_preflight\"
    exePath = CM_BASE_DRIVE & "pdf\merger\GiosPSMC.exe"
    merged  = work & "merged.pdf"

    'the close picks up whatever is sitting in \temp, so leftovers are a fault
    'in their own right - and they would make this test meaningless
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
            CM_PrintMergeTest = out & "[X] " & n & " leftover file(s) in " & tmp & _
                                " - clear them before closing." & vbCrLf
            CM_Done
            Exit Function
        End If
    End If

    On Error Resume Next
    fso.DeleteFolder work, True
    Err.Clear
    On Error GoTo 0
    fso.CreateFolder work

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
            GoTo CleanUp
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
            GoTo CleanUp
        End If

        On Error Resume Next
        fso.MoveFile p, work & "t" & i & ".pdf"
        If Err.Number <> 0 Then
            Err.Clear: On Error GoTo 0
            blocked = True
            out = out & "[X] The test PDF was created but could not be moved out of \pdf\temp." & vbCrLf
            GoTo CleanUp
        End If
        On Error GoTo 0

        out = out & "[OK] Test page " & i & ": printed from " & route & _
                    " and saved as PDF (" & CM_KB(fso, work & "t" & i & ".pdf") & " KB)." & vbCrLf
    Next i

    'now the merge, using the same command line the close uses
    If Not fso.FileExists(exePath) Then
        blocked = True
        out = out & "[X] PDF merger not present at " & exePath & " - merge not tested." & vbCrLf
        GoTo CleanUp
    End If

    CM_Step = "preflight: merging the two test pages"
    CM_Paint ""
    On Error Resume Next
    CreateObject("WScript.Shell").Run "%COMSPEC% /c " & exePath & " " & _
        Chr$(34) & work & "t1.pdf" & Chr$(34) & " " & _
        Chr$(34) & work & "t2.pdf" & Chr$(34) & " output " & merged, 0, False
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        blocked = True
        out = out & "[X] The PDF merger would not start." & vbCrLf
        GoTo CleanUp
    End If
    On Error GoTo 0

    If CM_WaitForFile(fso, work, 60, "merged test file") = "" Or Not fso.FileExists(merged) Then
        blocked = True
        out = out & "[X] The merger ran but produced no merged file within 60 seconds." & vbCrLf & _
                    "       " & exePath & vbCrLf
    ElseIf CM_KB(fso, merged) <= 0 Then
        blocked = True
        out = out & "[X] The merger produced an empty file." & vbCrLf
    Else
        out = out & "[OK] Merge works (" & CM_KB(fso, merged) & " KB from 2 pages)." & vbCrLf
    End If

CleanUp:
    On Error Resume Next
    fso.DeleteFolder work, True
    fso.DeleteFile tmp & "*.*", True
    Err.Clear
    On Error GoTo 0
    CM_Done
    CM_PrintMergeTest = out
End Function


'--- print SE16/T001 - a display, so nothing is created or changed -----------
Private Function CM_SapTestPrint(ByVal sess As Object) As String
    On Error Resume Next
    Err.Clear
    CM_ClearPopups sess

    sess.findById("wnd[0]/tbar[0]/okcd").Text = "/nse16"
    sess.findById("wnd[0]").sendVKey 0
    sess.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "T001"
    sess.findById("wnd[0]").sendVKey 0
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        CM_SapTestPrint = "SE16 did not open"
        Exit Function
    End If

    'keep the printed list to a single short page
    sess.findById("wnd[0]/usr/txtMAX_SEL").Text = "3"
    Err.Clear

    sess.findById("wnd[0]/tbar[1]/btn[8]").press          'F8 - display
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        CM_SapTestPrint = "SE16 would not display table T001"
        Exit Function
    End If

    sess.findById("wnd[0]").sendVKey 86                   'Ctrl+P - print
    sess.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCLX"
    sess.findById("wnd[1]").sendVKey 0
    If Err.Number <> 0 Then
        Err.Clear
        CM_SapBackOut sess
        On Error GoTo 0
        CM_SapTestPrint = "the SAP print dialog did not appear"
        Exit Function
    End If

    sess.findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
    If Err.Number <> 0 Then
        Err.Clear
        'the printer list can sit one window deeper, exactly as in Print_ZGLRME
        sess.findById("wnd[2]/tbar[0]/btn[0]").press
        sess.findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
        If Err.Number <> 0 Then
            Err.Clear
            CM_SapBackOut sess
            On Error GoTo 0
            CM_SapTestPrint = "SAP does not offer a front-end printer called 'PDFCreator'"
            Exit Function
        End If
    End If

    sess.findById("wnd[1]/tbar[0]/btn[13]").press         'print
    If Err.Number <> 0 Then
        Err.Clear
        CM_SapBackOut sess
        On Error GoTo 0
        CM_SapTestPrint = "SAP would not accept the print"
        Exit Function
    End If

    CM_SapBackOut sess
    Err.Clear
    On Error GoTo 0
End Function


'--- leave SAP on a clean screen ---------------------------------------------
Private Sub CM_SapBackOut(ByVal sess As Object)
    On Error Resume Next
    sess.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    sess.findById("wnd[0]").sendVKey 0
    Err.Clear
    On Error GoTo 0
End Sub


'--- fallback: print one page from Excel to the same printer -----------------
Private Function CM_ExcelTestPrint() As String
    Dim prn As String, ws As Object
    prn = CM_PdfCreatorPrinter()
    If prn = "" Then
        CM_ExcelTestPrint = "no Windows printer called 'PDFCreator'"
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    Set ws = ThisWorkbook.Sheets("Preflight")
    If ws Is Nothing Then Set ws = ThisWorkbook.Sheets(1)
    Err.Clear
    ws.PrintOut Copies:=1, ActivePrinter:=prn, Collate:=True
    If Err.Number <> 0 Then CM_ExcelTestPrint = "Excel could not print to " & prn
    Err.Clear
    On Error GoTo 0
End Function


'--- "PDFCreator on Ne00:" - the name Excel needs, "" if there is none -------
Public Function CM_PdfCreatorPrinter() As String
    Dim wmi As Object, col As Object, p As Object
    On Error Resume Next
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    If wmi Is Nothing Then
        Err.Clear: On Error GoTo 0
        Exit Function
    End If
    Set col = wmi.ExecQuery("SELECT Name, PortName FROM Win32_Printer")
    For Each p In col
        If InStr(1, CStr(p.Name), "PDFCreator", vbTextCompare) > 0 Then
            CM_PdfCreatorPrinter = CStr(p.Name) & " on " & CStr(p.PortName)
            Exit For
        End If
    Next
    Err.Clear
    On Error GoTo 0
End Function


'--- how many files are sitting in a folder ----------------------------------
Private Function CM_CountFiles(ByVal fso As Object, ByVal folder As String) As Long
    Dim f As Object
    On Error Resume Next
    If Not fso.FolderExists(folder) Then Exit Function
    For Each f In fso.GetFolder(folder).Files
        CM_CountFiles = CM_CountFiles + 1
    Next
    Err.Clear
    On Error GoTo 0
End Function


'--- file size in whole KB, -1 if it is not there ----------------------------
Private Function CM_KB(ByVal fso As Object, ByVal p As String) As Long
    On Error Resume Next
    CM_KB = -1
    CM_KB = CLng(fso.GetFile(p).Size \ 1024)
    Err.Clear
    On Error GoTo 0
End Function


'--- bounded wait used by the rehearsal; "" if nothing arrived ---------------
Private Function CM_WaitForFile(ByVal fso As Object, ByVal folder As String, _
                                ByVal maxSecs As Long, ByVal what As String) As String
    Dim base As String, nm As String, waited As Long
    Dim s1 As Double, s2 As Double

    base = folder
    If Right$(base, 1) <> "\" Then base = base & "\"

    Do
        nm = Dir(base & "*.*")
        If nm <> "" Then Exit Do
        If waited >= maxSecs Then Exit Function
        CM_Tick "waiting for " & what & " (" & waited & "/" & maxSecs & "s)"
        waited = waited + 1
    Loop

    'let it finish being written
    Do
        s1 = -1: s2 = -1
        On Error Resume Next
        s1 = fso.GetFile(base & nm).Size
        On Error GoTo 0
        CM_Tick "writing " & what & " (" & waited & "s)"
        On Error Resume Next
        s2 = fso.GetFile(base & nm).Size
        Err.Clear
        On Error GoTo 0
        If s1 >= 0 And s1 = s2 And s1 > 0 Then Exit Do
        nm = Dir(base & "*.*")
        If nm = "" Then Exit Function
        waited = waited + 1
        If waited >= maxSecs Then Exit Do
    Loop

    CM_WaitForFile = base & nm
End Function
'''

io.open(P, 'w', encoding='utf-8', newline='').write(env)
print('preflight: print + merge rehearsal added')
