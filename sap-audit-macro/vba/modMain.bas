Attribute VB_Name = "modMain"
'=======================================================================
' modMain -- entry points.
'
'   RunExtract        the whole job, all 56 samples, ten FEBAN periods
'   RunSingleMonth    one month only, for testing
'   CheckSetup        verify config and connection without touching data
'
' Start with CheckSetup, then a full RunExtract with Run mode = DRY RUN,
' and only then switch the Control sheet to EXTRACT.
'=======================================================================
Option Explicit

' Samples sheet columns
Private Const COL_IDX As Long = 1
Private Const COL_MONTH As Long = 2
Private Const COL_DATE_FROM As Long = 3
Private Const COL_DATE_TO As Long = 4
Private Const COL_PAY_DATE As Long = 5
Private Const COL_AMOUNT As Long = 6
Private Const COL_PARTY As Long = 7
Private Const COL_REF As Long = 8
Private Const COL_FLAGS As Long = 9
Private Const COL_STATUS As Long = 10
Private Const COL_STMT_ITEM As Long = 11
Private Const COL_FI_DOC As Long = 12
Private Const COL_INVOICES As Long = 13
Private Const COL_FILES As Long = 14
Private Const COL_MESSAGE As Long = 15
Private Const COL_INCLUDE As Long = 16

Private Type Sample
    Row As Long
    Idx As Long
    MonthTab As String
    DateFrom As Date
    DateTo As Date
    PayDate As Date
    Amount As Double
    Party As String
    Reference As String
End Type

'-----------------------------------------------------------------------
Public Sub CheckSetup()
    Dim report As String

    On Error GoTo Failed

    modConfig.LoadScreenMap
    modConfig.AssertScreenMapComplete
    modSapConnect.SapAttach

    report = "Setup looks usable." & vbCrLf & vbCrLf & _
             "SAP system   : " & modSapConnect.gSystemId & _
                              "  client " & modSapConnect.gClient & vbCrLf & _
             "SAP user     : " & modSapConnect.gSapUser & vbCrLf & _
             "SAP release  : " & IIf(Len(modSapConnect.gSapRelease) > 0, _
                                     modSapConnect.gSapRelease, "(not reported)") & vbCrLf & _
             "Transaction  : " & modSapConnect.CurrentTransaction() & vbCrLf & _
             "Company code : " & modConfig.Setting("Company code") & vbCrLf & _
             "Statement tx : " & modConfig.Setting("Transaction for statement search") & vbCrLf & _
             "Download root: " & modConfig.Setting("Download root folder") & vbCrLf & _
             "Run mode     : " & modConfig.Setting("Run mode") & vbCrLf & _
             "Date format  : " & modConfig.Setting("SAP date format") & _
                              "  (today would be typed as " & modUtil.SapDate(Date) & ")" & _
             vbCrLf & vbCrLf & _
             "Check that date -- if it does not match what SAP expects from your user, " & _
             "FEBAN will search the wrong period and report nothing found."

    MsgBox report, vbInformation, "Check setup"
    Exit Sub

Failed:
    MsgBox "Setup is not ready." & vbCrLf & vbCrLf & Err.Description, _
           vbExclamation, "Check setup"
End Sub

'-----------------------------------------------------------------------
Public Sub RunExtract()
    Dim samples() As Sample
    Dim count As Long
    Dim processed As Long, errored As Long, files As Long
    Dim monthTabs As Object
    Dim monthKey As Variant
    Dim i As Long
    Dim confirmation As VbMsgBoxResult

    On Error GoTo Failed

    modConfig.LoadScreenMap
    modConfig.AssertScreenMapComplete

    count = LoadSamples(samples)
    If count = 0 Then
        MsgBox "No samples found on the '" & modConfig.SHEET_SAMPLES & "' sheet.", _
               vbExclamation
        Exit Sub
    End If

    If Not modConfig.IsDryRun() Then
        confirmation = MsgBox( _
            "Run mode is EXTRACT, so files will be written to:" & vbCrLf & vbCrLf & _
            "    " & modConfig.Setting("Download root folder") & vbCrLf & vbCrLf & _
            count & " samples across " & CountDistinctMonths(samples, count) & _
            " FEBAN periods." & vbCrLf & vbCrLf & _
            "The extract will contain vendor and payment data. Make sure that folder " & _
            "is somewhere approved for it." & vbCrLf & vbCrLf & "Continue?", _
            vbQuestion + vbYesNo + vbDefaultButton2, "Confirm extract")
        If confirmation <> vbYes Then Exit Sub
    End If

    modSapConnect.SapAttach
    modSafety.gWriteAttemptBlocked = 0
    modLog.WriteHeaderBlock

    Application.ScreenUpdating = False

    ' FEBAN selects a period, not a payment, so open each month once and
    ' walk the samples inside it.
    Set monthTabs = DistinctMonths(samples, count)

    For Each monthKey In monthTabs.Keys
        ' TryOpenMonth swallows and logs its own errors, so the loop needs no
        ' inline handler -- 'Resume' back into a For Each is fragile.
        If TryOpenMonth(CStr(monthKey), CDate(monthTabs(monthKey)(0)), _
                        CDate(monthTabs(monthKey)(1))) Then

            ' The statement list is per period, so export it once here rather
            ' than once per sample -- otherwise Sep 25 alone would write seven
            ' copies of the same list.
            If Len(ExportMonthStatementList(CStr(monthKey))) > 0 Then
                files = files + 1
            End If

            For i = 1 To count
                If samples(i).MonthTab = CStr(monthKey) Then
                    If ProcessSample(samples(i), files) Then
                        processed = processed + 1
                    Else
                        errored = errored + 1
                        If modConfig.SettingIsYes("Stop on first error") Then GoTo Finished
                    End If
                End If
            Next i
        Else
            errored = errored + 1
            If modConfig.SettingIsYes("Stop on first error") Then GoTo Finished
        End If
    Next monthKey

Finished:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.StatusBar = False
    modLog.WriteFooterBlock processed, errored, files
    modUtil.CloseExportWorkbooksUnder modConfig.DownloadRoot()
    On Error GoTo 0

    MsgBox "Run finished." & vbCrLf & vbCrLf & _
           "Mode              : " & modConfig.Setting("Run mode") & vbCrLf & _
           "Samples processed : " & processed & vbCrLf & _
           "Errors            : " & errored & vbCrLf & _
           "Files written     : " & files & vbCrLf & _
           "Writes blocked    : " & modSafety.gWriteAttemptBlocked & vbCrLf & vbCrLf & _
           IIf(errored > 0, "Where they stopped: " & modLog.FailureSummary() & vbCrLf & vbCrLf, "") & _
           "See the 'Log' sheet for the full trail.", _
           IIf(errored = 0, vbInformation, vbExclamation), "FEBAN audit extract"
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    modLog.LogAction 0, "RUN ABORTED", Err.Description, "ERROR", vbNullString
    MsgBox "The run stopped." & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Nothing was posted -- the guard blocks write actions. See the 'Log' sheet.", _
           vbCritical, "FEBAN audit extract"
End Sub

' The FEBAN result list for the month now on screen, as period context. Best
' effort -- it is not part of the evidence chain, so a failure here is logged
' and the samples run anyway.
Private Function ExportMonthStatementList(ByVal monthTab As String) As String
    Dim folder As String

    On Error GoTo Failed

    folder = modUtil.JoinPath(modConfig.DownloadRoot(), _
                              modUtil.SafeFileName(monthTab))

    ExportMonthStatementList = modExport.ExportAlvGrid( _
        0, folder, "00_" & modUtil.SafeFileName(monthTab) & "_FEBAN_statement.xlsx")
    Exit Function

Failed:
    modLog.LogAction 0, "Export statement list", _
                 "Could not export the " & monthTab & " statement list: " & _
                 Err.Description & ". The samples in this month are unaffected.", _
                 "SKIPPED", vbNullString
End Function

' Open one FEBAN period, logging rather than raising on failure so that a
' bad month does not take the whole run down with it.
Private Function TryOpenMonth(ByVal monthTab As String, ByVal dateFrom As Date, _
                              ByVal dateTo As Date) As Boolean
    On Error GoTo Failed

    modFeban.OpenMonth dateFrom, dateTo
    TryOpenMonth = True
    Exit Function

Failed:
    modLog.LogAction 0, "FEBAN selection", _
                 "Month " & monthTab & " failed: " & Err.Description, _
                 "ERROR", vbNullString
End Function

'-----------------------------------------------------------------------
' One sample, end to end. Returns False on a handled failure.
'-----------------------------------------------------------------------
Private Function ProcessSample(ByRef item As Sample, ByRef filesTotal As Long) As Boolean
    Dim sheet As Worksheet
    Dim match As FebanMatch
    Dim chain As ChainResult
    Dim folder As String
    Dim stem As String
    Dim message As String

    Set sheet = ThisWorkbook.Worksheets(modConfig.SHEET_SAMPLES)

    On Error GoTo SampleFailed

    ' SAP opens each spreadsheet export in Excel once it has written it, and
    ' it does that asynchronously -- so the previous sample's export tends to
    ' surface during this one. Close anything sitting under the evidence
    ' folder before starting, or they stack up all run.
    modUtil.CloseExportWorkbooksUnder modConfig.DownloadRoot()

    ' Get back to the statement list BEFORE anything else. The previous sample
    ' may have ended several screens deep, or in FBL1N, or on an error -- and
    ' this used to sit after Exit Function, where it never ran at all. That is
    ' why one failure inside FBL1N took every later sample down with it: they
    ' looked for the FEBAN grid while the session was still in FBL1N.
    If Not modFeban.ReturnToStatementList(item.DateFrom, item.DateTo) Then
        Err.Raise vbObjectError + 550, "modMain.ProcessSample", _
                  "Could not get back to the FEBAN statement list before sample " & _
                  item.Idx & ". Stopping so that it is not read off the wrong screen."
    End If

    match = modFeban.FindSample(item.PayDate, item.Amount)

    If Not match.Found Then
        message = "No statement line in " & item.MonthTab & " matches " & _
                  Format$(item.PayDate, "dd/mm/yyyy") & " for " & _
                  Format$(item.Amount, "#,##0.00") & ". " & _
                  modFeban.WhyNoMatch(item.PayDate, item.Amount)
        WriteResult sheet, item.Row, "NOT FOUND", vbNullString, vbNullString, _
                    vbNullString, 0, message
        modLog.LogAction item.Idx, "Match", message, "ERROR", vbNullString
        Exit Function
    End If

    If match.Ambiguous Then
        message = match.CandidateCount & " statement lines share that date and amount. " & _
                  "Resolve by hand -- the macro will not guess which one the auditor means."
        WriteResult sheet, item.Row, "AMBIGUOUS", vbNullString, vbNullString, _
                    vbNullString, 0, message
        modLog.LogAction item.Idx, "Match", message, "ERROR", vbNullString
        Exit Function
    End If

    modLog.LogAction item.Idx, "Match", _
                 "Statement row " & match.GridRow & ", amount " & _
                 Format$(match.StatementAmount, "#,##0.00") & _
                 IIf(Len(match.PostingStatus) > 0, ", status " & match.PostingStatus, ""), _
                 "OK", vbNullString

    modFeban.SelectRow match.GridRow

    folder = modUtil.JoinPath(modConfig.DownloadRoot(), _
                              modUtil.SafeFileName(item.MonthTab))
    stem = modUtil.EvidenceStem(item.Idx, item.PayDate, item.Amount)

    ' statement item -> FI document -> clearing document -> cleared items with
    ' supplier names -> the invoice for the largest of them. The period's
    ' statement list was already exported once, back in RunExtract.
    chain = modChain.Walk(item.Idx, match, item.DateFrom, item.DateTo, folder, stem)

    filesTotal = filesTotal + CountFiles(chain)

    WriteResult sheet, item.Row, chain.Status, _
                match.StatementDate & " row " & match.GridRow, _
                chain.FiDocument, InvoiceTrail(chain), _
                CountFiles(chain), chain.Notes

    ' Only DONE counts as reaching the end. Anything else stopped somewhere,
    ' and reporting it as processed made "Errors: 0" true while seven chains
    ' had in fact failed.
    ProcessSample = (chain.Status = "DONE")
    Exit Function

SampleFailed:
    message = Err.Description
    WriteResult sheet, item.Row, "ERROR", vbNullString, vbNullString, vbNullString, 0, message
    modLog.LogAction item.Idx, "Sample failed", message, "ERROR", vbNullString

    ' Try to get back to a known screen so the next sample has a chance.
    On Error Resume Next
    modSafety.GuardedOkCode "/n" & modConfig.Setting("Transaction for statement search")
    On Error GoTo 0
End Function

Private Function CountFiles(ByRef chain As ChainResult) As Long
    If Len(chain.ZpListFile) > 0 Then CountFiles = CountFiles + 1
    If Len(chain.ZpExportFile) > 0 Then CountFiles = CountFiles + 1
    If Len(chain.InvoiceListFile) > 0 Then CountFiles = CountFiles + 1
    If Len(chain.InvoicePdfFile) > 0 Then CountFiles = CountFiles + 1
End Function

' The document trail, so the Samples sheet shows how the invoice was reached
' rather than only which one it landed on.
Private Function InvoiceTrail(ByRef chain As ChainResult) As String
    Dim trail As String

    If Len(chain.ClearingDocument) > 0 Then
        trail = "clearing " & chain.ClearingDocument
    End If

    If chain.ZpNumberCount > 0 Then
        trail = trail & IIf(Len(trail) > 0, " > ", "") & _
                chain.ZpNumberCount & " ZP"
    End If

    If Len(chain.ZpPaymentDocument) > 0 Then
        trail = trail & " > pmt " & chain.ZpPaymentDocument
    End If

    If Len(chain.InvoiceNumber) > 0 And chain.InvoiceNumber <> chain.ZpPaymentDocument Then
        trail = trail & " > inv " & chain.InvoiceNumber
    End If

    InvoiceTrail = trail
End Function

Private Sub WriteResult(ByVal sheet As Worksheet, ByVal row As Long, ByVal status As String, _
                        ByVal statementItem As String, ByVal fiDoc As String, _
                        ByVal invoices As String, ByVal fileCount As Long, _
                        ByVal message As String)
    sheet.Cells(row, COL_STATUS).Value = status
    sheet.Cells(row, COL_STMT_ITEM).Value = statementItem
    sheet.Cells(row, COL_FI_DOC).Value = fiDoc
    sheet.Cells(row, COL_INVOICES).Value = invoices
    sheet.Cells(row, COL_FILES).Value = fileCount
    sheet.Cells(row, COL_MESSAGE).Value = message

    Select Case status
        Case "DONE"
            sheet.Cells(row, COL_STATUS).Font.Color = RGB(0, 112, 48)
        Case "ERROR", "NOT FOUND", "AMBIGUOUS"
            sheet.Cells(row, COL_STATUS).Font.Color = RGB(192, 0, 0)
        Case Else
            ' PARTIAL, UNPOSTED, BLOCKED_SCF -- work left to do, not failures
            sheet.Cells(row, COL_STATUS).Font.Color = RGB(128, 96, 0)
    End Select
    sheet.Cells(row, COL_STATUS).Font.Bold = True
End Sub

'-----------------------------------------------------------------------
' Samples sheet -> array
'-----------------------------------------------------------------------
Private Function LoadSamples(ByRef samples() As Sample) As Long
    Dim sheet As Worksheet
    Dim row As Long, lastRow As Long
    Dim count As Long

    Set sheet = ThisWorkbook.Worksheets(modConfig.SHEET_SAMPLES)
    lastRow = sheet.Cells(sheet.Rows.Count, COL_IDX).End(xlUp).Row

    ReDim samples(1 To Application.Max(1, lastRow))

    For row = modConfig.SAMPLES_FIRST_ROW To lastRow
        If IsNumeric(sheet.Cells(row, COL_IDX).Value) And _
           IsDate(sheet.Cells(row, COL_PAY_DATE).Value) And _
           IsIncluded(sheet, row) Then

            count = count + 1
            samples(count).Row = row
            samples(count).Idx = CLng(sheet.Cells(row, COL_IDX).Value)
            samples(count).MonthTab = Trim$(CStr(sheet.Cells(row, COL_MONTH).Value))
            samples(count).PayDate = CDate(sheet.Cells(row, COL_PAY_DATE).Value)
            samples(count).Amount = CDbl(sheet.Cells(row, COL_AMOUNT).Value)
            samples(count).Party = Trim$(CStr(sheet.Cells(row, COL_PARTY).Value))
            samples(count).Reference = Trim$(CStr(sheet.Cells(row, COL_REF).Value))

            ' Columns C and D are formulas over the payment date. Fall back to
            ' deriving the range here if they have been cleared.
            If IsDate(sheet.Cells(row, COL_DATE_FROM).Value) Then
                samples(count).DateFrom = CDate(sheet.Cells(row, COL_DATE_FROM).Value)
            Else
                samples(count).DateFrom = modUtil.MonthStart(samples(count).PayDate)
            End If

            If IsDate(sheet.Cells(row, COL_DATE_TO).Value) Then
                samples(count).DateTo = CDate(sheet.Cells(row, COL_DATE_TO).Value)
            Else
                samples(count).DateTo = modUtil.MonthEnd(samples(count).PayDate)
            End If

            If Len(samples(count).MonthTab) = 0 Then
                samples(count).MonthTab = modUtil.MonthTabName(samples(count).PayDate)
            End If
        End If
    Next row

    LoadSamples = count
End Function

' Blank counts as included, so a freshly generated sheet runs everything and
' the operator only has to touch the rows they want left out.
Private Function IsIncluded(ByVal sheet As Worksheet, ByVal row As Long) As Boolean
    Dim flag As String

    flag = UCase$(Trim$(CStr(sheet.Cells(row, COL_INCLUDE).Value)))
    IsIncluded = (flag <> "NO" And flag <> "N" And flag <> "FALSE" And flag <> "0")
End Function

' month tab -> Array(dateFrom, dateTo), in first-seen order
Private Function DistinctMonths(ByRef samples() As Sample, ByVal count As Long) As Object
    Dim months As Object
    Dim i As Long

    Set months = CreateObject("Scripting.Dictionary")
    months.CompareMode = vbTextCompare

    For i = 1 To count
        If Not months.Exists(samples(i).MonthTab) Then
            months.Add samples(i).MonthTab, Array(samples(i).DateFrom, samples(i).DateTo)
        End If
    Next i

    Set DistinctMonths = months
End Function

Private Function CountDistinctMonths(ByRef samples() As Sample, ByVal count As Long) As Long
    CountDistinctMonths = DistinctMonths(samples, count).Count
End Function

'-----------------------------------------------------------------------
' Test one month without committing to the whole list.
'-----------------------------------------------------------------------
Public Sub RunSingleMonth()
    Dim answer As String
    Dim samples() As Sample
    Dim count As Long, i As Long
    Dim files As Long
    Dim matched As Boolean
    Dim done As Long, failed As Long

    answer = InputBox("Which month tab? For example: Sep 25", "Run one month", "Sep 25")
    If Len(Trim$(answer)) = 0 Then Exit Sub

    On Error GoTo Failed

    modConfig.LoadScreenMap
    modConfig.AssertScreenMapComplete
    count = LoadSamples(samples)

    modSapConnect.SapAttach
    modSafety.gWriteAttemptBlocked = 0
    modLog.WriteHeaderBlock

    For i = 1 To count
        If StrComp(samples(i).MonthTab, Trim$(answer), vbTextCompare) = 0 Then
            If Not matched Then
                modFeban.OpenMonth samples(i).DateFrom, samples(i).DateTo
                matched = True
            End If
            If ProcessSample(samples(i), files) Then
                done = done + 1
            Else
                failed = failed + 1
            End If
        End If
    Next i

    If Not matched Then
        MsgBox "No samples on the '" & modConfig.SHEET_SAMPLES & "' sheet for month tab '" & _
               answer & "'.", vbExclamation
        Exit Sub
    End If

    modLog.WriteFooterBlock done, failed, files
    modUtil.CloseExportWorkbooksUnder modConfig.DownloadRoot()

    ' Say where they stopped, read off the Log, rather than assuming. The old
    ' wording claimed none got past the statement match, which was wrong the
    ' moment a sample reached FBL1N and exported its batch list.
    MsgBox "Finished " & answer & "." & vbCrLf & vbCrLf & _
           "Samples reaching the end of the chain : " & done & vbCrLf & _
           "Samples that stopped early            : " & failed & vbCrLf & _
           "Files written                         : " & files & vbCrLf & vbCrLf & _
           IIf(failed > 0, "Where they stopped: " & modLog.FailureSummary() & vbCrLf & vbCrLf, "") & _
           "Files are under " & modConfig.DownloadRoot() & ".", _
           IIf(failed = 0, vbInformation, vbExclamation), "Run one month"
    Exit Sub

Failed:
    modLog.LogAction 0, "RUN ABORTED", Err.Description, "ERROR", vbNullString
    MsgBox "Stopped: " & Err.Description, vbCritical
End Sub
