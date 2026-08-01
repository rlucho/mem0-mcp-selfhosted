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
' Added when the macro stopped being about one audit request. Appended
' rather than inserted, so the Control sheet's formulas over columns A to P
' keep pointing at the same things.
Private Const COL_REQUEST As Long = 17
Private Const COL_COMPANY As Long = 18
Private Const COL_COMMENT As Long = 19
' Set when the row came from a SAP document list rather than a bank
' statement: the document to enter the chain at, and which rung that is.
Private Const COL_START_DOC As Long = 20
Private Const COL_START_AT As Long = 21
' The ZP the auditor picked, when their file showed its working. Nothing
' selects on it -- it exists to be disagreed with.
Private Const COL_EXPECTED_ZP As Long = 22

' The period's statement export, so each sample folder can be given a copy.
Private mMonthStatementFile As String

' Samples that finished with something less than a complete answer -- PARTIAL,
' NOT FOUND, AMBIGUOUS, BLOCKED_*. They no longer stop the run, so the summary
' has to say how many there were, or a run that answered half its samples reads
' as a clean sweep.
Private mIncomplete As Long

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
    Request As String          ' which audit request this row came from
    CompanyCode As String      ' asked once per request at import
    Comment As String          ' what the auditor asked for, in their words
    StartDocument As String    ' blank for a normal statement-driven sample
    StartAt As String          ' "", "CLEARING" or "PAYMENT"
    ExpectedZp As String       ' the auditor's own pick, to check ours against
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
    Dim remoteWas As Boolean

    On Error GoTo Failed

    modConfig.LoadScreenMap
    modConfig.AssertScreenMapComplete

    ' Scope first: every sample, or only the rows marked Include? The run
    ' works out which FEBAN periods it needs and opens them itself -- the
    ' operator picks samples, never months.
    confirmation = MsgBox(ScopeSummary(), vbQuestion + vbYesNoCancel, "Run samples")
    If confirmation = vbCancel Then Exit Sub

    count = LoadSamplesWhere(samples, confirmation = vbNo)

    If count = 0 Then
        MsgBox "Nothing to run." & vbCrLf & vbCrLf & _
               "Either the '" & modConfig.SHEET_SAMPLES & "' sheet is empty -- use " & _
               "'Import request' to add an auditor's file -- or every row is marked " & _
               "Include? = No.", vbExclamation, "Run samples"
        Exit Sub
    End If

    If Not modConfig.IsDryRun() Then
        confirmation = MsgBox( _
            "Run mode is EXTRACT, so files will be written to:" & vbCrLf & vbCrLf & _
            "    " & modConfig.Setting("Download root folder") & vbCrLf & vbCrLf & _
            count & " samples across " & CountDistinctMonths(samples, count) & _
            " FEBAN period(s)." & vbCrLf & vbCrLf & _
            "The extract will contain vendor and payment data. Make sure that folder " & _
            "is somewhere approved for it." & vbCrLf & vbCrLf & "Continue?", _
            vbQuestion + vbYesNo + vbDefaultButton2, "Confirm extract")
        If confirmation <> vbYes Then Exit Sub
    End If

    modSapConnect.SapAttach
    modSafety.gWriteAttemptBlocked = 0
    mIncomplete = 0
    modLog.WriteHeaderBlock

    ' SAP opens every export it writes. Refuse the request for the duration
    ' of the run -- the macro opens the file itself when it wants to read it.
    remoteWas = modUtil.SuppressRemoteOpen()

    ' And because refusing does not work on SAP's route in, every export is
    ' written to a scratch name first. Start from an empty scratch folder, in
    ' case a previous run was killed with files in it.
    modExport.SweepHandover 0

    Application.ScreenUpdating = False

    ' FEBAN selects a period, not a payment, so open each month once and
    ' walk the samples inside it.
    Set monthTabs = DistinctMonths(samples, count)

    For Each monthKey In monthTabs.Keys
        ' TryOpenMonth swallows and logs its own errors, so the loop needs no
        ' inline handler -- 'Resume' back into a For Each is fragile.
        ' The company code of this group has to be in force before FEBAN is
        ' opened, not just before the samples inside it are walked.
        modConfig.SetCompanyCode CStr(monthTabs(monthKey)(4))

        ' Rows that name their own starting document do not need a statement
        ' search at all -- FEBAN exists to find what they already carry.
        If Len(CStr(monthTabs(monthKey)(5))) > 0 Then
            For i = 1 To count
                If PeriodKey(samples(i)) = CStr(monthKey) Then
                    If ProcessSample(samples(i), files) Then
                        processed = processed + 1
                    Else
                        errored = errored + 1
                        If modConfig.SettingIsYes("Stop on first error") Then GoTo Finished
                    End If
                End If
            Next i
        ElseIf TryOpenMonth(CStr(monthTabs(monthKey)(2)), CDate(monthTabs(monthKey)(0)), _
                        CDate(monthTabs(monthKey)(1))) Then

            ' The statement list is per period, so export it once here rather
            ' than once per sample -- otherwise Sep 25 alone would write seven
            ' copies of the same list.
            If Len(ExportMonthStatementList(CStr(monthTabs(monthKey)(2)), _
                                            CStr(monthTabs(monthKey)(3)), _
                                            CStr(monthTabs(monthKey)(4)))) > 0 Then
                files = files + 1
            End If

            For i = 1 To count
                If PeriodKey(samples(i)) = CStr(monthKey) Then
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
    modUtil.RestoreRemoteOpen remoteWas
    Application.ScreenUpdating = True
    Application.StatusBar = False
    modLog.WriteFooterBlock processed, errored, files
    modUtil.CloseExportWorkbooksUnder modConfig.DownloadRoot()
    modExport.SweepHandover modExport.HANDOVER_GRACE_SECONDS
    On Error GoTo 0

    MsgBox "Run finished." & vbCrLf & vbCrLf & _
           "Mode              : " & modConfig.Setting("Run mode") & vbCrLf & _
           "Samples processed : " & processed & vbCrLf & _
           "  of those, short  : " & mIncomplete & _
                IIf(mIncomplete > 0, "  (see the Status column)", "") & vbCrLf & _
           "Errors            : " & errored & vbCrLf & _
           "Files written     : " & files & vbCrLf & _
           "Writes blocked    : " & modSafety.gWriteAttemptBlocked & vbCrLf & vbCrLf & _
           IIf(errored > 0, "Where they stopped: " & modLog.FailureSummary() & vbCrLf & vbCrLf, "") & _
           "See the 'Log' sheet for the full trail.", _
           IIf(errored = 0, vbInformation, vbExclamation), "FEBAN audit extract"
    Exit Sub

Failed:
    modUtil.RestoreRemoteOpen remoteWas
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
Private Function ExportMonthStatementList(ByVal monthTab As String, _
                                          ByVal request As String, _
                                          ByVal companyCode As String) As String
    Dim folder As String

    On Error GoTo Failed

    mMonthStatementFile = vbNullString

    folder = modUtil.JoinPath( _
                 modUtil.JoinPath(modConfig.DownloadRoot(), _
                                  modUtil.RequestFolderName(request, companyCode)), _
                 modUtil.SafeFileName(monthTab))

    ' The sample folders are cleared before each sample; this one was not,
    ' so three runs of Sep 25 left '1 - FEBAN statement list.xlsx', '_2' and
    ' '_3' side by side -- the export never overwrites. Only the numbered
    ' files go; the sample subfolders below are untouched.
    modUtil.ClearSampleFolder folder

    ExportMonthStatementList = modExport.ExportAlvGrid(0, folder, modUtil.FILE_FEBAN)
    mMonthStatementFile = ExportMonthStatementList
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
' DONE has its invoice. NO CLEARING is an internal transfer that settles
' nothing, NO VENDOR PAYMENTS a treasury or FX settlement that settles against
' the bank statement rather than a supplier -- both have no invoice to fetch
' and never did. Everything else stopped somewhere short.
'-----------------------------------------------------------------------
' Compare the payment this run picked with the one the auditor picked.
'
' A follow-up request shows its working: under each sample the auditor pastes
' the SAP extract they took, naming the ZP they treated as the largest payment
' of the batch. That is the same question this macro answers independently --
' export the batch from FBL1N, take the largest by value -- so the two can be
' held against each other.
'
' Agreement is the useful thing to be able to state: the pack is not just
' evidence, it is evidence that reaches the same document the auditor did.
' Disagreement is more useful still, and must be loud rather than silent --
' either the batch changed since they looked, or one of the two is reading a
' different batch, and nobody would spot it by eye across 15 samples.
'
' Only ever annotates. Nothing about the run changes on the outcome, because
' the auditor's pick is a second opinion, not an instruction.
'-----------------------------------------------------------------------
Private Sub CrossCheckAgainstAuditor(ByRef item As Sample, ByRef chain As ChainResult)
    Dim expected As String, actual As String
    Dim verdict As String

    expected = OnlyDigits(item.ExpectedZp)
    If Len(expected) = 0 Then Exit Sub

    actual = OnlyDigits(chain.ZpPaymentDocument)

    If Len(actual) = 0 Then
        verdict = "The auditor's file names ZP " & item.ExpectedZp & " as the largest " & _
                  "payment of this batch. This run did not reach a payment, so there is " & _
                  "nothing to compare it with."
        modLog.LogAction item.Idx, "Cross-check", verdict, "MANUAL", vbNullString
    ElseIf StrComp(expected, actual, vbTextCompare) = 0 Then
        verdict = "Agrees with the auditor: both reached ZP " & chain.ZpPaymentDocument & _
                  " as the largest payment of the batch."
        modLog.LogAction item.Idx, "Cross-check", verdict, "OK", vbNullString
    Else
        verdict = "DIFFERS from the auditor: their file names ZP " & item.ExpectedZp & _
                  " as the largest payment of this batch, this run reached ZP " & _
                  chain.ZpPaymentDocument & ". Both came from the same clearing " & _
                  "document, so one of them is reading a different batch -- worth " & _
                  "settling before the pack goes back."
        modLog.LogAction item.Idx, "Cross-check", verdict, "MANUAL", vbNullString
    End If

    chain.Notes = chain.Notes & IIf(Len(chain.Notes) > 0, " ", "") & verdict
End Sub

Private Function OnlyDigits(ByVal text As String) As String
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch >= "0" And ch <= "9" Then OnlyDigits = OnlyDigits & ch
    Next i
End Function

Private Function IsCompleteAnswer(ByVal status As String) As Boolean
    IsCompleteAnswer = (status = "DONE" Or status = "NO CLEARING" Or _
                        status = "NO VENDOR PAYMENTS")
End Function

Private Function ProcessSample(ByRef item As Sample, ByRef filesTotal As Long) As Boolean
    Dim sheet As Worksheet
    Dim match As FebanMatch
    Dim chain As ChainResult
    Dim folder As String
    Dim stem As String
    Dim message As String
    Dim cleared As Long

    Set sheet = ThisWorkbook.Worksheets(modConfig.SHEET_SAMPLES)

    On Error GoTo SampleFailed

    ' The company code travels with the sample now, so put it in force before
    ' anything reads it -- FEBAN's selection, FBL1N's, the report heading and
    ' the folder name all go through modConfig.CompanyCode().
    modConfig.SetCompanyCode item.CompanyCode

    ' SAP opens each spreadsheet export in Excel once it has written it, and
    ' it does that asynchronously -- so the previous sample's export tends to
    ' surface during this one. Close anything sitting under the evidence
    ' folder before starting, or they stack up all run. The scratch copies
    ' SAP actually opens are swept the same way, and deleted.
    modUtil.CloseExportWorkbooksUnder modConfig.DownloadRoot()
    modExport.CloseHandoverWorkbooks

    ' A row that names its own starting document skips the statement search
    ' entirely -- there is nothing to match, because the document it would
    ' have been looking for is already in the cell.
    If Len(item.StartAt) > 0 Then
        ProcessSample = ProcessDirectSample(item, filesTotal, sheet)
        Exit Function
    End If

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

    ' NOT FOUND and AMBIGUOUS are answers about the data, not failures of the
    ' macro, so both return True and the run carries on. They used to fall
    ' through to the error exit, which meant one line the auditor asked about
    ' that is not in FEBAN stopped everything queued behind it -- and with
    ' 'Stop on first error' set to YES that is the whole run. A sample whose
    ' payment sits on a bank account this company code's statements do not
    ' carry is a finding to report, in the same way NO CLEARING is.
    '
    ' Logged as MANUAL rather than ERROR for the same reason: a human has to
    ' look at it, but nothing went wrong on the screen. WhyNoMatch is what
    ' tells the two apart -- '1971 rows read, 71 on that date, none for that
    ' amount' is a real miss, while '0 rows on that date' means the period or
    ' the company code is wrong and every sample will say the same.
    If Not match.Found Then
        message = "No statement line in " & item.MonthTab & " matches " & _
                  Format$(item.PayDate, "dd/mm/yyyy") & " for " & _
                  Format$(item.Amount, "#,##0.00") & ". " & _
                  modFeban.WhyNoMatch(item.PayDate, item.Amount)
        WriteResult sheet, item.Row, "NOT FOUND", vbNullString, vbNullString, _
                    vbNullString, 0, message
        modLog.LogAction item.Idx, "Match", message, "MANUAL", vbNullString
        mIncomplete = mIncomplete + 1
        ProcessSample = True
        Exit Function
    End If

    If match.Ambiguous Then
        message = match.CandidateCount & " statement lines share that date and amount. " & _
                  "Resolve by hand -- the macro will not guess which one the auditor means."
        WriteResult sheet, item.Row, "AMBIGUOUS", vbNullString, vbNullString, _
                    vbNullString, 0, message
        modLog.LogAction item.Idx, "Match", message, "MANUAL", vbNullString
        mIncomplete = mIncomplete + 1
        ProcessSample = True
        Exit Function
    End If

    modLog.LogAction item.Idx, "Match", _
                 "Statement row " & match.GridRow & ", amount " & _
                 Format$(match.StatementAmount, "#,##0.00") & _
                 IIf(Len(match.PostingStatus) > 0, ", status " & match.PostingStatus, ""), _
                 "OK", vbNullString

    modFeban.SelectRow match.GridRow

    ' One folder per sample, inside one folder per month, inside one folder
    ' per audit request. The request folder carries the company code in its
    ' name, because several requests are in flight at once and they are not
    ' all the same company.
    folder = modUtil.JoinPath( _
                 modUtil.JoinPath( _
                     modUtil.JoinPath(modConfig.DownloadRoot(), _
                                      modUtil.RequestFolderName(item.Request, _
                                                                item.CompanyCode)), _
                     modUtil.SafeFileName(item.MonthTab)), _
                 modUtil.SampleFolderName(item.Idx, item.Amount))
    stem = modUtil.EvidenceStem(item.Idx, item.PayDate, item.Amount)

    ' A re-run starts this sample's folder clean, or the exports -- which
    ' never overwrite -- would leave '..._2.xlsx' beside every file and the
    ' pack would go to the auditor full of near duplicates.
    cleared = modUtil.ClearSampleFolder(folder)
    If cleared > 0 Then
        modLog.LogAction item.Idx, "Folder", _
                     "Removed " & cleared & " file(s) from an earlier run of this " & _
                     "sample before re-exporting into " & folder & ".", _
                     "OK", vbNullString
    End If

    ' statement item -> FI document -> clearing document -> cleared items with
    ' supplier names -> the invoice for the largest of them. The period's
    ' statement list was already exported once, back in RunExtract.
    chain = modChain.Walk(item.Idx, match, item.DateFrom, item.DateTo, folder, stem)

    ' The statement list is exported once per period, not once per sample, but
    ' each pack needs its own copy or it is not self-contained.
    CopyMonthStatementInto folder

    filesTotal = filesTotal + CountFiles(chain)

    WriteResult sheet, item.Row, chain.Status, _
                match.StatementDate & " row " & match.GridRow, _
                chain.FiDocument, InvoiceTrail(chain), _
                CountFiles(chain), chain.Notes

    WriteSampleReport item, match, chain, folder

    ' Every status the chain can name is a RESULT for this sample, so the run
    ' carries on to the next one. Only 'ERROR' -- raised below, when the macro
    ' does not recognise the screen it is on -- means it is no longer safe to
    ' continue, and only that trips 'Stop on first error'.
    '
    ' This used to answer True for DONE, NO CLEARING and NO VENDOR PAYMENTS
    ' only, so a single PARTIAL stopped everything queued behind it. One
    ' sample whose three ZP documents were not vendor items in that period
    ' ended a run with twelve samples still to go. A sample that did not reach
    ' an invoice is a finding to read on the sheet, not a reason to abandon
    ' the other eleven.
    CrossCheckAgainstAuditor item, chain

    If Not IsCompleteAnswer(chain.Status) Then mIncomplete = mIncomplete + 1
    ProcessSample = (chain.Status <> "ERROR")
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

' Give this sample's folder its own copy of the period's statement list, so
' the folder stands alone. Never fatal -- a missing copy costs context, not
' evidence, and the original is still one level up.
Private Sub CopyMonthStatementInto(ByVal folder As String)
    Dim fso As Object
    Dim target As String

    If Len(mMonthStatementFile) = 0 Then Exit Sub
    If Not modUtil.FileExists(mMonthStatementFile) Then Exit Sub

    target = modUtil.JoinPath(folder, modUtil.FILE_FEBAN)
    If modUtil.FileExists(target) Then Exit Sub

    On Error Resume Next
    modUtil.EnsureFolder folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile mMonthStatementFile, target, True
    On Error GoTo 0
End Sub

' The per-sample report. Never fatal: the evidence files are already on disk
' by this point, and a report that will not build must not undo them.
Private Sub WriteSampleReport(ByRef item As Sample, ByRef match As FebanMatch, _
                              ByRef chain As ChainResult, ByVal folder As String)
    Dim path As String

    On Error GoTo Failed

    path = modReport.BuildSampleReport(item.Idx, item.MonthTab, item.PayDate, item.Amount, _
                                       item.Party, item.Reference, match, chain, folder)

    modLog.LogAction item.Idx, "Report", _
                 "Wrote " & modUtil.ReportFileName(item.Idx, item.Amount, _
                                                   modConfig.CompanyCode()), _
                 "OK", path
    Exit Sub

Failed:
    modLog.LogAction item.Idx, "Report", _
                 "The sample report could not be written: " & Err.Description & _
                 ". The evidence files in " & folder & " are unaffected.", _
                 "SKIPPED", vbNullString
End Sub

'-----------------------------------------------------------------------
' A sample imported from a SAP document list rather than a bank statement.
'
' No FEBAN, no date-and-amount match, no drill through the FI document --
' the row already names the document, so the chain is entered at whichever
' rung that document sits on.
'-----------------------------------------------------------------------
Private Function ProcessDirectSample(ByRef item As Sample, ByRef filesTotal As Long, _
                                     ByVal sheet As Worksheet) As Boolean
    Dim chain As ChainResult
    Dim noMatch As FebanMatch     ' never populated: there was no statement search
    Dim folder As String
    Dim stem As String
    Dim cleared As Long

    On Error GoTo DirectFailed

    modConfig.SetCompanyCode item.CompanyCode
    modUtil.CloseExportWorkbooksUnder modConfig.DownloadRoot()
    modExport.CloseHandoverWorkbooks

    folder = modUtil.JoinPath( _
                 modUtil.JoinPath( _
                     modUtil.JoinPath(modConfig.DownloadRoot(), _
                                      modUtil.RequestFolderName(item.Request, _
                                                                item.CompanyCode)), _
                     modUtil.SafeFileName(item.MonthTab)), _
                 modUtil.SampleFolderName(item.Idx, item.Amount))
    stem = modUtil.EvidenceStem(item.Idx, item.PayDate, item.Amount)

    cleared = modUtil.ClearSampleFolder(folder)
    If cleared > 0 Then
        modLog.LogAction item.Idx, "Folder", _
                     "Removed " & cleared & " file(s) from an earlier run of this sample.", _
                     "OK", vbNullString
    End If

    Select Case item.StartAt
        Case "CLEARING"
            chain = modChain.WalkFromClearing(item.Idx, item.StartDocument, _
                                              item.DateFrom, item.DateTo, folder, stem)
        Case "PAYMENT"
            chain = modChain.WalkFromPayment(item.Idx, item.StartDocument, item.Party, _
                                             item.DateFrom, item.DateTo, folder, stem)
        Case Else
            Err.Raise vbObjectError + 590, "modMain.ProcessDirectSample", _
                      "Column U of row " & item.Row & " reads '" & item.StartAt & "'. " & _
                      "It must be CLEARING, PAYMENT, or blank for a normal " & _
                      "statement-driven sample."
    End Select

    filesTotal = filesTotal + CountFiles(chain)

    WriteResult sheet, item.Row, chain.Status, _
                "from " & item.StartAt & " " & item.StartDocument, _
                chain.FiDocument, InvoiceTrail(chain), _
                CountFiles(chain), chain.Notes

    WriteSampleReport item, noMatch, chain, folder

    ProcessDirectSample = (chain.Status = "DONE" Or _
                           chain.Status = "NO CLEARING" Or _
                           chain.Status = "NO VENDOR PAYMENTS")
    Exit Function

DirectFailed:
    WriteResult sheet, item.Row, "ERROR", vbNullString, vbNullString, vbNullString, _
                0, Err.Description
    modLog.LogAction item.Idx, "Sample failed", Err.Description, "ERROR", vbNullString
End Function

Private Function CountFiles(ByRef chain As ChainResult) As Long
    If Len(chain.FiDocumentFile) > 0 Then CountFiles = CountFiles + 1
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
    LoadSamples = LoadSamplesWhere(samples, False)
End Function

' includeEverything ignores the Include? column, for the operator who says
' 'run the lot' rather than marking rows one at a time.
Private Function LoadSamplesWhere(ByRef samples() As Sample, _
                                  ByVal includeEverything As Boolean) As Long
    Dim sheet As Worksheet
    Dim row As Long, lastRow As Long
    Dim count As Long

    Set sheet = ThisWorkbook.Worksheets(modConfig.SHEET_SAMPLES)
    lastRow = sheet.Cells(sheet.Rows.Count, COL_IDX).End(xlUp).Row

    ReDim samples(1 To Application.Max(1, lastRow))

    For row = modConfig.SAMPLES_FIRST_ROW To lastRow
        If IsNumeric(sheet.Cells(row, COL_IDX).Value) And _
           IsDate(sheet.Cells(row, COL_PAY_DATE).Value) And _
           (includeEverything Or IsIncluded(sheet, row)) Then

            count = count + 1
            samples(count).Row = row
            samples(count).Idx = CLng(sheet.Cells(row, COL_IDX).Value)
            samples(count).MonthTab = Trim$(CStr(sheet.Cells(row, COL_MONTH).Value))
            samples(count).PayDate = CDate(sheet.Cells(row, COL_PAY_DATE).Value)
            samples(count).Amount = CDbl(sheet.Cells(row, COL_AMOUNT).Value)
            samples(count).Party = Trim$(CStr(sheet.Cells(row, COL_PARTY).Value))
            samples(count).Reference = Trim$(CStr(sheet.Cells(row, COL_REF).Value))
            samples(count).Request = Trim$(CStr(sheet.Cells(row, COL_REQUEST).Value))
            samples(count).CompanyCode = Trim$(CStr(sheet.Cells(row, COL_COMPANY).Value))
            samples(count).Comment = Trim$(CStr(sheet.Cells(row, COL_COMMENT).Value))
            samples(count).StartDocument = Trim$(CStr(sheet.Cells(row, COL_START_DOC).Value))
            samples(count).StartAt = UCase$(Trim$(CStr(sheet.Cells(row, COL_START_AT).Value)))
            samples(count).ExpectedZp = Trim$(CStr(sheet.Cells(row, COL_EXPECTED_ZP).Value))

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

    LoadSamplesWhere = count
End Function

' What is on the sheet, per request, so the operator can see the scope
' before answering rather than trusting a single number.
Private Function ScopeSummary() As String
    Dim sheet As Worksheet
    Dim row As Long, lastUsed As Long
    Dim keys As Object, included As Object, total As Object
    Dim key As Variant
    Dim text As String
    Dim onSheet As Long, marked As Long

    Set sheet = ThisWorkbook.Worksheets(modConfig.SHEET_SAMPLES)
    lastUsed = sheet.Cells(sheet.Rows.Count, COL_PAY_DATE).End(xlUp).row

    Set keys = CreateObject("Scripting.Dictionary")
    Set included = CreateObject("Scripting.Dictionary")
    Set total = CreateObject("Scripting.Dictionary")
    keys.CompareMode = vbTextCompare
    included.CompareMode = vbTextCompare
    total.CompareMode = vbTextCompare

    For row = modConfig.SAMPLES_FIRST_ROW To lastUsed
        If IsDate(sheet.Cells(row, COL_PAY_DATE).Value) Then
            key = Trim$(CStr(sheet.Cells(row, COL_REQUEST).Value))
            If Len(key) = 0 Then key = "(no request name)"
            key = key & "  [" & Trim$(CStr(sheet.Cells(row, COL_COMPANY).Value)) & "]"

            If Not keys.Exists(key) Then
                keys.Add key, True
                total.Add key, 0
                included.Add key, 0
            End If

            total(key) = total(key) + 1
            onSheet = onSheet + 1

            If IsIncluded(sheet, row) Then
                included(key) = included(key) + 1
                marked = marked + 1
            End If
        End If
    Next row

    For Each key In keys.Keys
        text = text & "   " & key & "  --  " & included(key) & " of " & total(key) & vbCrLf
    Next key

    ScopeSummary = onSheet & " sample(s) on the '" & modConfig.SHEET_SAMPLES & _
                   "' sheet, " & marked & " marked Include." & vbCrLf & vbCrLf & _
                   text & vbCrLf & _
                   "The run opens whichever FEBAN periods these need, on its own." & _
                   vbCrLf & vbCrLf & _
                   "YES     run the " & marked & " marked Include" & vbCrLf & _
                   "NO      run all " & onSheet & ", ignoring the Include column" & vbCrLf & _
                   "CANCEL  stop"
End Function

'-----------------------------------------------------------------------
' Put Excel back the way it was, if a run ended badly.
'
' The run switches off Excel's response to DDE open requests. That is an
' APPLICATION setting: it survives closing the workbook, and while it is on,
' double-clicking a file in Explorer opens Excel with a blank window. Every
' exit path restores it, but a hard crash has no exit path -- hence a button.
'
' It also empties the scratch folder the exports pass through, for the same
' reason: a run that stopped hard leaves files in it.
'-----------------------------------------------------------------------
Public Sub RestoreExcelSettings()
    Dim swept As Long

    modUtil.RestoreRemoteOpen False

    On Error Resume Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    swept = modExport.SweepHandover(0)
    On Error GoTo 0

    MsgBox "Excel is back to normal:" & vbCrLf & vbCrLf & _
           "  responds to files opened from Explorer" & vbCrLf & _
           "  screen updating on" & vbCrLf & _
           "  alerts on" & vbCrLf & _
           "  " & swept & " leftover export(s) cleared from the scratch folder" & vbCrLf & vbCrLf & _
           "Only needed if a run stopped hard -- every normal exit does this " & _
           "on its own.", vbInformation, "Restore Excel settings"
End Sub

'-----------------------------------------------------------------------
' Wipe the result columns so the sheet reads as 'not started' again.
'
' A re-run does not need this -- every result column is overwritten as each
' sample finishes, and the evidence folders are cleared per sample before
' anything is exported. It is for the case where the sheet should stop
' claiming outcomes from a run that is no longer relevant: rows now marked
' Include? = No keep their old Status forever otherwise.
'-----------------------------------------------------------------------
Public Sub ClearResults()
    Dim sheet As Worksheet
    Dim row As Long, lastUsed As Long, cleared As Long

    Set sheet = ThisWorkbook.Worksheets(modConfig.SHEET_SAMPLES)
    lastUsed = sheet.Cells(sheet.Rows.Count, COL_PAY_DATE).End(xlUp).row

    If MsgBox("Clear Status, statement item, FI document, invoices, file count " & _
              "and message on every sample row?" & vbCrLf & vbCrLf & _
              "The samples themselves, the Include? column and the Log are left " & _
              "alone, and no files on disk are touched.", _
              vbQuestion + vbYesNo + vbDefaultButton2, "Clear results") <> vbYes Then
        Exit Sub
    End If

    For row = modConfig.SAMPLES_FIRST_ROW To lastUsed
        If IsDate(sheet.Cells(row, COL_PAY_DATE).Value) Then
            sheet.Range(sheet.Cells(row, COL_STATUS), _
                        sheet.Cells(row, COL_MESSAGE)).ClearContents
            sheet.Cells(row, COL_STATUS).Font.Color = RGB(0, 0, 0)
            sheet.Cells(row, COL_STATUS).Font.Bold = False
            cleared = cleared + 1
        End If
    Next row

    MsgBox cleared & " row(s) reset to not-started." & vbCrLf & vbCrLf & _
           "Use 'Clear the log' as well if you want the audit trail to start " & _
           "fresh -- it is kept separately because it is the record of what was " & _
           "pulled, and successive attempts staying visible is usually the point.", _
           vbInformation, "Clear results"
End Sub

'-----------------------------------------------------------------------
' Bulk-set the Include? column, so choosing what to run is not 141 edits.
'-----------------------------------------------------------------------
Public Sub IncludeAllSamples()
    SetIncludeAll "Yes"
End Sub

Public Sub ExcludeAllSamples()
    SetIncludeAll "No"
End Sub

' Mark the samples that already have a complete answer as excluded, so a
' re-run picks up only what is left.
'
' STATUS DOES NOT DECIDE WHAT RUNS -- Include? does. A sheet full of DONE
' still runs every row again, for hours, which is not what anyone expects
' the second time they press the button. This is the bridge between the two.
'
' Only the three complete answers are excluded. ERROR, PARTIAL and the
' BLOCKED statuses stay included, because those are the rows a re-run exists
' to retry.
Public Sub ExcludeFinishedSamples()
    Dim sheet As Worksheet
    Dim row As Long, lastUsed As Long
    Dim excluded As Long, retried As Long, notStarted As Long
    Dim status As String

    Set sheet = ThisWorkbook.Worksheets(modConfig.SHEET_SAMPLES)
    lastUsed = sheet.Cells(sheet.Rows.Count, COL_PAY_DATE).End(xlUp).row

    For row = modConfig.SAMPLES_FIRST_ROW To lastUsed
        If IsDate(sheet.Cells(row, COL_PAY_DATE).Value) Then
            status = UCase$(Trim$(CStr(sheet.Cells(row, COL_STATUS).Value)))

            Select Case status
                Case "DONE", "NO CLEARING", "NO VENDOR PAYMENTS"
                    sheet.Cells(row, COL_INCLUDE).Value = "No"
                    excluded = excluded + 1
                Case ""
                    sheet.Cells(row, COL_INCLUDE).Value = "Yes"
                    notStarted = notStarted + 1
                Case Else
                    sheet.Cells(row, COL_INCLUDE).Value = "Yes"
                    retried = retried + 1
            End Select
        End If
    Next row

    MsgBox excluded & " sample(s) already answered -- excluded." & vbCrLf & _
           retried & " stopped part-way -- left in, so a re-run retries them." & vbCrLf & _
           notStarted & " not started -- left in." & vbCrLf & vbCrLf & _
           "DONE, NO CLEARING and NO VENDOR PAYMENTS all count as answered: " & _
           "the first has its invoice, and the other two have no invoice to " & _
           "fetch and never did.", _
           vbInformation, "Exclude finished samples"
End Sub

Private Sub SetIncludeAll(ByVal value As String)
    Dim sheet As Worksheet
    Dim row As Long, lastUsed As Long, touched As Long

    Set sheet = ThisWorkbook.Worksheets(modConfig.SHEET_SAMPLES)
    lastUsed = sheet.Cells(sheet.Rows.Count, COL_PAY_DATE).End(xlUp).row

    For row = modConfig.SAMPLES_FIRST_ROW To lastUsed
        If IsDate(sheet.Cells(row, COL_PAY_DATE).Value) Then
            sheet.Cells(row, COL_INCLUDE).Value = value
            touched = touched + 1
        End If
    Next row

    MsgBox touched & " sample(s) set to Include? = " & value & "." & vbCrLf & vbCrLf & _
           "Change individual rows from here, then run.", _
           vbInformation, "Include " & LCase$(value)
End Sub

' Blank counts as included, so a freshly generated sheet runs everything and
' the operator only has to touch the rows they want left out.
Private Function IsIncluded(ByVal sheet As Worksheet, ByVal row As Long) As Boolean
    Dim flag As String

    flag = UCase$(Trim$(CStr(sheet.Cells(row, COL_INCLUDE).Value)))
    IsIncluded = (flag <> "NO" And flag <> "N" And flag <> "FALSE" And flag <> "0")
End Function

' month tab -> Array(dateFrom, dateTo), in first-seen order
' The unit of work is not a month, it is a request-and-company-and-month.
' FEBAN searches one company code at a time, and two requests can both hold
' a Sep 25 sample for different companies -- grouping on the month alone
' would run the second lot against the first one's company.
Private Function DistinctMonths(ByRef samples() As Sample, ByVal count As Long) As Object
    Dim months As Object
    Dim i As Long
    Dim key As String

    Set months = CreateObject("Scripting.Dictionary")
    months.CompareMode = vbTextCompare

    For i = 1 To count
        key = PeriodKey(samples(i))
        If Not months.Exists(key) Then
            months.Add key, Array(samples(i).DateFrom, samples(i).DateTo, _
                                  samples(i).MonthTab, samples(i).Request, _
                                  samples(i).CompanyCode, samples(i).StartAt)
        End If
    Next i

    Set DistinctMonths = months
End Function

Private Function PeriodKey(ByRef item As Sample) As String
    ' The entry mode is part of the key: rows that skip FEBAN must not be
    ' grouped with rows that need it, or a period that fails to open would
    ' take down samples that never needed it.
    PeriodKey = item.Request & "|" & item.CompanyCode & "|" & item.MonthTab & _
                "|" & item.StartAt
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
    Dim remoteWas As Boolean

    answer = InputBox("Which month tab? For example: Sep 25", "Run one month", "Sep 25")
    If Len(Trim$(answer)) = 0 Then Exit Sub

    On Error GoTo Failed

    modConfig.LoadScreenMap
    modConfig.AssertScreenMapComplete
    count = LoadSamples(samples)

    modSapConnect.SapAttach
    modSafety.gWriteAttemptBlocked = 0
    mIncomplete = 0
    modLog.WriteHeaderBlock
    remoteWas = modUtil.SuppressRemoteOpen()
    modExport.SweepHandover 0

    For i = 1 To count
        If StrComp(samples(i).MonthTab, Trim$(answer), vbTextCompare) = 0 Then
            If Not matched Then
                modFeban.OpenMonth samples(i).DateFrom, samples(i).DateTo

                ' Once per period, not per sample -- but every sample folder
                ' gets a copy, so the packs stand alone.
                modConfig.SetCompanyCode samples(i).CompanyCode
                If Len(ExportMonthStatementList(samples(i).MonthTab, _
                                                samples(i).Request, _
                                                samples(i).CompanyCode)) > 0 Then
                    files = files + 1
                End If

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
    modUtil.RestoreRemoteOpen remoteWas
    modUtil.CloseExportWorkbooksUnder modConfig.DownloadRoot()
    modExport.SweepHandover modExport.HANDOVER_GRACE_SECONDS

    ' Say where they stopped, read off the Log, rather than assuming. The old
    ' wording claimed none got past the statement match, which was wrong the
    ' moment a sample reached FBL1N and exported its batch list.
    MsgBox "Finished " & answer & "." & vbCrLf & vbCrLf & _
           "Samples reaching the end of the chain : " & done & vbCrLf & _
           "  of those, short of a full answer    : " & mIncomplete & vbCrLf & _
           "Samples that stopped early            : " & failed & vbCrLf & _
           "Files written                         : " & files & vbCrLf & vbCrLf & _
           IIf(failed > 0, "Where they stopped: " & modLog.FailureSummary() & vbCrLf & vbCrLf, "") & _
           "Files are under " & modConfig.DownloadRoot() & ".", _
           IIf(failed = 0, vbInformation, vbExclamation), "Run one month"
    Exit Sub

Failed:
    modUtil.RestoreRemoteOpen remoteWas
    modLog.LogAction 0, "RUN ABORTED", Err.Description, "ERROR", vbNullString
    MsgBox "Stopped: " & Err.Description, vbCritical
End Sub
