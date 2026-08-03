Attribute VB_Name = "modImport"
'=======================================================================
' modImport -- read an auditor's sample request into the Samples sheet.
'
' Seven request files, three layouts, and no two headers in the same
' place. What they have in common is only this: somewhere there is a row
' of column captions, and under it rows carrying a date, an amount, and
' usually a party name. Everything else varies.
'
'   Paper / Packaging   header on rows 3+4, amounts POSITIVE
'                       Month of payment | Payment reference | Date |
'                       Amount | Name of party
'
'   Jan / Feb / Mar     header on rows 2+3, or 4 in January, one column
'                       further left; amounts NEGATIVE
'                       Bank account number | Date | Customer Reference |
'                       Bank Reference | Transaction Description |
'                       Name of Party | Amount
'
'   Follow-up queries   one sheet per month, header on rows 10+11, 5+6
'                       and 6+7 respectively; amounts NEGATIVE; carries
'                       an EY Comments column saying what is wanted
'
' So nothing is addressed by position. The header row is found by looking
' for a row that names both a date and an amount; the columns are matched
' by caption against a list of synonyms; and because every one of these
' files splits its header over two rows -- a group caption like 'Payment
' details' above the real names -- each candidate row is read together
' with the one above it.
'
' Amounts are stored unsigned. The auditor's files disagree about the sign
' of a payment and the statement match compares magnitudes anyway.
'
' Nothing is written until the operator has seen what was matched.
'=======================================================================
Option Explicit

' Sheets belonging to a document-storage add-in, not to the request.
Private Const SKIP_SHEET_PREFIX As String = "DS_INTERNAL_"

' Caption synonyms, lower case. First match wins, so the more specific
' captions come first.
Private Const CAPTIONS_DATE As String = _
    "date|payment date|posting date|value date|document date"
Private Const CAPTIONS_AMOUNT As String = _
    "amount|amount in local currency|amount in lc|payment amount|value"
Private Const CAPTIONS_PARTY As String = _
    "name of party|name of the party|name 1|party|supplier|vendor|beneficiary|payee"
Private Const CAPTIONS_REFERENCE As String = _
    "transaction description|payment reference|description|narrative|" & _
    "customer reference|bank reference|reference"
Private Const CAPTIONS_COMMENT As String = _
    "ey comments|ey comment|comments|comment|query|request"

' A SAP document list rather than an auditor's statement extract. When all
' three are present the rows already name the documents that steps 1 to 5
' of the chain exist to discover.
Private Const CAPTIONS_DOCNUMBER As String = _
    "document number|documentno|document no.|doc. number|doc.no.|belnr"
Private Const CAPTIONS_DOCTYPE As String = _
    "document type|doc. type|documenttype|blart"
Private Const CAPTIONS_CLEARING As String = _
    "clearing document|clearing doc.|clrng doc.|augbl"

' Set when the operator confirmed this file is a SAP document list and the
' rows should enter the chain directly. Read by Commit.
Private mDirectEntry As Boolean

Private Type ColumnMap
    Found As Boolean
    HeaderRow As Long
    DateCol As Long
    AmountCol As Long
    PartyCol As Long
    ReferenceCol As Long
    CommentCol As Long
    ' Set only for a SAP document list -- a Payment Usage export or an
    ' FBL1N download. These are what let a row skip the statement search.
    DocNumberCol As Long
    DocTypeCol As Long
    ClearingCol As Long
End Type

'-----------------------------------------------------------------------
' The button. Pick a file, say which company code it belongs to, look at
' what was found, then commit.
'-----------------------------------------------------------------------
Public Sub ImportRequest()
    Dim path As String
    Dim companyCode As String
    Dim requestName As String
    Dim book As Workbook
    Dim previousAlerts As Boolean
    Dim preview As String
    Dim added As Long

    path = ChooseFile()
    If Len(path) = 0 Then Exit Sub

    requestName = DefaultRequestName(path)
    requestName = Trim$(InputBox( _
        "Name for this audit request." & vbCrLf & vbCrLf & _
        "It goes in the Samples sheet against every row from this file, and " & _
        "names the folder the evidence is written to.", _
        "Import request", requestName))
    If Len(requestName) = 0 Then Exit Sub

    companyCode = Trim$(InputBox( _
        "Company code for every sample in this request." & vbCrLf & vbCrLf & _
        "These requests are not all the same company -- the Paper samples are " & _
        "GBKM, and the Packaging ones are not -- so this is asked per file " & _
        "rather than taken from the Control sheet.", _
        "Import request", modConfig.Setting("Company code")))
    If Len(companyCode) = 0 Then Exit Sub

    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error GoTo Failed

    Set book = Application.Workbooks.Open(fileName:=path, UpdateLinks:=0, ReadOnly:=True)

    ' A SAP document list is not a statement extract, and treating it as one
    ' sends FEBAN looking for payment documents that were never statement
    ' lines. Offer the shortcut instead: the rows already name the documents.
    If Not AskAboutDocumentList(book) Then
        book.Close SaveChanges:=False
        Application.DisplayAlerts = previousAlerts
        Exit Sub
    End If

    preview = Describe(book, requestName, companyCode)

    If MsgBox(preview & vbCrLf & vbCrLf & "Add these to the Samples sheet?", _
              vbQuestion + vbYesNo, "Import request") <> vbYes Then
        book.Close SaveChanges:=False
        Application.DisplayAlerts = previousAlerts
        Exit Sub
    End If

    added = Commit(book, requestName, companyCode)

    book.Close SaveChanges:=False
    Application.DisplayAlerts = previousAlerts

    Renumber

    MsgBox added & " sample(s) added to the '" & modConfig.SHEET_SAMPLES & "' sheet " & _
           "for request '" & requestName & "' (" & companyCode & ")." & vbCrLf & vbCrLf & _
           "Check the Party and Reference columns before running -- and set " & _
           "Include? to No on any row you do not want." & vbCrLf & vbCrLf & _
           "Evidence for this request will go to:" & vbCrLf & _
           modUtil.JoinPath(modConfig.DownloadRoot(), _
                            modUtil.RequestFolderName(requestName, companyCode)), _
           vbInformation, "Import request"
    Exit Sub

Failed:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    Application.DisplayAlerts = previousAlerts
    On Error GoTo 0

    MsgBox "The request could not be imported." & vbCrLf & vbCrLf & Err.Description, _
           vbExclamation, "Import request"
End Sub

'-----------------------------------------------------------------------
' 'These are SAP documents, not statement lines.'
'
' A Payment Usage export has a document number, a document type and a
' clearing document, which an auditor's request never does. Those rows name
' what the first five steps of the chain exist to discover, so there is no
' reason to search a bank statement for them:
'
'   SB with a clearing document  -> enter at the clearing document
'   SB with none                 -> the FI line items are the evidence
'   ZP                           -> enter at the payment itself
'
' Returns False only if the operator cancels.
'-----------------------------------------------------------------------
Private Function AskAboutDocumentList(ByVal book As Workbook) As Boolean
    Dim sheet As Worksheet
    Dim map As ColumnMap
    Dim looksLikeDocuments As Boolean
    Dim answer As VbMsgBoxResult

    mDirectEntry = False

    For Each sheet In book.Worksheets
        If Not IsSkippable(sheet) Then
            map = FindColumns(sheet)
            If map.Found Then
                If map.DocNumberCol > 0 And map.DocTypeCol > 0 Then looksLikeDocuments = True
            End If
        End If
    Next sheet

    If Not looksLikeDocuments Then
        AskAboutDocumentList = True
        Exit Function
    End If

    answer = MsgBox( _
        "This file holds SAP DOCUMENTS, not bank statement lines -- it has a " & _
        "document number, a document type and a clearing document. A Payment " & _
        "Usage or FBL1N export looks like this." & vbCrLf & vbCrLf & _
        "Those rows already name what the first five steps of the chain exist " & _
        "to find, so FEBAN can be skipped:" & vbCrLf & vbCrLf & _
        "   ZP payment            -> straight to its invoices" & vbCrLf & _
        "   SB with a clearing doc -> the batch behind it, then its invoices" & vbCrLf & _
        "   SB with none           -> the FI line items, nothing settled" & vbCrLf & vbCrLf & _
        "Each one gets its invoice list exported and the largest invoice's PDF " & _
        "downloaded, the same as any other sample." & vbCrLf & vbCrLf & _
        "YES     enter the chain at each document (recommended)" & vbCrLf & _
        "NO      treat them as statement lines and search FEBAN by date and amount" & _
        vbCrLf & _
        "CANCEL  stop", _
        vbQuestion + vbYesNoCancel, "SAP documents, not statement lines")

    If answer = vbCancel Then Exit Function

    mDirectEntry = (answer = vbYes)
    AskAboutDocumentList = True
End Function

'-----------------------------------------------------------------------
' What the importer thinks it has found, before anything is written.
'-----------------------------------------------------------------------
Private Function Describe(ByVal book As Workbook, ByVal requestName As String, _
                          ByVal companyCode As String) As String
    Dim sheet As Worksheet
    Dim map As ColumnMap
    Dim text As String
    Dim rows As Long, total As Long
    Dim used As Long

    text = "Request : " & requestName & vbCrLf & _
           "Company : " & companyCode & vbCrLf & _
           "File    : " & book.Name & vbCrLf & vbCrLf

    For Each sheet In book.Worksheets
        If Not IsSkippable(sheet) Then
            used = used + 1
            map = FindColumns(sheet)

            If map.Found Then
                rows = CountRows(sheet, map)
                total = total + rows
                text = text & "'" & sheet.Name & "' -- header row " & map.HeaderRow & _
                       ", " & rows & " sample(s)" & vbCrLf & _
                       "     date=" & ColumnLetter(map.DateCol) & _
                       "  amount=" & ColumnLetter(map.AmountCol) & _
                       "  party=" & ColumnLetter(map.PartyCol) & _
                       "  reference=" & ColumnLetter(map.ReferenceCol) & _
                       IIf(map.CommentCol > 0, _
                           "  comment=" & ColumnLetter(map.CommentCol), "") & vbCrLf & _
                       FirstRows(sheet, map) & _
                       WrongShapeWarning(map) & vbCrLf
            Else
                text = text & "'" & sheet.Name & "' -- no date-and-amount header found, " & _
                       "skipped." & vbCrLf & vbCrLf
            End If
        End If
    Next sheet

    If used = 0 Then text = text & "No usable sheets in this file." & vbCrLf

    Describe = text & String$(50, "-") & vbCrLf & total & " sample(s) in total."
End Function

' Say so loudly when the sheet does not look like an audit request.
'
' A SAP Payment Usage export has a Document Date and an Amount in local
' currency, so the header probe accepts it and hands back rows -- but it has
' no party and no transaction description, because it is not a list of
' payments an auditor picked. It is the macro's OWN output. One got imported
' as a request; 'party=(none)' in the line above was too easy to skim past.
Private Function WrongShapeWarning(ByRef map As ColumnMap) As String
    If map.PartyCol > 0 Then Exit Function

    WrongShapeWarning = _
        "     *** No party column on this sheet. An auditor's request names " & vbCrLf & _
        "     *** the party for most rows. A sheet with a date and an amount " & vbCrLf & _
        "     *** but no party is usually a SAP export rather than a request " & vbCrLf & _
        "     *** -- check this is the file you meant before saying Yes." & vbCrLf
End Function

' Two rows of what was read, so a wrong column is obvious before it is
' committed rather than after a run.
Private Function FirstRows(ByVal sheet As Worksheet, ByRef map As ColumnMap) As String
    Dim row As Long, shown As Long
    Dim text As String

    For row = map.HeaderRow + 1 To LastRow(sheet)
        If shown >= 2 Then Exit For
        If IsSampleRow(sheet, row, map) Then
            shown = shown + 1
            text = text & "     " & _
                   Format$(RowDate(sheet, row, map.DateCol), "dd/mm/yyyy") & _
                   "  " & Format$(Abs(RowAmount(sheet, row, map.AmountCol)), "#,##0.00") & _
                   "  " & Left$(CellString(sheet, row, map.PartyCol), 28) & vbCrLf
        End If
    Next row

    FirstRows = text
End Function

'-----------------------------------------------------------------------
' Append the rows.
'-----------------------------------------------------------------------
Private Function Commit(ByVal book As Workbook, ByVal requestName As String, _
                        ByVal companyCode As String) As Long
    Dim target As Worksheet
    Dim sheet As Worksheet
    Dim map As ColumnMap
    Dim row As Long, outRow As Long
    Dim added As Long
    Dim payDate As Date

    Set target = ThisWorkbook.Worksheets(modConfig.SHEET_SAMPLES)
    outRow = NextFreeRow(target)

    ' Column V post-dates the workbook builder, so a sheet made before it has
    ' no heading there. Write one rather than leaving a column of values under
    ' a blank header.
    If Len(Trim$(CStr(target.Cells(4, 22).Value))) = 0 Then
        target.Cells(4, 22).Value = "Auditor's ZP"
        target.Columns(22).ColumnWidth = 14
    End If

    For Each sheet In book.Worksheets
        If Not IsSkippable(sheet) Then
            map = FindColumns(sheet)

            If map.Found Then
                For row = map.HeaderRow + 1 To LastRow(sheet)
                    If IsSampleRow(sheet, row, map) Then
                        payDate = RowDate(sheet, row, map.DateCol)

                        target.Cells(outRow, 2).Value = modUtil.MonthTabName(payDate)
                        target.Cells(outRow, 3).Value = modUtil.MonthStart(payDate)
                        target.Cells(outRow, 4).Value = modUtil.MonthEnd(payDate)
                        target.Cells(outRow, 5).Value = payDate
                        ' Unsigned: these files disagree about the sign and the
                        ' statement match compares magnitudes.
                        target.Cells(outRow, 6).Value = _
                            Abs(RowAmount(sheet, row, map.AmountCol))
                        target.Cells(outRow, 7).Value = _
                            NormaliseParty(CellString(sheet, row, map.PartyCol))
                        target.Cells(outRow, 8).Value = _
                            CellString(sheet, row, map.ReferenceCol)
                        target.Cells(outRow, 16).Value = "Yes"
                        target.Cells(outRow, 17).Value = requestName
                        target.Cells(outRow, 18).Value = companyCode
                        target.Cells(outRow, 19).Value = _
                            CellString(sheet, row, map.CommentCol)

                        ' A follow-up request often shows its own working: an
                        ' SAP extract pasted under each sample naming the ZP the
                        ' auditor picked. Capture it so the run can be checked
                        ' against their answer rather than merely trusted.
                        target.Cells(outRow, 22).Value = _
                            AuditorZpBelow(sheet, row, map)

                        ' Where this row enters the chain. Blank for a normal
                        ' statement-driven sample, which is every auditor
                        ' request; set only for a SAP document list.
                        If mDirectEntry Then
                            target.Cells(outRow, 20).Value = _
                                EntryDocument(sheet, row, map)
                            target.Cells(outRow, 21).Value = EntryRung(sheet, row, map)
                        End If

                        target.Cells(outRow, 3).NumberFormat = "dd/mm/yyyy"
                        target.Cells(outRow, 4).NumberFormat = "dd/mm/yyyy"
                        target.Cells(outRow, 5).NumberFormat = "dd/mm/yyyy"
                        target.Cells(outRow, 6).NumberFormat = "#,##0.00"

                        outRow = outRow + 1
                        added = added + 1
                    End If
                Next row
            End If
        End If
    Next sheet

    ApplyIncludeDropdown target

    Commit = added
End Function

'-----------------------------------------------------------------------
' Put the Yes/No dropdown back on the Include column.
'
' The workbook builder applied it once, over the rows that existed then --
' the original 56. Every row imported since has been plain free text, so the
' column reads as a dropdown at the top of the sheet and as anything you like
' further down. That matters more than it looks: the run tests the cell for
' 'No', so a typo there silently includes a sample the operator meant to
' leave out, and nothing anywhere reports it.
'
' Re-applied across every sample row rather than just the new ones, so one
' import repairs a sheet that has been drifting for several.
'-----------------------------------------------------------------------
Private Sub ApplyIncludeDropdown(ByVal target As Worksheet)
    Dim lastUsed As Long
    Dim area As Range

    lastUsed = target.Cells(target.Rows.Count, 5).End(xlUp).row
    If lastUsed < modConfig.SAMPLES_FIRST_ROW Then Exit Sub

    Set area = target.Range(target.Cells(modConfig.SAMPLES_FIRST_ROW, 16), _
                            target.Cells(lastUsed, 16))

    ' Never fatal. A sheet that will not take validation -- protected, or an
    ' Excel build that objects -- still imports perfectly well.
    On Error Resume Next
    area.Validation.Delete
    area.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                        Operator:=xlBetween, Formula1:="Yes,No"
    area.Validation.IgnoreBlank = True
    area.Validation.InCellDropdown = True
    area.HorizontalAlignment = xlCenter
    On Error GoTo 0
End Sub

'-----------------------------------------------------------------------
' Which rung of the chain a SAP document row belongs on, and which document
' to enter it with.
'
' A ZP IS the payment, so it enters at step 10 with its own number. An SB is
' a bank statement posting: if it carries a clearing document that is the
' batch, and the row enters at step 5 with the CLEARING number, not its own.
' An SB with no clearing document settled nothing, and WalkFromClearing on a
' blank number would go nowhere -- those are left as statement rows so the
' normal chain reports them as NO CLEARING with their line items exported.
'-----------------------------------------------------------------------
Private Function EntryRung(ByVal sheet As Worksheet, ByVal row As Long, _
                           ByRef map As ColumnMap) As String
    Dim docType As String

    docType = UCase$(CellString(sheet, row, map.DocTypeCol))

    If IsPaymentType(docType) Then
        EntryRung = "PAYMENT"
    ElseIf Len(CellString(sheet, row, map.ClearingCol)) > 0 Then
        EntryRung = "CLEARING"
    End If
End Function

Private Function EntryDocument(ByVal sheet As Worksheet, ByVal row As Long, _
                               ByRef map As ColumnMap) As String
    Select Case EntryRung(sheet, row, map)
        Case "PAYMENT":  EntryDocument = CellString(sheet, row, map.DocNumberCol)
        Case "CLEARING": EntryDocument = CellString(sheet, row, map.ClearingCol)
    End Select
End Function

' The Control sheet already names the payment document types, and takes a
' list, so this follows whatever is set there rather than hardcoding ZP.
Private Function IsPaymentType(ByVal docType As String) As Boolean
    Dim wanted() As String
    Dim i As Long

    If Len(docType) = 0 Then Exit Function

    wanted = Split(Replace(Replace(modConfig.Setting("Payment document type"), _
                                   ";", ","), "|", ","), ",")

    For i = LBound(wanted) To UBound(wanted)
        If Len(Trim$(wanted(i))) > 0 Then
            If StrComp(docType, UCase$(Trim$(wanted(i))), vbTextCompare) = 0 Then
                IsPaymentType = True
                Exit Function
            End If
        End If
    Next i
End Function

'-----------------------------------------------------------------------
' Finding the header
'-----------------------------------------------------------------------
Private Function FindColumns(ByVal sheet As Worksheet) As ColumnMap
    Dim result As ColumnMap
    Dim row As Long
    Dim dateCol As Long, amountCol As Long

    ' Twenty rows is generous: the deepest of these files puts its header on
    ' row 11, and looking further only risks matching a data row.
    For row = 1 To LesserOf(20, LastRow(sheet))
        dateCol = MatchColumn(sheet, row, CAPTIONS_DATE)
        amountCol = MatchColumn(sheet, row, CAPTIONS_AMOUNT)

        ' Both, on the same row, is what makes it a header rather than a
        ' stray label. A date column on its own turns up in plenty of places.
        If dateCol > 0 And amountCol > 0 Then
            result.Found = True
            result.HeaderRow = row
            result.DateCol = dateCol
            result.AmountCol = amountCol
            result.PartyCol = MatchColumn(sheet, row, CAPTIONS_PARTY)
            result.ReferenceCol = MatchColumn(sheet, row, CAPTIONS_REFERENCE)
            result.CommentCol = MatchColumn(sheet, row, CAPTIONS_COMMENT)
            result.DocNumberCol = MatchColumn(sheet, row, CAPTIONS_DOCNUMBER)
            result.DocTypeCol = MatchColumn(sheet, row, CAPTIONS_DOCTYPE)
            result.ClearingCol = MatchColumn(sheet, row, CAPTIONS_CLEARING)
            Exit For
        End If
    Next row

    FindColumns = result
End Function

' The column on this row whose caption matches one of these synonyms.
'
' Reads the cell together with the one above it, because every one of these
' files splits its header in two: 'Payment details' spanning the group on
' one row, 'Date' and 'Amount' underneath. Either half may carry the word.
Private Function MatchColumn(ByVal sheet As Worksheet, ByVal row As Long, _
                             ByVal synonyms As String) As Long
    Dim wanted() As String
    Dim col As Long, i As Long
    Dim caption As String

    wanted = Split(synonyms, "|")

    For i = LBound(wanted) To UBound(wanted)
        For col = 1 To LesserOf(40, LastColumn(sheet))
            caption = LCase$(modUtil.Squeeze(CellString(sheet, row, col)))

            If caption = wanted(i) Then
                MatchColumn = col
                Exit Function
            End If
        Next col
    Next i
End Function

'-----------------------------------------------------------------------
' Rows
'-----------------------------------------------------------------------
Private Function IsSampleRow(ByVal sheet As Worksheet, ByVal row As Long, _
                             ByRef map As ColumnMap) As Boolean
    If Not HasDate(sheet, row, map.DateCol) Then Exit Function

    ' A zero is a total line or a spacer, never a sample.
    IsSampleRow = (Abs(RowAmount(sheet, row, map.AmountCol)) > 0.005)
End Function

'-----------------------------------------------------------------------
' Dates and amounts that are not stored as dates and amounts.
'
' Three rows of the Paper request hold ' 24/10/2025 ' and '3480924.53' as
' TEXT -- someone retyped them at some point. Trusting IsDate on a string
' would make the import locale-dependent: 24/10/2025 reads fine on a UK
' machine and fails on a US one, and the only symptom would be three
' samples quietly missing from a 56-row request.
'
' So text is parsed here, day first, which is what these files are.
'-----------------------------------------------------------------------
Private Function HasDate(ByVal sheet As Worksheet, ByVal row As Long, _
                         ByVal col As Long) As Boolean
    Dim raw As Variant

    If col <= 0 Then Exit Function

    raw = sheet.Cells(row, col).Value
    If IsEmpty(raw) Then Exit Function

    If VarType(raw) = vbDate Then
        HasDate = True
    Else
        HasDate = (ParseDayFirst(CStr(raw)) > 0)
    End If
End Function

Private Function RowDate(ByVal sheet As Worksheet, ByVal row As Long, _
                         ByVal col As Long) As Date
    Dim raw As Variant

    raw = sheet.Cells(row, col).Value

    If VarType(raw) = vbDate Then
        RowDate = CDate(raw)
    Else
        RowDate = ParseDayFirst(CStr(raw))
    End If
End Function

Private Function RowAmount(ByVal sheet As Worksheet, ByVal row As Long, _
                           ByVal col As Long) As Double
    Dim raw As Variant

    If col <= 0 Then Exit Function

    raw = sheet.Cells(row, col).Value
    If IsEmpty(raw) Then Exit Function

    If IsNumeric(raw) And VarType(raw) <> vbString Then
        RowAmount = CDbl(raw)
    Else
        ' Handles thousands separators and a trailing minus, both of which
        ' turn up in retyped cells.
        RowAmount = modUtil.ParseSapAmount(CStr(raw))
    End If
End Function

' 'dd/mm/yyyy', 'dd.mm.yyyy', 'dd-mm-yyyy'. Returns 0 for anything else,
' which is how the caller tells a date from a caption.
Private Function ParseDayFirst(ByVal text As String) As Date
    Dim parts() As String
    Dim cleaned As String
    Dim d As Long, m As Long, y As Long

    cleaned = Trim$(text)
    If Len(cleaned) < 6 Then Exit Function

    cleaned = Replace(Replace(cleaned, ".", "/"), "-", "/")
    parts = Split(cleaned, "/")
    If UBound(parts) - LBound(parts) <> 2 Then Exit Function

    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then
        Exit Function
    End If

    d = CLng(Val(parts(0)))
    m = CLng(Val(parts(1)))
    y = CLng(Val(parts(2)))

    If y < 100 Then y = 2000 + y
    If d < 1 Or d > 31 Or m < 1 Or m > 12 Or y < 1990 Or y > 2100 Then Exit Function

    On Error Resume Next
    ParseDayFirst = DateSerial(y, m, d)
    On Error GoTo 0
End Function

Private Function CountRows(ByVal sheet As Worksheet, ByRef map As ColumnMap) As Long
    Dim row As Long

    For row = map.HeaderRow + 1 To LastRow(sheet)
        If IsSampleRow(sheet, row, map) Then CountRows = CountRows + 1
    Next row
End Function

' The auditor's files write '-' where no party is named. Blank says the
' same thing without looking like data.
Private Function NormaliseParty(ByVal text As String) As String
    Dim trimmed As String

    trimmed = Trim$(text)
    If trimmed = "-" Or trimmed = "--" Then Exit Function

    NormaliseParty = trimmed
End Function

'-----------------------------------------------------------------------
' Sheet and cell helpers
'-----------------------------------------------------------------------
Private Function IsSkippable(ByVal sheet As Worksheet) As Boolean
    If UCase$(Left$(sheet.Name, Len(SKIP_SHEET_PREFIX))) = SKIP_SHEET_PREFIX Then
        IsSkippable = True
        Exit Function
    End If

    ' The empty month tabs the auditor leaves in for their own working.
    IsSkippable = (LastRow(sheet) < 2)
End Function

Private Function LastRow(ByVal sheet As Worksheet) As Long
    On Error Resume Next
    LastRow = sheet.UsedRange.row + sheet.UsedRange.Rows.Count - 1
    On Error GoTo 0
End Function

Private Function LastColumn(ByVal sheet As Worksheet) As Long
    On Error Resume Next
    LastColumn = sheet.UsedRange.Column + sheet.UsedRange.Columns.Count - 1
    On Error GoTo 0
End Function

Private Function CellString(ByVal sheet As Worksheet, ByVal row As Long, _
                            ByVal col As Long) As String
    If col <= 0 Then Exit Function

    On Error Resume Next
    CellString = Trim$(CStr(sheet.Cells(row, col).Value))
    On Error GoTo 0
End Function

Private Function ColumnLetter(ByVal col As Long) As String
    If col <= 0 Then
        ColumnLetter = "(none)"
    Else
        ColumnLetter = Split(ThisWorkbook.Worksheets(1).Cells(1, col) _
                             .Address(True, False), "$")(0)
    End If
End Function

Private Function LesserOf(ByVal a As Long, ByVal b As Long) As Long
    LesserOf = IIf(a < b, a, b)
End Function

Private Function NextFreeRow(ByVal sheet As Worksheet) As Long
    Dim lastUsed As Long

    lastUsed = sheet.Cells(sheet.Rows.Count, 5).End(xlUp).row
    If lastUsed < modConfig.SAMPLES_FIRST_ROW Then
        NextFreeRow = modConfig.SAMPLES_FIRST_ROW
    Else
        NextFreeRow = lastUsed + 1
    End If
End Function

' Number every sample on the sheet from 1, so the folder names stay in step
' with the rows after an import appends to an existing list.
Private Sub Renumber()
    Dim sheet As Worksheet
    Dim row As Long, lastUsed As Long, n As Long

    Set sheet = ThisWorkbook.Worksheets(modConfig.SHEET_SAMPLES)
    lastUsed = sheet.Cells(sheet.Rows.Count, 5).End(xlUp).row

    For row = modConfig.SAMPLES_FIRST_ROW To lastUsed
        If IsDate(sheet.Cells(row, 5).Value) Then
            n = n + 1
            sheet.Cells(row, 1).Value = n
        End If
    Next row
End Sub

'-----------------------------------------------------------------------
' File dialog. GetOpenFilename rather than FileDialog, because it needs no
' reference and returns False rather than raising when cancelled.
'-----------------------------------------------------------------------
Private Function ChooseFile() As String
    Dim picked As Variant

    picked = Application.GetOpenFilename( _
        fileFilter:="Excel files (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
        Title:="Pick the auditor's sample request")

    If VarType(picked) = vbBoolean Then Exit Function
    ChooseFile = CStr(picked)
End Function

' 'SURL_Samples__February26260730_073000.713.xlsx' -> 'SURL Samples February26'.
' The auditor's export stamps a timestamp on the end of every file name;
' it is noise in a folder name.
Private Function DefaultRequestName(ByVal path As String) As String
    Dim name As String
    Dim i As Long
    Dim ch As String
    Dim digits As Long

    name = path
    i = InStrRev(name, "\")
    If i > 0 Then name = Mid$(name, i + 1)

    i = InStrRev(name, ".")
    If i > 0 Then name = Left$(name, i - 1)

    ' Cut at the first run of six or more digits -- that is the timestamp.
    For i = 1 To Len(name)
        ch = Mid$(name, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits + 1
            If digits >= 6 Then
                name = Left$(name, i - digits)
                Exit For
            End If
        Else
            digits = 0
        End If
    Next i

    name = Replace(name, "_", " ")
    Do While InStr(name, "  ") > 0
        name = Replace(name, "  ", " ")
    Loop

    DefaultRequestName = Trim$(name)
End Function

'-----------------------------------------------------------------------
' The ZP the auditor picked, out of the extract they pasted below the row.
'
' A follow-up request shows its working. Under the sample table there is a
' block per sample -- a heading repeating the sample's label, then a small
' SAP line-item extract, and in it the ZP document the auditor selected as
' the largest payment of that batch. That is the same document this macro
' derives independently from FBL1N, so capturing it turns the run into
' something checkable: two answers to the same question, from two routes.
'
' Read by caption, never by position. The columns are not in the same order
' on every block -- in one April sample 'Clearing Document' is column M and
' 'Document Header Text' is N, and two samples later they are the other way
' round. A positional read would have quietly returned the clearing document
' as though it were the payment.
'
' Returns "" whenever there is no such block, which is the normal case: an
' ordinary statement extract has nothing pasted underneath it.
'-----------------------------------------------------------------------
Private Function AuditorZpBelow(ByVal sheet As Worksheet, ByVal row As Long, _
                                ByRef map As ColumnMap) As String
    Dim label As String
    Dim blockRow As Long

    label = SampleLabel(sheet, row, map)
    If Len(label) = 0 Then Exit Function

    blockRow = LastRowSaying(sheet, label)

    ' The label's own row in the sample table is not a block. Only a later
    ' occurrence is, because the blocks sit below the table.
    If blockRow <= row Then Exit Function

    AuditorZpBelow = ZpInBlock(sheet, blockRow)
End Function

' The label that identifies this sample -- 'April Sample 3'. It sits to the
' left of the dated columns, so search the first few columns of the row for
' something with a word and a number in it rather than assuming a column.
Private Function SampleLabel(ByVal sheet As Worksheet, ByVal row As Long, _
                             ByRef map As ColumnMap) As String
    Dim col As Long
    Dim text As String

    For col = 1 To LesserOf(map.DateCol - 1, 6)
        text = CellString(sheet, row, col)
        If Len(text) >= 6 And Len(text) <= 40 Then
            If InStr(1, text, "sample", vbTextCompare) > 0 Then
                SampleLabel = text
                Exit Function
            End If
        End If
    Next col
End Function

' The LAST row carrying this exact text. The label appears twice -- once in
' the sample table, once heading the block -- and the block is always the
' later of the two.
Private Function LastRowSaying(ByVal sheet As Worksheet, ByVal label As String) As Long
    Dim row As Long, col As Long
    Dim limit As Long

    limit = LastRow(sheet)

    For row = limit To 1 Step -1
        For col = 1 To 8
            If StrComp(CellString(sheet, row, col), label, vbTextCompare) = 0 Then
                LastRowSaying = row
                Exit Function
            End If
        Next col
    Next row
End Function

' Walk down from the block heading: find its header row by caption, then read
' the rows under it, and answer with the document number of the first ZP.
Private Function ZpInBlock(ByVal sheet As Worksheet, ByVal blockRow As Long) As String
    Dim row As Long, col As Long
    Dim docCol As Long, typeCol As Long
    Dim headerRow As Long
    Dim caption As String
    Dim wanted As String

    wanted = FirstWord(modConfig.Setting("Payment document type"))
    If Len(wanted) = 0 Then wanted = "ZP"

    ' The header is within a few rows of the heading.
    For row = blockRow To LesserOf(blockRow + 6, LastRow(sheet))
        docCol = 0: typeCol = 0
        For col = 1 To 20
            caption = CellString(sheet, row, col)
            If InStr(1, caption, "Document Number", vbTextCompare) > 0 Then docCol = col
            If InStr(1, caption, "Document Type", vbTextCompare) > 0 Then typeCol = col
        Next col
        If docCol > 0 And typeCol > 0 Then
            headerRow = row
            Exit For
        End If
    Next row

    If headerRow = 0 Then Exit Function

    ' Then the rows under it, until the block runs out. Stop at the next
    ' heading rather than running into the following sample's block.
    For row = headerRow + 1 To LesserOf(headerRow + 30, LastRow(sheet))
        If StrComp(Normalise(CellString(sheet, row, typeCol)), wanted, vbTextCompare) = 0 Then
            ZpInBlock = OnlyDigitsOf(CellString(sheet, row, docCol))
            Exit Function
        End If
        If InStr(1, CellString(sheet, row, 2), "sample", vbTextCompare) > 0 Then Exit Function
    Next row
End Function

Private Function FirstWord(ByVal text As String) As String
    Dim parts() As String

    parts = Split(Replace(Trim$(text), ";", ","), ",")
    FirstWord = Trim$(parts(LBound(parts)))
End Function

Private Function Normalise(ByVal text As String) As String
    Normalise = UCase$(Trim$(text))
End Function

Private Function OnlyDigitsOf(ByVal text As String) As String
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch >= "0" And ch <= "9" Then OnlyDigitsOf = OnlyDigitsOf & ch
    Next i
End Function
