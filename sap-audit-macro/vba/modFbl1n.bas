Attribute VB_Name = "modFbl1n"
'=======================================================================
' modFbl1n -- steps 6 to 8: the ZP payments of one batch.
'
' Given the ZP document numbers pulled out of the Payment Usage export,
' this opens FBL1N for the company code and month, filters to those
' documents, sorts by amount descending, exports the list as evidence,
' and returns the largest payment in the batch.
'
' N1.vbs settled the selection screen, and corrected two predictions: the
' company code is KD_BUKRS, not DD_BUKRS, and the document number is not on
' the main selection screen at all -- it lives in dynamic selections, which
' have to be opened with tbar[1]/btn[16] before the multiple-selection
' arrow exists.
'
' It also showed that FBL1N renders as a CLASSIC LIST here, addressed as
' lbl[x,y], not as an ALV grid. That matters more than it sounds: there is
' no grid to read row by row, and the recording navigates it by clicking
' screen positions -- lbl[164,8], lbl[164,10]. Those coordinates are fixed
' screen rows. Replaying them across 56 batches of different sizes would
' open whichever document happened to land on row 8, silently, with no
' error to notice.
'
' So the list is not navigated by position. It is exported with the same
' List > Save/Send > File path used everywhere else, read back off disk to
' find the largest payment by value, and that payment is then reached by
' its document number. Slower by one export; correct regardless of how many
' rows the batch has.
'=======================================================================
Option Explicit

Public Type ZpPayment
    Found As Boolean
    GridRow As Long
    DocumentNumber As String
    Vendor As String
    VendorName As String
    Amount As Double
    RowsConsidered As Long
    ExportFile As String
    Skipped As Boolean            ' the stage is not configured
    Notes As String
End Type

'-----------------------------------------------------------------------
' True when enough of the Screen Map is filled in to attempt the stage.
'-----------------------------------------------------------------------
Public Function IsConfigured() As Boolean
    IsConfigured = (Len(modConfig.ElementIdOrBlank("Fbl1n.CompanyCode")) > 0) And _
                   (Len(modConfig.ElementIdOrBlank("Fbl1n.ExecuteButton")) > 0)
End Function

'-----------------------------------------------------------------------
' Steps 6 and 7, then return the largest ZP payment for step 8.
'-----------------------------------------------------------------------
Public Function LargestPaymentOfBatch(ByVal sampleIdx As Long, _
                                      ByVal zpNumbers As String, _
                                      ByVal dateFrom As Date, ByVal dateTo As Date, _
                                      ByVal folder As String, _
                                      ByVal fileStem As String) As ZpPayment
    Dim result As ZpPayment
    Dim count As Long

    count = CountNumbers(zpNumbers)

    If Not IsConfigured() Then
        result.Skipped = True
        result.Notes = "FBL1N is not mapped on the Screen Map, so the batch's ZP " & _
                       "payments were not listed. " & count & " ZP document number(s) " & _
                       "were collected and are in the Log, which is enough to run " & _
                       "FBL1N by hand: company code " & modConfig.Setting("Company code") & _
                       ", all items, posting date " & modUtil.SapDate(dateFrom) & " to " & _
                       modUtil.SapDate(dateTo) & ", document number = those numbers."
        modLog.LogAction sampleIdx, "FBL1N", result.Notes, "MANUAL", vbNullString
        modLog.LogAction sampleIdx, "ZP numbers", _
                     "The " & count & " number(s): " & Replace(zpNumbers, vbLf, ", "), _
                     "MANUAL", vbNullString
        LargestPaymentOfBatch = result
        Exit Function
    End If

    If count = 0 Then
        result.Notes = "No ZP document numbers were collected, so there is nothing to " & _
                       "look up in FBL1N."
        modLog.LogAction sampleIdx, "FBL1N", result.Notes, "ERROR", vbNullString
        LargestPaymentOfBatch = result
        Exit Function
    End If

    On Error GoTo Failed

    OpenSelectionScreen sampleIdx, dateFrom, dateTo
    EnterDocumentNumbers sampleIdx, zpNumbers, count

    modSafety.GuardedPress modConfig.ElementId("Fbl1n.ExecuteButton")

    If modSapConnect.StatusBarType() = "E" Or modSapConnect.StatusBarType() = "A" Then
        result.Notes = "FBL1N reported: " & modSapConnect.StatusBarText()
        modLog.LogAction sampleIdx, "FBL1N", result.Notes, "ERROR", vbNullString
        LargestPaymentOfBatch = result
        Exit Function
    End If

    ' Export the whole list, then decide from the file. Reading the file
    ' rather than the screen is what makes this safe across batches of
    ' different sizes -- see the note at the top of this module.
    result.ExportFile = modExport.ExportClassicList( _
        sampleIdx, folder, fileStem & "_ZP_payments.xlsx")

    If Len(result.ExportFile) = 0 Then
        result.Notes = "The FBL1N list did not export, so the largest payment of the " & _
                       "batch could not be identified."
        modLog.LogAction sampleIdx, "FBL1N", result.Notes, "ERROR", vbNullString
        LargestPaymentOfBatch = result
        Exit Function
    End If

    result = LargestInExport(sampleIdx, result)
    LargestPaymentOfBatch = result
    Exit Function

Failed:
    result.Notes = "The FBL1N stage stopped: " & Err.Description
    modLog.LogAction sampleIdx, "FBL1N failed", Err.Description, "ERROR", vbNullString
    LargestPaymentOfBatch = result
End Function

'-----------------------------------------------------------------------
' The largest payment in the exported list.
'-----------------------------------------------------------------------
Private Function LargestInExport(ByVal sampleIdx As Long, _
                                 ByRef partial As ZpPayment) As ZpPayment
    Dim result As ZpPayment
    Dim largest As ListRow

    result = partial

    largest = modExportRead.LargestRow(result.ExportFile, sampleIdx, _
                                       "ZP list amount column", _
                                       "ZP list vendor column", _
                                       "ZP list document column")

    If Not largest.Found Then
        result.Notes = "Exported " & result.ExportFile & " but no amount could be read " & _
                       "from it. Open the file and copy its column headings into the " & _
                       "'ZP list ...' settings on the Control sheet."
        modLog.LogAction sampleIdx, "FBL1N", result.Notes, "ERROR", result.ExportFile
        LargestInExport = result
        Exit Function
    End If

    result.Found = True
    result.Amount = largest.Amount
    result.DocumentNumber = largest.DocumentNumber
    result.VendorName = largest.Supplier
    result.RowsConsidered = largest.RowsConsidered

    modLog.LogAction sampleIdx, "FBL1N", _
                 "Largest of " & largest.RowsConsidered & " payments in the batch: " & _
                 Format$(largest.Amount, "#,##0.00") & _
                 IIf(Len(largest.DocumentNumber) > 0, _
                     ", document " & largest.DocumentNumber, "") & _
                 IIf(Len(largest.Supplier) > 0, ", " & largest.Supplier, ""), _
                 "OK", result.ExportFile

    LargestInExport = result
End Function

'-----------------------------------------------------------------------
' Step 6 -- the selection screen
'-----------------------------------------------------------------------
Private Sub OpenSelectionScreen(ByVal sampleIdx As Long, _
                                ByVal dateFrom As Date, ByVal dateTo As Date)
    Dim radioId As String

    modSafety.StartTransaction "FBL1N"

    modSapConnect.Element(modConfig.ElementId("Fbl1n.CompanyCode")).Text = _
        modConfig.Setting("Company code")

    ' 'All items' rather than open or cleared, so a payment that has since been
    ' reversed still shows up.
    radioId = modConfig.ElementIdOrBlank("Fbl1n.AllItemsRadio")
    If Len(radioId) > 0 Then
        If modSapConnect.Exists(radioId) Then
            On Error Resume Next
            modSapConnect.Element(radioId).Select
            On Error GoTo 0
        End If
    End If

    SetIfMapped "Fbl1n.PostingDateFrom", modUtil.SapDate(dateFrom)
    SetIfMapped "Fbl1n.PostingDateTo", modUtil.SapDate(dateTo)

    modLog.LogAction sampleIdx, "FBL1N", _
                 "Selection: company code " & modConfig.Setting("Company code") & _
                 ", all items, posting date " & modUtil.SapDate(dateFrom) & " to " & _
                 modUtil.SapDate(dateTo), "OK", vbNullString
End Sub

Private Sub SetIfMapped(ByVal mapKey As String, ByVal value As String)
    Dim elementId As String

    elementId = modConfig.ElementIdOrBlank(mapKey)
    If Len(elementId) = 0 Then Exit Sub
    If Not modSapConnect.Exists(elementId) Then Exit Sub

    modSapConnect.Element(elementId).Text = value
End Sub

'-----------------------------------------------------------------------
' Feed the ZP numbers into the document-number filter.
'
' One number goes straight into the field. Several need the multiple-
' selection dialog, and the only practical way to fill that from a script
' is its 'upload from clipboard' button -- typing rows one at a time is
' both slow and dependent on the dialog's scroll position.
'-----------------------------------------------------------------------
Private Sub EnterDocumentNumbers(ByVal sampleIdx As Long, ByVal zpNumbers As String, _
                                 ByVal count As Long)
    Dim singleFieldId As String
    Dim multiButtonId As String
    Dim pasteButtonId As String
    Dim confirmId As String

    singleFieldId = modConfig.ElementIdOrBlank("Fbl1n.DocNumberFrom")
    multiButtonId = modConfig.ElementIdOrBlank("Fbl1n.DocNumberMultiSelect")

    If count = 1 And Len(singleFieldId) > 0 Then
        If modSapConnect.Exists(singleFieldId) Then
            modSapConnect.Element(singleFieldId).Text = zpNumbers
            modLog.LogAction sampleIdx, "FBL1N", _
                         "One ZP document (" & zpNumbers & ") typed straight into the " & _
                         "document-number field.", "OK", vbNullString
            Exit Sub
        End If
    End If

    If Len(multiButtonId) = 0 Then
        Err.Raise vbObjectError + 570, "modFbl1n.EnterDocumentNumbers", _
                  count & " ZP documents need the multiple-selection dialog, but " & _
                  "Fbl1n.DocNumberMultiSelect is blank on the Screen Map. Record that " & _
                  "arrow button next to the document-number field."
    End If

    PutOnClipboard zpNumbers

    modSafety.GuardedPress multiButtonId
    modSafety.AssertPopupKnown

    pasteButtonId = modConfig.ElementIdOrBlank("MultiSel.PasteFromClipboard")
    confirmId = modConfig.ElementIdOrBlank("MultiSel.Confirm")

    If Len(pasteButtonId) = 0 Then
        Err.Raise vbObjectError + 571, "modFbl1n.EnterDocumentNumbers", _
                  "MultiSel.PasteFromClipboard is blank on the Screen Map, so the " & _
                  count & " ZP document numbers could not be pasted into the " & _
                  "multiple-selection dialog. They are listed in the Log."
    End If

    modSafety.GuardedPress pasteButtonId
    modSapConnect.WaitForSap

    If Len(confirmId) > 0 Then
        modSafety.GuardedPress confirmId
    Else
        modSafety.GuardedSendVKey "wnd[1]", 8      ' Execute/Copy on that dialog
    End If

    modLog.LogAction sampleIdx, "FBL1N", _
                 count & " ZP document numbers pasted into the multiple-selection " & _
                 "dialog via the clipboard.", "OK", vbNullString
End Sub

' Put newline-separated text on the clipboard without needing a reference to
' MSForms: write it down a column of a scratch sheet and copy the range.
Private Sub PutOnClipboard(ByVal text As String)
    Dim sheet As Worksheet
    Dim numbers() As String
    Dim i As Long
    Dim previousAlerts As Boolean

    numbers = Split(text, vbLf)

    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error Resume Next
    Set sheet = ThisWorkbook.Worksheets("_clipboard")
    On Error GoTo 0

    If sheet Is Nothing Then
        Set sheet = ThisWorkbook.Worksheets.Add
        sheet.Name = "_clipboard"
    End If

    sheet.Visible = xlSheetVeryHidden
    sheet.Cells.Clear

    For i = LBound(numbers) To UBound(numbers)
        ' Text format, so a long document number keeps its leading zeros and
        ' does not come back as 1.9E+09.
        sheet.Cells(i + 1, 1).NumberFormat = "@"
        sheet.Cells(i + 1, 1).Value = numbers(i)
    Next i

    sheet.Range(sheet.Cells(1, 1), sheet.Cells(UBound(numbers) - LBound(numbers) + 1, 1)).Copy

    Application.DisplayAlerts = previousAlerts
End Sub

Private Function CountNumbers(ByVal zpNumbers As String) As Long
    If Len(Trim$(zpNumbers)) = 0 Then Exit Function
    CountNumbers = UBound(Split(zpNumbers, vbLf)) - LBound(Split(zpNumbers, vbLf)) + 1
End Function

'-----------------------------------------------------------------------
' Step 7 -- sort descending on the amount column
'-----------------------------------------------------------------------
Private Sub SortByAmountDescending()
    Dim grid As Object
    Dim amountColumn As String

    Set grid = modSapConnect.Element(modConfig.ElementId("Fbl1n.ResultGrid"))
    amountColumn = modConfig.ElementIdOrBlank("Fbl1n.Col.Amount")
    If Len(amountColumn) = 0 Then Exit Sub

    ' Best effort: the export reads better sorted, but nothing downstream
    ' depends on the order.
    On Error Resume Next
    grid.setCurrentCell -1, amountColumn
    grid.selectColumn amountColumn
    grid.pressToolbarButton "&SORT_DSC"
    On Error GoTo 0

    modSapConnect.WaitForSap
End Sub

'-----------------------------------------------------------------------
' Step 8 -- the largest payment, by reading the grid
'-----------------------------------------------------------------------
Private Function FindLargest(ByVal sampleIdx As Long, ByRef partial As ZpPayment) As ZpPayment
    Dim result As ZpPayment
    Dim grid As Object
    Dim rows As Long, r As Long
    Dim amountColumn As String
    Dim value As Double, best As Double
    Dim ties As Long

    result = partial

    Set grid = modSapConnect.Element(modConfig.ElementId("Fbl1n.ResultGrid"))
    amountColumn = modConfig.ElementId("Fbl1n.Col.Amount")

    On Error Resume Next
    rows = grid.RowCount
    On Error GoTo 0

    For r = 0 To rows - 1
        value = Abs(modUtil.ParseSapAmount(GridCell(grid, r, amountColumn)))
        If value > 0 Then result.RowsConsidered = result.RowsConsidered + 1

        If value > best Then
            best = value
            ties = 1
            result.Found = True
            result.GridRow = r
            result.Amount = value
            result.DocumentNumber = GridCellIfMapped(grid, r, "Fbl1n.Col.DocNumber")
            result.Vendor = GridCellIfMapped(grid, r, "Fbl1n.Col.Vendor")
            result.VendorName = GridCellIfMapped(grid, r, "Fbl1n.Col.VendorName")
        ElseIf value = best And value > 0 Then
            ties = ties + 1
        End If
    Next r

    If Not result.Found Then
        result.Notes = "FBL1N returned " & rows & " row(s) but no readable amount in " & _
                       "column " & amountColumn & ". Check Fbl1n.Col.Amount."
        modLog.LogAction sampleIdx, "FBL1N", result.Notes, "ERROR", vbNullString
        FindLargest = result
        Exit Function
    End If

    If ties > 1 Then
        result.Notes = ties & " ZP payments in the batch share the largest amount " & _
                       Format$(best, "#,##0.00") & ". The first was taken; check by hand " & _
                       "whether that is the one the auditor means."
        modLog.LogAction sampleIdx, "FBL1N", result.Notes, "MANUAL", vbNullString
    End If

    modLog.LogAction sampleIdx, "FBL1N", _
                 "Largest of " & result.RowsConsidered & " ZP payments in the batch: " & _
                 Format$(result.Amount, "#,##0.00") & _
                 IIf(Len(result.DocumentNumber) > 0, ", document " & result.DocumentNumber, "") & _
                 IIf(Len(result.VendorName) > 0, ", vendor " & result.VendorName, _
                     IIf(Len(result.Vendor) > 0, ", vendor " & result.Vendor, "")), _
                 "OK", vbNullString

    FindLargest = result
End Function

' Open the largest ZP payment, so the invoice route can start from it.
Public Sub OpenPayment(ByVal gridRow As Long)
    Dim grid As Object
    Dim amountColumn As String

    Set grid = modSapConnect.Element(modConfig.ElementId("Fbl1n.ResultGrid"))
    amountColumn = modConfig.ElementId("Fbl1n.Col.Amount")

    grid.setCurrentCell gridRow, amountColumn
    modSapConnect.WaitForSap

    grid.doubleClickCurrentCell
    modSapConnect.WaitForSap
    modSafety.AssertPopupKnown
End Sub

Private Function GridCell(ByVal grid As Object, ByVal row As Long, _
                          ByVal columnName As String) As String
    On Error Resume Next
    GridCell = grid.GetCellValue(row, columnName)
    On Error GoTo 0
End Function

Private Function GridCellIfMapped(ByVal grid As Object, ByVal row As Long, _
                                  ByVal mapKey As String) As String
    Dim columnName As String

    columnName = modConfig.ElementIdOrBlank(mapKey)
    If Len(columnName) = 0 Then Exit Function

    GridCellIfMapped = Trim$(GridCell(grid, row, columnName))
End Function

' Column names as FBL1N actually reports them on this release. Run it once
' with an FBL1N result on screen and paste what it prints into the Screen Map.
'
' Named for its transaction rather than 'DumpGridColumns', which collided
' with modFeban's and left both showing ambiguously in the Macro dialog.
Public Sub DumpFbl1nColumns()
    modConfig.LoadScreenMap
    modSapConnect.SapAttach
    modFeban.DumpColumnsOf modConfig.ElementId("Fbl1n.ResultGrid"), "FBL1N result grid"
End Sub
