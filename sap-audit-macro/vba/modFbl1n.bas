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
' Open the largest payment of the batch, staying on the FBL1N list.
'
' The recordings reach a row with lbl[164,10] -- a fixed screen coordinate,
' which lands on whatever happens to be the tenth line and so picks the
' wrong document on a batch of a different size. But the fix is not to
' leave FBL1N: it is to find the row by its CONTENT rather than its
' position. The document number is already known from the export, so walk
' the labels on screen and focus the one whose text is that number, then
' F2 -- exactly the drill the recordings use, just aimed properly.
'
' Same principle as finding a column by what it holds rather than by its
' translated caption. Position and language are both accidents of display.
'
' FB03 is only the fallback, for when the number is not on screen at all --
' a long list that has scrolled, most likely. It is display-only and needs
' no extra rights, but it is not the normal path.
'-----------------------------------------------------------------------
Public Function OpenPaymentOnList(ByVal sampleIdx As Long, _
                                  ByVal documentNumber As String) As Boolean
    Dim labelId As String

    If Len(documentNumber) = 0 Then Exit Function

    labelId = LabelShowing(documentNumber)
    If Len(labelId) = 0 Then
        modLog.LogAction sampleIdx, "Open payment", _
                     "Document " & documentNumber & " is not among the labels on the " & _
                     "FBL1N list -- the list may have scrolled past it. Falling back to " & _
                     "opening it in FB03.", "MANUAL", vbNullString
        Exit Function
    End If

    On Error GoTo Failed

    modSapConnect.Element(labelId).SetFocus
    modSafety.GuardedSendVKey "wnd[0]", 2
    modSafety.AssertPopupKnown

    modLog.LogAction sampleIdx, "Open payment", _
                 "Opened payment " & documentNumber & " from the FBL1N list via " & _
                 labelId & " (found by its text, not by position)", "OK", vbNullString

    OpenPaymentOnList = True
    Exit Function

Failed:
    modLog.LogAction sampleIdx, "Open payment", _
                 "Could not open " & documentNumber & " from the list: " & Err.Description, _
                 "ERROR", vbNullString
End Function

' The id of the label on the current screen whose text is this document
' number. SAP pads document numbers with leading zeros in some columns and
' not others, so compare on digits.
Private Function LabelShowing(ByVal documentNumber As String) As String
    Dim area As Object, child As Object
    Dim wanted As String

    wanted = OnlyDigits(documentNumber)
    If Len(wanted) = 0 Then Exit Function

    If Not modSapConnect.Exists("wnd[0]/usr") Then Exit Function
    Set area = modSapConnect.Element("wnd[0]/usr")

    On Error Resume Next
    For Each child In area.Children
        If child.Type = "GuiLabel" Then
            If OnlyDigits(child.Text) = wanted Then
                LabelShowing = child.Id
                Exit For
            End If
        End If
    Next child
    On Error GoTo 0
End Function

Private Function OnlyDigits(ByVal text As String) As String
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch >= "0" And ch <= "9" Then OnlyDigits = OnlyDigits & ch
    Next i
End Function

'-----------------------------------------------------------------------
' Fallback: open it in FB03 by number. Display-only.
'-----------------------------------------------------------------------
Public Function OpenPaymentByDocument(ByVal sampleIdx As Long, _
                                      ByVal documentNumber As String, _
                                      ByVal fiscalYear As String) As Boolean
    Dim docId As String, bukrsId As String, yearId As String

    If Len(documentNumber) = 0 Then
        modLog.LogAction sampleIdx, "Open payment", _
                     "The largest payment carries no document number, so it could not be " & _
                     "opened. Check the document-number column of the exported list.", _
                     "ERROR", vbNullString
        Exit Function
    End If

    docId = modConfig.ElementIdOrBlank("FB03.DocNumber")
    If Len(docId) = 0 Then
        modLog.LogAction sampleIdx, "Open payment", _
                     "FB03.DocNumber is blank on the Screen Map, so payment " & _
                     documentNumber & " could not be opened. Run modProbe on the FB03 " & _
                     "entry screen and fill the FB03.* rows in.", "MANUAL", vbNullString
        Exit Function
    End If

    On Error GoTo Failed

    modSafety.StartTransaction "FB03"

    If Not modSapConnect.Exists(docId) Then
        modLog.LogAction sampleIdx, "Open payment", _
                     "FB03 opened but " & docId & " is not on it.", "ERROR", vbNullString
        Exit Function
    End If

    modSapConnect.Element(docId).Text = documentNumber

    bukrsId = modConfig.ElementIdOrBlank("FB03.CompanyCode")
    If Len(bukrsId) > 0 Then
        If modSapConnect.Exists(bukrsId) Then
            modSapConnect.Element(bukrsId).Text = modConfig.Setting("Company code")
        End If
    End If

    yearId = modConfig.ElementIdOrBlank("FB03.FiscalYear")
    If Len(yearId) > 0 And Len(fiscalYear) > 0 Then
        If modSapConnect.Exists(yearId) Then
            modSapConnect.Element(yearId).Text = fiscalYear
        End If
    End If

    modSafety.GuardedSendVKey "wnd[0]", 0
    modSafety.AssertPopupKnown

    If modSapConnect.StatusBarType() = "E" Then
        modLog.LogAction sampleIdx, "Open payment", _
                     "FB03 could not display " & documentNumber & ": " & _
                     modSapConnect.StatusBarText(), "ERROR", vbNullString
        Exit Function
    End If

    modLog.LogAction sampleIdx, "Open payment", _
                 "Opened payment document " & documentNumber & " in FB03", _
                 "OK", vbNullString

    OpenPaymentByDocument = True
    Exit Function

Failed:
    modLog.LogAction sampleIdx, "Open payment", Err.Description, "ERROR", vbNullString
End Function

'-----------------------------------------------------------------------
' There is no FBL1N grid to dump on this system.
'
' FBL1N returns a classic list here -- lbl[x,y] cells, not an ALV control --
' so it has no queryable column set. The column names that matter are the
' headings of the exported file, which is what the macro actually reads.
'-----------------------------------------------------------------------
Public Sub DumpFbl1nColumns()
    MsgBox "FBL1N returns a classic list on this system, not an ALV grid, so there " & _
           "are no grid columns to read." & vbCrLf & vbCrLf & _
           "The macro does not read that screen at all -- it exports the list and reads " & _
           "the file. To check the column names, open the exported " & _
           "*_ZP_payments.xlsx and look at its headings, then put them in the " & _
           "'ZP list ...' settings on the Control sheet if the Log says it picked " & _
           "the wrong ones." & vbCrLf & vbCrLf & _
           "modFeban.DumpGridColumns still works, because FEBAN does use an ALV grid.", _
           vbInformation, "FBL1N columns"
End Sub
