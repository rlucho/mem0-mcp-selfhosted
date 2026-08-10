Attribute VB_Name = "modScfVendors"
'=======================================================================
' modScfVendors -- who was actually paid in a supply-chain-finance batch.
'
' A confirming (SCF) settlement hides its suppliers behind the bank. The
' company pays ONE vendor -- Citibank, Santander -- and the invoices that
' payment covers sit on the BANK's vendor account as KA documents, not on
' the suppliers' accounts. Ask FBL1N who was paid and it answers 'Citibank
' Europe plc', which is true and useless.
'
' The link is the Reference (XBLNR) on each of those KA documents: it
' carries the ORIGINAL invoice's document number. So the walk is:
'
'   AB payment document
'     -> FB03, Environment > Payment Usage      the KA documents it settled
'       -> their Reference column               the original invoice numbers
'         -> FBL1N, no vendor, wide dates       the real suppliers
'
' Two things sink this if you do it by hand, and both are handled here:
' leaving the confirming party in FBL1N's vendor field, which just returns
' the bank again, and restricting the posting date to the month of the
' payment, when the invoices are from the months before it. In the batch
' this was built from, the payment is September and the invoices run May to
' July -- a September filter finds nothing at all and looks exactly like
' 'these documents do not exist'.
'
' Read-only. It runs FB03 and FBL1N, both display transactions, and writes
' nothing but files on your own filesystem.
'
' Standalone on purpose: no other module, no reference to set. Run
' FindVendors from Developer > Macros, or press the button AddButton makes.
'=======================================================================
Option Explicit

Private Const SHEET_INPUT As String = "Input"
Private Const SHEET_VENDORS As String = "Vendors"
Private Const SHEET_LOG As String = "Log"
Private Const SHEET_MAP As String = "Screen Map"

Private Const FIRST_DOC_ROW As Long = 12
Private Const HANDOVER_FOLDER As String = "scf-vendors-handover"

#If VBA7 Then
    Private Declare PtrSafe Sub SleepApi Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
#Else
    Private Declare Sub SleepApi Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
#End If

Private gSession As Object
Private mSequence As Long
Private mLogRow As Long
Private mVendorRow As Long

' The export just read, held as a plain array so the workbook can be closed
' before anything else opens one -- Excel will not hold two workbooks with
' the same name, and every export here is called the same thing.
Private mCells() As String
Private mRows As Long
Private mCols As Long

'=======================================================================
' Entry point
'=======================================================================
Public Sub FindVendors()
    Dim row As Long, lastRow As Long
    Dim document As String, fiscalYear As String
    Dim folder As String
    Dim walked As Long, failed As Long, found As Long

    On Error GoTo Failed

    Set gSession = Nothing
    mVendorRow = 0
    StartSheets

    folder = Setting("Output folder")
    If Len(folder) = 0 Then
        MsgBox "Fill in 'Output folder' on the " & SHEET_INPUT & " sheet first.", _
               vbExclamation, "SCF vendors"
        Exit Sub
    End If
    EnsureFolder folder

    If Not SapAttach() Then Exit Sub

    SweepHandover
    Application.ScreenUpdating = False

    lastRow = LastInputRow()
    For row = FIRST_DOC_ROW To lastRow
        document = OnlyDigits(CellText(Sheets(SHEET_INPUT), row, 2))
        If Len(document) > 0 Then
            fiscalYear = Trim$(CellText(Sheets(SHEET_INPUT), row, 3))
            Application.StatusBar = "AB document " & document & " ..."

            found = WalkOneDocument(document, fiscalYear, folder)
            Sheets(SHEET_INPUT).Cells(row, 4).Value = found & " supplier invoice(s)"

            If found > 0 Then walked = walked + 1 Else failed = failed + 1
            CloseExportWorkbooks
        End If
    Next row

    SweepHandover
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Sheets(SHEET_VENDORS).Activate

    MsgBox "Finished." & vbCrLf & vbCrLf & _
           "AB documents that produced suppliers : " & walked & vbCrLf & _
           "AB documents that produced none      : " & failed & vbCrLf & _
           "Supplier invoice rows written        : " & mVendorRow & vbCrLf & vbCrLf & _
           "The '" & SHEET_VENDORS & "' sheet has the list; the exports SAP produced " & _
           "are in " & folder & ". See '" & SHEET_LOG & "' if anything came back empty.", _
           IIf(failed = 0, vbInformation, vbExclamation), "SCF vendors"
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Note 0, "STOPPED", Err.Description
    MsgBox "Stopped: " & Err.Description, vbCritical, "SCF vendors"
End Sub

'-----------------------------------------------------------------------
' One AB payment document, end to end. Returns how many supplier invoice
' rows it produced.
'-----------------------------------------------------------------------
Private Function WalkOneDocument(ByVal document As String, ByVal fiscalYear As String, _
                                 ByVal folder As String) As Long
    Dim usagePath As String, invoicePath As String
    Dim references As String
    Dim count As Long

    On Error GoTo DocumentFailed

    If Not OpenInFb03(document, fiscalYear) Then
        Note document, "FB03", "Could not open the document in any fiscal year tried."
        Exit Function
    End If

    If Not OpenPaymentUsage(document) Then Exit Function

    usagePath = ExportCurrentList(document, folder, document & " - 1 - payment usage.xlsx")
    If Len(usagePath) = 0 Then
        Note document, "Export", "The payment usage list did not export."
        Exit Function
    End If

    references = ReferencesIn(document, usagePath, count)
    If count = 0 Then
        Note document, "References", _
             "No references were found in the payment usage list, so there is nothing " & _
             "to look up. Open " & usagePath & " and check the Reference column -- if " & _
             "it is headed something else here, add that heading to 'Reference column' " & _
             "on the " & SHEET_INPUT & " sheet."
        Exit Function
    End If

    Note document, "References", count & " reference(s) taken from the payment usage list."

    If Not RunFbl1n(document, references, count) Then Exit Function

    invoicePath = ExportCurrentList(document, folder, document & " - 2 - supplier invoices.xlsx")
    If Len(invoicePath) = 0 Then
        Note document, "Export", "The FBL1N list did not export."
        Exit Function
    End If

    WalkOneDocument = WriteVendors(document, invoicePath)
    Exit Function

DocumentFailed:
    Note document, "FAILED", Err.Description
End Function

'=======================================================================
' SAP
'=======================================================================
Private Function SapAttach() As Boolean
    Dim gui As Object, connection As Object
    Dim expected As String

    On Error GoTo NoSap

    Set gui = GetObject("SAPGUI").GetScriptingEngine
    Set connection = gui.Children(0)
    Set gSession = connection.Children(0)

    expected = Setting("Expected SAP system ID (SID)")
    If Len(expected) > 0 Then
        If StrComp(gSession.Info.SystemName, expected, vbTextCompare) <> 0 Then
            MsgBox "The attached SAP session is " & gSession.Info.SystemName & _
                   ", not " & expected & "." & vbCrLf & vbCrLf & _
                   "Nothing has been read. Log into the right system, or clear " & _
                   "'Expected SAP system ID (SID)' on the " & SHEET_INPUT & " sheet.", _
                   vbCritical, "SCF vendors"
            Exit Function
        End If
    End If

    Note 0, "SAP", "Attached to " & gSession.Info.SystemName & _
                   " client " & gSession.Info.Client & " as " & gSession.Info.user
    SapAttach = True
    Exit Function

NoSap:
    MsgBox "Could not attach to a running SAP GUI session." & vbCrLf & vbCrLf & _
           "Log into SAP first, and make sure scripting is enabled on both the " & _
           "server (sapgui/user_scripting) and this PC (Options > Accessibility & " & _
           "Scripting).", vbCritical, "SCF vendors"
End Function

Private Function Exists(ByVal elementId As String) As Boolean
    Dim probe As Object

    If Len(elementId) = 0 Then Exit Function
    On Error Resume Next
    Set probe = gSession.findById(elementId)
    On Error GoTo 0
    Exists = Not probe Is Nothing
End Function

Private Function El(ByVal elementId As String) As Object
    Set El = gSession.findById(elementId)
End Function

Private Sub WaitSap()
    On Error Resume Next
    Do While gSession.Busy
        SleepApi 200
    Loop
    On Error GoTo 0
    SleepApi 150
End Sub

Private Function StatusType() As String
    On Error Resume Next
    StatusType = gSession.findById("wnd[0]/sbar").MessageType
    On Error GoTo 0
End Function

Private Function StatusText() As String
    On Error Resume Next
    StatusText = gSession.findById("wnd[0]/sbar").Text
    On Error GoTo 0
End Function

Private Sub StartTransaction(ByVal code As String)
    gSession.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & code
    gSession.findById("wnd[0]").sendVKey 0
    WaitSap
End Sub

'-----------------------------------------------------------------------
' FB03 with the document number. The fiscal year is probed rather than
' assumed: GBHP posts September into FY 2026, so the calendar year is a
' guess and 'does not exist in fiscal year 2025' is the usual first answer.
'-----------------------------------------------------------------------
Private Function OpenInFb03(ByVal document As String, ByVal fiscalYear As String) As Boolean
    Dim baseYear As Long, tryYear As Long, attempt As Long
    Dim companyCode As String

    companyCode = Setting("Company code")
    baseYear = Val(fiscalYear)
    If baseYear = 0 Then baseYear = Year(Date)

    For attempt = 0 To 2
        Select Case attempt
            Case 0: tryYear = baseYear
            Case 1: tryYear = baseYear + 1
            Case 2: tryYear = baseYear - 1
        End Select

        StartTransaction "FB03"

        If Not Exists(ScreenId("FB03.DocNumber")) Then
            Note document, "FB03", "The FB03 entry screen did not appear."
            Exit Function
        End If

        El(ScreenId("FB03.DocNumber")).Text = document
        If Exists(ScreenId("FB03.CompanyCode")) Then _
            El(ScreenId("FB03.CompanyCode")).Text = companyCode
        If Exists(ScreenId("FB03.FiscalYear")) Then _
            El(ScreenId("FB03.FiscalYear")).Text = CStr(tryYear)

        gSession.findById("wnd[0]").sendVKey 0
        WaitSap

        If StatusType() <> "E" And StatusType() <> "A" Then
            Note document, "FB03", "Opened in fiscal year " & tryYear & "."
            OpenInFb03 = True
            Exit Function
        End If
    Next attempt
End Function

'-----------------------------------------------------------------------
' Environment > Payment Usage, by NAME first.
'
' The recorded position is right on the screen it was recorded from and
' wrong on any other -- on a line-item detail the same slot is the G/L
' account master, which answers 'You are not authorized to use transaction
' FS03' and leaves you exactly where you were.
'-----------------------------------------------------------------------
Private Function OpenPaymentUsage(ByVal document As String) As Boolean
    Dim menuId As String
    Dim caption As String

    caption = Setting("Payment usage menu text")
    If Len(caption) > 0 Then menuId = MenuNamed("wnd[0]/mbar", caption, 0)

    If Len(menuId) = 0 Then
        menuId = ScreenId("PaymentUsage.Menu")
        If Not Exists(menuId) Then
            Note document, "Payment usage", _
                 "Neither a menu entry called """ & caption & """ nor " & menuId & _
                 " is on this screen, so the payment usage list could not be opened."
            Exit Function
        End If
    End If

    El(menuId).Select
    WaitSap

    If StatusType() = "E" Or StatusType() = "A" Then
        Note document, "Payment usage", _
             "SAP refused that menu entry: " & StatusText() & " -- so it is not " & _
             "Payment Usage on this screen."
        Exit Function
    End If

    OpenPaymentUsage = True
End Function

Private Function MenuNamed(ByVal elementId As String, ByVal caption As String, _
                           ByVal depth As Long) As String
    Dim control As Object, child As Object
    Dim found As String

    If depth > 4 Then Exit Function
    If Not Exists(elementId) Then Exit Function

    Set control = El(elementId)

    On Error Resume Next
    For Each child In control.Children
        If StrComp(Trim$(child.Text), Trim$(caption), vbTextCompare) = 0 Then
            MenuNamed = child.Id
            Exit Function
        End If
        found = MenuNamed(child.Id, caption, depth + 1)
        If Len(found) > 0 Then
            MenuNamed = found
            Exit Function
        End If
    Next child
    On Error GoTo 0
End Function

'-----------------------------------------------------------------------
' FBL1N for the original invoices.
'
' Company code but NO vendor, because the whole point is that the suppliers
' are not the confirming party. Posting dates come from the Input sheet and
' should be left wide: the invoices predate the payment, often by months.
'-----------------------------------------------------------------------
Private Function RunFbl1n(ByVal document As String, ByVal references As String, _
                          ByVal count As Long) As Boolean
    Dim dateFrom As String, dateTo As String

    StartTransaction "FBL1N"

    If Not Exists(ScreenId("Fbl1n.CompanyCode")) Then
        Note document, "FBL1N", "The FBL1N selection screen did not appear."
        Exit Function
    End If

    El(ScreenId("Fbl1n.CompanyCode")).Text = Setting("Company code")

    ' Leave the vendor field alone. Anything in it and this returns the bank.
    If Exists(ScreenId("Fbl1n.AllItemsRadio")) Then El(ScreenId("Fbl1n.AllItemsRadio")).Select

    dateFrom = Trim$(Setting("Invoice posting date from"))
    dateTo = Trim$(Setting("Invoice posting date to"))
    If Exists(ScreenId("Fbl1n.PostingDateFrom")) Then _
        El(ScreenId("Fbl1n.PostingDateFrom")).Text = dateFrom
    If Exists(ScreenId("Fbl1n.PostingDateTo")) Then _
        El(ScreenId("Fbl1n.PostingDateTo")).Text = dateTo

    EnterReferences document, references, count

    El(ScreenId("Fbl1n.ExecuteButton")).press
    WaitSap

    If StatusType() = "E" Or StatusType() = "A" Then
        Note document, "FBL1N", "FBL1N reported: " & StatusText()
        Exit Function
    End If

    ' 'No items selected' is an ordinary message, not an error, and FBL1N
    ' simply stays on its selection screen. While the company-code field is
    ' still there, nothing was found.
    If Exists(ScreenId("Fbl1n.CompanyCode")) Then
        Note document, "FBL1N", _
             "FBL1N found no line items for those " & count & " reference(s) between " & _
             dateFrom & " and " & dateTo & ". " & StatusText() & " Widen the posting " & _
             "dates on the " & SHEET_INPUT & " sheet -- the invoices are older than the " & _
             "payment -- or check that the Reference column really holds SAP document " & _
             "numbers."
        Exit Function
    End If

    RunFbl1n = True
End Function

Private Sub EnterReferences(ByVal document As String, ByVal references As String, _
                            ByVal count As Long)
    Dim singleField As String, multiButton As String

    singleField = ScreenId("Fbl1n.DocNumberField")
    multiButton = ScreenId("Fbl1n.DocNumberMultiSelect")

    ' Dynamic selections are folded away until this is pressed, and the
    ' document-number field does not exist at all until they are open.
    If Not Exists(singleField) And Not Exists(multiButton) Then
        If Exists(ScreenId("Fbl1n.DynamicSelections")) Then
            El(ScreenId("Fbl1n.DynamicSelections")).press
            WaitSap
        End If
    End If

    If count = 1 And Exists(singleField) Then
        El(singleField).Text = references
        Exit Sub
    End If

    If Not Exists(multiButton) Then
        Err.Raise vbObjectError + 610, "modScfVendors.EnterReferences", _
                  count & " references need the multiple-selection dialog, but " & _
                  "Fbl1n.DocNumberMultiSelect is not on this screen."
    End If

    PutOnClipboard references

    El(multiButton).press
    WaitSap

    El(ScreenId("MultiSel.PasteFromClipboard")).press
    WaitSap
    El(ScreenId("MultiSel.Confirm")).press
    WaitSap

    Note document, "FBL1N", count & " reference(s) pasted into the multiple-selection dialog."
End Sub

Private Sub PutOnClipboard(ByVal text As String)
    Dim data As Object

    On Error Resume Next
    Set data = GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    data.SetText text
    data.PutInClipboard
    On Error GoTo 0
End Sub

'=======================================================================
' Export, and getting it back off disk
'=======================================================================
'-----------------------------------------------------------------------
' List > Save/Send > File, then the save dialog.
'
' SAP hands every export it writes to Excel through OLE, and Excel will not
' hold two workbooks with the same name whatever folder each is in. So SAP
' is given a scratch name that never repeats, and the file is copied to its
' proper name once written. Nothing here ever collides.
'-----------------------------------------------------------------------
Private Function ExportCurrentList(ByVal document As String, ByVal folder As String, _
                                   ByVal fileName As String) As String
    Dim menuId As String
    Dim scratchFolder As String, scratchName As String
    Dim written As String, target As String
    Dim bytes As Double

    menuId = ScreenId("Export.ListMenu")
    If Not Exists(menuId) Then
        Note document, "Export", "List > Save/Send > File is not on this screen (" & _
                                 menuId & "), so nothing could be exported."
        Exit Function
    End If

    El(menuId).Select
    WaitSap

    ' Some releases ask for a format first. Take the default.
    If Exists("wnd[1]") And Not Exists("wnd[1]/usr/ctxtDY_PATH") Then
        On Error Resume Next
        El("wnd[1]/tbar[0]/btn[0]").press
        On Error GoTo 0
        WaitSap
    End If

    If Not Exists(ScreenId("Save.Path")) Then
        Note document, "Export", "No save dialog appeared after the export command."
        Exit Function
    End If

    scratchFolder = HandoverFolder()
    EnsureFolder scratchFolder
    mSequence = mSequence + 1
    scratchName = Format$(mSequence, "0000") & " " & fileName

    El(ScreenId("Save.Path")).Text = scratchFolder
    El(ScreenId("Save.FileName")).Text = scratchName
    El(ScreenId("Save.GenerateButton")).press
    WaitSap

    written = JoinPath(scratchFolder, scratchName)
    If Not WaitForFile(written, 20) Then
        Note document, "Export", "SAP reported no error but " & written & " never appeared."
        Exit Function
    End If

    bytes = WaitUntilStable(written)
    If bytes = 0 Then
        Note document, "Export", "The export is empty -- the list had no rows."
        Exit Function
    End If

    target = JoinPath(folder, fileName)
    If Not CopyFileTo(written, target) Then
        Note document, "Export", "Wrote " & written & " but it could not be copied to " & target
        Exit Function
    End If

    Note document, "Export", "Wrote " & Format$(bytes / 1024, "0.0") & " KB to " & fileName
    ExportCurrentList = target
End Function

' Read an exported workbook into a plain string array, then close it. Held
' as an array so nothing stays open while the next export lands.
Private Function LoadExport(ByVal path As String) As Boolean
    Dim book As Workbook
    Dim area As Range
    Dim r As Long, c As Long
    Dim previousAlerts As Boolean

    Erase mCells
    mRows = 0
    mCols = 0

    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error GoTo Failed
    Set book = Application.Workbooks.Open(fileName:=path, UpdateLinks:=0, ReadOnly:=True)
    Set area = book.Worksheets(1).UsedRange

    mRows = area.Rows.Count
    mCols = area.Columns.Count
    ReDim mCells(1 To mRows, 1 To mCols)

    For r = 1 To mRows
        For c = 1 To mCols
            mCells(r, c) = Trim$(CStr(area.Cells(r, c).Text))
        Next c
    Next r

    book.Close SaveChanges:=False
    Application.DisplayAlerts = previousAlerts
    LoadExport = True
    Exit Function

Failed:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    Application.DisplayAlerts = previousAlerts
    On Error GoTo 0
End Function

Private Function ColumnHeaded(ByVal caption As String) As Long
    Dim r As Long, c As Long

    If Len(caption) = 0 Then Exit Function

    For r = 1 To LesserOf(mRows, 10)
        For c = 1 To mCols
            If StrComp(mCells(r, c), caption, vbTextCompare) = 0 Then
                ColumnHeaded = c
                Exit Function
            End If
        Next c
    Next r
End Function

Private Function HeaderRowOf(ByVal col As Long) As Long
    Dim r As Long

    For r = 1 To LesserOf(mRows, 10)
        If Len(mCells(r, col)) > 0 Then
            HeaderRowOf = r
            Exit Function
        End If
    Next r
End Function

'-----------------------------------------------------------------------
' The original invoice numbers, out of the payment usage list.
'
' Every row that carries a reference counts, whatever its document type --
' filtering on KA would be one more thing to get wrong on a system that
' calls it something else. Leading zeros go: SAP writes 0000243422 in the
' reference and 243422 in the document number, and FBL1N wants the latter.
'-----------------------------------------------------------------------
Private Function ReferencesIn(ByVal document As String, ByVal path As String, _
                              ByRef count As Long) As String
    Dim refCol As Long, headerRow As Long
    Dim r As Long
    Dim value As String
    Dim seen As Object
    Dim result As String

    count = 0
    If Not LoadExport(path) Then Exit Function

    refCol = ColumnHeaded(Setting("Reference column"))
    If refCol = 0 Then refCol = ColumnHeaded("Reference")
    If refCol = 0 Then Exit Function

    headerRow = HeaderRowOf(refCol)
    Set seen = CreateObject("Scripting.Dictionary")

    For r = headerRow + 1 To mRows
        value = OnlyDigits(mCells(r, refCol))
        Do While Len(value) > 1 And Left$(value, 1) = "0"
            value = Mid$(value, 2)
        Loop

        If Len(value) > 0 Then
            If Not seen.Exists(value) Then
                seen.Add value, True
                result = result & IIf(Len(result) > 0, vbLf, "") & value
                count = count + 1
            End If
        End If
    Next r

    ReferencesIn = result
End Function

'-----------------------------------------------------------------------
' The suppliers, out of the FBL1N export, onto the Vendors sheet.
'-----------------------------------------------------------------------
Private Function WriteVendors(ByVal document As String, ByVal path As String) As Long
    Dim sheet As Worksheet
    Dim supplierCol As Long, nameCol As Long, docCol As Long
    Dim amountCol As Long, typeCol As Long, refCol As Long, dateCol As Long
    Dim headerRow As Long, r As Long, out As Long
    Dim supplier As String

    If Not LoadExport(path) Then Exit Function

    supplierCol = ColumnHeaded(Setting("Supplier column"))
    If supplierCol = 0 Then supplierCol = ColumnHeaded("Supplier")
    nameCol = ColumnHeaded(Setting("Supplier name column"))
    If nameCol = 0 Then nameCol = ColumnHeaded("Name 1")
    docCol = ColumnHeaded("Document Number")
    amountCol = ColumnHeaded(Setting("Amount column"))
    If amountCol = 0 Then amountCol = ColumnHeaded("Amount in local currency")
    typeCol = ColumnHeaded("Document Type")
    refCol = ColumnHeaded("Reference")
    dateCol = ColumnHeaded("Document Date")

    If supplierCol = 0 And nameCol = 0 Then
        Note document, "Vendors", _
             "Neither a Supplier nor a Name 1 column is in " & path & ", so no vendors " & _
             "could be read. Open it and name the headings on the " & SHEET_INPUT & " sheet."
        Exit Function
    End If

    headerRow = HeaderRowOf(LesserOf(IIf(supplierCol > 0, supplierCol, nameCol), mCols))
    Set sheet = Sheets(SHEET_VENDORS)

    For r = headerRow + 1 To mRows
        supplier = mCells(r, LesserOf(IIf(supplierCol > 0, supplierCol, nameCol), mCols))
        If Len(supplier) > 0 Then
            mVendorRow = mVendorRow + 1
            out = mVendorRow + 4

            sheet.Cells(out, 1).Value = document
            If supplierCol > 0 Then sheet.Cells(out, 2).Value = mCells(r, supplierCol)
            If nameCol > 0 Then sheet.Cells(out, 3).Value = mCells(r, nameCol)
            If docCol > 0 Then sheet.Cells(out, 4).Value = mCells(r, docCol)
            If typeCol > 0 Then sheet.Cells(out, 5).Value = mCells(r, typeCol)
            If refCol > 0 Then sheet.Cells(out, 6).Value = mCells(r, refCol)
            If dateCol > 0 Then sheet.Cells(out, 7).Value = mCells(r, dateCol)
            If amountCol > 0 Then
                sheet.Cells(out, 8).Value = ParseAmount(mCells(r, amountCol))
                sheet.Cells(out, 8).NumberFormat = "#,##0.00"
            End If

            WriteVendors = WriteVendors + 1
        End If
    Next r

    Note document, "Vendors", WriteVendors & " supplier invoice row(s) written."
End Function

'=======================================================================
' Sheets
'=======================================================================
Private Sub StartSheets()
    Dim sheet As Worksheet

    Set sheet = Sheets(SHEET_VENDORS)
    If sheet.UsedRange.Rows.Count > 4 Then
        sheet.Range(sheet.Rows(5), sheet.Rows(sheet.Rows.Count)).ClearContents
    End If

    mLogRow = LastUsedRow(Sheets(SHEET_LOG), 1)
    If mLogRow < 4 Then mLogRow = 4
End Sub

Private Sub Note(ByVal document As Variant, ByVal step As String, ByVal detail As String)
    Dim sheet As Worksheet

    On Error Resume Next
    Set sheet = Sheets(SHEET_LOG)
    If sheet Is Nothing Then Exit Sub

    mLogRow = mLogRow + 1
    sheet.Cells(mLogRow, 1).Value = Now
    sheet.Cells(mLogRow, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    sheet.Cells(mLogRow, 2).Value = document
    sheet.Cells(mLogRow, 3).Value = step
    sheet.Cells(mLogRow, 4).Value = detail
    On Error GoTo 0
End Sub

Private Function Setting(ByVal label As String) As String
    Dim sheet As Worksheet
    Dim r As Long

    Set sheet = Sheets(SHEET_INPUT)
    For r = 1 To 30
        If StrComp(Trim$(CStr(sheet.Cells(r, 2).Value)), label, vbTextCompare) = 0 Then
            Setting = Trim$(CStr(sheet.Cells(r, 3).Value))
            Exit Function
        End If
    Next r
End Function

' A screen ID from the Screen Map sheet, falling back to the value this was
' built with. The fallbacks are what a live PP2 session actually answered
' to, so a missing or blank row costs nothing.
Private Function ScreenId(ByVal key As String) As String
    Dim sheet As Worksheet
    Dim r As Long

    On Error Resume Next
    Set sheet = Sheets(SHEET_MAP)
    On Error GoTo 0

    If Not sheet Is Nothing Then
        For r = 1 To 40
            If StrComp(Trim$(CStr(sheet.Cells(r, 1).Value)), key, vbTextCompare) = 0 Then
                ScreenId = Trim$(CStr(sheet.Cells(r, 2).Value))
                If Len(ScreenId) > 0 Then Exit Function
            End If
        Next r
    End If

    Select Case key
        Case "FB03.DocNumber":     ScreenId = "wnd[0]/usr/txtRF05L-BELNR"
        Case "FB03.CompanyCode":   ScreenId = "wnd[0]/usr/ctxtRF05L-BUKRS"
        Case "FB03.FiscalYear":    ScreenId = "wnd[0]/usr/txtRF05L-GJAHR"
        Case "PaymentUsage.Menu":  ScreenId = "wnd[0]/mbar/menu[5]/menu[3]"
        Case "Export.ListMenu":    ScreenId = "wnd[0]/mbar/menu[0]/menu[3]/menu[1]"
        Case "Save.Path":          ScreenId = "wnd[1]/usr/ctxtDY_PATH"
        Case "Save.FileName":      ScreenId = "wnd[1]/usr/ctxtDY_FILENAME"
        Case "Save.GenerateButton": ScreenId = "wnd[1]/tbar[0]/btn[0]"
        Case "Fbl1n.CompanyCode":  ScreenId = "wnd[0]/usr/ctxtKD_BUKRS-LOW"
        Case "Fbl1n.AllItemsRadio": ScreenId = "wnd[0]/usr/radX_AISEL"
        Case "Fbl1n.PostingDateFrom": ScreenId = "wnd[0]/usr/ctxtSO_BUDAT-LOW"
        Case "Fbl1n.PostingDateTo": ScreenId = "wnd[0]/usr/ctxtSO_BUDAT-HIGH"
        Case "Fbl1n.DynamicSelections": ScreenId = "wnd[0]/tbar[1]/btn[16]"
        Case "Fbl1n.ExecuteButton": ScreenId = "wnd[0]/tbar[1]/btn[8]"
        Case "MultiSel.PasteFromClipboard": ScreenId = "wnd[1]/tbar[0]/btn[24]"
        Case "MultiSel.Confirm":   ScreenId = "wnd[1]/tbar[0]/btn[8]"
        Case "Fbl1n.DocNumberField"
            ScreenId = "wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/" & _
                       "ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/" & _
                       "ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/txt%%DYN011-LOW"
        Case "Fbl1n.DocNumberMultiSelect"
            ScreenId = "wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/" & _
                       "ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/" & _
                       "ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/" & _
                       "btn%_%%DYN011_%_APP_%-VALU_PUSH"
    End Select
End Function

Private Function LastInputRow() As Long
    LastInputRow = LastUsedRow(Sheets(SHEET_INPUT), 2)
    If LastInputRow < FIRST_DOC_ROW Then LastInputRow = FIRST_DOC_ROW
End Function

Private Function LastUsedRow(ByVal sheet As Worksheet, ByVal col As Long) As Long
    LastUsedRow = sheet.Cells(sheet.Rows.Count, col).End(xlUp).row
End Function

Private Function CellText(ByVal sheet As Worksheet, ByVal row As Long, _
                          ByVal col As Long) As String
    CellText = Trim$(CStr(sheet.Cells(row, col).Value))
End Function

'=======================================================================
' Files
'=======================================================================
Private Function HandoverFolder() As String
    Dim base As String

    base = Environ$("TEMP")
    If Len(base) = 0 Then base = Setting("Output folder")
    HandoverFolder = JoinPath(base, HANDOVER_FOLDER)
End Function

Private Sub SweepHandover()
    Dim fso As Object, file As Object
    Dim folder As String
    Dim doomed As Collection, item As Variant

    folder = HandoverFolder()
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then Exit Sub

    CloseExportWorkbooks
    Set doomed = New Collection

    On Error Resume Next
    For Each file In fso.GetFolder(folder).Files
        doomed.Add file.path
    Next file
    For Each item In doomed
        fso.DeleteFile CStr(item), True
    Next item
    On Error GoTo 0
End Sub

' Close anything Excel opened out of the scratch folder. SAP hands each
' export over on its own schedule, so these arrive uninvited.
Private Sub CloseExportWorkbooks()
    Dim book As Workbook
    Dim doomed As Collection, item As Variant
    Dim folder As String
    Dim previousAlerts As Boolean

    folder = HandoverFolder()
    Set doomed = New Collection

    On Error Resume Next
    For Each book In Application.Workbooks
        If Not book Is ThisWorkbook Then
            If InStr(1, book.FullName, folder, vbTextCompare) = 1 Then doomed.Add book
        End If
    Next book

    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    For Each item In doomed
        item.Close SaveChanges:=False
    Next item
    Application.DisplayAlerts = previousAlerts
    On Error GoTo 0
End Sub

Private Function WaitForFile(ByVal path As String, ByVal maxSeconds As Double) As Boolean
    Dim waited As Double

    Do While waited < maxSeconds
        If FileThere(path) Then
            WaitForFile = True
            Exit Function
        End If
        SleepApi 500
        waited = waited + 0.5
    Loop
End Function

' SAP creates the file and then fills it, so 'it exists' is not 'it is
' finished'. Copying mid-write would put a truncated list in the folder.
Private Function WaitUntilStable(ByVal path As String) As Double
    Dim waited As Double
    Dim size As Double, previous As Double

    previous = -1
    Do While waited < 20
        size = FileSize(path)
        If size = previous And (size > 0 Or waited >= 3) Then
            WaitUntilStable = size
            Exit Function
        End If
        previous = size
        SleepApi 400
        waited = waited + 0.4
    Loop

    WaitUntilStable = FileSize(path)
End Function

Private Function CopyFileTo(ByVal source As String, ByVal target As String) As Boolean
    Dim fso As Object
    Dim attempt As Long
    Dim copied As Boolean

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(source) Then Exit Function

    For attempt = 1 To 4
        copied = False
        On Error Resume Next
        fso.CopyFile source, target, True
        copied = (fso.FileExists(target) And _
                  fso.GetFile(target).size = fso.GetFile(source).size)
        On Error GoTo 0

        If copied Then
            CopyFileTo = True
            Exit Function
        End If
        SleepApi 500
    Next attempt
End Function

Private Function FileThere(ByVal path As String) As Boolean
    FileThere = (Len(Dir$(path)) > 0)
End Function

Private Function FileSize(ByVal path As String) As Double
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then FileSize = fso.GetFile(path).size
End Function

Private Function EnsureFolder(ByVal path As String) As String
    Dim fso As Object
    Dim parent As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then
        parent = fso.GetParentFolderName(path)
        If Len(parent) > 0 And Not fso.FolderExists(parent) Then EnsureFolder parent
        fso.CreateFolder path
    End If
    EnsureFolder = path
End Function

Private Function JoinPath(ByVal folder As String, ByVal leaf As String) As String
    If Right$(folder, 1) = "\" Then
        JoinPath = folder & leaf
    Else
        JoinPath = folder & "\" & leaf
    End If
End Function

'=======================================================================
' Odds and ends
'=======================================================================
Private Function OnlyDigits(ByVal text As String) As String
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch >= "0" And ch <= "9" Then OnlyDigits = OnlyDigits & ch
    Next i
End Function

' SAP writes amounts in the logged-on user's format, so '-2.906.239,39' and
' '-2,906,239.39' are the same number. The last separator is the decimal one.
Private Function ParseAmount(ByVal text As String) As Double
    Dim cleaned As String
    Dim lastDot As Long, lastComma As Long
    Dim negative As Boolean

    cleaned = Replace(Replace(Trim$(text), " ", ""), Chr$(160), "")
    If Len(cleaned) = 0 Then Exit Function

    If Right$(cleaned, 1) = "-" Then
        negative = True
        cleaned = Left$(cleaned, Len(cleaned) - 1)
    ElseIf Left$(cleaned, 1) = "-" Then
        negative = True
        cleaned = Mid$(cleaned, 2)
    End If

    lastDot = InStrRev(cleaned, ".")
    lastComma = InStrRev(cleaned, ",")

    If lastComma > lastDot Then
        cleaned = Replace(cleaned, ".", "")
        cleaned = Replace(cleaned, ",", ".")
    Else
        cleaned = Replace(cleaned, ",", "")
    End If

    On Error Resume Next
    ParseAmount = CDbl(Val(cleaned))
    On Error GoTo 0

    If negative Then ParseAmount = -ParseAmount
End Function

Private Function LesserOf(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then LesserOf = a Else LesserOf = b
End Function

'=======================================================================
' One button, so the sheet carries its own trigger.
'=======================================================================
Public Sub AddButton()
    Dim sheet As Worksheet
    Dim shape As Object
    Dim i As Long

    Set sheet = Sheets(SHEET_INPUT)

    On Error Resume Next
    For i = sheet.Buttons.Count To 1 Step -1
        If Left$(sheet.Buttons(i).Name, 4) = "scf_" Then sheet.Buttons(i).Delete
    Next i
    On Error GoTo 0

    Set shape = sheet.Buttons.Add(430, 12, 190, 32)
    shape.Name = "scf_run"
    shape.Caption = "FIND THE VENDORS"
    shape.OnAction = "modScfVendors.FindVendors"

    On Error Resume Next
    shape.Placement = 3
    On Error GoTo 0

    MsgBox "Button added. Save the workbook so it stays.", vbInformation, "SCF vendors"
End Sub
