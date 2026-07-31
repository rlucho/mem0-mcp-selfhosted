Attribute VB_Name = "modReport"
'=======================================================================
' modReport -- one workbook per sample, for sending to the auditor.
'
' The evidence files are exports: raw SAP lists with thirty columns and no
' explanation of why they are in the folder. This builds the cover sheet
' that turns them into a pack -- what was asked for, what was found, how it
' was reached, and which file proves each step.
'
' Written into the sample's own folder as '01 - 8072447.42 - GBKM.xlsx', so
' the folder can be zipped and sent as it stands.
'
' Three sheets:
'   Report      the trail, top to bottom, with the files listed against it
'   Invoice     the PDF, embedded as an object, when there is one
'   Log         this sample's rows from the run's audit trail
'
' Nothing here talks to SAP. It runs after the chain has finished, reads
' what the chain returned, and never raises -- a report that will not build
' must not undo evidence already on disk.
'=======================================================================
Option Explicit

Private Const COL_LABEL As Long = 2
Private Const COL_VALUE As Long = 3

Private Const HEAD_FILL As Long = 15921906      ' pale blue
Private Const SECTION_FILL As Long = 14277081   ' pale grey

'-----------------------------------------------------------------------
' Build the report. Returns the path written, or "" if it could not be.
'-----------------------------------------------------------------------
Public Function BuildSampleReport(ByVal sampleIdx As Long, ByVal monthTab As String, _
                                  ByVal payDate As Date, ByVal amount As Double, _
                                  ByVal party As String, ByVal reference As String, _
                                  ByRef match As FebanMatch, ByRef chain As ChainResult, _
                                  ByVal folder As String) As String
    Dim book As Workbook
    Dim sheet As Worksheet
    Dim path As String
    Dim row As Long
    Dim previousAlerts As Boolean

    modUtil.EnsureFolder folder
    path = modUtil.JoinPath(folder, _
                            modUtil.ReportFileName(sampleIdx, amount, _
                                                   modConfig.Setting("Company code")))

    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error GoTo Failed

    Set book = Application.Workbooks.Add
    DropExtraSheets book

    Set sheet = book.Worksheets(1)
    sheet.Name = "Report"

    row = WriteHeader(sheet, sampleIdx, amount)
    row = WriteRequest(sheet, row, sampleIdx, monthTab, payDate, amount, party, reference)
    row = WriteTrail(sheet, row, match, chain)
    row = WriteOutcome(sheet, row, chain)
    row = WriteFiles(sheet, row, folder, chain)
    row = WriteProvenance(sheet, row, monthTab)

    Layout sheet

    AttachInvoice book, chain
    CopyLogRows book, sampleIdx

    sheet.Activate
    sheet.Range("A1").Select

    book.SaveAs fileName:=path, FileFormat:=51      ' xlOpenXMLWorkbook
    book.Close SaveChanges:=False
    Application.DisplayAlerts = previousAlerts

    BuildSampleReport = path
    Exit Function

Failed:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    Application.DisplayAlerts = previousAlerts
    On Error GoTo 0
End Function

'-----------------------------------------------------------------------
' Sections
'-----------------------------------------------------------------------
Private Function WriteHeader(ByVal sheet As Worksheet, ByVal sampleIdx As Long, _
                             ByVal amount As Double) As Long
    sheet.Cells(2, COL_LABEL).Value = "Audit sample " & Format$(sampleIdx, "00") & _
                                      " -- " & Format$(amount, "#,##0.00") & _
                                      " -- company code " & modConfig.Setting("Company code")
    sheet.Cells(2, COL_LABEL).Font.Bold = True
    sheet.Cells(2, COL_LABEL).Font.Size = 14

    sheet.Cells(3, COL_LABEL).Value = _
        "Extracted from SAP " & modSapConnect.gSystemId & " client " & _
        modSapConnect.gClient & " by " & modSapConnect.gSapUser & " on " & _
        Format$(Now, "dd/mm/yyyy hh:mm") & _
        IIf(modConfig.IsDryRun(), "  (DRY RUN)", "")
    sheet.Cells(3, COL_LABEL).Font.Italic = True

    WriteHeader = 5
End Function

Private Function WriteRequest(ByVal sheet As Worksheet, ByVal startRow As Long, _
                              ByVal sampleIdx As Long, ByVal monthTab As String, _
                              ByVal payDate As Date, ByVal amount As Double, _
                              ByVal party As String, ByVal reference As String) As Long
    Dim row As Long

    row = Section(sheet, startRow, "What the auditor asked for")

    row = Pair(sheet, row, "Sample number", Format$(sampleIdx, "00"))
    row = Pair(sheet, row, "Period", monthTab)
    row = Pair(sheet, row, "Payment date", Format$(payDate, "dd/mm/yyyy"))
    row = Pair(sheet, row, "Amount", Format$(amount, "#,##0.00"))
    row = Pair(sheet, row, "Named party", _
               IIf(Len(party) > 0, party, "(none named in the request)"))
    row = Pair(sheet, row, "Payment reference", reference)

    WriteRequest = row + 1
End Function

Private Function WriteTrail(ByVal sheet As Worksheet, ByVal startRow As Long, _
                            ByRef match As FebanMatch, ByRef chain As ChainResult) As Long
    Dim row As Long

    row = Section(sheet, startRow, "How it was traced in SAP")

    row = Pair(sheet, row, "1. Bank statement line", _
               IIf(match.Found, _
                   "FEBAN " & match.StatementDate & ", row " & match.GridRow & ", " & _
                   Format$(match.StatementAmount, "#,##0.00"), _
                   "no matching statement line"))
    row = Pair(sheet, row, "2. FI document", Blank(chain.FiDocument))
    row = Pair(sheet, row, "3. Clearing document", Blank(chain.ClearingDocument))
    row = Pair(sheet, row, "4. Payments in the batch", _
               IIf(chain.ZpNumberCount > 0, _
                   chain.ZpNumberCount & " " & modConfig.Setting("Payment document type") & _
                   " document(s)", "--"))
    row = Pair(sheet, row, "5. Largest payment of the batch", _
               IIf(Len(chain.ZpPaymentDocument) > 0, _
                   Format$(chain.ZpPaymentAmount, "#,##0.00") & _
                   "  document " & chain.ZpPaymentDocument & _
                   IIf(Len(chain.ZpPaymentVendor) > 0, "  " & chain.ZpPaymentVendor, "") & _
                   IIf(chain.IsConfirmingPayment, _
                       "   (confirming party -- supply-chain finance)", ""), "--"))
    row = Pair(sheet, row, "6. Largest invoice behind it", _
               IIf(Len(chain.InvoiceNumber) > 0, _
                   Format$(chain.InvoiceAmount, "#,##0.00") & _
                   "  document " & chain.InvoiceNumber & _
                   IIf(Len(chain.InvoiceSupplier) > 0, "  " & chain.InvoiceSupplier, ""), "--"))

    WriteTrail = row + 1
End Function

Private Function WriteOutcome(ByVal sheet As Worksheet, ByVal startRow As Long, _
                              ByRef chain As ChainResult) As Long
    Dim row As Long

    row = Section(sheet, startRow, "Outcome")

    row = Pair(sheet, row, "Status", chain.Status)
    Select Case chain.Status
        Case "DONE"
            sheet.Cells(row - 1, COL_VALUE).Font.Color = RGB(0, 112, 48)
        Case "NO CLEARING"
            sheet.Cells(row - 1, COL_VALUE).Font.Color = RGB(128, 96, 0)
        Case Else
            sheet.Cells(row - 1, COL_VALUE).Font.Color = RGB(192, 0, 0)
    End Select
    sheet.Cells(row - 1, COL_VALUE).Font.Bold = True

    row = Pair(sheet, row, "Notes", chain.Notes)
    sheet.Cells(row - 1, COL_VALUE).WrapText = True

    WriteOutcome = row + 1
End Function

Private Function WriteFiles(ByVal sheet As Worksheet, ByVal startRow As Long, _
                            ByVal folder As String, ByRef chain As ChainResult) As Long
    Dim row As Long

    row = Section(sheet, startRow, "Evidence in this folder")

    row = FileLine(sheet, row, folder, modUtil.FILE_FEBAN, _
                   "The FEBAN bank statement list for the whole period, with the " & _
                   "sample's line in it")
    row = FileLine(sheet, row, folder, modUtil.FILE_FIDOC, _
                   "The FI document's line items -- the evidence where nothing clears")
    row = FileLine(sheet, row, folder, modUtil.FILE_BATCH, _
                   "Payment usage of the clearing document: every payment in the batch")
    row = FileLine(sheet, row, folder, modUtil.FILE_ZPLIST, _
                   "FBL1N over those payments, with vendor names and amounts")
    row = FileLine(sheet, row, folder, modUtil.FILE_INVOICES, _
                   "Payment usage of the largest payment: the invoices it settles")
    row = FileLine(sheet, row, folder, modUtil.FILE_PDF, _
                   "The invoice document itself, as attached in SAP")

    WriteFiles = row + 1
End Function

Private Function WriteProvenance(ByVal sheet As Worksheet, ByVal startRow As Long, _
                                 ByVal monthTab As String) As Long
    Dim row As Long

    row = Section(sheet, startRow, "How to read this")

    row = Note(sheet, row, _
        "Every file here is a SAP export or a SAP attachment. Nothing was retyped, " & _
        "and nothing was posted or changed in SAP to produce it -- the extract only " & _
        "displays and downloads.")
    row = Note(sheet, row, _
        "The largest payment of the batch and the largest invoice behind it are read " & _
        "from the exported files, not from what was on screen. The invoice is the most " & _
        "negative row: a payment is a debit and the invoice it settles is a credit.")
    row = Note(sheet, row, _
        "The 'Log' sheet in this workbook is this sample's rows from the run's audit " & _
        "trail, timestamped and stamped with the SAP system, client and user.")

    WriteProvenance = row + 1
End Function

'-----------------------------------------------------------------------
' The PDF, embedded rather than only referenced, so the pack survives being
' forwarded as a single file. Embedding needs a PDF handler registered on
' the machine and is refused on some locked-down builds, so a failure falls
' back to a link and says so.
'-----------------------------------------------------------------------
Private Sub AttachInvoice(ByVal book As Workbook, ByRef chain As ChainResult)
    Dim sheet As Worksheet
    Dim embedded As Boolean

    If Len(chain.InvoicePdfFile) = 0 Then Exit Sub
    If Not modUtil.FileExists(chain.InvoicePdfFile) Then Exit Sub

    Set sheet = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count))
    sheet.Name = "Invoice"

    sheet.Cells(2, COL_LABEL).Value = "Invoice " & chain.InvoiceNumber & _
                                      IIf(Len(chain.InvoiceSupplier) > 0, _
                                          " -- " & chain.InvoiceSupplier, "") & _
                                      " -- " & Format$(chain.InvoiceAmount, "#,##0.00")
    sheet.Cells(2, COL_LABEL).Font.Bold = True

    On Error Resume Next
    sheet.OLEObjects.Add fileName:=chain.InvoicePdfFile, Link:=False, _
                         DisplayAsIcon:=True, _
                         IconLabel:=modUtil.FILE_PDF, _
                         Left:=40, Top:=60, Width:=64, Height:=64
    embedded = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    If embedded Then
        sheet.Cells(4, COL_LABEL).Value = "Double-click the icon below to open the PDF. " & _
                                          "It is embedded in this workbook, and the same " & _
                                          "file is in the folder beside it."
    Else
        sheet.Cells(4, COL_LABEL).Value = "The PDF could not be embedded on this machine, " & _
                                          "so it is linked instead. The file is in the " & _
                                          "folder beside this workbook."
        sheet.Hyperlinks.Add Anchor:=sheet.Cells(6, COL_LABEL), _
                             Address:=chain.InvoicePdfFile, _
                             TextToDisplay:=modUtil.FILE_PDF
    End If

    sheet.Columns(COL_LABEL).ColumnWidth = 90
    sheet.Cells(4, COL_LABEL).WrapText = True
End Sub

'-----------------------------------------------------------------------
' This sample's rows from the run's audit trail.
'-----------------------------------------------------------------------
Private Sub CopyLogRows(ByVal book As Workbook, ByVal sampleIdx As Long)
    Dim source As Worksheet, sheet As Worksheet
    Dim lastUsed As Long, r As Long, out As Long, c As Long

    On Error GoTo Failed

    Set source = ThisWorkbook.Worksheets(modConfig.SHEET_LOG)
    Set sheet = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count))
    sheet.Name = "Log"

    ' Column headings live on row 4 of the Log sheet.
    For c = 1 To 10
        sheet.Cells(1, c).Value = source.Cells(4, c).Value
        sheet.Cells(1, c).Font.Bold = True
    Next c

    out = 2
    lastUsed = source.Cells(source.Rows.Count, 1).End(xlUp).Row

    For r = 5 To lastUsed
        If source.Cells(r, 2).Value = sampleIdx Then
            For c = 1 To 10
                sheet.Cells(out, c).Value = source.Cells(r, c).Value
            Next c
            sheet.Cells(out, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
            out = out + 1
        End If
    Next r

    sheet.Columns("A:J").AutoFit
    sheet.Columns("H").ColumnWidth = 80
    sheet.Columns("H").WrapText = True
    Exit Sub

Failed:
End Sub

'-----------------------------------------------------------------------
' Small builders
'-----------------------------------------------------------------------
Private Function Section(ByVal sheet As Worksheet, ByVal row As Long, _
                         ByVal title As String) As Long
    sheet.Cells(row, COL_LABEL).Value = title
    sheet.Cells(row, COL_LABEL).Font.Bold = True
    sheet.Range(sheet.Cells(row, COL_LABEL), sheet.Cells(row, COL_VALUE)).Interior.Color = SECTION_FILL
    Section = row + 1
End Function

Private Function Pair(ByVal sheet As Worksheet, ByVal row As Long, _
                      ByVal label As String, ByVal value As String) As Long
    sheet.Cells(row, COL_LABEL).Value = label
    sheet.Cells(row, COL_VALUE).Value = value
    Pair = row + 1
End Function

Private Function Note(ByVal sheet As Worksheet, ByVal row As Long, _
                      ByVal text As String) As Long
    sheet.Cells(row, COL_VALUE).Value = text
    sheet.Cells(row, COL_VALUE).WrapText = True
    Note = row + 1
End Function

' One row per evidence file, marked present or not. A file that is missing
' is stated as missing rather than left off the list -- an auditor should
' see what was not obtained as well as what was.
Private Function FileLine(ByVal sheet As Worksheet, ByVal row As Long, _
                          ByVal folder As String, ByVal fileName As String, _
                          ByVal what As String) As Long
    Dim path As String

    path = modUtil.JoinPath(folder, fileName)

    If modUtil.FileExists(path) Then
        sheet.Hyperlinks.Add Anchor:=sheet.Cells(row, COL_LABEL), _
                             Address:=fileName, TextToDisplay:=fileName
        sheet.Cells(row, COL_VALUE).Value = what
    Else
        sheet.Cells(row, COL_LABEL).Value = fileName
        sheet.Cells(row, COL_LABEL).Font.Color = RGB(128, 128, 128)
        sheet.Cells(row, COL_VALUE).Value = "not produced for this sample -- " & what
        sheet.Cells(row, COL_VALUE).Font.Color = RGB(128, 128, 128)
    End If

    FileLine = row + 1
End Function

Private Function Blank(ByVal text As String) As String
    Blank = IIf(Len(text) > 0, text, "--")
End Function

Private Sub Layout(ByVal sheet As Worksheet)
    sheet.Columns(1).ColumnWidth = 2
    sheet.Columns(COL_LABEL).ColumnWidth = 34
    sheet.Columns(COL_VALUE).ColumnWidth = 95
    sheet.Cells.VerticalAlignment = xlTop
    sheet.Rows(2).RowHeight = 20
End Sub

' A new workbook can come with three sheets depending on the operator's
' Excel options. The report builds its own, so drop the rest.
Private Sub DropExtraSheets(ByVal book As Workbook)
    Dim i As Long
    Dim previousAlerts As Boolean

    previousAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error Resume Next
    For i = book.Worksheets.Count To 2 Step -1
        book.Worksheets(i).Delete
    Next i
    On Error GoTo 0

    Application.DisplayAlerts = previousAlerts
End Sub
