Attribute VB_Name = "modChain"
'=======================================================================
' modChain -- the end-to-end walk for one audit sample.
'
' Steps as the auditor described them, with what the recordings prove:
'
'   1  FEBAN, company code + statement dates from the audit row  RECORDED
'   2  export the list, find the row by date+amount, open it     RECORDED
'   3  Posting Area 1 Doc. number -> F2 -> FI document           RECORDED
'   4  first posting-key-40 line with a clearing doc, open it    RECORDED
'   5  clearing doc field -> F2 -> clearing document             RECORDED
'   6  Environment > Payment Usage -> the batch's ZP documents    RECORDED
'   7  export that list, read the ZP document numbers back       RECORDED
'   8  FBL1N filtered to those ZP numbers                       RECORDED
'   9  the largest ZP payment of the batch                       RECORDED
'  10  Payment usage on it -> the largest KR invoice -> its PDF  RECORDED
'
' Every step is now covered by a recording in recordings/. A stage whose
' Screen Map rows are blank still finishes at the last step that worked and
' says so, carrying the document numbers a person needs to take it on.
'
' ONE ROUTE, NOT TWO. N1.vbs walked a Santander SCF payment and B2.vbs a
' regular vendor, and they turned out to be the same shape -- Environment >
' Payment usage on the payment, export, open the invoice, save the
' attachment. The regular one only looked shorter because there were fewer
' rows to sort through. IsConfirmingPayment now only colours the report.
'
' THE INVOICE IS THE NEGATIVE ROW. The list step 10 picks from holds the
' payment and the invoices it settles: the payment is a debit and the
' invoices are credits, so the biggest invoice is the most negative row.
' This used to filter on document type and got it wrong -- KR is the SAP
' standard, this system posts RN -- and a type is configuration while a sign
' is arithmetic.
'
' Two 'largest' decisions, both read from exported data rather than assumed
' from a sort order: the largest ZP payment in the batch, then the largest
' invoice inside it. The recordings sort the list on screen and click a
' row by position -- lbl[164,8] -- which is a fixed screen coordinate and
' would silently pick the wrong document on a batch of a different size.
'
' NOT EVERY SAMPLE HAS AN INVOICE. A CHAPS transfer between the company's
' own bank accounts clears nothing and settles no supplier, so there is no
' batch, no payment and no invoice behind it. The chain exports the FI
' document's line items instead -- that IS the evidence for those -- and
' says so rather than reporting a failure.
'=======================================================================
Option Explicit

' Which window the Payment Usage list came up in. Set by OpenPaymentUsage and
' read by the export, because it is wnd[1] on this system and wnd[0] in the
' recordings -- neither can be assumed.
Private mListWindow As String

' What 'Environment > Payment usage' is called on this system, captured from
' the menu step 6 uses. Step 10 needs the same command from a screen where it
' sits at a different menu index.
Private mUsageMenuText As String

Public Type ChainResult
    ' steps 3-5
    FiDocument As String
    FiDocumentFile As String         ' the line items, when nothing clears
    ClearingDocument As String

    ' steps 6-7: the batch
    ZpListFile As String
    ZpNumbers As String              ' newline separated
    ZpNumberCount As Long

    ' steps 8-9: the largest payment of the batch
    ZpExportFile As String
    ZpPaymentDocument As String
    ZpPaymentVendor As String
    ZpPaymentAmount As Double
    ZpPaymentCount As Long
    IsConfirmingPayment As Boolean

    ' step 10: the largest invoice inside that payment
    InvoiceListFile As String
    InvoiceNumber As String
    InvoiceSupplier As String
    InvoiceAmount As Double
    InvoicePdfFile As String

    Status As String    ' DONE|BLOCKED_FBL1N|BLOCKED_INVOICE|PARTIAL|ERROR
    Notes As String
End Type

'-----------------------------------------------------------------------
' match is ByRef because VBA does not allow a user-defined type to be
' passed ByVal. It is read, never written.
Public Function Walk(ByVal sampleIdx As Long, ByRef match As FebanMatch, _
                     ByVal dateFrom As Date, ByVal dateTo As Date, _
                     ByVal folder As String, ByVal fileStem As String) As ChainResult
    Dim result As ChainResult
    Dim payment As ZpPayment
    Dim matched As Long, rejected As Long

    On Error GoTo Failed

    ' --- steps 2-3: open the statement item, read the FI document ---------
    modFeban.OpenStatementItem match.GridRow
    result.FiDocument = ReadMapped(sampleIdx, "Feban.Detail.DocNumber", "FI document")

    If Len(result.FiDocument) = 0 Then
        Finish result, "PARTIAL", _
               "The statement item shows no FI document, so it is not posted. Report " & _
               "the line as unposted rather than unmatched."
        modLog.LogAction sampleIdx, "Chain", result.Notes, "SKIPPED", vbNullString
        Walk = result
        Exit Function
    End If

    DrillWithF2 "Feban.Detail.DocNumber"
    modLog.LogAction sampleIdx, "Chain", _
                 "Step 3: opened FI document " & result.FiDocument, "OK", vbNullString

    ' --- steps 4-5: the clearing document --------------------------------
    If Not OpenClearingLine(sampleIdx) Then
        ' Nothing clears, so there is no batch, no payment and no invoice to
        ' reach -- an internal funding transfer between the company's own bank
        ' accounts looks exactly like this. The line items ARE the evidence
        ' for those samples, so export the screen we are standing on.
        result.FiDocumentFile = modExport.ExportAlvGridById( _
            sampleIdx, "Doc.BsegGrid", folder, modUtil.FILE_FIDOC)

        Finish result, "NO CLEARING", _
               "FI document " & result.FiDocument & " carries no clearing document on " & _
               "any line, so it settles nothing and there is no payment batch behind " & _
               "it -- a transfer between the company's own accounts posts exactly this " & _
               "way. Its line items are the evidence, and they are in " & _
               IIf(Len(result.FiDocumentFile) > 0, result.FiDocumentFile, _
                   "(the export did not run -- see the Log)") & "."
        modLog.LogAction sampleIdx, "Step 4", result.Notes, _
                     IIf(Len(result.FiDocumentFile) > 0, "OK", "ERROR"), result.FiDocumentFile
        Walk = result
        Exit Function
    End If

    result.ClearingDocument = ReadMapped(sampleIdx, "Doc.ClearingDocField", "clearing document")

    If Len(result.ClearingDocument) = 0 Then
        Finish result, "PARTIAL", _
               "No clearing document on the line the macro opened in FI document " & _
               result.FiDocument & "."
        Walk = result
        Exit Function
    End If

    DrillWithF2 "Doc.ClearingDocField"
    modLog.LogAction sampleIdx, "Chain", _
                 "Step 5: opened clearing document " & result.ClearingDocument, _
                 "OK", vbNullString

    ' --- steps 6-7: Payment Usage, and the ZP numbers in the batch -------
    OpenPaymentUsage sampleIdx

    result.ZpListFile = modExport.ExportListFrom( _
        sampleIdx, mListWindow, vbNullString, folder, modUtil.FILE_BATCH)

    If Len(result.ZpListFile) = 0 Then
        Finish result, "ERROR", _
               "The Payment Usage list did not export, so the batch's ZP document " & _
               "numbers could not be read. Clearing document is " & result.ClearingDocument & "."
        Walk = result
        Exit Function
    End If

    result.ZpNumbers = modExportRead.DocumentNumbersOfType( _
        result.ZpListFile, sampleIdx, modConfig.Setting("Payment document type"), _
        matched, rejected)
    result.ZpNumberCount = matched

    If matched = 0 Then
        ' Two very different things look the same here. A batch of 100+ rows
        ' with no ZP among them means the type column was read wrongly. A
        ' batch of two rows -- the SB statement line and the AB clearing that
        ' offsets it -- means this is not a vendor payment run at all, but a
        ' treasury or FX settlement, and there are no supplier invoices behind
        ' it to download. Say which one it is.
        If rejected <= 2 Then
            Finish result, "PARTIAL", _
                   "Clearing document " & result.ClearingDocument & " settles directly " & _
                   "against the bank statement -- the Payment Usage list holds only the " & _
                   "statement line and its offsetting entry, no vendor payments. This is " & _
                   "a treasury or FX settlement rather than a payment run, so there is no " & _
                   "supplier invoice behind it. See " & result.ZpListFile & "; the Log " & _
                   "names the document types it held."
        Else
            Finish result, "PARTIAL", _
                   "No " & modConfig.Setting("Payment document type") & " document " & _
                   "numbers were found among " & rejected & " rows in " & result.ZpListFile & _
                   ". That many rows is a payment run, so the document-type column was " & _
                   "probably read wrongly -- check the headings against the 'Payment " & _
                   "usage ...' settings on the Control sheet. The Log names the types it saw."
        End If
        Walk = result
        Exit Function
    End If

    ' --- steps 8-9: the largest ZP payment of the batch ------------------
    payment = modFbl1n.LargestPaymentOfBatch(sampleIdx, result.ZpNumbers, _
                                             dateFrom, dateTo, folder, fileStem)
    result.ZpExportFile = payment.ExportFile

    If payment.Skipped Then
        Finish result, "BLOCKED_FBL1N", _
               "Reached the batch: " & matched & " " & _
               modConfig.Setting("Payment document type") & " payment(s) behind clearing " & _
               "document " & result.ClearingDocument & ". " & payment.Notes
        Walk = result
        Exit Function
    End If

    If Not payment.Found Then
        Finish result, "PARTIAL", _
               "FBL1N ran but no largest payment could be read. " & payment.Notes
        Walk = result
        Exit Function
    End If

    result.ZpPaymentDocument = payment.DocumentNumber
    result.ZpPaymentAmount = payment.Amount
    result.ZpPaymentCount = payment.RowsConsidered
    result.ZpPaymentVendor = FirstNonEmpty(payment.VendorName, payment.Vendor)
    result.IsConfirmingPayment = NamesMatch(result.ZpPaymentVendor, _
                                            modConfig.Setting("Confirming party name"))

    ' Try the FBL1N list first -- that is where the recordings drill from, and
    ' finding the row by its document number keeps that navigation without
    ' inheriting its dependence on screen position. FB03 only if the number is
    ' not on screen.
    If Not modFbl1n.OpenPaymentOnList(sampleIdx, payment.DocumentNumber) Then
      If Not modFbl1n.OpenPaymentByDocument(sampleIdx, payment.DocumentNumber, _
                                            Format$(dateFrom, "yyyy")) Then
        Finish result, "BLOCKED_INVOICE", _
               "Largest payment of the batch is " & _
               Format$(result.ZpPaymentAmount, "#,##0.00") & _
               IIf(Len(result.ZpPaymentDocument) > 0, _
                   " (document " & result.ZpPaymentDocument & ")", "") & _
               ", but it could not be opened, so its invoices were not reached."
        Walk = result
        Exit Function
      End If
    End If

    ' --- step 10: the largest invoice inside it, and its PDF -------------
    FetchInvoicePdf sampleIdx, result, folder, fileStem

    Walk = result
    Exit Function

Failed:
    Finish result, "ERROR", Err.Description
    modLog.LogAction sampleIdx, "Chain failed", Err.Description, "ERROR", vbNullString
    Walk = result
End Function

'-----------------------------------------------------------------------
' Where 'Environment > Payment usage' is on THIS screen.
'
' The menu index is not a property of the command, it is a property of the
' screen: menu[5]/menu[3] from the clearing document, menu[4]/menu[3] in the
' recordings, and something else again on the FB03 overview. Two mapped
' guesses first, then a search of the menu bar for the caption step 6 read
' off this very system -- which is why the search needs no translation
' table and works on any logon language.
'-----------------------------------------------------------------------
Private Function FindUsageMenu() As String
    Dim mapped As String

    ' Caption first, indices second. A mapped index that happens to EXIST on
    ' this screen is not evidence it is the right command -- menu[4]/menu[3]
    ' is something different on every screen that has four menus. The caption
    ' was read off this system minutes earlier and is exact.
    If Len(mUsageMenuText) > 0 Then
        FindUsageMenu = MenuWithCaption("wnd[0]/mbar", mUsageMenuText, 0)
        If Len(FindUsageMenu) > 0 Then Exit Function
    End If

    mapped = modConfig.ElementIdOrBlank("Payment.UsageMenu")
    If Len(mapped) > 0 Then
        If modSapConnect.Exists(mapped) Then
            FindUsageMenu = mapped
            Exit Function
        End If
    End If

    mapped = modConfig.ElementIdOrBlank("PaymentUsage.Menu")
    If Len(mapped) > 0 Then
        If modSapConnect.Exists(mapped) Then FindUsageMenu = mapped
    End If
End Function

' Depth-first walk of the menu bar for an entry with this caption.
Private Function MenuWithCaption(ByVal elementId As String, ByVal caption As String, _
                                 ByVal depth As Long) As String
    Dim control As Object, child As Object
    Dim found As String

    If depth > 4 Then Exit Function
    If Not modSapConnect.Exists(elementId) Then Exit Function

    Set control = modSapConnect.Element(elementId)

    On Error Resume Next
    For Each child In control.Children
        If StrComp(Trim$(child.Text), Trim$(caption), vbTextCompare) = 0 Then
            MenuWithCaption = child.Id
            Exit Function
        End If

        found = MenuWithCaption(child.Id, caption, depth + 1)
        If Len(found) > 0 Then
            MenuWithCaption = found
            Exit Function
        End If
    Next child
    On Error GoTo 0
End Function

' The document-overview button, when the Screen Map carries one. Optional and
' never fatal: on a screen that is already the document it does nothing
' useful, and on one that is not it is the hop that makes Environment >
' Payment usage exist.
Private Sub ShowDocumentOverview(ByVal sampleIdx As Long)
    Dim buttonId As String
    Dim before As String

    buttonId = modConfig.ElementIdOrBlank("Doc.OverviewButton")
    If Len(buttonId) = 0 Then Exit Sub
    If Not modSapConnect.Exists(buttonId) Then Exit Sub

    before = modSapConnect.ScreenSignature()

    On Error Resume Next
    modSafety.GuardedPress buttonId
    On Error GoTo 0

    If modSapConnect.ScreenSignature() <> before Then
        modLog.LogAction sampleIdx, "Step 10", _
                     "Pressed the document-overview button to get from the line item to " & _
                     "the document, so Environment > Payment usage exists.", _
                     "OK", vbNullString
    End If
End Sub

' Every exit goes through here, so no path can leave Status unset -- the
' failure mode that let a sample with an unreadable list still report DONE.
Private Sub Finish(ByRef result As ChainResult, ByVal status As String, ByVal notes As String)
    result.Status = status
    result.Notes = Trim$(notes)
End Sub

'-----------------------------------------------------------------------
' Step 10, both routes.
'
' Regular vendor: open the payment and save the PDF hanging off it.
' Santander SCF:  a confirming payment settles the finance provider, so it
'                 needs extra navigation to reach the supplier invoices,
'                 and then the largest of those.
' Neither is recorded, so both are gated.
'-----------------------------------------------------------------------
Private Sub FetchInvoicePdf(ByVal sampleIdx As Long, ByRef result As ChainResult, _
                            ByVal folder As String, ByVal fileStem As String)
    Dim summary As String
    Dim usageMenuId As String
    Dim beforeMenu As String
    Dim largest As ListRow

    summary = "Largest of " & result.ZpPaymentCount & " ZP payments in the batch: " & _
              Format$(result.ZpPaymentAmount, "#,##0.00") & _
              IIf(Len(result.ZpPaymentDocument) > 0, _
                  " (document " & result.ZpPaymentDocument & ")", "") & _
              IIf(Len(result.ZpPaymentVendor) > 0, ", " & result.ZpPaymentVendor, "") & _
              IIf(result.IsConfirmingPayment, " -- the confirming party.", ".")

    ' N1.vbs (Santander SCF) and B2.vbs (a regular vendor) turned out to walk
    ' the same shape: Environment > Payment usage on the payment, export that
    ' list, open the invoice on it, then the attachment. The regular route
    ' only looked shorter because the operator had fewer rows to sort through.
    ' So there is one path here, not two, and IsConfirmingPayment now only
    ' colours the report.
    On Error GoTo Failed

    usageMenuId = FindUsageMenu()

    ' F2 on a vendor line-item list lands on the LINE ITEM, and Environment >
    ' Payment usage lives on the document. N1.vbs bridged that with the
    ' document-overview button, so do the same before giving up.
    If Len(usageMenuId) = 0 Then
        ShowDocumentOverview sampleIdx
        usageMenuId = FindUsageMenu()
    End If

    If Len(usageMenuId) = 0 Then
        Finish result, "BLOCKED_INVOICE", _
               summary & " Could not find Environment > Payment usage on the payment's " & _
               "screen. Tried " & modConfig.ElementIdOrBlank("Payment.UsageMenu") & ", " & _
               modConfig.ElementIdOrBlank("PaymentUsage.Menu") & " and a search of the " & _
               "whole menu bar for """ & mUsageMenuText & """. To finish by hand: open " & _
               "payment " & result.ZpPaymentDocument & " and take the largest negative " & _
               "(credit) document behind it -- that is the invoice."
        modLog.LogAction sampleIdx, "Step 10", result.Notes, "ERROR", vbNullString
        Exit Sub
    End If

    modLog.LogAction sampleIdx, "Step 10", _
                 "Opening the invoices behind payment " & result.ZpPaymentDocument & _
                 " through " & usageMenuId & " (""" & mUsageMenuText & """).", _
                 "OK", vbNullString

    beforeMenu = modSapConnect.ScreenSignature()

    modSapConnect.Element(usageMenuId).Select
    modSapConnect.WaitForSap
    modSafety.AssertPopupKnown

    ' If the menu did nothing, whatever list is still on screen is the one we
    ' came from -- and exporting it would produce a second copy of the ZP
    ' payment list under the invoices name, which is exactly what happened:
    ' same size, same 166 rows, then 'no KR row found in it'.
    If modSapConnect.ScreenSignature() = beforeMenu Then
        Finish result, "BLOCKED_INVOICE", _
               summary & " " & usageMenuId & " (""" & mUsageMenuText & """) left the " & _
               "screen unchanged, so the invoices behind the payment were never listed " & _
               "and nothing was exported. The command is there but did nothing from this " & _
               "screen -- record Environment > Payment usage from inside a payment " & _
               "document and correct Payment.UsageMenu on the Screen Map."
        modLog.LogAction sampleIdx, "Step 10", result.Notes, "ERROR", vbNullString
        Exit Sub
    End If

    result.InvoiceListFile = modExport.ExportClassicList( _
        sampleIdx, folder, modUtil.FILE_INVOICES)

    If Len(result.InvoiceListFile) = 0 Then
        Finish result, "PARTIAL", _
               summary & " The invoice list did not export."
        Exit Sub
    End If

    ' The invoice is the negative row: the payment is a debit, the invoice it
    ' settles is a credit, and the biggest invoice is the most negative. That
    ' holds whatever the document type happens to be called in this company
    ' code -- which is the point, because it was called RN here and the type
    ' filter this replaces was looking for KR.
    largest = modExportRead.MostNegativeRow(result.InvoiceListFile, sampleIdx, _
                                            "Invoice list amount column", _
                                            "Invoice list supplier column", _
                                            "Invoice list document column", _
                                            modConfig.Setting("Invoice document type"))

    If Not largest.Found Then
        Finish result, "PARTIAL", _
               summary & " Exported " & result.InvoiceListFile & " but no row in it " & _
               "carries a negative amount, so there is no invoice behind this payment " & _
               "to take. Open the file and check the amount column."
        Exit Sub
    End If

    result.InvoiceNumber = largest.DocumentNumber
    result.InvoiceSupplier = largest.Supplier
    result.InvoiceAmount = largest.Amount

    modLog.LogAction sampleIdx, "Step 10", _
                 "Largest of " & largest.RowsConsidered & " invoice(s) behind the " & _
                 "payment: " & Format$(largest.Amount, "#,##0.00") & _
                 IIf(Len(largest.DocumentNumber) > 0, _
                     ", document " & largest.DocumentNumber, "") & _
                 IIf(Len(largest.Supplier) > 0, ", " & largest.Supplier, ""), _
                 "OK", result.InvoiceListFile

    result.InvoicePdfFile = SaveAttachedPdf(sampleIdx, result, folder, fileStem)

    If Len(result.InvoicePdfFile) > 0 Then
        Finish result, "DONE", _
               summary & " Invoice " & result.InvoiceNumber & " for " & _
               Format$(result.InvoiceAmount, "#,##0.00") & ", PDF saved."
    Else
        Finish result, "BLOCKED_INVOICE", _
               summary & " Invoice " & result.InvoiceNumber & " identified, but its PDF " & _
               "was not written. " & result.Notes
    End If

    Exit Sub

Failed:
    Finish result, "PARTIAL", summary & " Step 10 stopped: " & Err.Description
    modLog.LogAction sampleIdx, "Step 10 failed", Err.Description, "ERROR", vbNullString
End Sub

'-----------------------------------------------------------------------
' The attachment list, and the PDF off it.
'
' Not recorded either. Where the Invoice.* rows are blank, or where SAP
' hands the document to the external viewer instead of offering a save
' dialog, this reports MANUAL with the document numbers rather than
' claiming a success that left no file behind.
'-----------------------------------------------------------------------
Private Function SaveAttachedPdf(ByVal sampleIdx As Long, ByRef result As ChainResult, _
                                 ByVal folder As String, ByVal fileStem As String) As String
    Dim toolboxId As String, gridId As String, exportItem As String
    Dim saveWindow As String, columnName As String
    Dim toolbox As Object, grid As Object
    Dim target As String

    toolboxId = FindGosToolbox()
    gridId = modConfig.ElementIdOrBlank("Invoice.AttachListGrid")
    exportItem = modConfig.ElementIdOrBlank("Invoice.ExportMenuItem")
    saveWindow = modConfig.ElementIdOrBlank("Invoice.SaveWindow")
    columnName = modConfig.ElementIdOrBlank("Invoice.AttachListColumn")

    If Len(toolboxId) = 0 Or Len(gridId) = 0 Or Len(exportItem) = 0 Then
        result.Notes = "The Invoice.* rows are not filled in, so no PDF download was " & _
                       "attempted. Everything needed to fetch it by hand is recorded: " & _
                       "payment " & result.ZpPaymentDocument & _
                       IIf(Len(result.InvoiceNumber) > 0, _
                           ", invoice " & result.InvoiceNumber, "") & _
                       IIf(Len(result.InvoiceSupplier) > 0, _
                           ", supplier " & result.InvoiceSupplier, "") & "."
        modLog.LogAction sampleIdx, "Invoice PDF", result.Notes, "MANUAL", vbNullString
        Exit Function
    End If

    target = modUtil.JoinPath(folder, modUtil.FILE_PDF)
    If modSafety.BlockedByDryRun("Would save the invoice PDF to " & target) Then Exit Function

    If Len(toolboxId) = 0 Then
        result.Notes = "The services-for-object toolbox is not on this screen, so the " & _
                       "attachment list could not be opened. The invoice may not be the " & _
                       "object the attachment hangs off."
        modLog.LogAction sampleIdx, "Invoice PDF", result.Notes, "MANUAL", vbNullString
        Exit Function
    End If

    Set toolbox = modSapConnect.Element(toolboxId)
    On Error Resume Next
    toolbox.pressContextButton "%GOS_TOOLBOX"
    toolbox.selectContextMenuItem "%GOS_VIEW_ATTA"
    On Error GoTo 0
    modSapConnect.WaitForSap

    If Not modSapConnect.Exists(gridId) Then
        result.Notes = "The attachment list did not open, or " & gridId & " is wrong " & _
                       "for it. Nothing was downloaded."
        modLog.LogAction sampleIdx, "Invoice PDF", result.Notes, "MANUAL", vbNullString
        CloseAttachmentList
        Exit Function
    End If

    Set grid = modSapConnect.Element(gridId)

    ' Take the first attachment. Where a document carries several, the run
    ' says so rather than silently picking one and calling it the invoice.
    On Error Resume Next
    If grid.RowCount > 1 Then
        modLog.LogAction sampleIdx, "Invoice PDF", _
                     grid.RowCount & " attachments on this document; the first was " & _
                     "taken. Check by hand which one the auditor wants.", _
                     "MANUAL", vbNullString
    End If
    If Len(columnName) > 0 Then grid.setCurrentCell 0, columnName
    grid.selectedRows = "0"
    On Error GoTo 0
    modSapConnect.WaitForSap

    ' The export is a context-menu item on the grid, not a toolbar button.
    On Error Resume Next
    grid.contextMenu
    grid.selectContextMenuItem exportItem
    On Error GoTo 0
    modSapConnect.WaitForSap

    If Len(saveWindow) = 0 Then saveWindow = "wnd[2]"

    If Not modSapConnect.Exists(saveWindow) Then
        result.Notes = "The attachment export produced no save dialog in " & saveWindow & _
                       ". SAP may have handed the document to the external viewer instead."
        modLog.LogAction sampleIdx, "Invoice PDF", result.Notes, "MANUAL", vbNullString
        CloseAttachmentList
        Exit Function
    End If

    ' The PDF dialog is one window deeper than the list exports, because the
    ' attachment list itself is already wnd[1].
    SaveAttachedPdf = modExport.CompleteSaveDialogIn( _
        sampleIdx, saveWindow, folder, modUtil.FILE_PDF)

    CloseAttachmentList
End Function

' Locate the services-for-object toolbox.
'
' Its container index is NOT stable: N1.vbs found it at titl/shellcont[2]
' and B2.vbs at titl/shellcont[1], on the same system, because the index
' depends on what else the title bar is carrying on that particular screen.
' A fixed ID would work on one document and quietly fail on the next, so
' probe the handful of possibilities instead. The Screen Map value is tried
' first, then the neighbours.
Private Function FindGosToolbox() As String
    Dim configured As String
    Dim candidate As String
    Dim i As Long

    configured = modConfig.ElementIdOrBlank("Invoice.GosToolbox")
    If Len(configured) > 0 Then
        If modSapConnect.Exists(configured) Then
            FindGosToolbox = configured
            Exit Function
        End If
    End If

    For i = 0 To 4
        candidate = "wnd[0]/titl/shellcont[" & i & "]/shell"
        If modSapConnect.Exists(candidate) Then
            FindGosToolbox = candidate
            Exit Function
        End If
    Next i

    ' Some screens expose it without the container index at all.
    If modSapConnect.Exists("wnd[0]/titl/shellcont/shell") Then
        FindGosToolbox = "wnd[0]/titl/shellcont/shell"
    End If
End Function

Private Sub CloseAttachmentList()
    Dim closeId As String

    closeId = modConfig.ElementIdOrBlank("Invoice.CloseAttachList")

    On Error Resume Next
    If Len(closeId) > 0 And modSapConnect.Exists(closeId) Then
        modSapConnect.Element(closeId).Press
    ElseIf modSapConnect.Exists("wnd[1]") Then
        modSapConnect.Element("wnd[1]").Close
    End If
    On Error GoTo 0
    modSapConnect.WaitForSap
End Sub

'-----------------------------------------------------------------------
' Screen steps
'-----------------------------------------------------------------------
Private Function ReadMapped(ByVal sampleIdx As Long, ByVal mapKey As String, _
                            ByVal what As String) As String
    Dim fieldId As String

    fieldId = modConfig.ElementId(mapKey)

    If Not modSapConnect.Exists(fieldId) Then
        modLog.LogAction sampleIdx, "Chain", _
                     mapKey & " is not on this screen, so the " & what & " could not be " & _
                     "read. Current transaction: " & modSapConnect.CurrentTransaction() & _
                     ". The previous step may not have landed where expected.", _
                     "ERROR", vbNullString
        Exit Function
    End If

    ReadMapped = Trim$(modSapConnect.Element(fieldId).Text)
    If ReadMapped = "0" Then ReadMapped = vbNullString
End Function

' The recordings drill in by focusing a field and pressing F2 -- choose/
' display, a read action, which the guard allows.
Private Sub DrillWithF2(ByVal mapKey As String)
    modSapConnect.Element(modConfig.ElementId(mapKey)).SetFocus
    modSafety.GuardedSendVKey "wnd[0]", 2
    modSafety.AssertPopupKnown
End Sub

'-----------------------------------------------------------------------
' Step 4 -- the first line with the wanted posting key AND a clearing doc.
'
' The recordings just double-clicked whichever row the cursor happened to
' be on, which is fine by hand and wrong unattended: the clearing document
' hangs off one specific line. So scan the grid for it.
'-----------------------------------------------------------------------
Private Function OpenClearingLine(ByVal sampleIdx As Long) As Boolean
    Dim grid As Object
    Dim rows As Long, r As Long
    Dim clearingColumn As String, keyColumn As String, wantedKey As String
    Dim clearingValue As String, keyValue As String
    Dim seenValues As String
    Dim chosen As Long

    Set grid = modSapConnect.Element(modConfig.ElementId("Doc.BsegGrid"))
    clearingColumn = modConfig.ElementId("Doc.Col.ClearingDoc")
    keyColumn = modConfig.ElementIdOrBlank("Doc.Col.PostingKey")
    wantedKey = Trim$(modConfig.Setting("Clearing line posting key"))

    On Error Resume Next
    rows = grid.RowCount
    On Error GoTo 0

    chosen = -1

    For r = 0 To rows - 1
        clearingValue = Trim$(GridCell(grid, r, clearingColumn))

        If Len(clearingValue) > 0 And clearingValue <> "0" Then
            If Len(keyColumn) = 0 Or Len(wantedKey) = 0 Then
                chosen = r                      ' no posting-key column mapped
                Exit For
            End If

            keyValue = Trim$(GridCell(grid, r, keyColumn))
            If keyValue = wantedKey Then
                chosen = r
                Exit For
            End If
        End If
    Next r

    ' Fall back to any line that carries a clearing document, and say so, so
    ' a wrong posting-key column shows up instead of stopping the sample.
    If chosen < 0 And Len(keyColumn) > 0 Then
        For r = 0 To rows - 1
            clearingValue = Trim$(GridCell(grid, r, clearingColumn))
            If Len(clearingValue) > 0 And clearingValue <> "0" Then
                chosen = r
                modLog.LogAction sampleIdx, "Step 4", _
                             "No line with posting key " & wantedKey & ", so row " & r & _
                             " was used because it carries a clearing document. Check " & _
                             "Doc.Col.PostingKey on the Screen Map.", "MANUAL", vbNullString
                Exit For
            End If
        Next r
    End If

    ' Nothing at all carries a clearing document. That is a fact about the
    ' document, not a mapping problem -- but only if the column really is the
    ' clearing-document one, so say what was read rather than just 'none'.
    If chosen < 0 Then
        For r = 0 To rows - 1
            seenValues = seenValues & IIf(Len(seenValues) > 0, "; ", "") & _
                         "row " & r & " key=" & _
                         IIf(Len(keyColumn) > 0, Trim$(GridCell(grid, r, keyColumn)), "?") & _
                         " clearing=[" & Trim$(GridCell(grid, r, clearingColumn)) & "]"
        Next r

        modLog.LogAction sampleIdx, "Step 4", _
                     "No line carries a clearing document. Over " & rows & " line(s), " & _
                     "column " & clearingColumn & " held: " & seenValues & _
                     ". If those cells are blank the document is genuinely not cleared -- " & _
                     "the payment may hang off Posting Area 2 of the statement item " & _
                     "rather than Posting Area 1, which is the field this macro reads. " & _
                     "Open the statement line by hand and check both areas.", _
                     "ERROR", vbNullString
        Exit Function
    End If

    grid.setCurrentCell chosen, clearingColumn
    modSapConnect.WaitForSap
    grid.doubleClickCurrentCell
    modSapConnect.WaitForSap
    modSafety.AssertPopupKnown

    modLog.LogAction sampleIdx, "Step 4", _
                 "Opened line-item row " & chosen & " of " & rows & _
                 " (clearing document " & clearingValue & ")", "OK", vbNullString

    OpenClearingLine = True
End Function

Private Sub OpenPaymentUsage(ByVal sampleIdx As Long)
    Dim anchorId As String
    Dim menuId As String

    menuId = modConfig.ElementId("PaymentUsage.Menu")

    ' Remember what this menu entry is CALLED on this system, before selecting
    ' it. Step 10 needs the same command from a different screen, where the
    ' menu sits at a different index -- menu[5]/menu[3] here, menu[4]/menu[3]
    ' in the recordings. Matching on the caption read from this system at run
    ' time finds it wherever it moved to, without hardcoding a translation.
    On Error Resume Next
    mUsageMenuText = modSapConnect.Element(menuId).Text
    On Error GoTo 0

    modSapConnect.Element(menuId).Select
    modSapConnect.WaitForSap
    modSafety.AssertPopupKnown

    ' SAP announces this step with an informational window titled after the
    ' document -- "Cleared Line Items for Document GBKM 0900722750 2026" --
    ' and then loads the list behind it. That window is not an obstacle and
    ' is not latched: which window holds the list is decided at export time,
    ' because the announcement may be gone by then.
    mListWindow = "wnd[0]"
    If modSapConnect.ModalWindowOpen() Then
        modLog.LogAction sampleIdx, "Step 6", _
                     "Payment Usage announced itself with """ & _
                     modSapConnect.ModalWindowTitle() & """. " & DescribeWindow("wnd[1]"), _
                     "OK", vbNullString
    Else
        anchorId = modConfig.ElementIdOrBlank("PaymentUsage.ListAnchor")
        If Len(anchorId) > 0 Then
            If Not modSapConnect.Exists(anchorId) Then
                modLog.LogAction sampleIdx, "Step 6", _
                             "Payment Usage produced no modal and the anchor " & _
                             anchorId & " is not on wnd[0] either. " & _
                             DescribeWindow("wnd[0]"), "ERROR", vbNullString
                Exit Sub
            End If
        End If

        modLog.LogAction sampleIdx, "Step 6", _
                     "Opened Environment > Payment Usage in the main window", _
                     "OK", vbNullString
    End If
End Sub

' A one-line inventory of what a window holds, so an unexpected screen is
' diagnosable from the Log instead of needing another run with the prober.
Private Function DescribeWindow(ByVal windowId As String) As String
    Dim grids As Long, labels As Long, fields As Long, buttons As Long
    Dim firstGrid As String

    CountControls windowId, 0, grids, labels, fields, buttons, firstGrid

    DescribeWindow = "It holds " & grids & " grid/shell, " & labels & " label, " & _
                     fields & " field and " & buttons & " button controls" & _
                     IIf(Len(firstGrid) > 0, ". First grid: " & firstGrid, "") & "."
End Function

Private Sub CountControls(ByVal elementId As String, ByVal depth As Long, _
                          ByRef grids As Long, ByRef labels As Long, _
                          ByRef fields As Long, ByRef buttons As Long, _
                          ByRef firstGrid As String)
    Dim control As Object, child As Object
    Dim kind As String

    If depth > 10 Then Exit Sub
    If Not modSapConnect.Exists(elementId) Then Exit Sub

    Set control = modSapConnect.Element(elementId)

    On Error Resume Next
    kind = control.Type
    On Error GoTo 0

    Select Case kind
        Case "GuiShell", "GuiGridView", "GuiTableControl"
            grids = grids + 1
            If Len(firstGrid) = 0 Then firstGrid = control.Id
        Case "GuiLabel"
            labels = labels + 1
        Case "GuiTextField", "GuiCTextField", "GuiComboBox"
            fields = fields + 1
        Case "GuiButton"
            buttons = buttons + 1
    End Select

    On Error Resume Next
    For Each child In control.Children
        CountControls child.Id, depth + 1, grids, labels, fields, buttons, firstGrid
    Next child
    On Error GoTo 0
End Sub

' A recorded step is a menu entry, a button, or a field to focus and F2.
' Which one is not knowable from the ID, so read the control's own type.
Private Sub ActivateElement(ByVal elementId As String)
    Dim control As Object
    Dim kind As String

    Set control = modSapConnect.Element(elementId)

    kind = vbNullString
    On Error Resume Next
    kind = control.Type
    On Error GoTo 0

    Select Case kind
        Case "GuiMenu"
            control.Select
            modSapConnect.WaitForSap
        Case "GuiButton", "GuiToolbarControl"
            modSafety.GuardedPress elementId
        Case Else
            On Error Resume Next
            control.SetFocus
            On Error GoTo 0
            modSafety.GuardedSendVKey "wnd[0]", 2
    End Select
End Sub

Private Sub CloseModal()
    On Error Resume Next
    modSapConnect.Element("wnd[1]").Close
    On Error GoTo 0
    modSapConnect.WaitForSap
End Sub

Private Function GridCell(ByVal grid As Object, ByVal row As Long, _
                          ByVal columnName As String) As String
    On Error Resume Next
    GridCell = grid.GetCellValue(row, columnName)
    On Error GoTo 0
End Function

Private Function FirstNonEmpty(ByVal a As String, ByVal b As String) As String
    FirstNonEmpty = IIf(Len(Trim$(a)) > 0, Trim$(a), Trim$(b))
End Function

'-----------------------------------------------------------------------
' 'SANTANDER SCF', 'SCF Santander' and 'Santander  SCF' are the same party.
' Every word of the configured name must appear in the supplier name, so
' 'SANTANDER UK PLC' does not match.
'-----------------------------------------------------------------------
Public Function NamesMatch(ByVal supplier As String, ByVal configured As String) As Boolean
    Dim tokens() As String
    Dim i As Long
    Dim haystack As String

    If Len(Trim$(configured)) = 0 Then Exit Function
    If Len(Trim$(supplier)) = 0 Then Exit Function

    haystack = LettersAndDigits(supplier)
    tokens = Split(modUtil.Squeeze(UCase$(configured)), " ")

    For i = LBound(tokens) To UBound(tokens)
        If Len(tokens(i)) > 0 Then
            If InStr(haystack, LettersAndDigits(tokens(i))) = 0 Then Exit Function
        End If
    Next i

    NamesMatch = True
End Function

Private Function LettersAndDigits(ByVal text As String) As String
    Dim i As Long
    Dim ch As String
    Dim upper As String

    upper = UCase$(text)
    For i = 1 To Len(upper)
        ch = Mid$(upper, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            LettersAndDigits = LettersAndDigits & ch
        End If
    Next i
End Function
