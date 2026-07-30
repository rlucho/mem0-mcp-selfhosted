Attribute VB_Name = "modChain"
'=======================================================================
' modChain -- the end-to-end drill-down, following recordings/Audit.vbs.
'
' The chain, exactly as recorded:
'
'   FEBAN result grid
'     double-click the matched row's amount cell
'   -> statement item detail        read txtD2201_BELNR  (the FI document)
'     F2
'   -> FI document overview         double-click the DMBTR line
'   -> line item detail             read txtBSEG-AUGBL   (the clearing document)
'     F2
'   -> clearing document            menu[5]/menu[3]
'   -> cleared items with supplier names
'     menu[0]/menu[3]/menu[1]       export to a local file
'
' The export is then read back off disk to find the largest cleared item
' and its supplier, which decides the invoice route:
'
'   supplier = SANTANDER SCF  -> confirming payment, Audit2 route
'   any other supplier        -> open the payment, download its PDF
'
' BLOCKED: the Santander SCF route cannot be implemented yet. Audit2.vbs
' captured no steps -- only the scripting boilerplate and a
' resizeWorkingPane call -- so the extra steps a confirming payment needs
' are unknown. Those samples are logged BLOCKED_SCF with everything
' needed to finish them by hand, and the run carries on.
'=======================================================================
Option Explicit

Public Type ChainResult
    FiDocument As String
    ClearingDocument As String
    ClearedItemsFile As String
    LargestSupplier As String
    LargestAmount As Double
    LargestDocument As String
    IsConfirmingPayment As Boolean
    InvoiceFile As String
    Status As String              ' DONE | BLOCKED_SCF | PARTIAL | ERROR
    Notes As String
End Type

'-----------------------------------------------------------------------
' Walk the whole chain for one matched statement line.
'-----------------------------------------------------------------------
Public Function Walk(ByVal sampleIdx As Long, ByVal match As FebanMatch, _
                     ByVal folder As String, ByVal fileStem As String) As ChainResult
    Dim result As ChainResult

    On Error GoTo Failed

    ' 1. statement item detail -> the FI document number
    modFeban.OpenStatementItem match.GridRow
    result.FiDocument = ReadFiDocument(sampleIdx)

    If Len(result.FiDocument) = 0 Then
        result.Status = "PARTIAL"
        result.Notes = "The statement item's detail screen shows no FI document, so it " & _
                       "appears not to be posted. Report the line as unposted."
        modLog.LogAction sampleIdx, "Chain", result.Notes, "SKIPPED", vbNullString
        Walk = result
        Exit Function
    End If

    ' 2. F2 into the FI document
    DrillWithF2 "Feban.Detail.DocNumber"
    modLog.LogAction sampleIdx, "Chain", _
                 "Opened FI document " & result.FiDocument, "OK", vbNullString

    ' 3. FI document -> line item -> the clearing document number
    OpenAmountLine
    result.ClearingDocument = ReadClearingDocument(sampleIdx)

    If Len(result.ClearingDocument) = 0 Then
        result.Status = "PARTIAL"
        result.Notes = "FI document " & result.FiDocument & " has no clearing document " & _
                       "on the line the macro opened, so the cleared invoices cannot be " & _
                       "listed. Check the line-item detail by hand."
        modLog.LogAction sampleIdx, "Chain", result.Notes, "ERROR", vbNullString
        Walk = result
        Exit Function
    End If

    ' 4. F2 into the clearing document, then the cleared-items list
    DrillWithF2 "Doc.ClearingDocField"
    modLog.LogAction sampleIdx, "Chain", _
                 "Opened clearing document " & result.ClearingDocument, "OK", vbNullString

    OpenClearedItemsList sampleIdx

    ' 5. export the cleared items, supplier names and all
    result.ClearedItemsFile = modExport.ExportClassicList( _
        sampleIdx, folder, fileStem & "_cleared_items.txt")

    If Len(result.ClearedItemsFile) = 0 Then
        result.Status = IIf(modConfig.IsDryRun(), "PARTIAL", "ERROR")
        result.Notes = "The cleared-items list was not exported, so the largest " & _
                       "supplier could not be identified."
        Walk = result
        Exit Function
    End If

    ' 6. read the export back to decide the invoice route
    IdentifyLargestItem sampleIdx, result

    ' 7. fetch the invoice down whichever route applies
    If result.IsConfirmingPayment Then
        result.Status = "BLOCKED_SCF"
        result.Notes = "Largest cleared item is " & result.LargestSupplier & " for " & _
                       Format$(result.LargestAmount, "#,##0.00") & _
                       IIf(Len(result.LargestDocument) > 0, _
                           " (document " & result.LargestDocument & ")", "") & _
                       ", so this is a confirming payment and needs the Audit2 route. " & _
                       "Those steps are unknown -- Audit2.vbs recorded nothing. " & _
                       "Re-record it and fill in the Scf.* rows on the Screen Map."
        modLog.LogAction sampleIdx, "Invoice route", result.Notes, "MANUAL", vbNullString
    Else
        result.InvoiceFile = DownloadRegularInvoice(sampleIdx, result, folder, fileStem)
        If Len(result.InvoiceFile) > 0 Then
            result.Status = "DONE"
        Else
            result.Status = "PARTIAL"
            If Len(result.Notes) = 0 Then
                result.Notes = "Cleared items exported, but the invoice PDF was not " & _
                               "written. See the Log for which step stopped."
            End If
        End If
    End If

    Walk = result
    Exit Function

Failed:
    result.Status = "ERROR"
    result.Notes = Err.Description
    modLog.LogAction sampleIdx, "Chain failed", Err.Description, "ERROR", vbNullString
    Walk = result
End Function

'-----------------------------------------------------------------------
' Screen steps
'-----------------------------------------------------------------------
Private Function ReadFiDocument(ByVal sampleIdx As Long) As String
    Dim fieldId As String

    fieldId = modConfig.ElementId("Feban.Detail.DocNumber")

    If Not modSapConnect.Exists(fieldId) Then
        modLog.LogAction sampleIdx, "Chain", _
                     "Feban.Detail.DocNumber is not on this screen. Either the item " & _
                     "detail did not open, or the recorded ID belongs to a different " & _
                     "sub-screen. Current transaction: " & _
                     modSapConnect.CurrentTransaction(), "ERROR", vbNullString
        Exit Function
    End If

    ReadFiDocument = Trim$(modSapConnect.Element(fieldId).Text)
    If ReadFiDocument = "0" Then ReadFiDocument = vbNullString
End Function

' The recording drills in by putting focus on a field and pressing F2, which
' is 'choose/display' -- a read action, and allowed by the guard.
Private Sub DrillWithF2(ByVal mapKey As String)
    Dim fieldId As String

    fieldId = modConfig.ElementId(mapKey)
    modSapConnect.Element(fieldId).SetFocus
    modSafety.GuardedSendVKey "wnd[0]", 2
    modSafety.AssertPopupKnown
End Sub

' Open the line item the payment amount sits on.
Private Sub OpenAmountLine()
    Dim grid As Object

    Set grid = modSapConnect.Element(modConfig.ElementId("Doc.BsegGrid"))
    grid.currentCellColumn = modConfig.ElementId("Doc.Col.Amount")
    modSapConnect.WaitForSap

    grid.doubleClickCurrentCell
    modSapConnect.WaitForSap
End Sub

Private Function ReadClearingDocument(ByVal sampleIdx As Long) As String
    Dim fieldId As String

    fieldId = modConfig.ElementId("Doc.ClearingDocField")

    If Not modSapConnect.Exists(fieldId) Then
        modLog.LogAction sampleIdx, "Chain", _
                     "Doc.ClearingDocField is not on this screen -- the line-item " & _
                     "detail may not have opened.", "ERROR", vbNullString
        Exit Function
    End If

    ReadClearingDocument = Trim$(modSapConnect.Element(fieldId).Text)
    If ReadClearingDocument = "0" Then ReadClearingDocument = vbNullString
End Function

Private Sub OpenClearedItemsList(ByVal sampleIdx As Long)
    Dim anchorId As String

    modSapConnect.Element(modConfig.ElementId("Cleared.Menu")).Select
    modSapConnect.WaitForSap
    modSafety.AssertPopupKnown

    ' The anchor is a label the recording touched on the cleared-items list. If
    ' it is configured and missing, the menu produced something else.
    anchorId = modConfig.ElementIdOrBlank("Cleared.ListAnchor")
    If Len(anchorId) > 0 Then
        If Not modSapConnect.Exists(anchorId) Then
            modLog.LogAction sampleIdx, "Chain", _
                         "Cleared.Menu did not produce the expected list -- the anchor " & _
                         anchorId & " is not on screen. The menu path may differ here. " & _
                         "Clear Cleared.ListAnchor on the Screen Map to skip this check.", _
                         "ERROR", vbNullString
        End If
    End If
End Sub

'-----------------------------------------------------------------------
' Read the exported cleared-items list back and find the biggest item.
'-----------------------------------------------------------------------
Private Sub IdentifyLargestItem(ByVal sampleIdx As Long, ByRef result As ChainResult)
    Dim largest As ListRow
    Dim confirmingName As String

    largest = modListFile.LargestRow(result.ClearedItemsFile, sampleIdx)

    If Not largest.Found Then
        result.Notes = "Could not read a supplier and amount out of " & _
                       result.ClearedItemsFile & ". Open it and check the column " & _
                       "captions against the 'Cleared list ...' settings on the " & _
                       "Control sheet."
        modLog.LogAction sampleIdx, "Largest item", result.Notes, "ERROR", _
                     result.ClearedItemsFile
        Exit Sub
    End If

    result.LargestSupplier = largest.Supplier
    result.LargestAmount = largest.Amount
    result.LargestDocument = largest.DocumentNumber

    confirmingName = modConfig.Setting("Confirming party name")
    result.IsConfirmingPayment = NamesMatch(largest.Supplier, confirmingName)

    modLog.LogAction sampleIdx, "Largest item", _
                 "Largest of " & largest.RowsConsidered & " cleared items: " & _
                 largest.Supplier & " " & Format$(largest.Amount, "#,##0.00") & _
                 IIf(Len(largest.DocumentNumber) > 0, _
                     ", document " & largest.DocumentNumber, "") & _
                 ". Route: " & IIf(result.IsConfirmingPayment, _
                                   "confirming (Audit2)", "regular supplier"), _
                 "OK", result.ClearedItemsFile
End Sub

' 'SANTANDER SCF', 'SCF Santander' and 'Santander  SCF' are the same party.
' Compared on letters and digits only, in either order.
Public Function NamesMatch(ByVal a As String, ByVal b As String) As Boolean
    Dim tokensB() As String
    Dim i As Long
    Dim normalisedA As String

    If Len(Trim$(b)) = 0 Then Exit Function

    normalisedA = LettersAndDigits(a)
    tokensB = Split(modUtil.Squeeze(UCase$(b)), " ")

    ' Every word of the configured name must appear in the supplier name.
    For i = LBound(tokensB) To UBound(tokensB)
        If Len(tokensB(i)) > 0 Then
            If InStr(normalisedA, LettersAndDigits(tokensB(i))) = 0 Then Exit Function
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

'-----------------------------------------------------------------------
' Regular-supplier invoice route: open the payment and save its PDF.
'
' The recording did not cover this either -- Audit2.vbs is empty -- so the
' Invoice.* IDs are standard-but-unconfirmed. When they are not filled in,
' this logs MANUAL with the document numbers already captured, which is
' enough for someone to fetch the PDF by hand.
'-----------------------------------------------------------------------
Private Function DownloadRegularInvoice(ByVal sampleIdx As Long, ByRef result As ChainResult, _
                                        ByVal folder As String, _
                                        ByVal fileStem As String) As String
    Dim toolboxId As String
    Dim listId As String
    Dim saveId As String
    Dim toolbox As Object

    toolboxId = modConfig.ElementIdOrBlank("Invoice.GosToolbox")
    listId = modConfig.ElementIdOrBlank("Invoice.AttachListGrid")
    saveId = modConfig.ElementIdOrBlank("Invoice.SaveButton")

    If Len(toolboxId) = 0 Or Len(saveId) = 0 Then
        result.Notes = "Invoice.GosToolbox / Invoice.SaveButton are not filled in, so " & _
                       "no PDF download was attempted. Everything needed to fetch it " & _
                       "by hand is recorded: clearing document " & _
                       result.ClearingDocument & ", supplier " & result.LargestSupplier & _
                       IIf(Len(result.LargestDocument) > 0, _
                           ", invoice document " & result.LargestDocument, "") & "."
        modLog.LogAction sampleIdx, "Invoice PDF", result.Notes, "MANUAL", vbNullString
        Exit Function
    End If

    If modSafety.BlockedByDryRun("Would download the invoice PDF for " & _
                                 result.LargestSupplier) Then Exit Function

    If Not modSapConnect.Exists(toolboxId) Then
        result.Notes = "The services-for-object toolbox is not on this screen, so the " & _
                       "attachment list could not be opened."
        modLog.LogAction sampleIdx, "Invoice PDF", result.Notes, "MANUAL", vbNullString
        Exit Function
    End If

    Set toolbox = modSapConnect.Element(toolboxId)
    On Error Resume Next
    toolbox.pressContextButton "%GOS_TOOLBOX"
    toolbox.selectContextMenuItem "%GOS_VIEW_ATTA"
    On Error GoTo 0
    modSapConnect.WaitForSap

    If Not modSapConnect.ModalWindowOpen() Then
        result.Notes = "No attachment list appeared. On many configurations SAP hands " & _
                       "the document straight to the external viewer, which a script " & _
                       "cannot capture. Fetch this one by hand, or ask Basis for a " & _
                       "content-server extract covering all the samples at once."
        modLog.LogAction sampleIdx, "Invoice PDF", result.Notes, "MANUAL", vbNullString
        Exit Function
    End If

    If Not modSapConnect.Exists(saveId) Then
        result.Notes = "The attachment list opened but Invoice.SaveButton (" & saveId & _
                       ") is not on it. Record this dialog and correct the Screen Map."
        modLog.LogAction sampleIdx, "Invoice PDF", result.Notes, "MANUAL", vbNullString
        CloseModal
        Exit Function
    End If

    modSafety.GuardedPress saveId
    DownloadRegularInvoice = modExport.CompleteSaveDialogPublic( _
        sampleIdx, folder, fileStem & "_invoice.pdf")
End Function

Private Sub CloseModal()
    On Error Resume Next
    modSapConnect.Element("wnd[1]").Close
    On Error GoTo 0
    modSapConnect.WaitForSap
End Sub
