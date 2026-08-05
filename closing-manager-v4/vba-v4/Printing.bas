Attribute VB_Name = "Printing"
Option Base 1
Dim shellX, fsoX
Public FPathReport As String, FName As String
Sub SetPDFCreator()

Set shellX = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

'settings
'defPrinter = ""
'defDir = ""
'defName = ""
'
'device = "PDFCreator,winspool,Ne00:"    'nazwa drukarki PDF
'autoFilename = "<DateTime>" 'nazwa wyjsciowego pliku PDF => nazwa pliku, ktory przedrukujemy
'openAfterPrinting = 0   'otwieranie po wydruku
'
'shellX.regwrite "HKEY_CURRENT_USER\Software\PDFCreator\Program\DisableUpdateCheck", 1, "REG_SZ" 'sprawdzanie aktualizacji = FALSE
'shellX.regwrite "HKEY_CURRENT_USER\Software\PDFCreator\Program\useAutosave", 1, "REG_SZ"        'opcja autosave = TRUE
'shellX.regwrite "HKEY_CURRENT_USER\Software\PDFCreator\Program\NoProcessingAtStartup", 0, "REG_SZ"  'nie drukuj zaraz po wlaczeniu aplikacji = FALSE
'shellX.regwrite "HKEY_CURRENT_USER\Software\PDFCreator\Program\AutosaveStartStandardProgram", openAfterPrinting, "REG_SZ"   'otworz plik po wydruku = FALSE
'shellX.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\2500", 3
'
'defPrinter = shellX.regread("HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows\device")
'defDir = shellX.regread("HKEY_CURRENT_USER\Software\PDFCreator\Program\AutosaveDirectory")
'defName = shellX.regread("HKEY_CURRENT_USER\Software\PDFCreator\Program\AutosaveFilename")
'
'shellX.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows\device", device, "REG_SZ"   'drukarka PDF
'shellX.regwrite "HKEY_CURRENT_USER\Software\PDFCreator\Program\AutosaveDirectory", FTemp, "REG_SZ"            'sciezka wydruku
'shellX.regwrite "HKEY_CURRENT_USER\Software\PDFCreator\Program\AutosaveFilename", autoFilename, "REG_SZ"            'nazwa pliku

If fso.FileExists("C:\Program Files\PDFCreator\PDFCreator.exe") Then
    varProc = Shell("C:\Program Files\PDFCreator\PDFCreator.exe", 1)
ElseIf fso.FileExists("C:\Program Files (x86)\PDFCreator\PDFCreator.exe") Then
    varProc = Shell("C:\Program Files (x86)\PDFCreator\PDFCreator.exe", 1)
Else
    'V4-CIO FIX: was an unguarded Shell of the (x86) path -> run-time error if PDFCreator absent
    MsgBox "PDFCreator is not installed in the expected location." & vbCrLf & _
           "Install PDFCreator (the printer must be named 'PDFCreator') before running the close.", _
           vbCritical, "Closing Manager"
End If

End Sub
Sub SetPrinter()

Dim sess As Object
Set sess = SAPsess

With sess
    .findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    .findById("wnd[0]").sendVKey 0
    On Error Resume Next
    .findById("wnd[0]/usr/btnSTARTBUTTON").press
    On Error GoTo 0
    .findById("wnd[0]/tbar[0]/btn[86]").press
    .findById("wnd[0]/tbar[0]/btn[86]").press
    .findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCLX"
    .findById("wnd[1]").sendVKey 0
    On Error Resume Next
    .findById("wnd[1]/tbar[0]/btn[6]").press
    If Err.Number <> 0 Then
        On Error GoTo 0
        .findById("wnd[2]/tbar[0]/btn[0]").press
        .findById("wnd[1]/tbar[0]/btn[6]").press
    End If
    On Error GoTo 0
    .findById("wnd[2]/tbar[0]/btn[6]").press
    RowN = .findById("wnd[3]/usr/cntlCONTAINER/shellcont/shell").visiblerowcount
    Do Until RowN = 0
        .findById("wnd[3]/usr/cntlCONTAINER/shellcont/shell").selectedRows = "0"
        .findById("wnd[3]/usr/btnBUTTON2").press
        RowN = .findById("wnd[3]/usr/cntlCONTAINER/shellcont/shell").visiblerowcount
    Loop
    
    'ALV Selections
    .findById("wnd[3]/usr/cmbPRI_DEF-FNAME").Key = "ALVSL"
    .findById("wnd[3]/usr/btnBUTTON3").press
    .findById("wnd[4]/usr/subSUBSCREEN:SAPLSPRI:0600/chkPRIPAR_DYN-ALVSL").Selected = False
    .findById("wnd[4]/tbar[0]/btn[0]").press
    .findById("wnd[3]/usr/btnBUTTON1").press
    
    'ALV Statistics
    .findById("wnd[3]/usr/cmbPRI_DEF-FNAME").Key = "ALVST"
    .findById("wnd[3]/usr/btnBUTTON3").press
    .findById("wnd[4]/usr/subSUBSCREEN:SAPLSPRI:0600/chkPRIPAR_DYN-ALVST").Selected = False
    .findById("wnd[4]/tbar[0]/btn[0]").press
    .findById("wnd[3]/usr/btnBUTTON1").press
    
    'Delete immediately after printing
    .findById("wnd[3]/usr/cmbPRI_DEF-FNAME").Key = "PRREL"
    .findById("wnd[3]/usr/btnBUTTON3").press
    .findById("wnd[4]/usr/subSUBSCREEN:SAPLSPRI:0600/chkPRI_PARAMS-PRREL").Selected = True
    .findById("wnd[4]/tbar[0]/btn[0]").press
    .findById("wnd[3]/usr/btnBUTTON1").press
    
    'Operating System Cover Page
    .findById("wnd[3]/usr/cmbPRI_DEF-FNAME").Key = "PRUNX"
    .findById("wnd[3]/usr/btnBUTTON3").press
    .findById("wnd[4]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRUNX").Key = ""
    .findById("wnd[4]/tbar[0]/btn[0]").press
    .findById("wnd[3]/usr/btnBUTTON1").press
    
    'Print Time
    .findById("wnd[3]/usr/cmbPRI_DEF-FNAME").Key = "PRIMM"
    .findById("wnd[3]/usr/btnBUTTON3").press
    .findById("wnd[4]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRIMM2").Key = "X"
    .findById("wnd[4]/tbar[0]/btn[0]").press
    .findById("wnd[3]/usr/btnBUTTON1").press
    
    'SAP Cover Page
    .findById("wnd[3]/usr/cmbPRI_DEF-FNAME").Key = "PRSAP"
    .findById("wnd[3]/usr/btnBUTTON3").press
    .findById("wnd[4]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRSAP").Key = ""
    .findById("wnd[4]/tbar[0]/btn[0]").press
    .findById("wnd[3]/usr/btnBUTTON1").press
    
    .findById("wnd[3]/tbar[0]/btn[0]").press
    .findById("wnd[2]/tbar[0]/btn[0]").press
    .findById("wnd[1]/tbar[0]/btn[12]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
End With

End Sub
Sub Print_ZGLRME(sess, printN)

Set shellX = CreateObject("WScript.Shell")
Set fso = CreateObject("scripting.filesystemobject")

Call SetPDFCreator

'delete all from temp file
fso.DeleteFile (FTemp & "\*.*"), True

With sess
    'only transmitted
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nzglrme"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/chkP_XMIT").Selected = True
    .findById("wnd[0]/usr/ctxtS_BUKRS-LOW").Text = CC
    .findById("wnd[0]/usr/txtP_MONAT").Text = Monthx
    .findById("wnd[0]/usr/txtP_GJAHR").Text = Yearx
    .findById("wnd[0]/usr/ctxtP_VARID").Text = "/default"
    .findById("wnd[0]/usr/ctxtP_VARIE").Text = "/default"
    .findById("wnd[0]/tbar[1]/btn[8]").press
    On Error Resume Next
    .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").pressToolbarContextButton "&MB_FILTER"
    If Err.Number = 0 Then
        On Error GoTo 0
        .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").selectContextMenuItem "&DELETE_FILTER"
        .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").setCurrentCell -1, "SACCT"
        .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").selectColumn "SACCT"
        .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").pressToolbarButton "&MB_FILTER"
        On Error Resume Next
        .findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
        If Err.Number <> 0 Then
            On Error GoTo 0
            .findById("wnd[1]/tbar[0]/btn[0]").press
            .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").pressToolbarButton "&PRINT_BACK"
            .findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCLX"
            .findById("wnd[1]").sendVKey 0
            On Error Resume Next
            .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
            If Err.Number <> 0 Then
                On Error GoTo 0
                .findById("wnd[2]/tbar[0]/btn[0]").press
                .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
            End If
            On Error GoTo 0
            .findById("wnd[1]/tbar[0]/btn[13]").press
            
        Else
            On Error GoTo 0
            
            .findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").Select
            .findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = "85500080"
            .findById("wnd[2]").sendVKey 8
            .findById("wnd[1]").sendVKey 0
            .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").setCurrentCell -1, "MESSNUM"
            .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").selectColumn "MESSNUM"
            .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").pressToolbarButton "&MB_FILTER"
            .findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN002_%_APP_%-VALU_PUSH").press
            .findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").Select
            .findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = "E10"
            .findById("wnd[2]/tbar[0]/btn[8]").press
            .findById("wnd[1]/tbar[0]/btn[0]").press
            .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").setCurrentCell -1, "MTYPE"
            .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").selectColumn "PRCTR"
            .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").selectColumn "CTYPE"
            .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").selectColumn "MTYPE"
            .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").pressToolbarButton "&MB_SUBTOT"
            .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").pressToolbarContextButton "&MB_VIEW"
            .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").selectContextMenuItem "&PRINT_BACK_PREVIEW"
            .findById("wnd[0]/mbar/menu[3]/menu[6]/menu[0]").Select
            .findById("wnd[0]/tbar[0]/btn[86]").press
            .findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCLX"
            .findById("wnd[1]").sendVKey 0
            On Error Resume Next
            .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
            If Err.Number <> 0 Then
                On Error GoTo 0
                .findById("wnd[2]/tbar[0]/btn[0]").press
                .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
            End If
            On Error GoTo 0
            .findById("wnd[1]/tbar[0]/btn[13]").press
            .findById("wnd[0]/tbar[0]/btn[3]").press
        End If
        .findById("wnd[0]/tbar[0]/btn[3]").press
        
        File = ""
        'V4-CIO FIX: bounded, responsive wait for the printed PDF. The old loop
        'rebuilt a Shell.Application on every pass with no pause, so an empty
        'temp folder became a tight spin: Excel showed "Not Responding" and the
        'shell object eventually dropped out with error 80010108.
        File = CM_WaitForPrint(FTemp, fso, "ZGLRME (errors only)")
        
        fso.MoveFile File, FFinal & "\" & CC & "-" & Yearx & "-" & Right("0" & Monthx, 2) & "-" & printN & ".pdf"
        printN = printN + 1
    End If
    On Error GoTo 0

    'all
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nzglrme"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/chkP_XMIT").Selected = False
    .findById("wnd[0]/usr/ctxtS_BUKRS-LOW").Text = CC
    .findById("wnd[0]/usr/txtP_MONAT").Text = Monthx
    .findById("wnd[0]/usr/txtP_GJAHR").Text = Yearx
    .findById("wnd[0]/usr/ctxtP_VARID").Text = "/default"
    .findById("wnd[0]/usr/ctxtP_VARIE").Text = "/default"
    .findById("wnd[0]/tbar[1]/btn[8]").press
    .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").pressToolbarContextButton "&MB_FILTER"
    .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").selectContextMenuItem "&DELETE_FILTER"
    .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").setCurrentCell -1, "SACCT"
    .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").selectColumn "SACCT"
    .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").pressToolbarButton "&MB_FILTER"
    On Error Resume Next
    .findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
    If Err.Number <> 0 Then
        On Error GoTo 0
        .findById("wnd[1]/tbar[0]/btn[0]").press
        .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").pressToolbarButton "&PRINT_BACK"
        .findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCLX"
        .findById("wnd[1]").sendVKey 0
        On Error Resume Next
        .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
        If Err.Number <> 0 Then
            On Error GoTo 0
            .findById("wnd[2]/tbar[0]/btn[0]").press
            .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
        End If
        On Error GoTo 0
        .findById("wnd[1]/tbar[0]/btn[13]").press
        
    Else
        On Error GoTo 0
        
        .findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").Select
        .findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = "85500080"
        .findById("wnd[2]").sendVKey 8
        .findById("wnd[1]").sendVKey 0
        .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").setCurrentCell -1, "MESSNUM"
        .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").selectColumn "MESSNUM"
        .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").pressToolbarButton "&MB_FILTER"
        .findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN002_%_APP_%-VALU_PUSH").press
        .findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").Select
        .findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = "E10"
        .findById("wnd[2]/tbar[0]/btn[8]").press
        .findById("wnd[1]/tbar[0]/btn[0]").press
        .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").setCurrentCell -1, "MTYPE"
        .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").selectColumn "PRCTR"
        .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").selectColumn "CTYPE"
        .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").selectColumn "MTYPE"
        .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").pressToolbarButton "&MB_SUBTOT"
        .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").pressToolbarContextButton "&MB_VIEW"
        .findById("wnd[0]/usr/tabsTAB100/tabpTAB100_FC1/ssubTAB100_SCA:ZGLE10669_MNTH_END_CLOSE_CHECK:0101/cntlC100TAB1/shellcont/shell").selectContextMenuItem "&PRINT_BACK_PREVIEW"
        .findById("wnd[0]/mbar/menu[3]/menu[6]/menu[0]").Select
        .findById("wnd[0]/tbar[0]/btn[86]").press
        .findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCLX"
        .findById("wnd[1]").sendVKey 0
        On Error Resume Next
        .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
        If Err.Number <> 0 Then
            On Error GoTo 0
            .findById("wnd[2]/tbar[0]/btn[0]").press
            .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
        End If
        On Error GoTo 0
        .findById("wnd[1]/tbar[0]/btn[13]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
    End If
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    
    File = ""
    'V4-CIO FIX: bounded, responsive wait for the printed PDF. The old loop
    'rebuilt a Shell.Application on every pass with no pause, so an empty
    'temp folder became a tight spin: Excel showed "Not Responding" and the
    'shell object eventually dropped out with error 80010108.
    File = CM_WaitForPrint(FTemp, fso, "ZGLRME (full report)")
    
    fso.MoveFile File, FFinal & "\" & CC & "-" & Yearx & "-" & Right("0" & Monthx, 2) & "-" & printN & ".pdf"
    printN = printN + 1
End With

End Sub
Function Print_EIS4(sess, printN)

Dim Am As Double

Set fso = CreateObject("scripting.filesystemobject")

Call SetPDFCreator

'delete all from temp file
fso.DeleteFile (FTemp & "\*.*"), True

With sess
    .findById("wnd[0]").maximize
    .findById("wnd[0]/tbar[0]/okcd").Text = "/ngr55"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/ctxtRGRWJ-JOB").Text = "eis4"
    .findById("wnd[0]").sendVKey 8
    .findById("wnd[0]/usr/txt$CUR-YR").Text = Yearx
    .findById("wnd[0]/usr/txt$CUR-PER").Text = Monthx
    .findById("wnd[0]/usr/ctxt_COCODES-LOW").Text = CC
    .findById("wnd[0]").sendVKey 8
    On Error Resume Next
    .findById("wnd[0]/shellcont/shell/shellcont[2]/shell").hierarchyHeaderWidth = 448
    If Err.Number = 0 Then
        .findById("wnd[0]/tbar[0]/btn[86]").press
        .findById("wnd[1]/tbar[0]/btn[0]").press
        .findById("wnd[1]/tbar[0]/btn[0]").press
        .findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCLX"
        .findById("wnd[1]").sendVKey 0
        On Error Resume Next
        .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
        If Err.Number <> 0 Then
            On Error GoTo 0
            .findById("wnd[2]/tbar[0]/btn[0]").press
            .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
        End If
        On Error GoTo 0
        .findById("wnd[1]/tbar[0]/btn[13]").press
        .findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").Select
        .findById("wnd[1]/tbar[0]/btn[0]").press
        .findById("wnd[1]/usr/ctxtDY_PATH").Text = FPath
        .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "eis4.txt"
        .findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").Text = "4120"
        .findById("wnd[1]/tbar[0]/btn[11]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
        On Error Resume Next
        .findById("wnd[1]/usr/btnBUTTON_YES").press
        On Error GoTo 0
        .findById("wnd[0]/tbar[0]/btn[3]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
    Else
        Print_EIS4 = 0
        Exit Function
    End If
End With

File = ""
'V4-CIO FIX: bounded, responsive wait for the printed PDF. The old loop
'rebuilt a Shell.Application on every pass with no pause, so an empty
'temp folder became a tight spin: Excel showed "Not Responding" and the
'shell object eventually dropped out with error 80010108.
File = CM_WaitForPrint(FTemp, fso, "report group EIS4")

fso.MoveFile File, FFinal & "\" & CC & "-" & Yearx & "-" & Right("0" & Monthx, 2) & "-" & printN & ".pdf"
printN = printN + 1

Set strim = New ADODB.Stream
strim.Charset = "utf-8"
strim.Open

strim.LoadFromFile (FPath & "eis4.txt")

Do Until strim.EOS
    Dim line As String
    line = strim.ReadText(-2)
Loop

n = InStr(1, line, "  ", vbBinaryCompare)
Do Until n = 0
    line = Replace(line, "  ", " ")
    n = InStr(1, line, "  ", vbBinaryCompare)
Loop
arr = Split(line, " ")

Print_EIS4 = Round(Replace(Replace(arr(UBound(arr, 1) - 1), "(", "-"), ")", ""), 2)

Kill FPath & "eis4.txt"

End Function
Function Print_GIS4(sess, printN)

Dim Am As Double

Set fso = CreateObject("scripting.filesystemobject")

Call SetPDFCreator

'delete all from temp file
fso.DeleteFile (FTemp & "\*.*"), True

With sess
    .findById("wnd[0]").maximize
    .findById("wnd[0]/tbar[0]/okcd").Text = "/ngr55"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/ctxtRGRWJ-JOB").Text = "gis4"
    .findById("wnd[0]").sendVKey 8
    .findById("wnd[0]/usr/txt$CUR-YR").Text = Yearx
    .findById("wnd[0]/usr/txt$CUR-PER").Text = Monthx
    .findById("wnd[0]/usr/ctxt_COCODES-LOW").Text = CC
    .findById("wnd[0]").sendVKey 8
    .findById("wnd[0]/tbar[0]/btn[86]").press
    On Error Resume Next
    .findById("wnd[1]/tbar[0]/btn[0]").press
    If Err.Number = 0 Then
        .findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCLX"
        .findById("wnd[1]").sendVKey 0
        On Error Resume Next
        .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
        If Err.Number <> 0 Then
            On Error GoTo 0
            .findById("wnd[2]/tbar[0]/btn[0]").press
            .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
        End If
        On Error GoTo 0
        .findById("wnd[1]/tbar[0]/btn[13]").press
        .findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").Select
        .findById("wnd[1]/tbar[0]/btn[0]").press
        .findById("wnd[1]/usr/ctxtDY_PATH").Text = FPath
        .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "gis4.txt"
        .findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").Text = "4120"
        .findById("wnd[1]/tbar[0]/btn[11]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
        On Error Resume Next
        .findById("wnd[1]/usr/btnBUTTON_YES").press
        On Error GoTo 0
        .findById("wnd[0]/tbar[0]/btn[3]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
    Else
        Print_GIS4 = 0
        Exit Function
    End If
End With

File = ""
'V4-CIO FIX: bounded, responsive wait for the printed PDF. The old loop
'rebuilt a Shell.Application on every pass with no pause, so an empty
'temp folder became a tight spin: Excel showed "Not Responding" and the
'shell object eventually dropped out with error 80010108.
File = CM_WaitForPrint(FTemp, fso, "report group GIS4")

fso.MoveFile File, FFinal & "\" & CC & "-" & Yearx & "-" & Right("0" & Monthx, 2) & "-" & printN & ".pdf"
printN = printN + 1

Set strim = New ADODB.Stream
strim.Charset = "utf-8"
strim.Open

strim.LoadFromFile (FPath & "gis4.txt")

Do Until strim.EOS
    Dim line As String
    line = strim.ReadText(-2)
Loop

n = InStr(1, line, "  ", vbBinaryCompare)
Do Until n = 0
    line = Replace(line, "  ", " ")
    n = InStr(1, line, "  ", vbBinaryCompare)
Loop
arr = Split(line, " ")

Print_GIS4 = Round(Replace(Replace(arr(UBound(arr, 1) - 1), "(", "-"), ")", ""), 2)

Kill FPath & "gis4.txt"

End Function
Function Print_ZGE132(sess, ArrPC, ArrInd)

Dim Am As Double, AmL As Double, AmEIS4 As Double, AmG As Double, AmGIS4 As Double
Dim rng As Range

Set fso = CreateObject("scripting.filesystemobject")

Call SetPDFCreator

'delete all from temp file
fso.DeleteFile (FTemp & "\*.*"), True

With sess
    'local currency
    .findById("wnd[0]").maximize
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nzge132"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/chkP_CLENT").Selected = True
    .findById("wnd[0]/usr/chkP_BCENT").Selected = True
    .findById("wnd[0]/usr/ctxtP_CCODE").Text = CC
    .findById("wnd[0]/usr/txtP_MONTH").Text = Monthx
    .findById("wnd[0]/usr/txtP_YEAR").Text = Yearx
    .findById("wnd[0]/usr/chkP_BCENT").SetFocus
    .findById("wnd[0]/tbar[1]/btn[8]").press
    .findById("wnd[0]/tbar[0]/btn[86]").press
    .findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCLX"
    .findById("wnd[1]").sendVKey 0
    On Error Resume Next
    .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
    If Err.Number <> 0 Then
        On Error GoTo 0
        .findById("wnd[2]/tbar[0]/btn[0]").press
        .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
    End If
    On Error GoTo 0
    .findById("wnd[1]/tbar[0]/btn[13]").press
    .findById("wnd[0]/mbar/menu[0]/menu[5]/menu[2]/menu[2]").Select
    .findById("wnd[1]").sendVKey 0
    .findById("wnd[1]/usr/ctxtDY_PATH").Text = FPath
    .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "zge132.txt"
    .findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").Text = "4120"
    .findById("wnd[1]/tbar[0]/btn[11]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    
    File = ""
    'V4-CIO FIX: bounded, responsive wait for the printed PDF. The old loop
    'rebuilt a Shell.Application on every pass with no pause, so an empty
    'temp folder became a tight spin: Excel showed "Not Responding" and the
    'shell object eventually dropped out with error 80010108.
    File = CM_WaitForPrint(FTemp, fso, "ZGE132 (before posting)")
    
    fso.MoveFile File, FFinal & "\" & CC & "-" & Yearx & "-" & Right("0" & Monthx, 2) & "-" & printN & ".pdf"
    printN = printN + 1
    
    'group currency
    .findById("wnd[0]/tbar[1]/btn[25]").press
    .findById("wnd[0]/usr/chkP_BLC").Selected = True
    .findById("wnd[0]/usr/ctxtP_BUKRS").Text = CC
    .findById("wnd[0]/usr/txtP_MONTH").Text = Monthx
    .findById("wnd[0]/usr/txtP_YEAR").Text = Yearx
    .findById("wnd[0]/usr/chkP_BLC").SetFocus
    .findById("wnd[0]/tbar[1]/btn[8]").press
    .findById("wnd[0]/tbar[0]/btn[86]").press
    .findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCLX"
    .findById("wnd[1]").sendVKey 0
    On Error Resume Next
    .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
    If Err.Number <> 0 Then
        On Error GoTo 0
        .findById("wnd[2]/tbar[0]/btn[0]").press
        .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
    End If
    On Error GoTo 0
    .findById("wnd[1]/tbar[0]/btn[13]").press
    .findById("wnd[0]/tbar[1]/btn[45]").press
    .findById("wnd[1]/tbar[0]/btn[0]").press
    .findById("wnd[1]/usr/ctxtDY_PATH").Text = FPath
    .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "zge132G.txt"
    .findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").Text = "4120"
    .findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").SetFocus
    .findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").caretPosition = 4
    .findById("wnd[1]/tbar[0]/btn[11]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    
    File = ""
    'V4-CIO FIX: bounded, responsive wait for the printed PDF. The old loop
    'rebuilt a Shell.Application on every pass with no pause, so an empty
    'temp folder became a tight spin: Excel showed "Not Responding" and the
    'shell object eventually dropped out with error 80010108.
    File = CM_WaitForPrint(FTemp, fso, "ZGE132 (after posting)")
    
    fso.MoveFile File, FFinal & "\" & CC & "-" & Yearx & "-" & Right("0" & Monthx, 2) & "-" & printN & ".pdf"
    printN = printN + 1
    
End With

'----------------------------------------------------------------
'local currency
'----------------------------------------------------------------
LastRow = FindLastRow(1, 1, 0, 0, "ZGE132")
If LastRow > 1 Then Sheets("ZGE132").Range("A2", Sheets("ZGE132").Cells(LastRow, 3)).ClearContents

Set strim = New ADODB.Stream
strim.Charset = "utf-8"
strim.Open

strim.LoadFromFile (FPath & "zge132.txt")

Do Until strim.EOS
    Dim line As String
    line = strim.ReadText(-2)

    If IsNumeric(Left(Trim(line), 1)) Then
        n = InStr(1, line, "  ", vbBinaryCompare)
        Do Until n = 0
            line = Replace(line, "  ", " ")
            n = InStr(1, line, "  ", vbBinaryCompare)
        Loop
        arr = Split(line, " ")
        If UBound(arr) = 1 Then
            If Right(arr(LBound(arr) + 1), 1) = "-" Then
                AmL = -Round(Left(arr(LBound(arr) + 1), Len(arr(LBound(arr) + 1)) - 1), 2)
            Else
                AmL = Round(arr(LBound(arr) + 1), 2)
            End If
        ElseIf arr(UBound(arr) - 1) = "||" Then
            If Right(arr(UBound(arr)), 1) = "-" Then
                Am = -Round(Left(arr(UBound(arr)), Len(arr(UBound(arr))) - 1), 2)
            Else
                Am = Round(arr(UBound(arr)), 2)
            End If
            PC = arr(LBound(arr))
            EmptRow = FindLastRow(1, 1, 1, 0, "ZGE132")
            Sheets("ZGE132").Cells(EmptRow, 1) = PC
            Sheets("ZGE132").Cells(EmptRow, 2) = Am
        End If
        
    End If
Loop

'----------------------------------------------------------------
'check errors in ZGE132 (local currency)
'----------------------------------------------------------------
Call CheckZGE132("L", ArrPC, ArrInd)

'----------------------------------------------------------------
'group currency
'----------------------------------------------------------------
LastRow = FindLastRow(1, 1, 0, 0, "ZGE132")
If LastRow > 1 Then Sheets("ZGE132").Range("A2", Sheets("ZGE132").Cells(LastRow, 3)).ClearContents

Set strim = New ADODB.Stream
strim.Charset = "utf-8"
strim.Open

strim.LoadFromFile (FPath & "zge132G.txt")

Do Until strim.EOS
    line = strim.ReadText(-2)
    If Left(Trim(line), 1) = "|" Then
        arr = Split(line, "|")
        If UBound(arr) = 6 And Left(arr(LBound(arr) + 1), 7) = "* Total" Then
            If Right(arr(UBound(arr) - 3), 1) = "-" Then
                AmG = -Round(Left(arr(UBound(arr) - 3), Len(arr(UBound(arr) - 3)) - 1), 2)
            Else
                AmG = Round(arr(UBound(arr) - 3), 2)
            End If
        End If
        If UBound(arr) = 7 And Trim(arr(LBound(arr) + 1)) <> "Profit Ctr" And Left(arr(LBound(arr) + 1), 7) <> "* Total" Then
            EmptRow = FindLastRow(1, 1, 1, 0, "ZGE132")
            Sheets("ZGE132").Cells(EmptRow, 1) = Trim(arr(LBound(arr) + 1))
            If Right(arr(UBound(arr) - 2), 1) = "-" Then
                Am = -Round(Left(arr(UBound(arr) - 2), Len(arr(UBound(arr) - 2)) - 1), 2)
            Else
                Am = Round(arr(UBound(arr) - 2), 2)
            End If
            Sheets("ZGE132").Cells(EmptRow, 2) = Am
        End If
    End If
Loop

'----------------------------------------------------------------
'check errors in ZGE132 (group currency)
'----------------------------------------------------------------
Call CheckZGE132("G", ArrPC, ArrInd)

'clearing Error tab if no errors
If BPCL = False And CPCL = False And BPCG = False And CPCG = False Then
    LastRow = FindLastRow(1, 2, 0, 0, "Errors")
    Set rng = Sheets("Errors").Range("A1", Sheets("Errors").Cells(LastRow, 30))
    Call ApplyBorders(0, 0, rng)
    With Sheets("Errors").Cells
        .ClearContents
        .ClearFormats
        .Font.Bold = False
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.ThemeColor = xlThemeColorDark1
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
    End With
Else
    Print_ZGE132 = False
    Sheets("Errors").Visible = True
    Sheets("Errors").Select
    Range("A1").Select
End If

'check if ZGE132 is equal to EIS4 and GIS4
AmEIS4 = Sheets("config").Range("AA4")
AmGIS4 = Sheets("config").Range("AA10")

If Round(AmEIS4 + AmL, 2) <> 0 Then
    Print_ZGE132 = False
    ZGEISF = True
Else
    Sheets("config").Range("AA5") = AmEIS4
End If
If Round(AmGIS4 + AmG, 2) <> 0 Then
    Print_ZGE132 = False
    ZGGISF = True
Else
    Sheets("config").Range("AA11") = AmGIS4
End If

If ZGEISF = False And ZGGISF = False And BPCG = False And BPCL = False And CPCG = False And CPCL = False Then
    Print_ZGE132 = True
End If

End Function
Sub CheckZGE132(CType, ArrPC, ArrInd)

Dim rng As Range

Sheets("Errors").Visible = True

LastRow = FindLastRow(1, 1, 0, 0, "ZGE132")
ArrZG = Sheets("ZGE132").Range("A2", Sheets("ZGE132").Cells(LastRow, 5))

For i = 1 To UBound(ArrZG, 1)
    For j = 1 To UBound(ArrPC, 1)
        If CStr(ArrZG(i, 1)) = CStr(ArrPC(j, 2)) Then
            ArrZG(i, 3) = ArrPC(j, 3)
            ArrZG(i, 4) = ArrPC(j, 4)
            For l = 1 To UBound(ArrInd, 1)
                If CStr(ArrZG(i, 4)) = CStr(ArrInd(l, 1)) Then
                    ArrZG(i, 5) = ArrInd(l, 2)
                    Exit For
                End If
            Next l
            Exit For
        End If
    Next j
Next i

'check blanks
If CType = "L" Then
    Sheets("Errors").Range("A10") = "Blank Profit Center in local currency:"
    Sheets("Errors").Range("A10").Font.Bold = True
ElseIf CType = "G" Then
    EmptRow = FindLastRow(1, 2, 2, 0, "Errors")
    Sheets("Errors").Cells(EmptRow, 1) = "Blank Profit Center in group currency:"
    Sheets("Errors").Cells(EmptRow, 1).Font.Bold = True
End If

AmB = 0
For i = 1 To UBound(ArrZG, 1)
    If Round(ArrZG(i, 2), 2) <> 0 Then AmB = AmB + Round(ArrZG(i, 2), 2)
Next i

If Round(AmB, 2) <> 0 Then
    EmptRow = FindLastRow(1, 1, 1, 0, "Errors")
    Sheets("Errors").Cells(EmptRow, 2) = Format(AmB, "#,##0.00") & " " & Cur
    If CType = "L" Then BPCL = True Else BPCG = True
Else
    EmptRow = FindLastRow(1, 1, 1, 0, "Errors")
    Sheets("Errors").Cells(EmptRow, 2) = "No errors."
End If

'check cross errors
EmptRow = FindLastRow(1, 2, 2, 0, "Errors")
If CType = "L" Then
    Sheets("Errors").Cells(EmptRow, 1) = "Cross Profit Center errors in local currency:"
ElseIf CType = "G" Then
    Sheets("Errors").Cells(EmptRow, 1) = "Cross Profit Center errors in group currency:"
End If
Sheets("Errors").Cells(EmptRow, 1).Font.Bold = True
Sheets("Errors").Cells(EmptRow + 1, 2) = "CoCode"
Sheets("Errors").Cells(EmptRow + 1, 3) = "Profit Ctr"
Sheets("Errors").Cells(EmptRow + 1, 4) = "Amount"
Sheets("Errors").Cells(EmptRow + 1, 5) = "Transmit"
Sheets("Errors").Cells(EmptRow + 1, 6) = "GAAP ind"
Sheets("Errors").Cells(EmptRow + 1, 7) = "GAAP ind description"

If PB = 0 And GAAP = 0 And NA = 0 Then
    Sheets("Errors").Cells(EmptRow + 1, 1).EntireRow.ClearContents
    Sheets("Errors").Cells(EmptRow + 1, 2) = "Cross Profit Center postings allowed in all cases."
Else

    LastRow = FindLastRow(1, 1, 0, 0, "ZGE132")
    Dim ArrZGG()
    ReDim ArrZGG(LastRow, 2)
    
    'making subtotals by GAAP indicator
    l = 1
    For i = 1 To UBound(ArrZG, 1)
        found = False
        For j = 1 To UBound(ArrZGG, 1)
            If CStr(ArrZG(i, 4)) = CStr(ArrZGG(j, 1)) Then
                found = True
                ArrZGG(j, 2) = ArrZGG(j, 2) + ArrZG(i, 2)
                Exit For
            End If
        Next j
        If found = False Then
            ArrZGG(l, 1) = ArrZG(i, 4)
            ArrZGG(l, 2) = ArrZG(i, 2)
            l = l + 1
        End If
    Next i
    
    'checking cross-PC
    If PB = 1 Then
        'no cross allowed
        For i = 1 To UBound(ArrZG, 1)
            If Round(ArrZG(i, 2), 2) <> 0 Then
                EmptRow = FindLastRow(1, 2, 1, 0, "Errors")
                Sheets("Errors").Cells(EmptRow, 2) = CC
                For k = 1 To UBound(ArrZG, 2)
                    Sheets("Errors").Cells(EmptRow, k + 2) = ArrZG(i, k)
                Next k
            End If
        Next i
    ElseIf GAAP = 0 Then
        'no cross allowed
        For i = 1 To UBound(ArrZG, 1)
            If Round(ArrZG(i, 2), 2) <> 0 Then
                EmptRow = FindLastRow(1, 2, 1, 0, "Errors")
                Sheets("Errors").Cells(EmptRow, 2) = CC
                For k = 1 To UBound(ArrZG, 2)
                    Sheets("Errors").Cells(EmptRow, k + 2) = ArrZG(i, k)
                Next k
            End If
        Next i
    ElseIf NA = 1 Then
        'no cross allowed
        For i = 1 To UBound(ArrZG, 1)
            If Round(ArrZG(i, 2), 2) <> 0 Then
                EmptRow = FindLastRow(1, 2, 1, 0, "Errors")
                Sheets("Errors").Cells(EmptRow, 2) = CC
                For k = 1 To UBound(ArrZG, 2)
                    Sheets("Errors").Cells(EmptRow, k + 2) = ArrZG(i, k)
                Next k
            End If
        Next i
    ElseIf GAAP = 1 Then
        'cross within one GAAP ind allowed
        For i = 1 To UBound(ArrZGG, 1)
            If Round(ArrZGG(i, 2), 2) <> 0 Then
                For j = 1 To UBound(ArrZG, 1)
                    If Round(ArrZG(j, 2), 2) <> 0 Then
                        EmptRow = FindLastRow(1, 2, 1, 0, "Errors")
                        Sheets("Errors").Cells(EmptRow, 2) = CC
                        For k = 1 To UBound(ArrZG, 2)
                            Sheets("Errors").Cells(EmptRow, k + 2) = ArrZG(j, k)
                        Next k
                    End If
                Next j
                Exit For
            End If
        Next i
    End If
End If

LastRow = FindLastRow(1, 2, 0, 0, "Errors")
FirstRow = FindLastRow(1, 1, 1, 0, "Errors")
If LastRow = FirstRow Then
    Sheets("Errors").Cells(LastRow, 1).EntireRow.ClearContents
    Sheets("Errors").Cells(LastRow, 2) = "No cross Profit Center errors."
Else
    If CType = "L" Then CPCL = True Else CPCG = True
    Set rng = Sheets("Errors").Range(Sheets("Errors").Cells(FirstRow, 2), Sheets("Errors").Cells(LastRow, 7))
    Call ApplyBorders(1, 1, rng)
    Call ApplyBorders(2, 2, rng)
    With Sheets("Errors").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("F" & FirstRow & ":F" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange rng
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    rng.HorizontalAlignment = xlCenter
    With Sheets("Errors").Range(Sheets("Errors").Cells(FirstRow, 2), Sheets("Errors").Cells(FirstRow, 7))
        .EntireColumn.AutoFit
        .Font.Bold = True
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 14071936
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
    End With
    Sheets("Errors").Range(Sheets("Errors").Cells(FirstRow + 1, 4), Sheets("Errors").Cells(LastRow, 4)).NumberFormat = "#,##0.00"
    w = FirstRow + 1
    With Sheets("Errors")
        Do Until .Cells(w, 2) = ""
            If .Cells(w, 6) <> .Cells(w + 1, 6) Then
                .Range(.Cells(w, 2), .Cells(w, 7)).Borders(xlEdgeBottom).Weight = xlMedium
            End If
            w = w + 1
        Loop
    End With
End If

End Sub
Sub Print_ZGR215(sess)

Set fso = CreateObject("Scripting.FileSystemObject")

With sess
    .findById("wnd[0]").maximize
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nzgr215"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/btn%_SBELNR_%_APP_%-VALU_PUSH").press
    .findById("wnd[1]/tbar[0]/btn[23]").press
    .findById("wnd[2]/usr/ctxtDY_PATH").Text = FPath
    .findById("wnd[2]/usr/ctxtDY_FILENAME").Text = "DocL.csv"
    .findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 8
    .findById("wnd[2]/tbar[0]/btn[0]").press
    .findById("wnd[1]/tbar[0]/btn[8]").press
    .findById("wnd[0]/usr/ctxtSBUKRS").Text = CC
    .findById("wnd[0]/usr/txtSYEAR").Text = Yearx
    .findById("wnd[0]/usr/chkP_PALV").SetFocus
    .findById("wnd[0]/usr/chkP_PALV").Selected = True
    .findById("wnd[0]/usr/ctxtP_ALV").Text = "/arek2" 'utworzyc nowy layout
    .findById("wnd[0]/tbar[1]/btn[8]").press
    .findById("wnd[0]/mbar/menu[4]/menu[1]/menu[0]").Select
    .findById("wnd[0]/tbar[0]/btn[86]").press
    .findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCLX"
    .findById("wnd[1]").sendVKey 0
    On Error Resume Next
    .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
    If Err.Number <> 0 Then
        On Error GoTo 0
        .findById("wnd[2]/tbar[0]/btn[0]").press
        .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
    End If
    On Error GoTo 0
    .findById("wnd[1]/tbar[0]/btn[13]").press
    
    File = ""
    'V4-CIO FIX: bounded, responsive wait for the printed PDF. The old loop
    'rebuilt a Shell.Application on every pass with no pause, so an empty
    'temp folder became a tight spin: Excel showed "Not Responding" and the
    'shell object eventually dropped out with error 80010108.
    File = CM_WaitForPrint(FTemp, fso, "ZGR215 (document list)")
    
    fso.MoveFile File, FFinal & "\" & CC & "-" & Yearx & "-" & Right("0" & Monthx, 2) & "-" & printN & ".pdf"
    printN = printN + 1
    
    If Sheets("config").Range("AA12") <> "No posting" Then
        .findById("wnd[0]").maximize
        .findById("wnd[0]/tbar[0]/okcd").Text = "/nzgr215"
        .findById("wnd[0]").sendVKey 0
        .findById("wnd[0]/usr/btn%_SBELNR_%_APP_%-VALU_PUSH").press
        .findById("wnd[1]/tbar[0]/btn[23]").press
        .findById("wnd[2]/usr/ctxtDY_PATH").Text = FPath
        .findById("wnd[2]/usr/ctxtDY_FILENAME").Text = "DocG.csv"
        .findById("wnd[2]/usr/ctxtDY_FILENAME").caretPosition = 8
        .findById("wnd[2]/tbar[0]/btn[0]").press
        .findById("wnd[1]/tbar[0]/btn[8]").press
        .findById("wnd[0]/usr/ctxtSBUKRS").Text = CC
        .findById("wnd[0]/usr/txtSYEAR").Text = Yearx
        .findById("wnd[0]/usr/chkP_PALV").SetFocus
        .findById("wnd[0]/usr/chkP_PALV").Selected = True
        .findById("wnd[0]/usr/ctxtP_ALV").Text = "/arek2" 'utworzyc nowy layout
        .findById("wnd[0]/tbar[1]/btn[8]").press
        .findById("wnd[0]/mbar/menu[4]/menu[1]/menu[0]").Select
        .findById("wnd[0]/tbar[0]/btn[86]").press
        .findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCLX"
        .findById("wnd[1]").sendVKey 0
        On Error Resume Next
        .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
        If Err.Number <> 0 Then
            On Error GoTo 0
            .findById("wnd[2]/tbar[0]/btn[0]").press
            .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
        End If
        On Error GoTo 0
        .findById("wnd[1]/tbar[0]/btn[13]").press
        
        File = ""
        'V4-CIO FIX: bounded, responsive wait for the printed PDF. The old loop
        'rebuilt a Shell.Application on every pass with no pause, so an empty
        'temp folder became a tight spin: Excel showed "Not Responding" and the
        'shell object eventually dropped out with error 80010108.
        File = CM_WaitForPrint(FTemp, fso, "ZGR215 (documents)")
        
        fso.MoveFile File, FFinal & "\" & CC & "-" & Yearx & "-" & Right("0" & Monthx, 2) & "-" & printN & ".pdf"
        printN = printN + 1
        
    End If

End With


End Sub
Sub Print_GTB1(sess, printN)

Set fso = CreateObject("scripting.filesystemobject")
Call SetPDFCreator

'delete all from temp file
fso.DeleteFile (FTemp & "\*.*"), True

With sess
    .findById("wnd[0]/tbar[0]/okcd").Text = "/ngr55"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/ctxtRGRWJ-JOB").Text = "gtb1"
    .findById("wnd[0]").sendVKey 8
    .findById("wnd[0]/usr/txt$CUR-YR").Text = Yearx
    .findById("wnd[0]/usr/txt$CUR-PER").Text = Monthx
    .findById("wnd[0]/usr/ctxt_COCODES-LOW").Text = CC
    .findById("wnd[0]").sendVKey 8
    .findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "000001"
    .findById("wnd[0]/shellcont/shell/shellcont[0]/shell").hierarchyHeaderWidth = 421
    .findById("wnd[0]/tbar[0]/btn[86]").press
    .findById("wnd[1]").sendVKey 0
    .findById("wnd[1]").sendVKey 0
    .findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCLX"
    .findById("wnd[1]").sendVKey 0
    On Error Resume Next
    .findById("wnd[2]").sendVKey 0
    On Error GoTo 0
    .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
    .findById("wnd[1]/tbar[0]/btn[13]").press
    .findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").Select
    .findById("wnd[1]").sendVKey 0
    .findById("wnd[1]/usr/ctxtDY_PATH").Text = FPath
    .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "gtb1.txt"
    .findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").Text = "4120"
    .findById("wnd[1]/tbar[0]/btn[11]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    On Error Resume Next
    .findById("wnd[1]/usr/btnBUTTON_YES").press
    On Error GoTo 0
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
End With

File = ""
'V4-CIO FIX: bounded, responsive wait for the printed PDF. The old loop
'rebuilt a Shell.Application on every pass with no pause, so an empty
'temp folder became a tight spin: Excel showed "Not Responding" and the
'shell object eventually dropped out with error 80010108.
File = CM_WaitForPrint(FTemp, fso, "report group GTB1")

fso.MoveFile File, FFinal & "\" & CC & "-" & Yearx & "-" & Right("0" & Monthx, 2) & "-" & printN & ".pdf"
printN = printN + 1

Set strim = New ADODB.Stream
strim.Charset = "utf-8"
strim.Open

strim.LoadFromFile (FPath & "gtb1.txt")

Do Until strim.EOS
    Dim line As String
    line = strim.ReadText(-2)
    
    n = InStr(1, line, "  ", vbBinaryCompare)
    Do Until n = 0
        line = Replace(line, "  ", " ")
        n = InStr(1, line, "  ", vbBinaryCompare)
    Loop
    arr = Split(line, " ")
    If UBound(arr, 1) > 2 Then
        If CStr(arr(LBound(arr) + 1)) = "44400200" Then
            If Right(arr(UBound(arr) - 1), 1) = "-" Then
                Am = -CDbl(Left(arr(UBound(arr) - 1), Len(arr(UBound(arr) - 1)) - 1))
            Else
                Am = CDbl(arr(UBound(arr) - 1))
            End If
        ElseIf CStr(Trim(arr(LBound(arr) + 1))) = "****" Then
            If Right(arr(UBound(arr) - 1), 1) = "-" Then
                AmT = -CDbl(Left(arr(UBound(arr) - 1), Len(arr(UBound(arr) - 1)) - 1))
            Else
                AmT = CDbl(arr(UBound(arr) - 1))
            End If
        End If
    End If
Loop

Sheets("config").Range("AA18") = Am

If Round(AmT, 2) <> 0 Then
    List = "ClosingTracker"
    Date1 = Format(Now(), "yyyy-MM-ddTHH:mm:ssZ")
    
    updates = "<Batch> <Method ID='1' Cmd='New'>" & _
                "<Field Name='UName'>" & UCase(Environ("username")) & "</Field>" & _
                "<Field Name='LogType'>" & "Closing reports - error" & "</Field>" & _
                "<Field Name='CC'>" & CC & "</Field>" & _
                "<Field Name='Periodx'>" & Month(LastDay) & "</Field>" & _
                "<Field Name='Yearx'>" & Year(LastDay) & "</Field>" & _
                "<Field Name='Success'>0</Field>" & _
                "<Field Name='Timestamp'>" & Date1 & "</Field>" & _
                "<Field Name='Comment'>" & "GTB1 not equal to zero" & "</Field>" & _
                "</Method></Batch>"

    Call spAddToList(updates, List)
    
    Sheets("Errors").Range("A10") = "Error in GTB1 - Total of all accounts not equal to zero. Amount: " & AmT & " USD."
    Sheets("Errors").Range("A10").Font.Bold = True
    
End If

Kill FPath & "gtb1.txt"

End Sub
Sub Print_ZGE1174(sess, printN)

Set fso = CreateObject("scripting.filesystemobject")
Call SetPDFCreator

'delete all from temp file
fso.DeleteFile (FTemp & "\*.*"), True

With sess
    .findById("wnd[0]").maximize
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nzge1174"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/ctxtP_BUKRS").Text = CC
    .findById("wnd[0]/usr/ctxtP_BUDAT").Text = Day(LastDay) & "." & Monthx & "." & Yearx
    .findById("wnd[0]/usr/ctxtP_WWERT").Text = "15." & Monthx & "." & Yearx
    .findById("wnd[0]/tbar[1]/btn[8]").press
    .findById("wnd[1]/usr/btnSPOPUP-OPTION1").press
    .findById("wnd[0]/tbar[0]/btn[86]").press
    .findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCLX"
    .findById("wnd[1]").sendVKey 0
    On Error Resume Next
    .findById("wnd[2]").sendVKey 0
    On Error GoTo 0
    .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
    .findById("wnd[1]/tbar[0]/btn[13]").press
    On Error Resume Next
    .findById("wnd[0]/mbar/menu[3]/menu[5]/menu[2]/menu[2]").Select
    If Err.Number <> 0 Then
        On Error GoTo 0
        .findById("wnd[0]/mbar/menu[0]/menu[5]/menu[2]/menu[2]").Select
    End If
    On Error GoTo 0
    .findById("wnd[1]").sendVKey 0
    .findById("wnd[1]/usr/ctxtDY_PATH").Text = FPath
    .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "zge1174.txt"
    .findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").Text = "4120"
    .findById("wnd[1]/tbar[0]/btn[11]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
End With

File = ""
'V4-CIO FIX: bounded, responsive wait for the printed PDF. The old loop
'rebuilt a Shell.Application on every pass with no pause, so an empty
'temp folder became a tight spin: Excel showed "Not Responding" and the
'shell object eventually dropped out with error 80010108.
File = CM_WaitForPrint(FTemp, fso, "ZGE1174")

fso.MoveFile File, FFinal & "\" & CC & "-" & Yearx & "-" & Right("0" & Monthx, 2) & "-" & printN & ".pdf"
printN = printN + 1

Set strim = New ADODB.Stream
strim.Charset = "utf-8"
strim.Open

strim.LoadFromFile (FPath & "zge1174.txt")

Do Until strim.EOS
    Dim line As String
    line = strim.ReadText(-2)
Loop

If InStr(1, line, "No Records") = 0 Then

    List = "ClosingTracker"
    Date1 = Format(Now(), "yyyy-MM-ddTHH:mm:ssZ")
    
    updates = "<Batch> <Method ID='1' Cmd='New'>" & _
                "<Field Name='UName'>" & UCase(Environ("username")) & "</Field>" & _
                "<Field Name='LogType'>" & "Closing reports - error" & "</Field>" & _
                "<Field Name='CC'>" & CC & "</Field>" & _
                "<Field Name='Periodx'>" & Month(LastDay) & "</Field>" & _
                "<Field Name='Yearx'>" & Year(LastDay) & "</Field>" & _
                "<Field Name='Success'>0</Field>" & _
                "<Field Name='Timestamp'>" & Date1 & "</Field>" & _
                "<Field Name='Comment'>" & "ZGE1174" & "</Field>" & _
                "</Method></Batch>"

    Call spAddToList(updates, List)
    
    If Sheets("Errors").Range("A10") = "" Then
        Sheets("Errors").Range("A10") = "Error in ZGE1174 - posting possible."
        Sheets("Errors").Range("A10").Font.Bold = True
    Else
        EmptRow = FindLastRow(1, 1, 2, 0, "Errors")
        Sheets("Errors").Cells(EmptRow, 1) = "Error in ZGE1174 - posting possible."
        Sheets("Errors").Range("A10").Font.Bold = True
    End If
End If

Sheets("config").Range("AA19") = "Yes"

End Sub
Sub CombinePDF(printN)
'V4-CIO: same merge behaviour, but the waits are now bounded (cannot hang Excel),
'the merge tool is checked before use, and the destination-name loop no longer
'spins forever. Functional logic and the GiosPSMC command line are unchanged.

Dim waited As Long
Const MAX_WAIT As Long = 300     'safety cap in seconds for each wait loop

Set shellX = CreateObject("WScript.Shell")
Set fso = CreateObject("scripting.filesystemobject")

Call CreateVariants
Call CreatePaths

paramsource = ""
paramOutput = Fmerged & "\" & Yearx & Monthx & CC & ".pdf"
       
For i = 1 To printN
    paramsource = paramsource & Chr(34) & FFinal & "\" & CC & "-" & Yearx & "-" & Right("0" & Monthx, 2) & "-" & i & ".pdf" & Chr(34) & " "
Next i
          
paramsource = Left(paramsource, Len(paramsource) - 1)

'V4-CIO FIX: fail fast with a clear message if the merge tool is missing
If Not fso.FileExists(Fmerger & "\GiosPSMC.exe") Then
    MsgBox "PDF merger not found:" & vbCrLf & Fmerger & "\GiosPSMC.exe" & vbCrLf & vbCrLf & _
           "Run 'Preflight Check' or restore it from the network share, then try again.", _
           vbCritical, "Closing Manager"
    Exit Sub
End If

shellX.Run "%COMSPEC% /c " & Fmerger & "\GiosPSMC.exe" & " " & paramsource & " output " & paramOutput

'V4-CIO FIX: bounded wait for the merged file to appear (was an unbounded Do..Loop)
waited = 0
Do
    If fso.FileExists(paramOutput) Then
        File = fso.GetFile(paramOutput)
        Exit Do
    Else
        Application.Wait (Now + TimeValue("0:00:01"))
        waited = waited + 1
    End If
    If waited >= MAX_WAIT Then
        MsgBox "PDF merge timed out - the merged file was not created:" & vbCrLf & paramOutput, _
               vbCritical, "Closing Manager"
        Exit Sub
    End If
Loop

'V4-CIO FIX: bounded wait for the file to stop growing (was an unbounded Do..Loop)
Do
    size1 = fso.GetFile(File).Size
    Application.Wait (Now + TimeValue("0:00:01"))
    size2 = fso.GetFile(File).Size
    waited = waited + 1

    If size1 = size2 And size1 <> 0 And size2 <> 0 Then
        Exit Do
    End If
    If waited >= MAX_WAIT * 2 Then Exit Do
Loop

fso.DeleteFile (FFinal & "\*.*"), True

Call EnsureFolders   'V4-CIO: make sure the year\month tree exists before the move

fN = 1
FPathReport = CM_REPORT_ROOT & Yearx & "\" & Right("0" & Monthx, 2) & "\"
FName = CC & " Closure " & Right("0" & Monthx, 2) & " " & Yearx & ".pdf"
FName1 = CC & " Closure " & Right("0" & Monthx, 2) & " " & Yearx
Do Until fso.FileExists(FPathReport & FName) = False
    FName = FName1 & "-" & fN & ".pdf"
    fN = fN + 1                          'V4-CIO FIX: was missing -> infinite loop when a same-named report already existed
Loop

'V4-CIO FIX: report, don't crash, if the final move fails
On Error Resume Next
fso.MoveFile File, FPathReport & FName
If Err.Number <> 0 Then
    MsgBox "Could not move the final report to:" & vbCrLf & FPathReport & FName & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Closing Manager"
End If
On Error GoTo 0

End Sub
