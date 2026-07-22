Option Base 1
Sub Post_ZGE132(sess)

With sess
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nzge132"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/chkP_CLENT").Selected = True
    .findById("wnd[0]/usr/ctxtP_CCODE").Text = CC
    .findById("wnd[0]/usr/txtP_MONTH").Text = Monthx
    .findById("wnd[0]/usr/txtP_YEAR").Text = Yearx
    .findById("wnd[0]").sendVKey 2
    .findById("wnd[0]/usr/chkP_BCENT").Selected = True
    .findById("wnd[0]/tbar[1]/btn[8]").press
    .findById("wnd[0]/tbar[1]/btn[5]").press
    On Error Resume Next
    .findById("wnd[1]/usr/btnBUTTON_1").press
    On Error GoTo 0
    .findById("wnd[1]/tbar[0]/btn[0]").press
    
    Application.Wait (Now + TimeValue("0:00:10"))
    
    .findById("wnd[1]/tbar[0]/btn[0]").press
    
    Application.Wait (Now + TimeValue("0:00:10"))
    
    .findById("wnd[1]/tbar[0]/btn[0]").press
    
    Application.Wait (Now + TimeValue("0:00:10"))
    
    .findById("wnd[0]/tbar[0]/btn[3]").press
    On Error Resume Next
    .findById("wnd[1]/tbar[0]/btn[0]").press
    If Err.Number = 0 Then
        On Error GoTo 0
        On Error Resume Next
        .findById("wnd[1]/usr/btnBUTTON_1").press
        On Error GoTo 0
        On Error Resume Next
        .findById("wnd[1]/tbar[0]/btn[0]").press
        If Err.Number = 0 Then
            On Error GoTo 0
            .findById("wnd[0]/tbar[0]/btn[3]").press
            .findById("wnd[0]/tbar[0]/btn[3]").press
            .findById("wnd[0]/tbar[0]/btn[3]").press
        Else
            On Error GoTo 0
            Sheets("config").Range("AA12") = "No posting"
        End If
    Else
        On Error GoTo 0
        Sheets("config").Range("AA12") = "No posting"
    End If
    On Error GoTo 0
End With

End Sub
Sub Post_ZGE132GC(sess)

With sess
    .findById("wnd[0]").maximize
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nzge132"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/tbar[1]/btn[25]").press
    .findById("wnd[0]/usr/chkP_BLC").Selected = True
    .findById("wnd[0]/usr/ctxtP_BUKRS").Text = CC
    .findById("wnd[0]/usr/txtP_MONTH").Text = Monthx
    .findById("wnd[0]/usr/txtP_YEAR").Text = Yearx
    .findById("wnd[0]/usr/chkP_BLC").SetFocus
    .findById("wnd[0]/tbar[1]/btn[8]").press
    .findById("wnd[0]/tbar[1]/btn[36]").press
    .findById("wnd[1]/tbar[0]/btn[0]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
End With

End Sub
Function Check_ZGE132AP(sess)

Set shellX = CreateObject("WScript.Shell")
Set fso = CreateObject("scripting.filesystemobject")

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
    .findById("wnd[1]/tbar[0]/btn[0]").press
    .findById("wnd[1]/usr/ctxtDY_PATH").Text = FPath
    .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "zge132.txt"
    .findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
    .findById("wnd[1]/tbar[0]/btn[11]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    
    File = ""
    Do Until File <> ""
        Set objShell = CreateObject("Shell.Application")
        Set objFolder = objShell.Namespace(FTemp & "\")
        Set colItems = objFolder.Items
        For Each objitem In colItems
            Do
                If fso.FileExists(FTemp & "\" & objitem) Then
                    File = fso.GetFile(FTemp & "\" & objitem)
                    Exit Do
                Else
                    Application.Wait (Now + TimeValue("0:00:01"))
                End If
            Loop
            Do
                size1 = fso.GetFile(File).Size
                Application.Wait (Now + TimeValue("0:00:01"))
                size2 = fso.GetFile(File).Size
                If size1 = size2 And size1 <> 0 And size2 <> 0 Then
                    Exit Do
                End If
            Loop
        Next
    Loop
    
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
    .findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
    .findById("wnd[1]").sendVKey 0
    .findById("wnd[1]/usr/ctxtDY_PATH").Text = FPath
    .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ZGE132G.txt"
    .findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 11
    .findById("wnd[1]/tbar[0]/btn[11]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    
    File = ""
    Do Until File <> ""
        Set objShell = CreateObject("Shell.Application")
        Set objFolder = objShell.Namespace(FTemp & "\")
        Set colItems = objFolder.Items
        For Each objitem In colItems
            Do
                If fso.FileExists(FTemp & "\" & objitem) Then
                    File = fso.GetFile(FTemp & "\" & objitem)
                    Exit Do
                Else
                    Application.Wait (Now + TimeValue("0:00:01"))
                End If
            Loop
            Do
                size1 = fso.GetFile(File).Size
                Application.Wait (Now + TimeValue("0:00:01"))
                size2 = fso.GetFile(File).Size
                If size1 = size2 And size1 <> 0 And size2 <> 0 Then
                    Exit Do
                End If
            Loop
        Next
    Loop
    
    fso.MoveFile File, FFinal & "\" & CC & "-" & Yearx & "-" & Right("0" & Monthx, 2) & "-" & printN & ".pdf"
    printN = printN + 1
    
End With

'----------------------------------------------------------------
'local currency
'----------------------------------------------------------------
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
                Am = Am + -Round(Left(arr(UBound(arr)), Len(arr(UBound(arr))) - 1), 2)
            Else
                Am = Am + Round(arr(UBound(arr)), 2)
            End If
        End If
    End If
Loop

If Round(AmL, 2) <> 0 Then
    Check_ZGE132AP = False
    BL = True
Else
    Sheets("config").Range("AA8") = AmL
End If
If Round(Am, 2) <> 0 Then
    Check_ZGE132AP = False
    CPCL = True
End If

'----------------------------------------------------------------
'group currency
'----------------------------------------------------------------
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
                Am = Am + -Round(Left(arr(UBound(arr) - 2), Len(arr(UBound(arr) - 2)) - 1), 2)
            Else
                Am = Am + Round(arr(UBound(arr) - 2), 2)
            End If
            If Round(Am, 2) <> 0 Then
                Check_ZGE132AP = False
                CPCG = True
            End If
            Sheets("ZGE132").Cells(EmptRow, 2) = Am
        End If
    End If
Loop

If Round(AmG, 2) <> 0 Then
    Check_ZGE132AP = False
    BG = True
Else
    Sheets("config").Range("AA14") = AmG
End If
'If Round(Am, 2) <> 0 Then
'    Check_ZGE132AP = False
'    CPCG = True
'End If

If BL = False And CPCL = False And BG = False And CPCG = False Then
    Check_ZGE132AP = True
End If

End Function
Sub GetSM35(sess, sessN)

DocL = ""
DocG = ""

With sess
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nsm35"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/subD1000_HEADER:SAPMSBDC_CC:1005/txtD0100-CREATOR").Text = UCase(SAPID)
    .findById("wnd[0]").sendVKey 0
    j = 0
    LC = False
    GC = False
    Do
        VR = .findById("wnd[0]/usr/tabsD1000_TABSTRIP/tabpALLE/ssubD1000_SUBSCREEN:SAPMSBDC_CC:1010/tblSAPMSBDC_CCTC_APQI/").visiblerowcount
        For i = 0 To VR
            TXT = .findById("wnd[0]/usr/tabsD1000_TABSTRIP/tabpALLE/ssubD1000_SUBSCREEN:SAPMSBDC_CC:1010/tblSAPMSBDC_CCTC_APQI/txtITAB_APQI-GROUPID[0," & i & "]").Text
            If InStr(1, TXT, CC) > 0 Then
                j = j + 1
                If InStr(1, TXT, "P&B_") > 0 Then
                    LC = True
                    .findById("wnd[0]/usr/subD1000_FOOT:SAPMSBDC_CC:1015/btnPB_DESELECT_ALL").press
                    .findById("wnd[0]/usr/tabsD1000_TABSTRIP/tabpALLE/ssubD1000_SUBSCREEN:SAPMSBDC_CC:1010/tblSAPMSBDC_CCTC_APQI").getAbsoluteRow(i).Selected = True
                    .findById("wnd[0]/tbar[1]/btn[2]").press
                    .findById("wnd[0]/usr/tabsTAB_DYNPRO/tabpTAB_PROTO").Select
                    VR1 = .findById("wnd[0]/usr/tabsTAB_DYNPRO/tabpTAB_PROTO/ssubSCR_DYNPRO:RSBDC_ANALYSE:0400/tblRSBDC_ANALYSETC_PROTOCOL/").visiblerowcount
                    FR = .findById("wnd[0]/usr/tabsTAB_DYNPRO/tabpTAB_PROTO/ssubSCR_DYNPRO:RSBDC_ANALYSE:0400/tblRSBDC_ANALYSETC_PROTOCOL/txtBDC_PROTOCOL-LONGTEXT[1,0]").Text
                    Do Until FR = FRCh
                        For l = 0 To VR1 - 1
                            
                            TXT2 = .findById("wnd[0]/usr/tabsTAB_DYNPRO/tabpTAB_PROTO/ssubSCR_DYNPRO:RSBDC_ANALYSE:0400/tblRSBDC_ANALYSETC_PROTOCOL/txtBDC_PROTOCOL-LONGTEXT[1," & l & "]").Text
                            If InStr(1, TXT2, "was posted in") > 0 Then
                                TXT2 = Replace(TXT2, "Document ", "")
                                n = InStr(1, TXT2, " was")
                                TXT2 = Left(TXT2, n - 1)
                                If DocL = "" Then
                                    DocL = TXT2
                                Else
                                    DocL = DocL & ";" & TXT2
                                End If
                            End If
                        Next l
                        FR = .findById("wnd[0]/usr/tabsTAB_DYNPRO/tabpTAB_PROTO/ssubSCR_DYNPRO:RSBDC_ANALYSE:0400/tblRSBDC_ANALYSETC_PROTOCOL/txtBDC_PROTOCOL-LONGTEXT[1,0]").Text
                        .findById("wnd[0]").sendVKey 82
                        FRCh = .findById("wnd[0]/usr/tabsTAB_DYNPRO/tabpTAB_PROTO/ssubSCR_DYNPRO:RSBDC_ANALYSE:0400/tblRSBDC_ANALYSETC_PROTOCOL/txtBDC_PROTOCOL-LONGTEXT[1,0]").Text
                    Loop
                    .findById("wnd[0]/tbar[0]/btn[3]").press
                End If
                If InStr(1, TXT, "P&BL_") > 0 Then
                    GC = True
                    .findById("wnd[0]/usr/subD1000_FOOT:SAPMSBDC_CC:1015/btnPB_DESELECT_ALL").press
                    .findById("wnd[0]/usr/tabsD1000_TABSTRIP/tabpALLE/ssubD1000_SUBSCREEN:SAPMSBDC_CC:1010/tblSAPMSBDC_CCTC_APQI").getAbsoluteRow(i).Selected = True
                    .findById("wnd[0]/tbar[1]/btn[2]").press
                    .findById("wnd[0]/usr/tabsTAB_DYNPRO/tabpTAB_PROTO").Select
                    VR1 = .findById("wnd[0]/usr/tabsTAB_DYNPRO/tabpTAB_PROTO/ssubSCR_DYNPRO:RSBDC_ANALYSE:0400/tblRSBDC_ANALYSETC_PROTOCOL/").visiblerowcount
                    FR = .findById("wnd[0]/usr/tabsTAB_DYNPRO/tabpTAB_PROTO/ssubSCR_DYNPRO:RSBDC_ANALYSE:0400/tblRSBDC_ANALYSETC_PROTOCOL/txtBDC_PROTOCOL-LONGTEXT[1,0]").Text
                    Do Until FR = FRCh
                        For l = 0 To VR1 - 1
                            
                            TXT2 = .findById("wnd[0]/usr/tabsTAB_DYNPRO/tabpTAB_PROTO/ssubSCR_DYNPRO:RSBDC_ANALYSE:0400/tblRSBDC_ANALYSETC_PROTOCOL/txtBDC_PROTOCOL-LONGTEXT[1," & l & "]").Text
                            If InStr(1, TXT2, "was posted in") > 0 Then
                                TXT2 = Replace(TXT2, "Document ", "")
                                n = InStr(1, TXT2, " was")
                                TXT2 = Left(TXT2, n - 1)
                                If DocG = "" Then
                                    DocG = TXT2
                                ElseIf Left(DocG, 2) = "XY" Then
                                    DocG = DocG
                                Else
                                    DocG = DocG & ";" & TXT2
                                End If
                            End If
                            If TXT2 = "CO Validation Error:  ZGXMIT Status Set to X or Y." Then
                                DocG = "XY-" & printN
                            End If
                        Next l
                        FR = .findById("wnd[0]/usr/tabsTAB_DYNPRO/tabpTAB_PROTO/ssubSCR_DYNPRO:RSBDC_ANALYSE:0400/tblRSBDC_ANALYSETC_PROTOCOL/txtBDC_PROTOCOL-LONGTEXT[1,0]").Text
                        .findById("wnd[0]").sendVKey 82
                        FRCh = .findById("wnd[0]/usr/tabsTAB_DYNPRO/tabpTAB_PROTO/ssubSCR_DYNPRO:RSBDC_ANALYSE:0400/tblRSBDC_ANALYSETC_PROTOCOL/txtBDC_PROTOCOL-LONGTEXT[1,0]").Text
                    Loop
                    .findById("wnd[0]/tbar[0]/btn[3]").press
                End If
            End If
            If j = sessN Then Exit For
        Next i
        
    Loop Until j = sessN
    .findById("wnd[0]/tbar[0]/btn[3]").press
End With

If DocL <> "" Then
    Sheets("config").Range("AA6") = DocL
    arr = Split(DocL, ";")
    Set alloutput = New ADODB.Stream
        
    alloutput.Charset = "utf-8"
    alloutput.Open
    
    For i = 0 To UBound(arr)
        alloutput.writetext (arr(i) & vbCrLf)
    Next i
                                        
    alloutput.SaveToFile FPath & "\DocL.csv", 2
    alloutput.Close
End If
If DocG <> "" And Left(DocG, 2) <> "XY" Then
    Sheets("config").Range("AA12") = DocG
    arr = Split(DocG, ";")
    Set alloutput = New ADODB.Stream
        
    alloutput.Charset = "utf-8"
    alloutput.Open
    
    For i = 0 To UBound(arr)
        alloutput.writetext (arr(i) & vbCrLf)
    Next i
                                        
    alloutput.SaveToFile FPath & "\DocG.csv", 2
    alloutput.Close
ElseIf Left(DocG, 2) = "XY" Then
    Sheets("config").Range("AA12") = DocG
End If

End Sub
Sub Post_ZGLGWUL(sess)

ErrTxt = ""
Set shellX = CreateObject("WScript.Shell")
Set fso = CreateObject("scripting.filesystemobject")

With sess
    .findById("wnd[0]").maximize
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nzglgwul"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/ctxtS_RBUKRS-LOW").Text = CC
    .findById("wnd[0]/usr/txtP_MONAT").Text = Monthx
    .findById("wnd[0]/usr/txtP_RYEAR").Text = Yearx
    .findById("wnd[0]/usr/chkP_POST").Selected = True
    .findById("wnd[0]/usr/chkP_SCURR").Selected = True
    .findById("wnd[0]/tbar[1]/btn[8]").press
    On Error Resume Next
    .findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
    If Err.Number <> 0 Then
        ErrTxt = .findById("wnd[1]/usr/txtMESSTXT1").Text
    Else
        .findById("wnd[1]").sendVKey 0
        .findById("wnd[1]/usr/ctxtDY_PATH").Text = FPath
        .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "zglgwul.txt"
        .findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").Text = "4120"
        .findById("wnd[1]/tbar[0]/btn[11]").press
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
        Do Until File <> ""
            Set objShell = CreateObject("Shell.Application")
            Set objFolder = objShell.Namespace(FTemp & "\")
            Set colItems = objFolder.Items
            For Each objitem In colItems
                Do
                    If fso.FileExists(FTemp & "\" & objitem) Then
                        File = fso.GetFile(FTemp & "\" & objitem)
                        Exit Do
                    Else
                        Application.Wait (Now + TimeValue("0:00:01"))
                    End If
                Loop
                Do
                    size1 = fso.GetFile(File).Size
                    Application.Wait (Now + TimeValue("0:00:01"))
                    size2 = fso.GetFile(File).Size
                    If size1 = size2 And size1 <> 0 And size2 <> 0 Then
                        Exit Do
                    End If
                Loop
            Next
        Loop
        
        fso.MoveFile File, FFinal & "\" & CC & "-" & Yearx & "-" & Right("0" & Monthx, 2) & "-" & printN & ".pdf"
        printN = printN + 1
    End If
End With

If ErrTxt = "" Then

    Set strim = New ADODB.Stream
    strim.Charset = "utf-8"
    strim.Open
    
    strim.LoadFromFile (FPath & "zglgwul.txt")
    
    Do Until strim.EOS
        line = strim.ReadText(-2)
        If Left(Trim(line), 1) = "|" Then
            n = InStr(1, line, "  ", vbBinaryCompare)
            Do Until n = 0
                line = Replace(line, "  ", " ")
                n = InStr(1, line, "  ", vbBinaryCompare)
            Loop
            arr = Split(line, " ")
            If CStr(arr(LBound(arr) + 1)) = "44400200" Then
                If Right(arr(UBound(arr)), 1) = "-" Then
                    Am = -Round(Left(arr(UBound(arr)), Len(arr(UBound(arr) - 1))), 2)
                ElseIf arr(UBound(arr)) = "|" Then
                    Am = 0
                Else
                    Am = Round(arr(UBound(arr)), 2)
                End If
                Exit Do
            End If
        End If
    Loop
    Sheets("config").Range("AA16") = Am
Else
    Sheets("config").Range("AA16") = ErrTxt
End If

End Sub
Sub Post_ZGE132AG(sess)

Set shellX = CreateObject("WScript.Shell")
Set fso = CreateObject("scripting.filesystemobject")

With sess
    .findById("wnd[0]/tbar[0]/btn[3]").press
    On Error Resume Next
    .findById("wnd[1]/usr/btnBUTTON_1").press
    .findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo 0
    .findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
    .findById("wnd[1]").sendVKey 0
    .findById("wnd[1]/usr/ctxtDY_PATH").Text = FPath
    .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "zge132gwul.txt"
    .findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
    .findById("wnd[1]/tbar[0]/btn[11]").press
    .findById("wnd[0]/tbar[0]/btn[86]").press
    .findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "LOCLX"
    .findById("wnd[1]").sendVKey 0
    On Error Resume Next
    .findById("wnd[2]").sendVKey 0
    On Error GoTo 0
    .findById("wnd[1]/usr/cmbPRIPAR_EXT-OSPRINTER").Key = "PDFCreator"
    .findById("wnd[1]/tbar[0]/btn[13]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    .findById("wnd[0]/tbar[0]/btn[3]").press
    
    File = ""
    Do Until File <> ""
        Set objShell = CreateObject("Shell.Application")
        Set objFolder = objShell.Namespace(FTemp & "\")
        Set colItems = objFolder.Items
        For Each objitem In colItems
            Do
                If fso.FileExists(FTemp & "\" & objitem) Then
                    File = fso.GetFile(FTemp & "\" & objitem)
                    Exit Do
                Else
                    Application.Wait (Now + TimeValue("0:00:01"))
                End If
            Loop
            Do
                size1 = fso.GetFile(File).Size
                Application.Wait (Now + TimeValue("0:00:01"))
                size2 = fso.GetFile(File).Size
                If size1 = size2 And size1 <> 0 And size2 <> 0 Then
                    Exit Do
                End If
            Loop
        Next
    Loop
    
    fso.MoveFile File, FFinal & "\" & CC & "-" & Yearx & "-" & Right("0" & Monthx, 2) & "-" & printN & ".pdf"
    printN = printN + 1
    
End With

Set strim = New ADODB.Stream
strim.Charset = "utf-8"
strim.Open

Call CreateArray(FPath & "zge132gwul.txt")
strim.LoadFromFile (FPath & "zge132gwul.txt")

CheckPosting = False
Do Until strim.EOS Or CheckPosting = True
    Dim line As String, tablica() As String
    line = strim.ReadText(-2)

    If Left(Trim(line), 1) = "|" And Trim(getLineData(line, FirstColumn, 1)) <> FirstColumn Then
        If Left(line, 15) = "| List does not" Then
        Else
            If Right(Trim(getLineData(line, "Amount", 1)), 1) = "-" Then
                CheckPosting = True
            ElseIf Round(Trim(getLineData(line, "Amount", 1)), 2) <> 0 Then
                CheckPosting = True
            End If
        End If
    End If
Loop

If CheckPosting = True Then
    Call GetSM35AG(sess)
    
    If Left(Sheets("config").Range("AA17"), 2) = "XY" Then
    
        strErr = ""
        List = "ClosingTracker"
        Date1 = Format(Now(), "yyyy-MM-ddTHH:mm:ssZ")
        
        strErr = "XY"
        
        updates = "<Batch> <Method ID='1' Cmd='New'>" & _
                "<Field Name='UName'>" & UCase(Environ("username")) & "</Field>" & _
                "<Field Name='LogType'>" & "Closing reports - error" & "</Field>" & _
                "<Field Name='CC'>" & CC & "</Field>" & _
                "<Field Name='Periodx'>" & Month(LastDay) & "</Field>" & _
                "<Field Name='Yearx'>" & Year(LastDay) & "</Field>" & _
                "<Field Name='Success'>0</Field>" & _
                "<Field Name='Timestamp'>" & Date1 & "</Field>" & _
                "<Field Name='Comment'>" & "ZGE132: " & strErr & "</Field>" & _
                "</Method></Batch>"

        Call spAddToList(updates, List)
    
        'info
        strError = ""
        MsgBox "XY error. Send e-mail to GFIM team and run macro again."
    
        Exit Sub
    End If
Else
    Sheets("config").Range("AA17") = "No posting."
End If

End Sub
Sub GetSM35AG(sess)

DocG = ""
    
Set alloutput = New ADODB.Stream
    
alloutput.Charset = "utf-8"
alloutput.Open
DocG = ""
With sess
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nsm35"
    .findById("wnd[0]").sendVKey 0
    .findById("wnd[0]/usr/subD1000_HEADER:SAPMSBDC_CC:1005/txtD0100-CREATOR").Text = UCase(SAPID)
    .findById("wnd[0]").sendVKey 0
    j = 0
    VR = .findById("wnd[0]/usr/tabsD1000_TABSTRIP/tabpALLE/ssubD1000_SUBSCREEN:SAPMSBDC_CC:1010/tblSAPMSBDC_CCTC_APQI/").visiblerowcount
    For i = 0 To VR
        TXT = .findById("wnd[0]/usr/tabsD1000_TABSTRIP/tabpALLE/ssubD1000_SUBSCREEN:SAPMSBDC_CC:1010/tblSAPMSBDC_CCTC_APQI/txtITAB_APQI-GROUPID[0," & i & "]").Text
        If InStr(1, TXT, CC) > 0 Then
            If InStr(1, TXT, "P&BL_") > 0 Then
                .findById("wnd[0]/usr/subD1000_FOOT:SAPMSBDC_CC:1015/btnPB_DESELECT_ALL").press
                .findById("wnd[0]/usr/tabsD1000_TABSTRIP/tabpALLE/ssubD1000_SUBSCREEN:SAPMSBDC_CC:1010/tblSAPMSBDC_CCTC_APQI").getAbsoluteRow(i).Selected = True
                .findById("wnd[0]/tbar[1]/btn[2]").press
                .findById("wnd[0]/usr/tabsTAB_DYNPRO/tabpTAB_PROTO").Select
                VR1 = .findById("wnd[0]/usr/tabsTAB_DYNPRO/tabpTAB_PROTO/ssubSCR_DYNPRO:RSBDC_ANALYSE:0400/tblRSBDC_ANALYSETC_PROTOCOL/").visiblerowcount
                FR = .findById("wnd[0]/usr/tabsTAB_DYNPRO/tabpTAB_PROTO/ssubSCR_DYNPRO:RSBDC_ANALYSE:0400/tblRSBDC_ANALYSETC_PROTOCOL/txtBDC_PROTOCOL-LONGTEXT[1,0]").Text
                Do Until FR = FRCh
                    For l = 0 To VR1 - 1
                        TXT2 = .findById("wnd[0]/usr/tabsTAB_DYNPRO/tabpTAB_PROTO/ssubSCR_DYNPRO:RSBDC_ANALYSE:0400/tblRSBDC_ANALYSETC_PROTOCOL/txtBDC_PROTOCOL-LONGTEXT[1," & l & "]").Text
                        If InStr(1, TXT2, "was posted in") > 0 Then
                            TXT2 = Replace(TXT2, "Document ", "")
                            n = InStr(1, TXT2, " was")
                            TXT2 = Left(TXT2, n - 1)
                            If DocG = "" Then
                                DocG = TXT2
                            Else
                                DocG = DocG & ";" & TXT2
                            End If
                            alloutput.writetext (TXT2 & vbCrLf)
                        End If
                        If TXT2 = "CO Validation Error:  ZGXMIT Status Set to X or Y." Then
                            DocG = "XY-" & printN
                        End If
                    Next l
                    FR = .findById("wnd[0]/usr/tabsTAB_DYNPRO/tabpTAB_PROTO/ssubSCR_DYNPRO:RSBDC_ANALYSE:0400/tblRSBDC_ANALYSETC_PROTOCOL/txtBDC_PROTOCOL-LONGTEXT[1,0]").Text
                    .findById("wnd[0]").sendVKey 82
                    FRCh = .findById("wnd[0]/usr/tabsTAB_DYNPRO/tabpTAB_PROTO/ssubSCR_DYNPRO:RSBDC_ANALYSE:0400/tblRSBDC_ANALYSETC_PROTOCOL/txtBDC_PROTOCOL-LONGTEXT[1,0]").Text
                Loop
                .findById("wnd[0]/tbar[0]/btn[3]").press
                Exit For
            End If
        End If
    Next i
End With

Sheets("config").Range("AA17") = DocG
    
alloutput.SaveToFile FPath & "\ZGE132AG.csv", 2
alloutput.Close

Set fso = CreateObject("Scripting.FileSystemObject")

If Left(DocG, 2) <> "XY" Then
    With sess
        .findById("wnd[0]/tbar[0]/btn[3]").press
        
        .findById("wnd[0]").maximize
        .findById("wnd[0]/tbar[0]/okcd").Text = "/nzgr215"
        .findById("wnd[0]").sendVKey 0
        .findById("wnd[0]/usr/btn%_SBELNR_%_APP_%-VALU_PUSH").press
        .findById("wnd[1]/tbar[0]/btn[23]").press
        .findById("wnd[2]/usr/ctxtDY_PATH").Text = FPath
        .findById("wnd[2]/usr/ctxtDY_FILENAME").Text = "ZGE132AG.csv"
        .findById("wnd[2]/tbar[0]/btn[0]").press
        .findById("wnd[1]/tbar[0]/btn[8]").press
        .findById("wnd[0]/usr/ctxtSBUKRS").Text = CC
        .findById("wnd[0]/usr/txtSYEAR").Text = Yearx
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
        Do Until File <> ""
            Set objShell = CreateObject("Shell.Application")
            Set objFolder = objShell.Namespace(FTemp & "\")
            Set colItems = objFolder.Items
            For Each objitem In colItems
                Do
                    If fso.FileExists(FTemp & "\" & objitem) Then
                        File = fso.GetFile(FTemp & "\" & objitem)
                        Exit Do
                    Else
                        Application.Wait (Now + TimeValue("0:00:01"))
                    End If
                Loop
                Do
                    size1 = fso.GetFile(File).Size
                    Application.Wait (Now + TimeValue("0:00:01"))
                    size2 = fso.GetFile(File).Size
                    If size1 = size2 And size1 <> 0 And size2 <> 0 Then
                        Exit Do
                    End If
                Loop
            Next
        Loop
        
        fso.MoveFile File, FFinal & "\" & CC & "-" & Yearx & "-" & Right("0" & Monthx, 2) & "-" & printN & ".pdf"
        printN = printN + 1
        
    End With
End If

End Sub
