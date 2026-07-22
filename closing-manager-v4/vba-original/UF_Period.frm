
Private Sub BT_CANCEL_Click()

Unload Me

End Sub

Private Sub BT_OK_Click()

Date1 = DateAdd("m", -1, Date)
If Len(Month(Date1)) = 1 Then Monthx = "0" & Month(Date1) Else Monthx = Month(Date1)

If Len(TB_period) <> 2 Or IsNumeric(TB_period) = False Or Len(TB_year) <> 4 Or IsNumeric(TB_year) = False Then
    strError = "You need to provide correct period and year, e.g." & vbNewLine & vbNewLine & "Period: " & Monthx & vbNewLine & "Year: " & Year(Date1)
    With UF_Error
        .Lbl_Error.Caption = strError
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
    Exit Sub
End If

Sheets("config").Range("Y6") = TB_period & "/" & TB_year
Unload Me

End Sub
Private Sub UserForm_Initialize()

Date1 = DateAdd("m", -1, Date)
If Len(Month(Date1)) = 1 Then Monthx = "0" & Month(Date1) Else Monthx = Month(Date1)
TB_period = Monthx
TB_year = Year(Date1)

End Sub
