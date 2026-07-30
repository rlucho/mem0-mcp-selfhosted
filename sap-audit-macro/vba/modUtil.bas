Attribute VB_Name = "modUtil"
'=======================================================================
' modUtil -- date, amount, string and filesystem helpers.
'
' The date and amount formatting here is the part most likely to bite.
' SAP renders dates and numbers in the logged-on user's own format, and
' the sample workbook was written by someone whose Excel uses
' dd/mm/yyyy with '.' thousands and ',' decimals. Both are handled
' explicitly rather than left to CStr.
'=======================================================================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub SleepApi Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
#Else
    Private Declare Sub SleepApi Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
#End If

Public Sub SleepSeconds(ByVal seconds As Double)
    SleepApi CLng(seconds * 1000)
End Sub

'-----------------------------------------------------------------------
' Dates
'-----------------------------------------------------------------------
' SAP input fields accept the format of the logged-on user, so read that
' from the session rather than assuming. Falls back to dd.mm.yyyy, the
' most common setting on European systems.
Public Function SapDate(ByVal value As Date) As String
    Dim pattern As String

    pattern = SapDatePattern()

    Select Case pattern
        Case "DMY":  SapDate = Format$(value, "dd.mm.yyyy")   ' 31.12.2025
        Case "DMY/": SapDate = Format$(value, "dd/mm/yyyy")   ' 31/12/2025
        Case "MDY":  SapDate = Format$(value, "mm/dd/yyyy")   ' 12/31/2025
        Case "YMD":  SapDate = Format$(value, "yyyy-mm-dd")   ' 2025-12-31
        Case Else
            Err.Raise vbObjectError + 560, "modUtil.SapDate", _
                      "'SAP date format' on the Control sheet reads '" & pattern & _
                      "'. It must be one of DMY, DMY/, MDY or YMD -- see SU3 > Defaults."
    End Select
End Function

' Probe the session's own date format. GuiSession does not expose it
' directly on every release, so this reads it from a known date field
' where possible and otherwise returns the default.
Private Function SapDatePattern() As String
    Dim override As String

    ' An operator who knows their system's setting can pin it on the
    ' Control sheet rather than relying on the probe.
    On Error Resume Next
    override = UCase$(Trim$(modConfig.Setting("SAP date format")))
    On Error GoTo 0

    If Len(override) > 0 Then
        SapDatePattern = override
    Else
        SapDatePattern = "DMY"
    End If
End Function

Public Function MonthStart(ByVal value As Date) As Date
    MonthStart = DateSerial(Year(value), Month(value), 1)
End Function

Public Function MonthEnd(ByVal value As Date) As Date
    MonthEnd = DateSerial(Year(value), Month(value) + 1, 0)
End Function

' 'Sep 25' -- matches the auditor's per-month tab names.
Public Function MonthTabName(ByVal value As Date) As String
    MonthTabName = Format$(value, "mmm yy")
End Function

'-----------------------------------------------------------------------
' Amounts
'-----------------------------------------------------------------------
' Parse an amount as SAP rendered it. Handles '8.072.447,42' and
' '8,072,447.42', and a trailing minus, which SAP uses for credits.
Public Function ParseSapAmount(ByVal text As String) As Double
    Dim cleaned As String
    Dim isNegative As Boolean
    Dim lastDot As Long, lastComma As Long

    cleaned = Trim$(text)
    If Len(cleaned) = 0 Then Exit Function

    If Right$(cleaned, 1) = "-" Then
        isNegative = True
        cleaned = Trim$(Left$(cleaned, Len(cleaned) - 1))
    ElseIf Left$(cleaned, 1) = "-" Then
        isNegative = True
        cleaned = Trim$(Mid$(cleaned, 2))
    End If

    cleaned = Replace(cleaned, " ", vbNullString)
    cleaned = Replace(cleaned, Chr$(160), vbNullString)   ' non-breaking space

    ' Whichever separator appears last is the decimal separator.
    lastDot = InStrRev(cleaned, ".")
    lastComma = InStrRev(cleaned, ",")

    If lastComma > lastDot Then
        cleaned = Replace(cleaned, ".", vbNullString)
        cleaned = Replace(cleaned, ",", ".")
    Else
        cleaned = Replace(cleaned, ",", vbNullString)
    End If

    If Not IsNumeric(cleaned) Then Exit Function

    ParseSapAmount = CDbl(Val(cleaned))                   ' Val is locale-independent
    If isNegative Then ParseSapAmount = -ParseSapAmount
End Function

' The sample list holds unsigned amounts even where the statement shows a
' debit, so compare on magnitude.
Public Function AmountsMatch(ByVal a As Double, ByVal b As Double, _
                             ByVal tolerance As Double) As Boolean
    AmountsMatch = (Abs(Abs(a) - Abs(b)) <= tolerance)
End Function

'-----------------------------------------------------------------------
' Strings and filenames
'-----------------------------------------------------------------------
Public Function Squeeze(ByVal text As String) As String
    Dim result As String
    result = Trim$(Replace(Replace(text, vbTab, " "), vbCrLf, " "))
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    Squeeze = result
End Function

Public Function SafeFileName(ByVal text As String) As String
    Dim illegal As Variant
    Dim item As Variant
    Dim result As String

    illegal = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)
    result = Squeeze(text)

    For Each item In illegal
        result = Replace(result, CStr(item), "_")
    Next item

    ' Windows will not accept a name ending in a dot or space.
    Do While Len(result) > 0 And (Right$(result, 1) = "." Or Right$(result, 1) = " ")
        result = Left$(result, Len(result) - 1)
    Loop

    If Len(result) > 120 Then result = Left$(result, 120)
    If Len(result) = 0 Then result = "unnamed"

    SafeFileName = result
End Function

' Create a folder and any missing parents. Returns the path.
Public Function EnsureFolder(ByVal path As String) As String
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

Public Function JoinPath(ByVal folder As String, ByVal leaf As String) As String
    If Right$(folder, 1) = "\" Then
        JoinPath = folder & leaf
    Else
        JoinPath = folder & "\" & leaf
    End If
End Function

Public Function FileExists(ByVal path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(path)
End Function

Public Function FileSizeBytes(ByVal path As String) As Double
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then FileSizeBytes = fso.GetFile(path).Size
End Function

' Stem for every file belonging to one sample, so the extract sorts in
' sample order and each file names the evidence it is.
Public Function EvidenceStem(ByVal sampleIdx As Long, ByVal paymentDate As Date, _
                             ByVal amount As Double) As String
    EvidenceStem = Format$(sampleIdx, "00") & "_" & _
                   Format$(paymentDate, "yyyy-mm-dd") & "_" & _
                   Format$(Abs(amount), "0")
End Function
