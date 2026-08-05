#!/usr/bin/env python3
"""Fixes from the independent review of the shipping V4-CIO VBA."""
import io, os
SRC = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'build_v4')
P = os.path.join(SRC, 'mCloseEnv_V4.bas')
s = io.open(P, encoding='utf-8', newline='').read()

# ---- 1. the both-separators branch never recorded what it had just proved ----
old = '''    If pDot > 0 And pCom > 0 Then
        'both kinds present: whichever comes last is the decimal separator
        If pDot > pCom Then
            s = Replace(s, ",", "")
        Else
            s = Replace(s, ".", "")
            s = Replace(s, ",", ".")
        End If
'''
assert s.count(old) == 1
new = '''    If pDot > 0 And pCom > 0 Then
        'Both kinds present: whichever comes last is the decimal separator. This
        'is the most decisive evidence there is, so it is what the run learns from.
        If pDot > pCom Then
            If Not CM_Learn(".") Then reason = "CLASH": Exit Function
            s = Replace(s, ",", "")
        Else
            If Not CM_Learn(",") Then reason = "CLASH": Exit Function
            s = Replace(s, ".", "")
            s = Replace(s, ",", ".")
        End If
'''
s = s.replace(old, new)

for a, b in [
 ("            CM_DecSeen = other\n            s = Replace(s, sep, \"\")",
  "            If Not CM_Learn(other) Then reason = \"CLASH\": Exit Function\n            s = Replace(s, sep, \"\")"),
 ("            CM_DecSeen = sep\n            If sep = \",\" Then s = Replace(s, \",\", \".\")",
  "            If Not CM_Learn(sep) Then reason = \"CLASH\": Exit Function\n            If sep = \",\" Then s = Replace(s, \",\", \".\")"),
]:
    assert s.count(a) == 1, a[:40]
    s = s.replace(a, b)

# ---- 2. contradictory evidence must stop, not silently overwrite -------------
anchor = "'--- which character SAP is using as the decimal point"
assert s.count(anchor) == 1
s = s.replace(anchor, r'''
'--- record the convention; False if it contradicts what we already knew -----
' One SAP user writes amounts one way, so two extracts in one close cannot
' disagree. If they do, something is wrong that guessing would only hide - the
' close stops rather than read one literal two different ways in one run.
Private Function CM_Learn(ByVal sep As String) As Boolean
    If CM_SAP_DECIMAL = "." Or CM_SAP_DECIMAL = "," Then
        CM_Learn = (CM_SAP_DECIMAL = sep)      'pinned by configuration
        Exit Function
    End If
    If CM_DecSeen <> "" And CM_DecSeen <> sep Then Exit Function
    CM_DecSeen = sep
    CM_Learn = True
End Function


''' + anchor)

# ---- 3. explain the clash, and stop showing a one-digit "alternative" --------
old = '''            Case "AMBIG"'''
assert s.count(old) == 1
new = '''            Case "CLASH"
                what = "Two SAP reports in this run wrote amounts in different" & vbCrLf & _
                       "        number formats, which cannot both be right." & vbCrLf & _
                       "        The value that disagreed:  [" & parts(1) & "]"
            Case "AMBIG"'''
s = s.replace(old, new)

old = '''                       "        That is either " & Replace(Replace(parts(1), ".", ""), ",", "") & _
                       " or about " & Left$(parts(1), 1) & "." & vbCrLf & _'''
assert s.count(old) == 1
new = '''                       "        That is either " & Replace(Replace(parts(1), ".", ""), ",", "") & _
                       " or about " & Split(Replace(parts(1), ",", "."), ".")(0) & "." & vbCrLf & _'''
s = s.replace(old, new)

old = '''        If parts(0) = "AMBIG" Then'''
assert s.count(old) == 1
new = '''        If parts(0) = "CLASH" Then
            todo = "Send this message to the CI Team. One of the SAP reports is" & vbCrLf & _
                   "        being produced with a different decimal notation from the" & vbCrLf & _
                   "        others, which usually means a report variant or a user" & vbCrLf & _
                   "        default was changed. Setting CM_SAP_DECIMAL does not fix" & vbCrLf & _
                   "        this - the reports themselves disagree."
        ElseIf parts(0) = "AMBIG" Then'''
s = s.replace(old, new)

# ---- 4. CreateFolder with error handling off, right after a best-effort delete
old = '''    On Error Resume Next
    fso.DeleteFolder work, True
    Err.Clear
    On Error GoTo 0
    fso.CreateFolder work
'''
assert s.count(old) == 1
new = '''    On Error Resume Next
    fso.DeleteFolder work, True
    Err.Clear
    On Error GoTo 0
    'guarded: if the delete above could not remove it (open in Explorer, say),
    'an unguarded CreateFolder would raise error 58 and kill the preflight
    EnsureFolderChain fso, work
'''
s = s.replace(old, new)

io.open(P, 'w', encoding='utf-8', newline='').write(s)
print("mCloseEnv: learn-from-both-separators, clash check, CreateFolder guard, AMBIG wording")
