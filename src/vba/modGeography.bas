Attribute VB_Name = "modGeography"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit

Public Function HundredBlock(ByVal raw As Variant, ByRef status As String) As String
    Dim s As String, re As Object, m As Object, street As String, number As Long
    status = "Unsupported address; review"
    If IsError(raw) Then status = "Source cell error": Exit Function
    s = NormalText(raw)
    If Len(s) = 0 Then status = "Missing address": Exit Function
    If InStr(s, "&") > 0 Or InStr(s, " AND ") > 0 Or InStr(s, " AT ") > 0 Or InStr(s, "BLOCK") > 0 Then Exit Function
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^([0-9]{1,7}) ([A-Z].*)$"
    If Not re.Test(s) Then Exit Function
    Set m = re.Execute(s)(0)
    number = CLng(m.SubMatches(0)): street = m.SubMatches(1)
    If Left$(street, 6) = "PO BOX" Or Left$(street, 8) = "P.O. BOX" Then Exit Function
    re.Pattern = "\s+((APT|UNIT|SUITE|STE)(\s+|$)|#\s*).*$": re.IgnoreCase = True
    street = Trim$(re.Replace(street, ""))
    If Len(street) = 0 Then Exit Function
    HundredBlock = CStr((number \ 100) * 100) & " BLOCK " & street
    status = "Parsed; street aliases not standardized"
End Function

Public Sub BuildHundredBlocksLegacy(Optional ByVal source As Range, Optional ByVal addressHeader As String = "Address", Optional ByVal cityHeader As String = "City", Optional ByVal stateHeader As String = "State")
    On Error GoTo failed
    Dim r As Range, a As Variant, ac As Long, cc As Long, sc As Long, ws As Worksheet
    Dim out() As Variant, i As Long, block As String, status As String, city As String, state As String
    Set r = source
    If r Is Nothing Then Set r = PickSource()
    If r Is Nothing Then Exit Sub
    a = r.Value2
    If source Is Nothing Then
        ac = AskColumn(a, "Street address.", "Address", False, "ADDRESS")
        cc = AskColumn(a, "City or municipality.", "City", False, "CITY")
        sc = AskColumn(a, "State.", "State", False, "STATE")
    Else
        ac = HeaderColumn(a, addressHeader): cc = HeaderColumn(a, cityHeader): sc = HeaderColumn(a, stateHeader)
    End If
    ReDim out(1 To UBound(a, 1) - 1, 1 To 3)
    For i = 2 To UBound(a, 1)
        block = HundredBlock(a(i, ac), status)
        city = NormalText(a(i, cc)): state = NormalText(a(i, sc))
        out(i - 1, 1) = block
        If Len(block) > 0 Then
            If Len(city) = 0 Or Len(state) = 0 Or IsError(a(i, cc)) Or IsError(a(i, sc)) Then
                status = "Missing/error city or state; no block key"
            Else
                out(i - 1, 2) = KeyPart(state) & KeyPart(city) & KeyPart(block)
            End If
        End If
        out(i - 1, 3) = status
    Next
    Set ws = NewOutput(r.Parent.Parent, "HundredBlocks", r, "Address-number bins are not Census blocks or street segments. Original addresses remain sensitive. Intersections, ranges and masked addresses require review.")
    CopySourceValues r, ws
    WriteHeaders ws, 6, UBound(a, 2) + 1, Array("CAP_HundredBlock", "CAP_BlockKey", "CAP_BlockStatus")
    ws.Cells(7, UBound(a, 2) + 1).Resize(UBound(out, 1), 3).NumberFormat = "@"
    ws.Cells(7, UBound(a, 2) + 1).Resize(UBound(out, 1), 3).Value2 = out
    FinishOutput ws, UBound(a, 1) + 5, UBound(a, 2) + 3
    Exit Sub
failed: ToolkitError "Hundred-Block Builder"
End Sub
