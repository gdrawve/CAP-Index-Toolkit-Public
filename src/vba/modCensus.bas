Attribute VB_Name = "modCensus"
Option Explicit

Public Sub PrepareCensus(Optional ByVal source As Range, Optional ByVal addressHeader As String = "Address", Optional ByVal cityHeader As String = "City", Optional ByVal stateHeader As String = "State", Optional ByVal zipHeader As String = "")
    On Error GoTo failed
    Dim r As Range, a As Variant, ac As Long, cc As Long, sc As Long, zc As Long
    Dim out() As Variant, i As Long, eligible As Long, ws As Worksheet, status As String
    Dim address As String, city As String, state As String, zip As String, parsed As String, zipRE As Object
    Set zipRE = CreateObject("VBScript.RegExp")
    zipRE.Pattern = "^\d{5}(-\d{4})?$"
    Set r = source
    If r Is Nothing Then Set r = PickSource()
    If r Is Nothing Then Exit Sub
    a = r.Value2
    If source Is Nothing Then
        ac = AskColumn(a, "Street address.", "Address", False, "ADDRESS")
        cc = AskColumn(a, "City.", "City", False, "CITY")
        sc = AskColumn(a, "State.", "State", False, "STATE")
        zc = AskColumn(a, "ZIP; preserve leading zeros as text. Optional.", "", True, "ZIP")
    Else
        ac = HeaderColumn(a, addressHeader): cc = HeaderColumn(a, cityHeader)
        sc = HeaderColumn(a, stateHeader): zc = HeaderColumn(a, zipHeader)
    End If
    ReDim out(1 To UBound(a, 1) - 1, 1 To 10)
    For i = 2 To UBound(a, 1)
        address = SafeText(a(i, ac)): city = SafeText(a(i, cc)): state = SafeText(a(i, sc)): zip = ""
        If zc > 0 Then zip = SafeText(a(i, zc))
        parsed = HundredBlock(address, status)
        status = "Ready for privacy review"
        If Len(parsed) = 0 Then status = "Unsupported address; review"
        If (Len(city) = 0 Or Len(state) = 0) And Len(zip) = 0 Then status = "Missing city/state or ZIP; review"
        If IsError(a(i, ac)) Or IsError(a(i, cc)) Or IsError(a(i, sc)) Then status = "Source cell error"
        If zc > 0 Then
            If IsError(a(i, zc)) Then status = "ZIP cell error"
            If Len(zip) > 0 And Not zipRE.Test(zip) Then status = "ZIP must be text with 5 digits or ZIP+4; review"
            If Len(zip) > 0 And VarType(a(i, zc)) <> vbString Then status = "ZIP stored numerically; review leading zeros"
        End If
        out(i - 1, 1) = "R" & Format$(i - 1, "000000")
        out(i - 1, 2) = address: out(i - 1, 3) = city: out(i - 1, 4) = state: out(i - 1, 5) = zip
        If status = "Ready for privacy review" Then
            eligible = eligible + 1
            out(i - 1, 6) = (eligible - 1) \ 10000 + 1
            out(i - 1, 7) = (eligible - 1) Mod 10000 + 1
        End If
        out(i - 1, 8) = r.Row + i - 1: out(i - 1, 9) = status
        out(i - 1, 10) = "Not submitted"
    Next
    Set ws = NewOutput(r.Parent.Parent, "CensusStaging", r, "Local staging only; no network requests. Privacy review required before external submission. CSV export: eligible rows, columns A:E only, no header, one file per BatchNo. SourceRow is internal; never upload case IDs or offense details.")
    ws.Cells(4, 1).Value2 = "Selected rows=" & UBound(a, 1) - 1 & "; ready for privacy review=" & eligible & "; batch limit=10000. Request IDs apply only to this saved run; retain the sheet for result joins."
    WriteHeaders ws, 6, 1, Array("RequestID", "Address", "City", "State", "ZIP", "BatchNo", "BatchRow", "SourceRow", "ReviewStatus", "SubmissionStatus")
    ws.Cells(7, 1).Resize(UBound(out, 1), 5).NumberFormat = "@"
    ws.Cells(7, 1).Resize(UBound(out, 1), 10).Value2 = out
    FinishOutput ws, UBound(a, 1) + 5, 10
    Exit Sub
failed: ToolkitError "Census batch staging"
End Sub
