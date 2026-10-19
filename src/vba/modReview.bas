Attribute VB_Name = "modReview"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit

Public Function RowSignature(ByVal a As Variant, ByVal rowNo As Long) As String
    Dim j As Long, s As String, v As Variant
    For j = 1 To UBound(a, 2)
        v = a(rowNo, j)
        If IsError(v) Then
            s = s & KeyPart("Error:" & CStr(v))
        ElseIf IsEmpty(v) Then
            s = s & KeyPart("Empty:")
        Else
            s = s & KeyPart(CStr(VarType(v)) & ":" & CStr(v))
        End If
    Next
    RowSignature = s
End Function

Public Function ExactRowEqual(ByVal a As Variant, ByVal row1 As Long, ByVal row2 As Long) As Boolean
    Dim j As Long, x As Variant, y As Variant
    For j = 1 To UBound(a, 2)
        x = a(row1, j): y = a(row2, j)
        If VarType(x) <> VarType(y) Then Exit Function
        If IsError(x) Then
            If CStr(x) <> CStr(y) Then Exit Function
        ElseIf Not IsEmpty(x) Then
            If x <> y Then Exit Function
        End If
    Next
    ExactRowEqual = True
End Function

Public Sub ReviewCases(Optional ByVal source As Range, Optional ByVal incidentHeader As String = "IncidentID", Optional ByVal crimeHeader As String = "")
    On Error GoTo failed
    Dim crimeCol As Long, types As Object, lists As Object, oneIncident As Object, crime As String, crimeKey As String, item As Variant
    Dim errNo As Long, description As String
    Dim r As Range, a As Variant, c As Long, i As Long, ws As Worksheet, out() As Variant
    Dim rows As Object, cases As Object, representatives As Object, sig As String, baseSig As String, id As String, collision As Long, rowKeys() As String
    If source Is Nothing Then OpenToolSetup "capReview": Exit Sub
    Set r = source
    If r Is Nothing Then Set r = PickSource("Include IncidentID and Crime Type plus all offense records. Repeated IDs may represent several offenses; no rows are removed.")
    If r Is Nothing Then Exit Sub
    ValidateOutputSource r
    a = ScopedValues(r)
    If source Is Nothing Then
        c = AskColumn(a, "Stable incident key; include agency/year when case numbers repeat.", "IncidentID", False, "INCIDENTID")
        crimeCol = AskColumn(a, "Crime-type/offense-description field for the incident offense list. Optional if absent.", "CrimeType", True, "CRIME")
    Else
        c = HeaderColumn(a, incidentHeader)
        If Len(crimeHeader) > 0 Then crimeCol = HeaderColumn(a, crimeHeader) Else crimeCol = OptionalHeader(a, "CrimeType")
    End If
    Set rows = CreateObject("Scripting.Dictionary"): Set cases = CreateObject("Scripting.Dictionary")
    Set types = CreateObject("Scripting.Dictionary"): Set lists = CreateObject("Scripting.Dictionary")
    Set representatives = CreateObject("Scripting.Dictionary")
    ReDim rowKeys(1 To UBound(a, 1) - 1)
    For i = 2 To UBound(a, 1)
        baseSig = RowSignature(a, i): sig = baseSig: collision = 0
        Do While rows.Exists(sig)
            If ExactRowEqual(a, i, representatives(sig)) Then Exit Do
            collision = collision + 1: sig = baseSig & KeyPart("collision" & collision)
        Loop
        If Not rows.Exists(sig) Then
            rows.Add sig, 0: representatives.Add sig, i
        End If
        rowKeys(i - 1) = sig
        rows(sig) = rows(sig) + 1
        id = SafeText(a(i, c))
        If Len(id) > 0 And Not IsError(a(i, c)) Then
            If Not cases.Exists(id) Then cases.Add id, 0
            cases(id) = cases(id) + 1
            If Not types.Exists(id) Then Set oneIncident = CreateObject("Scripting.Dictionary"): types.Add id, oneIncident
            Set oneIncident = types(id)
            If crimeCol > 0 Then
                If Not IsError(a(i, crimeCol)) Then
                    crime = SafeText(a(i, crimeCol)): crimeKey = NormalText(crime)
                    If Len(crime) > 0 And Not oneIncident.Exists(crimeKey) Then oneIncident.Add crimeKey, crime
                End If
            End If
        End If
    Next
    For Each item In types.Keys
        Set oneIncident = types(item): lists.Add item, IncidentTypeList(oneIncident)
    Next
    ReDim out(1 To UBound(a, 1) - 1, 1 To 6)
    For i = 2 To UBound(a, 1)
        sig = rowKeys(i - 1): id = SafeText(a(i, c))
        out(i - 1, 1) = ScopeSourceRow(r, i)
        out(i - 1, 2) = rows(sig)
        out(i - 1, 3) = rows(sig) > 1
        If Len(id) = 0 Or IsError(a(i, c)) Then
            out(i - 1, 5) = "Missing/error incident ID"
        Else
            out(i - 1, 4) = cases(id)
            out(i - 1, 5) = IIf(cases(id) > 1, "Repeated ID: offense rows or duplicate; review", "Single record for ID")
        End If
        If crimeCol = 0 Then
            out(i - 1, 6) = "[Crime type field not mapped]"
        ElseIf Len(id) = 0 Or IsError(a(i, c)) Then
            out(i - 1, 6) = "[No valid incident ID]"
        Else
            out(i - 1, 6) = lists(id)
        End If
    Next
    Set ws = NewOutput(r.Parent.Parent, "CaseReview", r, "No records removed. Exact equality uses all selected columns and value types. Repeated incident IDs may represent legitimate offense records. IDs are trimmed, case-sensitive. Incident crime types are distinct nonblank descriptions, normalized for case/spacing and listed on every offense row; missing/error types are not invented.")
    AnnotateAnalysisScope ws, r
    ws.Range("A6").Resize(UBound(a, 1), UBound(a, 2)).NumberFormat = "@"
    ws.Range("A6").Resize(UBound(a, 1), UBound(a, 2)).Value2 = a
    CopySourceFormats r, ws.Range("A6").Resize(UBound(a, 1), UBound(a, 2))
    WriteHeaders ws, 6, UBound(a, 2) + 1, Array("CAP_SourceRow", "CAP_ExactRowCount", "CAP_ExactDuplicate", "CAP_IncidentRowCount", "CAP_CaseReview", "CAP_IncidentCrimeTypes")
    ws.Cells(7, UBound(a, 2) + 1).Resize(UBound(out, 1), 6).NumberFormat = "@"
    ws.Cells(7, UBound(a, 2) + 1).Resize(UBound(out, 1), 6).Value2 = out
    ws.Cells(4, 1).Value2 = "Records: " & CStr(UBound(a, 1) - 1) & " | Distinct valid incident IDs: " & CStr(cases.Count)
    FinishOutput ws, UBound(a, 1) + 5, UBound(a, 2) + 6
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    If Not source Is Nothing Then Err.Raise errNo, "Duplicate/case review", description
    MsgBox description, vbExclamation, "CAP Index"
End Sub

Private Function IncidentTypeList(ByVal descriptions As Object) As String
    Dim keys As Variant, i As Long, result As String, entry As String
    If descriptions.Count = 0 Then IncidentTypeList = "[No recorded crime type]": Exit Function
    keys = descriptions.Keys
    SortCrimeTypeKeys keys, LBound(keys), UBound(keys)
    For i = LBound(keys) To UBound(keys)
        entry = IIf(Len(result) > 0, "; ", "") & descriptions(keys(i))
        If Len(result) + Len(entry) > 30000 Then
            result = result & "; [list shortened; full offense rows retained]": Exit For
        End If
        result = result & entry
    Next
    IncidentTypeList = result
End Function

Private Sub SortCrimeTypeKeys(ByRef keys As Variant, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long, pivot As String, temporary As Variant
    i = first: j = last: pivot = keys((first + last) \ 2)
    Do While i <= j
        Do While StrComp(keys(i), pivot, vbTextCompare) < 0: i = i + 1: Loop
        Do While StrComp(keys(j), pivot, vbTextCompare) > 0: j = j - 1: Loop
        If i <= j Then temporary = keys(i): keys(i) = keys(j): keys(j) = temporary: i = i + 1: j = j - 1
    Loop
    If first < j Then SortCrimeTypeKeys keys, first, j
    If i < last Then SortCrimeTypeKeys keys, i, last
End Sub
