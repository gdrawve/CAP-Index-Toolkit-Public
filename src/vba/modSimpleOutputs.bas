Attribute VB_Name = "modSimpleOutputs"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit

Public Sub BuildHundredBlocks(Optional ByVal source As Range, Optional ByVal addressHeader As String = "Address", Optional ByVal cityHeader As String = "", Optional ByVal stateHeader As String = "", Optional ByVal incidentHeader As String = "", Optional ByVal generateSynthetic As Boolean = False, Optional ByVal seed As Long = 20261013, Optional ByVal showSummary As Boolean = True)
    Dim r As Range, ws As Worksheet, a As Variant, output() As Variant, ac As Long, ic As Long, cc As Long, sc As Long
    Dim i As Long, j As Long, cols As Long, base As Long, street As String, precision As String, status As String, masked As Boolean
    Dim fieldPos() As Long, finalCols As Long
    Dim identity As String, key As String, response As Variant, headers As Variant, errNo As Long, description As String
    On Error GoTo failed
    If source Is Nothing Then OpenToolSetup IIf(generateSynthetic, "capBlockRandom", "capBlocks"): Exit Sub
    Set r = source
    If r Is Nothing Then Set r = PickSource("Select your working dataset headers/data. Only selected rows and fields are copied. Include the headers and every required field. Include Address; randomization can use IncidentID, City and State.")
    If r Is Nothing Then Exit Sub
    Set r = FullWorkingSource(r)
    ValidateOutputSource r
    a = r.Value2: cols = UBound(a, 2)
    If source Is Nothing Then
        ac = AskColumn(a, "Address or agency hundred-block field.", "Address", False, "ADDRESS")
        If generateSynthetic Then
            ic = AskColumn(a, "Optional IncidentID for consistent synthetic numbers across offense rows.", "", True, "INCIDENTID")
            cc = AskColumn(a, "Optional City to distinguish identical block addresses.", "", True, "CITY")
            sc = AskColumn(a, "Optional State.", "", True, "STATE")
            response = Application.InputBox("Repeatable seed: 1 to 2147483646. The seed is recorded in the output sheet's run notes.", "CAP Index | Randomized Addresses", CStr(seed), Type:=1)
            If VarType(response) = vbBoolean Then Exit Sub
            If response < 1 Or response > 2147483646# Or response <> Fix(response) Then Err.Raise vbObjectError + 700, , "Invalid random seed."
            seed = CLng(response)
        End If
    Else
        ac = HeaderColumn(a, addressHeader): ic = HeaderColumn(a, incidentHeader)
        cc = HeaderColumn(a, cityHeader): sc = HeaderColumn(a, stateHeader)
    End If
    If seed < 1 Or seed > 2147483646 Then Err.Raise vbObjectError + 700, , "Invalid random seed."
    If generateSynthetic Then headers = Array("CAP_HundredBlock", "CAP_SyntheticAddress", "Review Status") Else headers = Array("CAP_HundredBlock", "Review Status")
    ReDim fieldPos(0 To UBound(headers)): finalCols = cols
    For j = 0 To UBound(headers)
        fieldPos(j) = OptionalHeader(a, CStr(headers(j)))
        If fieldPos(j) = 0 Then finalCols = finalCols + 1: fieldPos(j) = finalCols
    Next
    ReDim output(1 To UBound(a, 1), 1 To finalCols)
    For i = 1 To UBound(a, 1): For j = 1 To cols: output(i, j) = a(i, j): Next: Next
    For j = 0 To UBound(headers): output(1, fieldPos(j)) = headers(j): Next
    For i = 2 To UBound(a, 1)
        For j = 0 To UBound(headers): output(i, fieldPos(j)) = Empty: Next
        If ParseBlockAddress(a(i, ac), base, street, precision, status, masked) Then
            output(i, fieldPos(0)) = CStr(base) & " BLOCK " & street
            status = "Ready"
            If generateSynthetic Then
                If masked Then
                    identity = "": If ic > 0 Then identity = SafeText(a(i, ic))
                    If Len(identity) = 0 Or identity = "[CELL ERROR]" Then identity = "ROW:" & r.Row + i - 1 Else identity = "INCIDENT:" & identity
                    key = KeyPart(output(i, fieldPos(0)))
                    If cc > 0 Then key = key & KeyPart(NormalText(a(i, cc)))
                    If sc > 0 Then key = key & KeyPart(NormalText(a(i, sc)))
                    output(i, fieldPos(1)) = CStr(base + BlockRandomSuffix(seed, KeyPart(identity) & key)) & " " & street
                    status = "Synthetic block estimate; not the observed incident address"
                Else
                    output(i, fieldPos(1)) = SafeText(a(i, ac)): status = "Original address retained; not a masked block"
                End If
            End If
        End If
        output(i, fieldPos(UBound(headers))) = status
    Next
    Set ws = NewOutput(r.Parent.Parent, IIf(generateSynthetic, "Randomized_Addresses", "Hundred_Blocks"), r, IIf(generateSynthetic, "Seed=" & seed & ". 00/XX block inputs generate numbers from base through base+99. Synthetic numbers do not recover true locations.", "Hundred-block addresses and review status only. City/state-qualified analytical keys are constructed internally. Street aliases require review."))
    ws.Range("A6").Resize(UBound(output, 1), UBound(output, 2)).NumberFormat = "@"
    ws.Range("A6").Resize(UBound(output, 1), UBound(output, 2)).Value2 = output
    CopySourceFormats r, ws.Range("A6").Resize(UBound(output, 1), cols)
    For j = 0 To UBound(headers): ws.Cells(6, fieldPos(j)).Resize(UBound(output, 1)).NumberFormat = "@": Next
    FinishOutput ws, UBound(output, 1) + 5, UBound(output, 2)
    MakeWorkingTable ws, UBound(output, 1) + 5, UBound(output, 2)
    If showSummary Then MsgBox "Created " & ws.Name & ". Original source data is unchanged. Select the table starting at row 6 for another tool.", vbInformation, "CAP Index"
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    If Not source Is Nothing Then Err.Raise errNo, "Hundred Blocks", description
    MsgBox description, vbExclamation, "CAP Index"
End Sub

Public Sub ValidateOutputSource(ByVal r As Range)
    If r.Areas.Count <> 1 Or r.Rows.Count < 2 Or r.Rows.Count > 200001 Or r.Columns.Count > 200 Then Err.Raise vbObjectError + 701, , "Select headers and 1 to 200,000 data rows in one rectangle, at most 200 columns."
    If r.Parent.Parent Is ThisWorkbook Then Err.Raise vbObjectError + 702, , "Select your data workbook."
    If r.Parent.Parent.ReadOnly Then Err.Raise vbObjectError + 703, , "Save a writable copy of the data workbook."
    If r.Parent.Parent.ProtectStructure Then Err.Raise vbObjectError + 704, , "Unprotect workbook structure so the output sheet can be added. The source sheet itself can remain protected."
End Sub

Public Sub GeocodeCensus(Optional ByVal source As Range, Optional ByVal addressHeader As String = "Address", Optional ByVal cityHeader As String = "City", Optional ByVal stateHeader As String = "State", Optional ByVal zipHeader As String = "", Optional ByVal transport As Object, Optional ByVal batchRows As Long = 10000, Optional ByVal showSummary As Boolean = True, Optional ByVal batchBytes As Long = 4900000)
    Dim r As Range, ws As Worksheet, a As Variant, result As Variant, output() As Variant, headers As Variant
    Dim ac As Long, cc As Long, sc As Long, zc As Long, anchor As Long, cols As Long, n As Long, i As Long, j As Long, srcCol As Long
    Dim fieldPos(1 To 7) As Long, finalCols As Long
    Dim errNo As Long, description As String, engineError As Long, engineDescription As String, named As Name
    On Error GoTo failed
    If source Is Nothing Then OpenToolSetup "capCensus": Exit Sub
    Set r = source
    If r Is Nothing Then Set r = PickSource("Select headers and data containing street Address plus City/State or ZIP. Only address components and generated request IDs are sent to Census. Results go to a new Geocoded_Addresses sheet.")
    If r Is Nothing Then Exit Sub
    Set r = FullWorkingSource(r)
    ValidateOutputSource r: a = r.Value2: n = UBound(a, 1): cols = UBound(a, 2)
    If source Is Nothing Then
        ac = AskColumn(a, "Street address or CAP_SyntheticAddress.", "Address", False, "ADDRESS")
        cc = AskColumn(a, "Optional City when ZIP is available.", "City", True, "CITY")
        sc = AskColumn(a, "Optional State when ZIP is available.", "State", True, "STATE")
        zc = AskColumn(a, "Optional ZIP when City and State are available.", "", True, "ZIP")
    Else
        ac = HeaderColumn(a, addressHeader): cc = HeaderColumn(a, cityHeader): sc = HeaderColumn(a, stateHeader): zc = HeaderColumn(a, zipHeader)
    End If
    addressHeader = SafeText(a(1, ac)): cityHeader = "": stateHeader = "": zipHeader = ""
    If cc > 0 Then cityHeader = SafeText(a(1, cc))
    If sc > 0 Then stateHeader = SafeText(a(1, sc))
    If zc > 0 Then zipHeader = SafeText(a(1, zc))
    anchor = Application.Max(ac, cc, sc, zc)
    Set ws = NewOutput(r.Parent.Parent, "Geocoded_Addresses", r, "Census benchmark=" & CENSUS_BENCHMARK & " | Run=" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & ". Match type is not a numeric confidence score. Coordinates interpolate address ranges; synthetic addresses remain approximate.")
    ws.Range("A6").Resize(n, cols).NumberFormat = "@": ws.Range("A6").Resize(n, cols).Value2 = a
    On Error Resume Next
    GeocodeCensusInPlace ws.Range("A6").Resize(n, cols), addressHeader, cityHeader, stateHeader, zipHeader, transport, batchRows, False, batchBytes
    engineError = Err.Number: engineDescription = Err.Description: Err.Clear: On Error GoTo failed
    If SafeText(ws.Cells(6, anchor + 1).Value2) <> "Full Address" Then
        If engineError <> 0 Then Err.Raise engineError, , engineDescription
        Err.Raise vbObjectError + 705, , "Geocoder output was not created."
    End If
    result = ws.Range("A6").Resize(n, cols + 9).Value2
    headers = Array("Full Address", "Matched Address", "Latitude", "Longitude", "Match Status", "Match Type", "Review Flag")
    finalCols = cols
    For j = 1 To 7
        fieldPos(j) = OptionalHeader(a, CStr(headers(j - 1)))
        If fieldPos(j) = 0 Then finalCols = finalCols + 1: fieldPos(j) = finalCols
    Next
    ReDim output(1 To n, 1 To finalCols)
    For i = 1 To n
        For j = 1 To cols
            srcCol = j: If j > anchor Then srcCol = j + 9
            output(i, j) = result(i, srcCol)
        Next
        For j = 1 To 7: output(i, fieldPos(j)) = result(i, anchor + j): Next
    Next
    For Each named In ws.Names
        If InStr(named.Name, "_CAPCensus_") > 0 Then named.Delete
    Next
    ws.Range("A6").Resize(n, cols + 9).Clear
    ws.Range("A6").Resize(n, finalCols).NumberFormat = "@": ws.Range("A6").Resize(n, finalCols).Value2 = output
    CopySourceFormats r, ws.Range("A6").Resize(n, cols)
    ws.Cells(7, fieldPos(3)).Resize(n - 1).NumberFormat = "0.000000"
    ws.Cells(7, fieldPos(4)).Resize(n - 1).NumberFormat = "0.000000"
    FinishOutput ws, n + 5, finalCols
    MakeWorkingTable ws, n + 5, finalCols
    If engineError <> 0 Then Err.Raise engineError, , engineDescription & ". Completed results retained in " & ws.Name
    If showSummary Then MsgBox "Created " & ws.Name & ". Review match status, match type and review flag. Source data is unchanged.", vbInformation, "CAP Index"
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    If Not source Is Nothing Then Err.Raise errNo, "Census Geocoder", description
    MsgBox description, vbExclamation, "CAP Index"
End Sub

Public Function OptionalHeader(ByVal a As Variant, ByVal header As String) As Long
    Dim j As Long
    For j = 1 To UBound(a, 2)
        If StrComp(SafeText(a(1, j)), header, vbTextCompare) = 0 Then
            If OptionalHeader > 0 Then Err.Raise vbObjectError + 706, , "Duplicate source header: " & header
            OptionalHeader = j
        End If
    Next
End Function

Public Sub CopySourceFormats(ByVal source As Range, ByVal target As Range)
    Dim j As Long, sourceFormat As Variant
    For j = 1 To source.Columns.Count
        sourceFormat = source.Columns(j).NumberFormat
        If IsNull(sourceFormat) Then sourceFormat = source.Cells(2, j).NumberFormat
        target.Columns(j).NumberFormat = sourceFormat
    Next
End Sub

Public Function FullWorkingSource(ByVal selected As Range) As Range
    Dim lo As ListObject
    If selected.Areas.Count <> 1 Or selected.Rows.Count < 2 Then Err.Raise vbObjectError + 707, , "Select the working dataset including its header and data rows."
    On Error Resume Next: Set lo = selected.Cells(1).ListObject: On Error GoTo 0
    If Not lo Is Nothing Then
        If selected.Row <> lo.HeaderRowRange.Row Then Err.Raise vbObjectError + 708, , "Include the table header in the selection."
        If lo.ListRows.Count = 0 Then Err.Raise vbObjectError + 709, , "The table has no records."
        If lo.ShowTotals Then
            If Not Intersect(selected, lo.TotalsRowRange) Is Nothing Then Err.Raise vbObjectError + 709, , "Exclude the table totals row."
        End If
    End If
    Set FullWorkingSource = selected
End Function

Private Sub MakeWorkingTable(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim table As ListObject
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    Set table = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(6, 1), ws.Cells(lastRow, lastCol)), , xlYes)
    table.TableStyle = "TableStyleMedium2"
End Sub
