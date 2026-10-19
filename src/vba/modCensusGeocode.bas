Attribute VB_Name = "modCensusGeocode"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit
Public Function RepairZip(ByVal raw As Variant, ByRef repaired As Boolean) As String
    Dim original As String, digits As String, i As Long, ch As String
    original = SafeText(raw): RepairZip = original: repaired = False
    If IsError(raw) Or IsNull(raw) Or IsEmpty(raw) Then Exit Function
    digits = original
    If VarType(raw) <> vbString Then
        If VarType(raw) = vbBoolean Or VarType(raw) = vbDate Then Exit Function
        If Not IsNumeric(raw) Then Exit Function
        If CDbl(raw) < 0 Or CDbl(raw) <> Fix(CDbl(raw)) Then Exit Function
        If CDbl(raw) > 999999999# Then Exit Function
        digits = Format$(CDbl(raw), "0")
    End If
    For i = 1 To Len(digits)
        ch = Mid$(digits, i, 1): If ch < "0" Or ch > "9" Then Exit Function
    Next
    Select Case Len(digits)
        Case 3 To 5: RepairZip = Right$("00000" & digits, 5)
        Case 6 To 9
            digits = Right$("000000000" & digits, 9)
            RepairZip = Left$(digits, 5) & "-" & Right$(digits, 4)
    End Select
    repaired = (RepairZip <> original)
End Function

Public Function GeocodeHeaders() As Variant
    GeocodeHeaders = Array("Full Address", "Matched Address", "Latitude", "Longitude", "Match Status", "Match Type", "Review Flag", "Census Benchmark", "Geocoded At")
End Function

Public Function FullAddress(ByVal street As String, ByVal city As String, ByVal state As String, ByVal zip As String) As String
    Dim s As String: s = Trim$(street)
    If Len(Trim$(city)) > 0 Then s = s & IIf(Len(s) > 0, ", ", "") & Trim$(city)
    If Len(Trim$(state)) > 0 Then s = s & IIf(Len(s) > 0, ", ", "") & Trim$(state)
    If Len(Trim$(zip)) > 0 Then s = s & IIf(Len(s) > 0, " ", "") & Trim$(zip)
    FullAddress = s
End Function

Public Function AddressProblem(ByVal street As String, ByVal city As String, ByVal state As String, ByVal zip As String) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    If Len(street) = 0 Then AddressProblem = "Missing street address": Exit Function
    If InStr(street, vbCr) Or InStr(street, vbLf) Or InStr(city, vbCr) Or InStr(city, vbLf) Or InStr(state, vbCr) Or InStr(state, vbLf) Or InStr(zip, vbCr) Or InStr(zip, vbLf) Then AddressProblem = "Address contains a line break; review": Exit Function
    re.IgnoreCase = True: re.Pattern = "(^P\.?\s*O\.?\s+BOX\b|\bBLOCK\b|\sAND\s|&|\sAT\s)"
    If re.Test(street) Then AddressProblem = "PO box, intersection or masked block; review": Exit Function
    If Len(zip) > 0 Then
        re.Pattern = "^\d{5}(-\d{4})?$"
        If Not re.Test(zip) Then AddressProblem = "Invalid ZIP; preserve leading zeros": Exit Function
    End If
    If Len(zip) = 0 And (Len(city) = 0 Or Len(state) = 0) Then AddressProblem = "Provide street plus city/state or ZIP"
End Function

Private Function ExistingGeocodeOutput(ByVal r As Range, ByVal anchorCol As Long) As Range
    Dim named As Name, marked As Range, headers As Variant, j As Long
    headers = GeocodeHeaders()
    For Each named In r.Parent.Names
        If InStr(1, named.Name, "_CAPCensus_", vbTextCompare) > 0 Then
            Set marked = Nothing
            On Error Resume Next: Set marked = named.RefersToRange: On Error GoTo 0
            If Not marked Is Nothing Then
                If marked.Row = r.Row And marked.Column = anchorCol And marked.Columns.Count = 10 Then
                    For j = 0 To 8
                        If SafeText(marked.Cells(1, j + 2).Value2) <> headers(j) Then Err.Raise vbObjectError + 320, , "Generated geocoder headers were changed; restore them before refreshing."
                    Next
                    Set ExistingGeocodeOutput = marked.Cells(1, 2): Exit Function
                End If
            End If
        End If
    Next
End Function

Private Sub WriteGeocodeOutput(ByVal first As Range, ByVal out As Variant)
    Dim n As Long: n = UBound(out, 1)
    first.Offset(1).Resize(n, 9).NumberFormat = "@"
    first.Offset(1).Resize(n, 9).Value2 = out
    first.Offset(1, 2).Resize(n, 2).NumberFormat = "0.000000"
    first.Offset(1, 8).Resize(n, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
End Sub

Public Sub GeocodeCensusInPlace(Optional ByVal source As Range, Optional ByVal addressHeader As String = "Address", Optional ByVal cityHeader As String = "City", Optional ByVal stateHeader As String = "State", Optional ByVal zipHeader As String = "", Optional ByVal transport As Object, Optional ByVal batchRows As Long = 10000, Optional ByVal showSummary As Boolean = True, Optional ByVal batchBytes As Long = 4900000)
    Dim r As Range, ws As Worksheet, wb As Workbook, a As Variant, ac As Long, cc As Long, sc As Long, zc As Long
    Dim n As Long, i As Long, j As Long, id As String, street As String, city As String, state As String, zip As String, problem As String, key As String
    Dim unique As Object, ids As Object, lines As Collection, rowIDs() As String, out() As Variant
    Dim first As Range, anchor As Long, lo As ListObject, pos As Long, inserted As Long, marked As Name, headers As Variant
    Dim csv As String, line As String, start As Long, finish As Long, bytes As Long, expected As Object, results As Object, content As String, value As Variant
    Dim complete As Long, batches As Long, processedRows As Long, beganMutation As Boolean, offset As Double
    Dim oldEvents As Boolean, oldScreen As Boolean, oldInteractive As Boolean, oldCancel As Long, oldStatus As Variant
    Dim failure As String, failedRun As Boolean, errNo As Long
    Dim matches As Long, nonExact As Long, ties As Long, noMatches As Long, invalidInputs As Long, unfinished As Long
    Dim locationNote() As String, precisionCol As Long, zipRepaired As Boolean
    On Error GoTo failed
    Set r = source: If r Is Nothing Then Set r = PickSource()
    If r Is Nothing Then Exit Sub
    Set ws = r.Parent: Set wb = ws.Parent
    If r.Areas.Count <> 1 Or r.Rows.Count < 2 Or r.Rows.Count > 200001 Or r.Columns.Count > 200 Then Err.Raise vbObjectError + 321, , "Select a rectangular header/data range with up to 200,000 rows and 200 columns."
    If wb Is ThisWorkbook Then Err.Raise vbObjectError + 322, , "Select a data workbook."
    If wb.ReadOnly Or ws.ProtectContents Then Err.Raise vbObjectError + 323, , "The workbook must be writable and the worksheet unprotected."
    If IsNull(r.MergeCells) Then Err.Raise vbObjectError + 324, , "Unmerge source cells before geocoding."
    If r.MergeCells Then Err.Raise vbObjectError + 324, , "Unmerge source cells before geocoding."
    If batchRows < 1 Or batchRows > CENSUS_MAX_ROWS Or batchBytes < 256 Or batchBytes > CENSUS_MAX_BYTES Then Err.Raise vbObjectError + 325, , "Invalid Census batch limit."
    a = r.Value2: n = UBound(a, 1) - 1
    If source Is Nothing Then
        ac = AskColumn(a, "Street address.", "Address", False, "ADDRESS")
        cc = AskColumn(a, "City. Optional when ZIP is available.", "City", True, "CITY")
        sc = AskColumn(a, "State. Optional when ZIP is available.", "State", True, "STATE")
        zc = AskColumn(a, "ZIP. Optional when city/state are available.", "", True, "ZIP")
    Else
        ac = HeaderColumn(a, addressHeader): cc = HeaderColumn(a, cityHeader)
        sc = HeaderColumn(a, stateHeader): zc = HeaderColumn(a, zipHeader)
    End If
    anchor = r.Column - 1 + Application.Max(ac, cc, sc, zc)
    If ac + 4 <= UBound(a, 2) Then
        If StrComp(SafeText(a(1, ac + 4)), "CAP_AddressPrecision", vbTextCompare) = 0 Then precisionCol = ac + 4
    End If
    If anchor > ws.Columns.Count - 9 Then Err.Raise vbObjectError + 326, , "There is no room for nine geocoder fields."
    Set unique = CreateObject("Scripting.Dictionary"): Set ids = CreateObject("Scripting.Dictionary"): Set lines = New Collection
    ReDim out(1 To n, 1 To 9): ReDim rowIDs(1 To n): ReDim locationNote(1 To n)
    For i = 2 To UBound(a, 1)
        street = SafeText(a(i, ac)): city = "": state = "": zip = "": problem = ""
        If cc > 0 Then city = SafeText(a(i, cc))
        If sc > 0 Then state = SafeText(a(i, sc))
        zipRepaired = False
        If zc > 0 Then zip = RepairZip(a(i, zc), zipRepaired)
        out(i - 1, 1) = FullAddress(street, city, state, zip)
        If StrComp(SafeText(a(1, ac)), "CAP_SyntheticAddress", vbTextCompare) = 0 Then
            locationNote(i - 1) = "Synthetic hundred-block estimate; not an observed incident location"
        ElseIf precisionCol > 0 Then
            If InStr(1, SafeText(a(i, precisionCol)), "Hundred block", vbTextCompare) = 1 Then locationNote(i - 1) = "Masked hundred-block input; approximate location only"
        End If
        If zipRepaired Then
            If Len(locationNote(i - 1)) > 0 Then locationNote(i - 1) = locationNote(i - 1) & "; "
            locationNote(i - 1) = locationNote(i - 1) & "ZIP restored from a number or short value; verify"
            If Len(zip) = 10 Then locationNote(i - 1) = locationNote(i - 1) & "; ZIP+4 formatting and any missing leading zeros were inferred; verify"
        End If
        problem = AddressProblem(street, city, state, zip)
        If IsError(a(i, ac)) Then problem = "Source cell error"
        If cc > 0 Then
            If IsError(a(i, cc)) Then problem = "Source city cell error"
        End If
        If sc > 0 Then
            If IsError(a(i, sc)) Then problem = "Source state cell error"
        End If
        If zc > 0 Then
            If IsError(a(i, zc)) Then problem = "Source ZIP cell error"
        End If
        If Len(problem) > 0 Then
            out(i - 1, 5) = "Invalid Input": out(i - 1, 7) = problem
        Else
            key = KeyPart(NormalText(street)) & KeyPart(NormalText(city)) & KeyPart(NormalText(state)) & KeyPart(zip)
            If Not unique.Exists(key) Then
                id = "A" & Format$(unique.Count + 1, "000000")
                line = CsvQuote(id) & "," & CsvQuote(street) & "," & CsvQuote(city) & "," & CsvQuote(state) & "," & CsvQuote(zip) & vbCrLf
                If UTF8Size(line) > batchBytes Then
                    out(i - 1, 5) = "Invalid Input": out(i - 1, 7) = "Address exceeds batch byte limit"
                    GoTo nextRow
                End If
                unique.Add key, id: ids.Add id, unique.Count: lines.Add line
            End If
            rowIDs(i - 1) = unique(key): out(i - 1, 5) = "Pending"
        End If
nextRow:
    Next
    headers = GeocodeHeaders(): Set first = ExistingGeocodeOutput(r, anchor)
    On Error Resume Next: Set lo = ws.Cells(r.Row, anchor).ListObject: On Error GoTo failed
    If Not lo Is Nothing Then
        If r.Row <> lo.HeaderRowRange.Row Or n <> lo.ListRows.Count Or r.Column < lo.Range.Column Or r.Column + r.Columns.Count - 1 > lo.Range.Column + lo.Range.Columns.Count - 1 Then Err.Raise vbObjectError + 327, , "Select the table header and all its data rows, excluding totals."
        pos = anchor - lo.Range.Column + 1
        If first Is Nothing Then
            For j = 0 To 8
                For i = 1 To lo.ListColumns.Count
                    If StrComp(lo.ListColumns(i).Name, headers(j), vbTextCompare) = 0 Then Err.Raise vbObjectError + 328, , "Table field already exists: " & headers(j)
                Next
            Next
        End If
    End If
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    oldInteractive = Application.Interactive: oldCancel = Application.EnableCancelKey: oldStatus = Application.StatusBar
    beganMutation = True: Application.EnableEvents = False: Application.ScreenUpdating = False
    Application.EnableCancelKey = xlErrorHandler: Application.Interactive = False
    If first Is Nothing Then
        If lo Is Nothing Then
            ws.Columns(anchor + 1).Resize(, 9).Insert Shift:=xlToRight: inserted = 9
        Else
            For j = 1 To 9: lo.ListColumns.Add Position:=pos + j: inserted = inserted + 1: Next
        End If
        Set first = ws.Cells(r.Row, anchor + 1)
        Set marked = ws.Names.Add(Name:="_CAPCensus_" & anchor & "_" & r.Row, RefersTo:="='" & Replace(ws.Name, "'", "''") & "'!" & ws.Cells(r.Row, anchor).Resize(1, 10).Address)
        marked.Visible = False
    End If
    WriteHeaders ws, r.Row, first.Column, headers
    WriteGeocodeOutput first, out
    With first.Resize(1, 9)
        .Font.Bold = True: .Interior.Color = RGB(26, 55, 82): .Font.Color = RGB(255, 255, 255): .WrapText = True
    End With
    If wb.Date1904 Then offset = 1462
    start = 1
    Do While start <= lines.Count
        csv = "": bytes = 0: finish = start: Set expected = CreateObject("Scripting.Dictionary")
        Do While finish <= lines.Count And finish - start < batchRows
            line = lines(finish)
            If bytes + UTF8Size(line) > batchBytes Then Exit Do
            bytes = bytes + UTF8Size(line): csv = csv & line
            expected.Add "A" & Format$(finish, "000000"), True: finish = finish + 1
        Loop
        batches = batches + 1
        Application.StatusBar = "CAP Index Census | Batch " & batches & " | " & complete & " / " & lines.Count & " distinct addresses complete | Esc to cancel"
        If transport Is Nothing Then
            content = CensusPost(csv)
        Else
            content = transport.PostBatch(csv)
        End If
        Set results = CensusResponse(content, expected)
        For i = 1 To n
            If results.Exists(rowIDs(i)) Then
                value = results(rowIDs(i))
                For j = 0 To 5: out(i, j + 2) = value(j): Next
                If Len(locationNote(i)) > 0 Then out(i, 7) = locationNote(i) & IIf(Len(out(i, 7)) > 0, "; " & out(i, 7), "")
                out(i, 8) = CENSUS_BENCHMARK: out(i, 9) = CDbl(Now) - offset
                processedRows = processedRows + 1
            End If
        Next
        WriteGeocodeOutput first, out
        complete = complete + expected.Count: start = finish
        DoEvents
    Loop
    GoTo done
failed:
    errNo = Err.Number: failure = Err.Description: failedRun = True
    On Error Resume Next
    If beganMutation And Not first Is Nothing Then
        For i = 1 To n
            If out(i, 5) = "Pending" Then out(i, 5) = IIf(errNo = 18, "Cancelled", "Batch Failed"): out(i, 7) = Left$(failure, 240)
        Next
        WriteGeocodeOutput first, out
    ElseIf inserted > 0 Then
        If lo Is Nothing Then
            ws.Columns(anchor + 1).Resize(, inserted).Delete
        Else
            For j = 1 To inserted: lo.ListColumns(pos + 1).Delete: Next
        End If
    End If
    On Error GoTo 0
done:
    On Error Resume Next
    If Not first Is Nothing Then
        For i = 1 To n
            Select Case out(i, 5)
                Case "Match"
                    matches = matches + 1
                    If out(i, 6) = "Non-Exact" Then nonExact = nonExact + 1
                Case "Tie": ties = ties + 1
                Case "No Match": noMatches = noMatches + 1
                Case "Invalid Input": invalidInputs = invalidInputs + 1
                Case Else: unfinished = unfinished + 1
            End Select
        Next
        first.Resize(n + 1, 9).Columns.AutoFit
        For j = 0 To 8
            If first.Offset(0, j).EntireColumn.ColumnWidth > 40 Then
                first.Offset(0, j).EntireColumn.ColumnWidth = 40
                first.Offset(0, j).Resize(n + 1).WrapText = True
            End If
        Next
        first.Resize(n + 1, 9).Rows.AutoFit
    End If
    If beganMutation Then
        Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
        Application.Interactive = oldInteractive: Application.EnableCancelKey = oldCancel: Application.StatusBar = oldStatus
    End If
    On Error GoTo 0
    If failedRun And Not source Is Nothing Then Err.Raise errNo, "Census Geocoder", failure
    If showSummary Then
        If failedRun Then
            MsgBox "Census Geocoder stopped: " & failure & vbCrLf & "Completed results retained: " & complete & " distinct addresses / " & processedRows & " source rows. Unfinished rows are flagged.", vbExclamation, "CAP Index"
        Else
            MsgBox "Census geocoding complete." & vbCrLf & "Distinct addresses processed: " & complete & " | Batches: " & batches & vbCrLf & "Source rows: " & n & vbCrLf & "Matched: " & matches & " (non-exact: " & nonExact & ")" & vbCrLf & "No match: " & noMatches & " | Ties: " & ties & vbCrLf & "Invalid input: " & invalidInputs & " | Unfinished: " & unfinished & vbCrLf & "Review Match Status and Review Flag. Coordinates estimate positions along address ranges.", vbInformation, "CAP Index"
        End If
    End If
End Sub
