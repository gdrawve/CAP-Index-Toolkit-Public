Attribute VB_Name = "modDatasetCheck"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
Option Explicit

Public Sub BuildDatasetCheck(Optional ByVal source As Range, Optional ByVal dateHeader As String = "Date", Optional ByVal crimeHeader As String = "CrimeType", Optional ByVal incidentHeader As String = "IncidentID", Optional ByVal addressHeader As String = "Address", Optional ByVal latitudeHeader As String = "Latitude", Optional ByVal longitudeHeader As String = "Longitude", Optional ByVal showSummary As Boolean = True)
    Dim r As Range, ws As Worksheet, a As Variant, ids As Object, signatures As Object, crimes As Object, months As Object, monthIDs As Object
    Dim dc As Long, cc As Long, ic As Long, ac As Long, lc As Long, oc As Long, n As Long, i As Long, j As Long, rowNo As Long
    Dim missingDates As Long, invalidDates As Long, missingIDs As Long, missingCrimes As Long, missingAddresses As Long, duplicates As Long
    Dim missingCoords As Long, invalidCoords As Long, zeroCoords As Long, swaps As Long, conflicts As Long, blankRows As Long, errors As Long
    Dim d As Date, ht As Boolean, status As String, id As String, category As String, signature As String, monthKey As String, k As Variant
    Dim firstMonth As Date, lastMonth As Date, span As Long, dates() As Double, validDates() As Boolean, blank As Boolean, lat As Double, lon As Double
    Dim fieldBlanks() As Long, fieldErrors() As Long, headerIssues As Long, headers As Object, oldEvents As Boolean, errNo As Long, description As String
    If source Is Nothing Then OpenToolSetup "capCheck": Exit Sub
    oldEvents = Application.EnableEvents
    On Error GoTo failed
    Set r = source
    If r Is Nothing Then Set r = PickSource("Dataset Check + Coverage: select headers/data. Map optional Date, Crime Type, IncidentID, Address, Latitude and Longitude. All selected rows are checked.")
    If r Is Nothing Then
        If showSummary Then MsgBox "No dataset selected. No changes were made.", vbInformation, "CAP Index | Dataset Check"
        Exit Sub
    End If
    ValidateOutputSource r: a = ScopedValues(r): n = UBound(a, 1) - 1
    If source Is Nothing Then
        dc = CheckField(a, "Date or timestamp", "Date"): cc = CheckField(a, "Crime Type", "CrimeType")
        ic = CheckField(a, "IncidentID", "IncidentID"): ac = CheckField(a, "Address", "Address")
        lc = CheckField(a, "Latitude (decimal degrees)", "Latitude"): oc = CheckField(a, "Longitude (decimal degrees)", "Longitude")
    Else
        dc = CheckColumn(a, dateHeader): cc = CheckColumn(a, crimeHeader): ic = CheckColumn(a, incidentHeader)
        ac = CheckColumn(a, addressHeader): lc = CheckColumn(a, latitudeHeader): oc = CheckColumn(a, longitudeHeader)
    End If
    Set ids = CreateObject("Scripting.Dictionary"): Set signatures = CreateObject("Scripting.Dictionary")
    Set crimes = CreateObject("Scripting.Dictionary"): Set months = CreateObject("Scripting.Dictionary"): Set monthIDs = CreateObject("Scripting.Dictionary")
    Set headers = CreateObject("Scripting.Dictionary")
    ReDim fieldBlanks(1 To UBound(a, 2)): ReDim fieldErrors(1 To UBound(a, 2))
    ReDim dates(2 To UBound(a, 1)): ReDim validDates(2 To UBound(a, 1))
    For j = 1 To UBound(a, 2)
        signature = NormalText(a(1, j))
        If Len(signature) = 0 Or IsError(a(1, j)) Then
            headerIssues = headerIssues + 1
        ElseIf headers.Exists(signature) Then
            headerIssues = headerIssues + 1
        Else
            headers.Add signature, True
        End If
    Next
    For i = 2 To UBound(a, 1)
        signature = "": blank = True
        For j = 1 To UBound(a, 2)
            If IsError(a(i, j)) Then
                fieldErrors(j) = fieldErrors(j) + 1: errors = errors + 1: blank = False
            ElseIf Len(SafeText(a(i, j))) = 0 Then
                fieldBlanks(j) = fieldBlanks(j) + 1
            Else
                blank = False
            End If
            signature = signature & KeyPart(CheckCellSignature(a(i, j)))
        Next
        If blank Then blankRows = blankRows + 1
        If Not blank Then
            If signatures.Exists(signature) Then duplicates = duplicates + 1 Else signatures.Add signature, True
        End If
        If dc > 0 Then
            If Len(SafeText(a(i, dc))) = 0 Then
                missingDates = missingDates + 1
            ElseIf ParseTimestamp(a(i, dc), r.Parent.Parent.Date1904, d, ht, status) Then
                validDates(i) = True: dates(i) = CDbl(DateValue(d)): monthKey = Format$(d, "yyyy-mm")
                If Not months.Exists(monthKey) Then months.Add monthKey, 0&
                months(monthKey) = months(monthKey) + 1
                d = DateSerial(Year(d), Month(d), 1)
                If firstMonth = 0 Or d < firstMonth Then firstMonth = d
                If d > lastMonth Then lastMonth = d
            Else
                invalidDates = invalidDates + 1
            End If
        End If
        If ic > 0 Then
            id = SafeText(a(i, ic))
            If Len(id) = 0 Or IsError(a(i, ic)) Then
                missingIDs = missingIDs + 1
            Else
                signature = IIf(validDates(i), CStr(dates(i)), "INVALID/ABSENT DATE")
                If Not ids.Exists(id) Then
                    ids.Add id, signature
                ElseIf ids(id) <> signature Then
                    ids(id) = "CONFLICT"
                End If
            End If
        End If
        If cc > 0 Then
            category = NormalText(a(i, cc))
            If Len(category) = 0 Or IsError(a(i, cc)) Then missingCrimes = missingCrimes + 1: category = "[Missing/error crime type]"
            If Not crimes.Exists(category) Then crimes.Add category, 0&
            crimes(category) = crimes(category) + 1
        End If
        If ac > 0 Then
            If Len(SafeText(a(i, ac))) = 0 Or IsError(a(i, ac)) Then missingAddresses = missingAddresses + 1
        End If
        If lc > 0 And oc > 0 Then
            If Len(SafeText(a(i, lc))) = 0 Or Len(SafeText(a(i, oc))) = 0 Then
                missingCoords = missingCoords + 1
            ElseIf IsError(a(i, lc)) Or IsError(a(i, oc)) Then
                invalidCoords = invalidCoords + 1
            ElseIf Not IsNumeric(a(i, lc)) Or Not IsNumeric(a(i, oc)) Then
                invalidCoords = invalidCoords + 1
            Else
                lat = CDbl(a(i, lc)): lon = CDbl(a(i, oc))
                If Abs(lat) > 90 Or Abs(lon) > 180 Then
                    invalidCoords = invalidCoords + 1
                    If Abs(lat) <= 180 And Abs(lon) <= 90 Then swaps = swaps + 1
                ElseIf lat = 0 And lon = 0 Then
                    zeroCoords = zeroCoords + 1
                End If
            End If
        End If
    Next
    For Each k In ids.Keys: If ids(k) = "CONFLICT" Then conflicts = conflicts + 1
    Next
    For i = 2 To UBound(a, 1)
        If validDates(i) And ic > 0 Then
            id = SafeText(a(i, ic))
            If ids.Exists(id) Then
                If ids(id) <> "CONFLICT" Then
                    signature = KeyPart(Format$(CDate(dates(i)), "yyyy-mm")) & KeyPart(id)
                    If Not monthIDs.Exists(signature) Then monthIDs.Add signature, True
                End If
            End If
        End If
    Next
    If firstMonth <> 0 Then
        If DateDiff("m", firstMonth, lastMonth) + 1 > 1200 Then Err.Raise vbObjectError + 901, , "Date span exceeds 100 years. Review date values before running monthly coverage."
    End If
    Application.EnableEvents = False
    Set ws = NewOutput(r.Parent.Parent, "Dataset_Check", r, "Dataset quality + coverage. Source records are retained; no source data changed.")
    AnnotateAnalysisScope ws, r
    ws.Range("A3:H3").Merge: ws.Range("A3").WrapText = True: ws.Rows(3).RowHeight = 30
    ws.Range("A4:H4").Merge: ws.Range("A4").Value2 = "Record grain is not inferred: rows may be offenses. Repeated IDs may represent valid multi-offense incidents. Duplicate counts do not remove rows."
    ws.Range("A5:H5").Merge: ws.Range("A5").Value2 = "Coverage is observed activity, not proof of complete reporting. Zero-record months may be gaps. Coordinate ranges do not establish jurisdiction or positional accuracy."
    ws.Range("A4:H5").WrapText = True: ws.Rows("4:5").RowHeight = 30
    WriteHeaders ws, 6, 1, Array("Check", "Count / status", "Interpretation")
    rowNo = 7
    CheckMetric ws, rowNo, "Source records", n, "All selected rows, including hidden/filtered and empty rows"
    CheckMetric ws, rowNo, "Blank rows", blankRows, "Fully empty selected data rows"
    CheckMetric ws, rowNo, "Header issues", headerIssues, "Blank/error headers and repeated normalized names"
    CheckMetric ws, rowNo, "Cell errors", errors, "Excel error values across all selected fields"
    CheckMetric ws, rowNo, "Exact duplicate extra rows", duplicates, "Value/type-identical rows beyond the first; blank rows excluded"
    CheckMetric ws, rowNo, "Recorded unique incident IDs", IIf(ic > 0, ids.Count, "Not mapped"), "Nonblank IDs; assumes one agency/key namespace"
    CheckMetric ws, rowNo, "Missing/error incident IDs", IIf(ic > 0, missingIDs, "Not mapped"), "Cannot support unique-incident counting"
    CheckMetric ws, rowNo, "IDs with conflicting dates", IIf(ic > 0 And dc > 0, conflicts, "Date/ID not mapped"), "Excluded from monthly incident counts; investigate key scope or dates"
    CheckMetric ws, rowNo, "Missing dates", IIf(dc > 0, missingDates, "Not mapped"), "Excluded from monthly coverage"
    CheckMetric ws, rowNo, "Invalid dates", IIf(dc > 0, invalidDates, "Not mapped"), "US month/day/year or ISO; native Excel dates supported"
    CheckMetric ws, rowNo, "Missing/error crime types", IIf(cc > 0, missingCrimes, "Not mapped"), "Shown explicitly in crime-type coverage"
    CheckMetric ws, rowNo, "Missing/error addresses", IIf(ac > 0, missingAddresses, "Not mapped"), "No address text supplied"
    CheckMetric ws, rowNo, "Missing coordinate pairs", IIf(lc > 0 And oc > 0, missingCoords, "Pair not mapped"), "At least one coordinate is blank"
    CheckMetric ws, rowNo, "Invalid coordinate pairs", IIf(lc > 0 And oc > 0, invalidCoords, "Pair not mapped"), "Nonnumeric/error or outside latitude +/-90, longitude +/-180"
    CheckMetric ws, rowNo, "Possible reversed pairs", IIf(lc > 0 And oc > 0, swaps, "Pair not mapped"), "Invalid as supplied but valid if swapped; not corrected automatically"
    CheckMetric ws, rowNo, "Zero-zero coordinate pairs", IIf(lc > 0 And oc > 0, zeroCoords, "Pair not mapped"), "Likely placeholders; individual zero coordinates can be valid"
    rowNo = rowNo + 2: ws.Cells(rowNo, 1).Value2 = "Monthly coverage": rowNo = rowNo + 1
    WriteHeaders ws, rowNo, 1, Array("Month", "Source records", "Unique dated incidents", "Coverage note"): rowNo = rowNo + 1
    If firstMonth = 0 Then
        ws.Cells(rowNo, 1).Value2 = "Date not mapped or no valid dates": rowNo = rowNo + 1
    Else
        span = DateDiff("m", firstMonth, lastMonth) + 1
        If span > 1200 Then Err.Raise vbObjectError + 901, , "Date span exceeds 100 years. Review date values before running monthly coverage."
        For i = 0 To span - 1
            monthKey = Format$(DateAdd("m", i, firstMonth), "yyyy-mm")
            ws.Cells(rowNo, 1).Value2 = monthKey: ws.Cells(rowNo, 2).Value2 = 0
            If months.Exists(monthKey) Then ws.Cells(rowNo, 2).Value2 = months(monthKey)
            ws.Cells(rowNo, 3).Value2 = IIf(ic > 0, 0, "Not mapped")
            ws.Cells(rowNo, 4).Value2 = IIf(ws.Cells(rowNo, 2).Value2 = 0, "No records; possible gap", "Observed records; reporting completeness unverified")
            If i = 0 Or i = span - 1 Then ws.Cells(rowNo, 4).Value2 = ws.Cells(rowNo, 4).Value2 & "; observed edge month"
            months(monthKey) = rowNo: rowNo = rowNo + 1
        Next
        For i = 2 To UBound(a, 1)
            If validDates(i) And ic > 0 Then
                id = SafeText(a(i, ic)): monthKey = Format$(CDate(dates(i)), "yyyy-mm"): signature = KeyPart(monthKey) & KeyPart(id)
                If monthIDs.Exists(signature) Then
                    ws.Cells(months(monthKey), 3).Value2 = ws.Cells(months(monthKey), 3).Value2 + 1: monthIDs.Remove signature
                End If
            End If
        Next
    End If
    rowNo = rowNo + 2: ws.Cells(rowNo, 1).Value2 = "Crime-type coverage (source records; normalized labels)": rowNo = rowNo + 1
    WriteHeaders ws, rowNo, 1, Array("Crime type", "Source records", "Share of all source rows"): rowNo = rowNo + 1
    For Each k In crimes.Keys
        ws.Cells(rowNo, 1).NumberFormat = "@": ws.Cells(rowNo, 1).Value2 = k: ws.Cells(rowNo, 2).Value2 = crimes(k)
        ws.Cells(rowNo, 3).Value2 = crimes(k) / n: ws.Cells(rowNo, 3).NumberFormat = "0.0%": rowNo = rowNo + 1
    Next
    If cc = 0 Then ws.Cells(rowNo, 1).Value2 = "Crime Type not mapped": rowNo = rowNo + 1
    rowNo = rowNo + 2: ws.Cells(rowNo, 1).Value2 = "Field completeness": rowNo = rowNo + 1
    WriteHeaders ws, rowNo, 1, Array("Source column", "Field", "Blank rows", "Error cells", "Blank share"): rowNo = rowNo + 1
    For j = 1 To UBound(a, 2)
        ws.Cells(rowNo, 1).Value2 = j: ws.Cells(rowNo, 2).NumberFormat = "@": ws.Cells(rowNo, 2).Value2 = SafeText(a(1, j))
        ws.Cells(rowNo, 3).Value2 = fieldBlanks(j): ws.Cells(rowNo, 4).Value2 = fieldErrors(j)
        ws.Cells(rowNo, 5).Value2 = fieldBlanks(j) / n: ws.Cells(rowNo, 5).NumberFormat = "0.0%": rowNo = rowNo + 1
    Next
    FinishOutput ws, rowNo - 1, 5: ws.AutoFilterMode = False
    ws.Columns("A:B").ColumnWidth = 32: ws.Columns("C:E").ColumnWidth = 38
    ws.Range("A6:E" & rowNo - 1).WrapText = True: ws.Range("A6:E" & rowNo - 1).Rows.AutoFit
    ws.Names.Add Name:="_CAPExportSummary", RefersTo:="='" & ws.Name & "'!$A$1:$E$23"
    ws.Names.Add Name:="_CAPReportBody", RefersTo:="='" & ws.Name & "'!$A$1:$E$" & rowNo - 1
    Application.EnableEvents = oldEvents
    If showSummary Then MsgBox "Created " & ws.Name & " with dataset checks and coverage. Source data is unchanged.", vbInformation, "CAP Index"
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description: Application.EnableEvents = oldEvents
    If Not source Is Nothing Then Err.Raise errNo, "Dataset Check", description
    MsgBox description, vbExclamation, "CAP Index | Dataset Check"
End Sub

Private Function CheckField(ByVal a As Variant, ByVal label As String, ByVal suggested As String) As Long
    If CheckColumn(a, suggested) = 0 Then suggested = ""
    Dim answer As Variant
    suggested = SavedField(ActiveWorkbook, label, suggested)
    If Len(suggested) > 0 Then
        If CheckColumn(a, suggested) = 0 Then suggested = ""
    End If
    answer = Application.InputBox("Optional " & label & ". Enter its exact unique header, or leave blank. Absent/ambiguous fields are marked Not mapped.", "CAP Index | Dataset Check", suggested, Type:=2)
    If VarType(answer) = vbBoolean Then Err.Raise vbObjectError + 6, , "Operation cancelled."
    CheckField = CheckColumn(a, SafeText(answer))
    If Len(SafeText(answer)) = 0 Or CheckField > 0 Then RememberField ActiveWorkbook, label, SafeText(answer)
End Function

Private Sub CheckMetric(ByVal ws As Worksheet, ByRef rowNo As Long, ByVal label As String, ByVal count As Variant, ByVal explanation As String)
    ws.Cells(rowNo, 1).Value2 = label: ws.Cells(rowNo, 2).Value2 = count: ws.Cells(rowNo, 3).Value2 = explanation: rowNo = rowNo + 1
End Sub

Private Function CheckColumn(ByVal a As Variant, ByVal header As String) As Long
    Dim j As Long, matches As Long
    If Len(header) = 0 Then Exit Function
    For j = 1 To UBound(a, 2)
        If StrComp(SafeText(a(1, j)), header, vbTextCompare) = 0 Then
            matches = matches + 1: CheckColumn = j
        End If
    Next
    If matches <> 1 Then CheckColumn = 0
End Function

Private Function CheckCellSignature(ByVal value As Variant) As String
    If IsNull(value) Then
        CheckCellSignature = "NULL"
    Else
        CheckCellSignature = CStr(VarType(value)) & ":" & CStr(value)
    End If
End Function
