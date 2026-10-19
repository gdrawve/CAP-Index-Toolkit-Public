Attribute VB_Name = "modSingleSite"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
Option Explicit

Public Function CoordinatePair(ByVal latValue As Variant, ByVal lonValue As Variant, ByRef lat As Double, ByRef lon As Double) As Boolean
    If IsError(latValue) Or IsError(lonValue) Then Exit Function
    If Len(SafeText(latValue)) = 0 Or Len(SafeText(lonValue)) = 0 Then Exit Function
    If Not IsNumeric(latValue) Or Not IsNumeric(lonValue) Then Exit Function
    lat = CDbl(latValue): lon = CDbl(lonValue)
    If Abs(lat) > 90 Or Abs(lon) > 180 Or (lat = 0 And lon = 0) Then Exit Function
    CoordinatePair = True
End Function

Public Function DistanceMiles(ByVal siteLat As Double, ByVal siteLon As Double, ByVal lat As Double, ByVal lon As Double) As Double
    Dim pi As Double, x As Double, a As Double, radians As Double
    pi = 4 * Atn(1): radians = pi / 180
    a = Sin((lat - siteLat) * radians / 2) ^ 2 + Cos(siteLat * radians) * Cos(lat * radians) * Sin((lon - siteLon) * radians / 2) ^ 2
    a = Application.Min(1, Application.Max(0, a))
    If a >= 1 Then DistanceMiles = pi * 3958.7613: Exit Function
    x = Sqr(a / (1 - a)): DistanceMiles = 2 * 3958.7613 * Atn(x)
End Function

Public Function DistanceBand(ByVal distance As Double) As Long
    If distance <= 0.25 + 0.000000001 Then
        DistanceBand = 1
    ElseIf distance <= 0.5 + 0.000000001 Then
        DistanceBand = 2
    ElseIf distance <= 1 + 0.000000001 Then
        DistanceBand = 3
    End If
End Function

Public Function SingleSiteCache(ByVal ws As Worksheet) As Range
    Set SingleSiteCache = CAPNamedRange(ws, "_CAPSiteData")
End Function

Public Sub BuildSingleSite(Optional ByVal source As Range, Optional ByVal siteLat As Double = 999, Optional ByVal siteLon As Double = 999, _
    Optional ByVal dateHeader As String = "Date", Optional ByVal crimeHeader As String = "CrimeType", Optional ByVal incidentHeader As String = "IncidentID", _
    Optional ByVal latitudeHeader As String = "Latitude", Optional ByVal longitudeHeader As String = "Longitude", Optional ByVal measure As String = "", _
    Optional ByVal firstDate As Double = 0, Optional ByVal lastDate As Double = 0, Optional ByVal showSummary As Boolean = True)
    Dim r As Range, a As Variant, ctx As Object, data As Variant, output() As Variant, ws As Worksheet, cache As Worksheet, nearby As Worksheet
    Dim lc As Long, oc As Long, i As Long, j As Long, lat As Double, lon As Double, status As String, ids As Object, id As String, signature As String
    Dim review As Long, flag As Long, matched As Long, synthetic As Boolean, oldEvents As Boolean, errNo As Long, description As String, named As Name
    Dim choice As Variant
    If source Is Nothing Then OpenToolSetup "capSingleSite": Exit Sub
    oldEvents = Application.EnableEvents
    On Error GoTo failed
    Set r = source
    If r Is Nothing Then
        MsgBox "Single Site: select headers/data containing Latitude, Longitude, Date, Crime Type and optional IncidentID. Enter site coordinates next. Results include a radius report and nearby source records.", vbInformation, "CAP Index"
        Set r = PickSource("Single Site inputs: Latitude, Longitude, Date, Crime Type and optional IncidentID. Distances are approximate straight-line miles; synthetic points are excluded.")
    End If
    If r Is Nothing Then Exit Sub
    ValidateOutputSource r: a = ScopedValues(r)
    If source Is Nothing Then
        lc = AskColumn(a, "Incident Latitude (decimal degrees).", "Latitude", False, "LATITUDE")
        oc = AskColumn(a, "Incident Longitude (decimal degrees).", "Longitude", False, "LONGITUDE")
        i = AskColumn(a, "Optional occurrence DATE or timestamp. Leave blank for an undated dataset.", IIf(OptionalHeader(a, "Date") > 0, "Date", ""), True, "DATE")
        dateHeader = "": If i > 0 Then dateHeader = SafeText(a(1, i))
        crimeHeader = SafeText(a(1, AskColumn(a, "Crime type or offense description.", "CrimeType", False, "CRIME")))
        i = AskColumn(a, "Optional stable IncidentID. Missing/conflicting IDs are excluded in incident mode.", "", True, "INCIDENTID")
        incidentHeader = "": If i > 0 Then incidentHeader = SafeText(a(1, i))
        If siteLat = 999 Then siteLat = AskSiteCoordinate("Site latitude (-90 to 90)", 0)
        If siteLon = 999 Then siteLon = AskSiteCoordinate("Site longitude (-180 to 180)", 0)
        choice = Application.InputBox("Count measure: 1 = Unique incidents; 2 = Source/offense records.", "CAP Index | Single Site", IIf(Len(incidentHeader) > 0, "1", "2"), Type:=1)
        If VarType(choice) = vbBoolean Then Exit Sub
        If choice <> 1 And choice <> 2 Then Err.Raise vbObjectError + 960, , "Choose measure 1 or 2."
        measure = IIf(choice = 1, "Unique incidents", "Source records")
    Else
        lc = HeaderColumn(a, latitudeHeader): oc = HeaderColumn(a, longitudeHeader)
    End If
    If Not CoordinatePair(siteLat, siteLon, lat, lon) Then Err.Raise vbObjectError + 960, , "Enter valid site coordinates; 0,0 is treated as a placeholder."
    If Len(dateHeader) = 0 Then
        Set ctx = UndatedSiteInput(r, crimeHeader, incidentHeader, measure)
    Else
        Set ctx = ReadReportInput(r, dateHeader, crimeHeader, incidentHeader, "", "All crime types", measure, firstDate, lastDate, False)
    End If
    If source Is Nothing And Len(dateHeader) > 0 Then
        ctx("Start") = AskReportDate("Start date (inclusive).", ctx("SourceStart"))
        ctx("End") = AskReportDate("End date (inclusive).", ctx("SourceEnd"))
        If ctx("Start") > ctx("End") Then Err.Raise vbObjectError + 960, , "Start must precede end."
    End If
    data = ctx("Data"): ReDim output(1 To UBound(data, 1), 1 To 10)
    Set ids = CreateObject("Scripting.Dictionary")
    review = OptionalHeader(a, "Review Flag"): If review = 0 Then review = OptionalHeader(a, "Review Status")
    flag = OptionalHeader(a, "CAP_SyntheticFlag"): matched = OptionalHeader(a, "Match Status")
    For i = 2 To UBound(data, 1)
        For j = 1 To 6: output(i, j) = data(i, j): Next
        output(i, 8) = CoordinatePair(a(i, lc), a(i, oc), lat, lon)
        output(i, 9) = "Supplied coordinates; accuracy unverified"
        If matched > 0 Then
            If Len(SafeText(a(i, matched))) > 0 Then output(i, 9) = "Geocoding status: " & SafeText(a(i, matched)) & "; accuracy unverified"
        End If
        synthetic = False
        If review > 0 Then synthetic = InStr(1, SafeText(a(i, review)), "Synthetic", vbTextCompare) > 0
        If flag > 0 Then synthetic = synthetic Or (NormalText(a(i, flag)) = "TRUE" Or NormalText(a(i, flag)) = "YES")
        If synthetic Then output(i, 8) = False: output(i, 9) = "Synthetic location excluded"
        signature = "INVALID POSITION"
        If output(i, 8) Then
            output(i, 7) = DistanceMiles(siteLat, siteLon, lat, lon)
            output(i, 10) = CStr(lat) & "|" & CStr(lon)
            signature = output(i, 10)
        End If
        id = output(i, 3)
        If Len(id) > 0 Then
            If Not ids.Exists(id) Then
                ids.Add id, signature
            ElseIf ids(id) <> signature Then
                ids(id) = "CONFLICT"
            End If
        End If
        If i Mod 1000 = 0 Then Application.StatusBar = "CAP Index: preparing Single Site " & i - 1 & " / " & UBound(data, 1) - 1: DoEvents
    Next
    For i = 2 To UBound(data, 1)
        id = output(i, 3)
        If Len(id) > 0 Then output(i, 5) = output(i, 5) And ids(id) <> "CONFLICT" And output(i, 8)
    Next
    Application.EnableEvents = False
    Set ws = NewOutput(r.Parent.Parent, "Single_Site", r, "")
    AnnotateAnalysisScope ws, r
    Set nearby = NewOutput(r.Parent.Parent, "Site_Nearby", r, "Nearby source/offense rows, regardless of summary count measure. Synthetic and invalid locations excluded. No source data changed.")
    AnnotateAnalysisScope nearby, r
    Set cache = r.Parent.Parent.Worksheets.Add(After:=r.Parent.Parent.Worksheets(r.Parent.Parent.Worksheets.Count))
    cache.Name = "_CAPS_" & ws.Name
    cache.Range("A1").Resize(UBound(output, 1), 10).NumberFormat = "@": cache.Range("A1").Resize(UBound(output, 1), 10).Value2 = output
    cache.Range("T1").Resize(UBound(a, 1), UBound(a, 2)).NumberFormat = "@": cache.Range("T1").Resize(UBound(a, 1), UBound(a, 2)).Value2 = a
    CopySourceFormats r, cache.Range("T1").Resize(UBound(a, 1), UBound(a, 2))
    cache.Range("M1").Value2 = ctx("Start"): cache.Range("M2").Value2 = ctx("End"): cache.Range("M3").Value2 = ctx("Measure")
    cache.Range("M4").Value2 = siteLat: cache.Range("M5").Value2 = siteLon
    cache.Range("M6").Value2 = Len(dateHeader) > 0
    Set named = ws.Names.Add(Name:="_CAPSiteData", RefersTo:="='" & cache.Name & "'!" & cache.Range("A1").Resize(UBound(output, 1), 10).Address): named.Visible = False
    Set named = ws.Names.Add(Name:="_CAPSiteRaw", RefersTo:="='" & cache.Name & "'!" & cache.Range("T1").Resize(UBound(a, 1), UBound(a, 2)).Address): named.Visible = False
    ws.Names.Add Name:="_CAPSiteNearby", RefersTo:="='" & nearby.Name & "'!$A$1"
    BuildSiteControls ws, output, ctx, siteLat, siteLon
    ws.Activate: cache.Visible = xlSheetVeryHidden
    RefreshSingleSite ws, True
    Application.EnableEvents = oldEvents: Application.StatusBar = False
    If showSummary Then MsgBox "Created " & ws.Name & " and " & nearby.Name & ". See excluded rows and location-accuracy notes in the report. Hidden snapshot retains source fields; share exports when appropriate.", vbInformation, "CAP Index | Single Site"
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description: Application.EnableEvents = oldEvents: Application.StatusBar = False
    If Not source Is Nothing Then Err.Raise errNo, "Single Site", description
    MsgBox description, vbExclamation, "CAP Index | Single Site"
End Sub

Private Function AskSiteCoordinate(ByVal prompt As String, ByVal suggested As Double) As Double
    Dim answer As Variant
    answer = Application.InputBox(prompt, "CAP Index | Single Site", suggested, Type:=1)
    If VarType(answer) = vbBoolean Then Err.Raise vbObjectError + 6, , "Operation cancelled."
    AskSiteCoordinate = CDbl(answer)
End Function

Private Sub BuildSiteControls(ByVal ws As Worksheet, ByVal data As Variant, ByVal ctx As Object, ByVal siteLat As Double, ByVal siteLon As Double)
    Dim categories As Object, i As Long, key As Variant, cache As Range
    Set cache = SingleSiteCache(ws): Set categories = CreateObject("Scripting.Dictionary"): categories.Add "All crime types", True
    For i = 2 To UBound(data, 1): If Not categories.Exists(data(i, 2)) Then categories.Add data(i, 2), True
    Next
    i = 0: cache.Parent.Columns("O").NumberFormat = "@"
    For Each key In categories.Keys: i = i + 1: cache.Parent.Cells(i, 15).Value2 = key: Next
    ws.Names.Add Name:="_CAPSiteTypes", RefersTo:="='" & cache.Parent.Name & "'!$O$1:$O$" & i
    ws.Range("A1:H1").Merge: ws.Range("A1").Font.Size = 20: ws.Range("A1:H1").Interior.Color = RGB(26, 55, 82): ws.Range("A1").Font.Color = vbWhite: ws.Rows(1).RowHeight = 34
    ws.Range("A2:H2").Merge: ws.Range("A2").WrapText = True: ws.Rows(2).RowHeight = 34
    ws.Range("A3").Value2 = "Crime type / blue field": ws.Range("B3:D3").Merge: ws.Range("B3").Value2 = "All crime types"
    ws.Range("B3").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=_CAPSiteTypes"
    ws.Range("B3").Validation.ShowError = True: ws.Range("B3").Validation.InCellDropdown = True
    ws.Range("E3:H3").Merge: ws.Range("E3").Value2 = ctx("Measure")
    ws.Range("A4").Value2 = "Site latitude": ws.Range("B4").Value2 = siteLat: ws.Range("C4").Value2 = "Site longitude": ws.Range("D4").Value2 = siteLon
    ws.Range("E4").Value2 = "From": ws.Range("F4").Value2 = ctx("Start"): ws.Range("G4").Value2 = "Through": ws.Range("H4").Value2 = ctx("End")
    If ctx("Start") = 0 Then
        ws.Range("F4").Value2 = "Not mapped": ws.Range("H4").Value2 = "Not mapped"
    ElseIf ws.Parent.Date1904 Then
        ws.Range("F4").Value2 = ctx("Start") - 1462: ws.Range("H4").Value2 = ctx("End") - 1462
    End If
    ws.Range("F4,H4").NumberFormat = "mm/dd/yyyy": ws.Range("B4,D4").NumberFormat = "0.000000"
    ws.Range("B3:D3,B4,D4,F4,H4").Interior.Color = RGB(224, 239, 250)
    If ctx("Start") = 0 Then ws.Range("F4,H4").Interior.Color = RGB(242, 246, 250)
    ws.Range("A5:H5").Merge: ws.Range("A5").Value2 = "Approximate spherical straight-line miles; not driving distance. Count on unrounded distance. Coordinate error and Earth-model approximation affect boundary cases. Synthetic locations excluded. Rebuild to reload source."
    ws.Range("A3:H5").WrapText = True: ws.Rows(3).RowHeight = 36: ws.Rows(5).RowHeight = 42
    ws.Columns("A").ColumnWidth = 25: ws.Columns("B:H").ColumnWidth = 18
    InitializeCAPEvents
End Sub

Public Sub RefreshSingleSite(ByVal ws As Worksheet, Optional ByVal propagateErrors As Boolean = False)
    Dim cache As Range, raw As Range, data As Variant, original As Variant, nearby As Worksheet, counts(1 To 3) As Long
    Dim categories As Object, seen As Object, seenTypes As Object, key As Variant, id As String, signature As String, crime As String, measure As String
    Dim i As Long, j As Long, b As Long, firstDate As Double, lastDate As Double, siteLat As Double, siteLon As Double, position As Variant
    Dim eligible As Long, selected As Long, invalidDate As Long, invalidPosition As Long, invalidID As Long, outside As Long, rowNo As Long
    Dim distance As Double, records As Collection, result() As Variant, values As Variant, oldEvents As Boolean, errNo As Long, description As String, allowed As Boolean, hasDates As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo failed
    Set cache = SingleSiteCache(ws): If cache Is Nothing Then Err.Raise vbObjectError + 961, , "Select a Single Site report."
    If Not CoordinatePair(ws.Range("B4").Value2, ws.Range("D4").Value2, siteLat, siteLon) Then Err.Raise vbObjectError + 961, , "Invalid site coordinates."
    hasDates = cache.Parent.Range("M6").Value2
    If hasDates Then
        firstDate = SiteReportDate(ws.Range("F4").Value2, ws.Parent.Date1904): lastDate = SiteReportDate(ws.Range("H4").Value2, ws.Parent.Date1904)
    Else
        If ws.Range("F4").Value2 <> "Not mapped" Or ws.Range("H4").Value2 <> "Not mapped" Then Err.Raise vbObjectError + 961, , "Dates are not mapped. Rebuild Single Site with a date field to filter by dates."
    End If
    If firstDate > lastDate Then Err.Raise vbObjectError + 961, , "Start must precede end."
    crime = SafeText(ws.Range("B3").Value2): measure = cache.Parent.Range("M3").Value2: data = cache.Value2
    allowed = crime = "All crime types"
    For i = 2 To UBound(data, 1): If crime = data(i, 2) Then allowed = True
    Next
    If Not allowed Then Err.Raise vbObjectError + 961, , "Choose a crime from the dropdown."
    Set seen = CreateObject("Scripting.Dictionary"): Set seenTypes = CreateObject("Scripting.Dictionary"): Set categories = CreateObject("Scripting.Dictionary"): Set records = New Collection
    For i = 2 To UBound(data, 1)
        If crime = "All crime types" Or crime = data(i, 2) Then
            selected = selected + 1
            If hasDates And IsEmpty(data(i, 1)) Then
                invalidDate = invalidDate + 1
            ElseIf Not hasDates Or (data(i, 1) >= firstDate And data(i, 1) <= lastDate) Then
                If Not data(i, 8) Then
                    invalidPosition = invalidPosition + 1
                ElseIf measure = "Unique incidents" And Not data(i, 5) Then
                    invalidID = invalidID + 1
                Else
                    position = Split(data(i, 10), "|")
                    distance = DistanceMiles(siteLat, siteLon, CDbl(position(0)), CDbl(position(1)))
                    b = DistanceBand(distance): data(i, 7) = distance
                    If b = 0 Then
                        outside = outside + 1
                    Else
                        records.Add i
                        id = data(i, 3): signature = id
                        If measure = "Source records" Then signature = "ROW:" & i
                        If Not seen.Exists(signature) Then
                            seen.Add signature, True: counts(b) = counts(b) + 1: eligible = eligible + 1
                        End If
                        key = data(i, 2)
                        signature = KeyPart(CStr(key)) & KeyPart(signature)
                        If Not seenTypes.Exists(signature) Then
                            seenTypes.Add signature, True
                            If Not categories.Exists(key) Then categories.Add key, Array(0&, 0&, 0&)
                            values = categories(key): values(b - 1) = values(b - 1) + 1: categories(key) = values
                        End If
                    End If
                End If
            End If
        End If
        If i Mod 1000 = 0 Then Application.StatusBar = "CAP Index: Single Site distances " & i - 1 & " / " & UBound(data, 1) - 1: DoEvents
    Next
    Application.EnableEvents = False
    ws.Range("A6:H" & Application.Max(14, ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1)).ClearContents
    For i = ws.ChartObjects.Count To 1 Step -1
        If Left$(ws.ChartObjects(i).Name, 10) = "CAPReport_" Then ws.ChartObjects(i).Delete
    Next
    WriteHeaders ws, 6, 1, Array("Radius (miles)", "Within radius", "Separate band", "Band count", "Count measure")
    For i = 1 To 3
        ws.Cells(i + 6, 1).Value2 = Choose(i, 0.25, 0.5, 1)
        ws.Cells(i + 6, 2).Value2 = counts(1): If i > 1 Then ws.Cells(i + 6, 2).Value2 = ws.Cells(i + 6, 2).Value2 + counts(2)
        If i = 3 Then ws.Cells(i + 6, 2).Value2 = ws.Cells(i + 6, 2).Value2 + counts(3)
        ws.Cells(i + 6, 3).Value2 = Choose(i, "0-0.25 miles", ">0.25-0.5 miles", ">0.5-1 mile")
        ws.Cells(i + 6, 4).Value2 = counts(i): ws.Cells(i + 6, 5).Value2 = measure
    Next
    ws.Range("A10:H10").Merge: ws.Range("A10").Value2 = "Selected source rows=" & selected & "; invalid dates=" & invalidDate & "; invalid/synthetic positions in period=" & invalidPosition & "; invalid/conflicting ID rows in period=" & invalidID & "; valid rows beyond 1 mile=" & outside
    ws.Range("A11:H11").Merge: ws.Range("A11").Value2 = "Nearby sheet retains " & records.Count & " source/offense rows. Summary counts " & eligible & " " & LCase$(measure) & ". Crime groups can overlap for multi-offense incidents; cumulative radii are nested. Source coverage is unverified."
    ws.Range("A10:H11").WrapText = True: ws.Rows("10:11").RowHeight = 38
    WriteHeaders ws, 13, 1, Array("Crime type", "Within 0.25 mi", "Within 0.5 mi", "Within 1 mi", "0-0.25 mi", ">0.25-0.5 mi", ">0.5-1 mi")
    rowNo = 14
    For Each key In categories.Keys
        values = categories(key): ws.Cells(rowNo, 1).NumberFormat = "@": ws.Cells(rowNo, 1).Value2 = key
        ws.Cells(rowNo, 2).Value2 = values(0): ws.Cells(rowNo, 3).Value2 = values(0) + values(1): ws.Cells(rowNo, 4).Value2 = values(0) + values(1) + values(2)
        For j = 0 To 2: ws.Cells(rowNo, j + 5).Value2 = values(j): Next
        rowNo = rowNo + 1
    Next
    If categories.Count = 0 Then ws.Range("A14").Value2 = "No eligible activity within 1 mile."
    SiteHeaderStyle ws.Range("A6:H6"): SiteHeaderStyle ws.Range("A13:H13")
    ws.Range("A6:H" & Application.Max(14, rowNo - 1)).WrapText = True
    ws.Range("A6:H9").Rows.AutoFit: ws.Range("A13:H" & Application.Max(14, rowNo - 1)).Rows.AutoFit
    Set raw = CAPNamedRange(ws, "_CAPSiteRaw"): original = raw.Value2
    Set nearby = CAPNamedRange(ws, "_CAPSiteNearby").Parent
    For i = nearby.ListObjects.Count To 1 Step -1: nearby.ListObjects(i).Unlist: Next
    nearby.Range("A6").Resize(Application.Max(2, nearby.UsedRange.Rows.Count), UBound(original, 2) + 3).Clear
    ReDim result(1 To records.Count + 1, 1 To UBound(original, 2) + 3)
    For j = 1 To UBound(original, 2): result(1, j) = original(1, j): Next
    result(1, UBound(original, 2) + 1) = "CAP_DistanceMiles": result(1, UBound(original, 2) + 2) = "CAP_DistanceBand": result(1, UBound(original, 2) + 3) = "CAP_LocationPrecision"
    rowNo = 1
    For Each key In records
        rowNo = rowNo + 1
        For j = 1 To UBound(original, 2): result(rowNo, j) = original(key, j): Next
        result(rowNo, UBound(original, 2) + 1) = data(key, 7)
        result(rowNo, UBound(original, 2) + 2) = Choose(DistanceBand(data(key, 7)), "0-0.25 miles", ">0.25-0.5 miles", ">0.5-1 mile")
        result(rowNo, UBound(original, 2) + 3) = data(key, 9)
    Next
    nearby.Range("A6").Resize(UBound(result, 1), UBound(result, 2)).NumberFormat = "@": nearby.Range("A6").Resize(UBound(result, 1), UBound(result, 2)).Value2 = result
    If records.Count > 0 Then
        For j = 1 To UBound(original, 2): nearby.Cells(7, j).Resize(records.Count).NumberFormat = raw.Cells(2, j).NumberFormat: Next
        nearby.Cells(7, UBound(original, 2) + 1).Resize(records.Count).NumberFormat = "0.0000"
    End If
    nearby.ListObjects.Add(xlSrcRange, nearby.Range("A6").Resize(UBound(result, 1), UBound(result, 2)), , xlYes).TableStyle = "TableStyleMedium2"
    nearby.Range("A6").Resize(UBound(result, 1), UBound(result, 2)).Columns.AutoFit
    description = "Dates not mapped; all selected rows"
    If hasDates Then description = Format$(CDate(firstDate), "mm/dd/yyyy") & " - " & Format$(CDate(lastDate), "mm/dd/yyyy")
    nearby.Range("A3").Value2 = "Site=" & siteLat & ", " & siteLon & "; Crime=" & crime & "; Period=" & description & "; nearby source records=" & records.Count
    ws.Names.Add Name:="_CAPExportSummary", RefersTo:="='" & ws.Name & "'!$A$1:$H$" & Application.Min(30, Application.Max(14, 13 + categories.Count))
    ws.Names.Add Name:="_CAPReportBody", RefersTo:="='" & ws.Name & "'!$A$1:$H$" & Application.Max(14, 13 + categories.Count)
    CreateReportChart ws, "CAPReport_SingleSite", "Within radius | " & crime & " | " & measure, ws.Range("A7:A9"), ws.Range("B7:B9"), "J6", xlColumnClustered
    cache.Parent.Range("R1").Value2 = crime: cache.Parent.Range("R2").Value2 = siteLat: cache.Parent.Range("R3").Value2 = siteLon
    cache.Parent.Range("R4").Value2 = firstDate: cache.Parent.Range("R5").Value2 = lastDate
    StampReportStatus ws: ws.Activate: ActiveWindow.DisplayGridlines = False: ws.Range("B3").Select
    Application.EnableEvents = oldEvents: Application.StatusBar = False
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description: Application.EnableEvents = oldEvents: Application.StatusBar = False
    If propagateErrors Then Err.Raise errNo, "Single Site refresh", description
    MsgBox description, vbExclamation, "CAP Index | Single Site"
End Sub

Public Function SiteReportDate(ByVal value As Variant, ByVal date1904 As Boolean) As Double
    Dim parsed As Date, timed As Boolean, status As String
    If Not ParseTimestamp(value, date1904, parsed, timed, status) Then Err.Raise vbObjectError + 961, , "Enter valid start/end dates."
    If timed Then Err.Raise vbObjectError + 961, , "Use date-only boundaries."
    SiteReportDate = CDbl(DateValue(parsed))
End Function

Private Sub SiteHeaderStyle(ByVal r As Range)
    r.Interior.Color = RGB(26, 55, 82): r.Font.Color = vbWhite: r.Font.Bold = True
End Sub

Private Function UndatedSiteInput(ByVal r As Range, ByVal crimeHeader As String, ByVal incidentHeader As String, ByVal measure As String) As Object
    Dim ctx As Object, a As Variant, data() As Variant, i As Long, cc As Long, ic As Long
    a = ScopedValues(r): cc = HeaderColumn(a, crimeHeader): ic = HeaderColumn(a, incidentHeader)
    If Len(measure) = 0 Then measure = IIf(ic > 0, "Unique incidents", "Source records")
    If measure <> "Unique incidents" And measure <> "Source records" Then Err.Raise vbObjectError + 960, , "Invalid count measure."
    If measure = "Unique incidents" And ic = 0 Then Err.Raise vbObjectError + 960, , "Map an IncidentID field or choose Source records."
    ReDim data(1 To UBound(a, 1), 1 To 6)
    For i = 2 To UBound(a, 1)
        data(i, 2) = CrimeCategory(a(i, cc)): data(i, 3) = "": data(i, 5) = False: data(i, 6) = ScopeSourceRow(r, i)
        If ic > 0 Then
            If Not IsError(a(i, ic)) Then data(i, 3) = SafeText(a(i, ic)): data(i, 5) = Len(data(i, 3)) > 0
        End If
    Next
    Set ctx = CreateObject("Scripting.Dictionary")
    ctx.Add "Source", r: ctx.Add "Data", data: ctx.Add "Start", 0#: ctx.Add "End", 0#: ctx.Add "Measure", measure
    Set UndatedSiteInput = ctx
End Function

Public Function SitePeriodLabel(ByVal ws As Worksheet) As String
    If Not SingleSiteCache(ws).Parent.Range("M6").Value2 Then
        SitePeriodLabel = "Dates not mapped; all selected rows"
    Else
        SitePeriodLabel = Format$(CDate(SiteReportDate(ws.Range("F4").Value2, ws.Parent.Date1904)), "mm/dd/yyyy") & " - " & Format$(CDate(SiteReportDate(ws.Range("H4").Value2, ws.Parent.Date1904)), "mm/dd/yyyy")
    End If
End Function
