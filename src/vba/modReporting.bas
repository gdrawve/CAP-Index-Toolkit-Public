Attribute VB_Name = "modReporting"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit

Public Function ReportChange(ByVal previous As Double, ByVal current As Double) As Variant
    If previous = 0 Then ReportChange = "N/A - zero baseline" Else ReportChange = (current - previous) / previous
End Function

Public Sub BuildTrendReport(Optional ByVal source As Range, Optional ByVal dateHeader As String = "Date", Optional ByVal crimeHeader As String = "CrimeType", Optional ByVal incidentHeader As String = "IncidentID", Optional ByVal crimeFilter As String = "All crime types", Optional ByVal measure As String = "", Optional ByVal unit As String = "Months", Optional ByVal firstDate As Double = 0, Optional ByVal lastDate As Double = 0, Optional ByVal showSummary As Boolean = True, Optional ByVal reportTarget As Worksheet, Optional ByVal reportContext As Object)
    Dim ctx As Object, rows As Collection, data As Variant, output() As Variant, ws As Worksheet, r As Range
    Dim origin As Date, boundary As Date, nextBoundary As Date, p As Long, i As Long, index As Long, row As Variant, offset As Double
    Dim choice As Variant, complete() As Boolean, counts() As Long, days As Long, description As String, errNo As Long
    On Error GoTo failed
    If source Is Nothing And reportContext Is Nothing Then OpenToolSetup "capTrends": Exit Sub
    If reportContext Is Nothing Then
    Set ctx = ReadReportInput(source, dateHeader, crimeHeader, incidentHeader, "", crimeFilter, measure, firstDate, lastDate, False)
    Else
        Set ctx = reportContext
    End If
    If ctx Is Nothing Then Exit Sub
    If source Is Nothing Then
        choice = Application.InputBox("Trend interval: 1 = Months; 2 = Years.", "CAP Index | Trends", "1", Type:=1)
        If VarType(choice) = vbBoolean Then Exit Sub
        If choice <> 1 And choice <> 2 Then Err.Raise vbObjectError + 820, , "Choose Months or Years."
        unit = IIf(choice = 1, "Months", "Years")
    End If
    If unit <> "Months" And unit <> "Years" Then Err.Raise vbObjectError + 820, , "Invalid trend interval."
    If unit = "Months" Then origin = DateSerial(Year(CDate(ctx("Start"))), Month(CDate(ctx("Start"))), 1) Else origin = DateSerial(Year(CDate(ctx("Start"))), 1, 1)
    p = DateDiff(IIf(unit = "Months", "m", "yyyy"), origin, CDate(ctx("End"))) + 1
    If p > 1000 Then Err.Raise vbObjectError + 821, , "More than 1000 periods; choose Years or a shorter range."
    ReDim counts(1 To p): ReDim output(1 To p, 1 To 6): ReDim complete(1 To p)
    Set rows = ReportRows(ctx): data = ctx("Data")
    For Each row In rows
        index = DateDiff(IIf(unit = "Months", "m", "yyyy"), origin, CDate(data(row, 1))) + 1: counts(index) = counts(index) + 1
    Next
    Set r = ctx("Source")
    If reportTarget Is Nothing Then
    Set r = ctx("Source"): Set ws = NewOutput(r.Parent.Parent, IIf(unit = "Months", "Monthly_Trends", "Annual_Trends"), r, "")
    AnnotateAnalysisScope ws, r
    Else
        Set ws = reportTarget: ClearReportingResults ws
    End If
    If ws.Parent.Date1904 Then offset = 1462
    For i = 1 To p
        boundary = DateAdd(IIf(unit = "Months", "m", "yyyy"), i - 1, origin)
        nextBoundary = DateAdd(IIf(unit = "Months", "m", "yyyy"), 1, boundary)
        days = Application.Min(CDbl(nextBoundary) - 1, ctx("End")) - Application.Max(CDbl(boundary), ctx("Start")) + 1
        complete(i) = CDbl(boundary) >= ctx("Start") And CDbl(nextBoundary) - 1 <= ctx("End") And CDbl(boundary) >= ctx("SourceStart") And CDbl(nextBoundary) - 1 <= ctx("SourceEnd")
        output(i, 1) = CDbl(boundary) - offset: output(i, 2) = counts(i): output(i, 3) = days: output(i, 4) = counts(i) / days
        output(i, 5) = IIf(complete(i), "Calendar bin within observed span; coverage unverified", "Partial requested bin / observed edge")
        output(i, 6) = "N/A - first period"
        If i > 1 Then
            If complete(i) And complete(i - 1) Then output(i, 6) = ReportChange(counts(i - 1), counts(i)) Else output(i, 6) = "N/A - partial/edge comparison"
        End If
    Next
    WriteHeaders ws, 6, 1, Array("Period", "Event count", "Requested days", "Events/day", "Period coverage", "Change vs previous")
    ws.Range("A7").Resize(p, 6).NumberFormat = "@": ws.Range("A7").Resize(p, 6).Value2 = output
    ws.Range("A7:A" & p + 6).NumberFormat = IIf(unit = "Months", "mmm yyyy", "yyyy")
    ws.Range("B7:C" & p + 6).NumberFormat = "0": ws.Range("D7:D" & p + 6).NumberFormat = "0.000": ws.Range("F7:F" & p + 6).NumberFormat = "0.0%"
    FinishOutput ws, p + 6, 6
    ConfigureReport ws, ctx, ws.Range("A1:H" & Application.Min(p + 6, 30))
    CreateReportChart ws, "CAPReport_Trend", "Recorded crime by " & LCase$(unit) & " | " & ctx("Crime") & " | " & ctx("Measure"), ws.Range("A7:A" & p + 6), ws.Range("B7:B" & p + 6), "J6", xlLineMarkers
    AttachReportingFilter ws, ctx, "Trend", unit
    If showSummary Then MsgBox "Output: " & ws.Name & ". Missing/invalid date rows: " & ctx("MissingDates") & ". Trend report created. Edge/partial comparisons are flagged; percentage changes use comparable complete calendar bins only.", vbInformation, "CAP Index"
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    If Not source Is Nothing Then Err.Raise errNo, "Trends", description
    MsgBox description, vbExclamation, "CAP Index"
End Sub

Public Sub BuildPeriodComparison(Optional ByVal source As Range, Optional ByVal dateHeader As String = "Date", Optional ByVal crimeHeader As String = "CrimeType", Optional ByVal incidentHeader As String = "IncidentID", Optional ByVal locationHeader As String = "", Optional ByVal crimeFilter As String = "All crime types", Optional ByVal measure As String = "", Optional ByVal aStart As Double = 0, Optional ByVal aEnd As Double = 0, Optional ByVal bStart As Double = 0, Optional ByVal bEnd As Double = 0, Optional ByVal showSummary As Boolean = True, Optional ByVal reportTarget As Worksheet, Optional ByVal reportContext As Object)
    Dim ctx As Object, rowsA As Collection, rowsB As Collection, data As Variant, groups As Object, key As Variant, row As Variant, counts As Variant
    Dim ws As Worksheet, r As Range, output() As Variant, i As Long, daysA As Long, daysB As Long, safe As Boolean
    Dim errNo As Long, description As String
    On Error GoTo failed
    If source Is Nothing And reportContext Is Nothing Then OpenToolSetup "capCompare": Exit Sub
    If reportContext Is Nothing Then
    Set ctx = ReadReportInput(source, dateHeader, crimeHeader, incidentHeader, locationHeader, crimeFilter, measure, bStart, bEnd, False)
    Else
        Set ctx = reportContext
    End If
    If ctx Is Nothing Then Exit Sub
    bStart = ctx("Start"): bEnd = ctx("End")
    If aStart = 0 Then aEnd = bStart - 1: aStart = aEnd - (bEnd - bStart)
    If source Is Nothing Then
        aStart = AskReportDate("Comparison A start date (reporting dates already selected define period B).", aStart)
        aEnd = AskReportDate("Comparison A end date (inclusive).", aEnd)
    End If
    If aStart > aEnd Or aEnd >= bStart Then Err.Raise vbObjectError + 822, , "Period A must end before period B begins."
    Set rowsA = ReportRows(ctx, aStart, aEnd): Set rowsB = ReportRows(ctx, bStart, bEnd): data = ctx("Data")
    daysA = aEnd - aStart + 1: daysB = bEnd - bStart + 1
    safe = aStart >= ctx("SourceStart") And bEnd <= ctx("SourceEnd")
    Set groups = CreateObject("Scripting.Dictionary")
    Dim countsA As Object, countsB As Object
    Set countsA = PeriodGroupCounts(ctx, aStart, aEnd): Set countsB = PeriodGroupCounts(ctx, bStart, bEnd)
    For Each key In countsA.Keys: groups.Add key, Array(countsA(key), 0&): Next
    For Each key In countsB.Keys
        If Not groups.Exists(key) Then groups.Add key, Array(0&, 0&)
        counts = groups(key): counts(1) = countsB(key): groups(key) = counts
    Next
    Set r = ctx("Source")
    If reportTarget Is Nothing Then
    Set r = ctx("Source"): Set ws = NewOutput(r.Parent.Parent, "Period_Comparison", r, "")
    AnnotateAnalysisScope ws, r
    Else
        Set ws = reportTarget: ClearReportingResults ws
    End If
    WriteHeaders ws, 6, 1, Array(IIf(ctx("LocationMapped"), "Location", "Crime type"), "B event count", "A event count", "Count difference", "A events/day", "B events/day", "Count change", "Daily-rate change")
    If groups.Count > 0 Then
        ReDim output(1 To groups.Count, 1 To 8)
        For Each key In groups.Keys
            i = i + 1: counts = groups(key)
            output(i, 1) = key: output(i, 2) = counts(1): output(i, 3) = counts(0): output(i, 4) = counts(1) - counts(0)
            output(i, 5) = counts(0) / daysA: output(i, 6) = counts(1) / daysB
            output(i, 7) = "N/A - unequal periods": output(i, 8) = "N/A - outside observed span"
            If safe Then
                If daysA = daysB Then output(i, 7) = ReportChange(counts(0), counts(1))
                output(i, 8) = ReportChange(counts(0) / daysA, counts(1) / daysB)
            Else
                output(i, 7) = "N/A - outside observed span"
            End If
        Next
        SortConcentration output, 1, groups.Count
        ws.Range("A7").Resize(groups.Count, 8).NumberFormat = "@": ws.Range("A7").Resize(groups.Count, 8).Value2 = output
        ws.Range("B7:D" & groups.Count + 6).NumberFormat = "0": ws.Range("E7:F" & groups.Count + 6).NumberFormat = "0.000"
        ws.Range("G7:H" & groups.Count + 6).NumberFormat = "0.0%"
    Else
        ws.Range("A7").Value2 = "No eligible activity in either period."
    End If
    FinishOutput ws, Application.Max(7, groups.Count + 6), 8
    ConfigureReport ws, ctx, ws.Range("A1:H" & Application.Min(Application.Max(7, groups.Count + 6), 30))
    ws.Range("A3").Value2 = ctx("Crime") & " | " & ctx("Measure") & " | A: " & Format$(CDate(aStart), "mm/dd/yyyy") & " - " & Format$(CDate(aEnd), "mm/dd/yyyy") & " | B: " & Format$(CDate(bStart), "mm/dd/yyyy") & " - " & Format$(CDate(bEnd), "mm/dd/yyyy")
    If Not ctx("LocationMapped") And ctx("Measure") = "Unique incidents" Then ws.Range("A5").Value2 = "Crime-type incident counts overlap when incidents contain multiple offenses. Group counts are not mutually exclusive. Reporting completeness remains unverified."
    ws.Range("A4").Value2 = "Period totals A=" & rowsA.Count & "; B=" & rowsB.Count & ". Compare events/day for unequal lengths. Both periods require verified reporting coverage for interpretation."
    If groups.Count > 0 Then
        CreateReportChart ws, "CAPReport_Comparison", "Period A vs B | " & ctx("Crime") & " | ranked by period B", ws.Range("A7:A" & Application.Min(groups.Count + 6, 26)), ws.Range("B7:B" & Application.Min(groups.Count + 6, 26)), "J6", xlBarClustered
        Dim second As Series
        With ws.ChartObjects("CAPReport_Comparison").Chart
            .SeriesCollection(1).Name = "Period B"
            Set second = .SeriesCollection.NewSeries: second.Name = "Period A"
            second.XValues = ws.Range("A7:A" & Application.Min(groups.Count + 6, 26)): second.Values = ws.Range("C7:C" & Application.Min(groups.Count + 6, 26))
            .HasLegend = True: .Legend.Position = xlLegendPositionBottom
        End With
    End If
    AttachReportingFilter ws, ctx, "Comparison", "", aStart, aEnd
    If showSummary Then MsgBox "Output: " & ws.Name & ". Missing/invalid date rows: " & ctx("MissingDates") & ". Period comparison created. Raw percentage change is suppressed for unequal period lengths or dates outside the observed source span.", vbInformation, "CAP Index"
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    If Not source Is Nothing Then Err.Raise errNo, "Period Comparison", description
    MsgBox description, vbExclamation, "CAP Index"
End Sub

Public Sub BuildCrimeProfile(Optional ByVal source As Range, Optional ByVal dateHeader As String = "Date", Optional ByVal crimeHeader As String = "CrimeType", Optional ByVal incidentHeader As String = "IncidentID", Optional ByVal locationHeader As String = "", Optional ByVal crimeFilter As String = "All crime types", Optional ByVal measure As String = "", Optional ByVal firstDate As Double = 0, Optional ByVal lastDate As Double = 0, Optional ByVal showSummary As Boolean = True, _
    Optional ByVal reportTarget As Worksheet, Optional ByVal reportContext As Object)
    Dim ctx As Object, rows As Collection, typeRows As Collection, data As Variant, categories As Object, locations As Object, seen As Object
    Dim row As Variant, key As Variant, ws As Worksheet, r As Range, output() As Variant, detail() As Variant, i As Long
    Dim records As Long, eligible As Long, missingIDs As Long, missingLocations As Long, unique As Long, excluded As Long, errNo As Long, description As String
    On Error GoTo failed
    If source Is Nothing And reportContext Is Nothing Then OpenToolSetup "capProfile": Exit Sub
    If reportContext Is Nothing Then
    Set ctx = ReadReportInput(source, dateHeader, crimeHeader, incidentHeader, locationHeader, crimeFilter, measure, firstDate, lastDate, False)
    Else
        Set ctx = reportContext
    End If
    If ctx Is Nothing Then Exit Sub
    data = ctx("Data"): Set rows = ReportRows(ctx): Set categories = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
    Set locations = CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(data, 1)
        If Not IsEmpty(data(i, 1)) Then
            If data(i, 1) >= ctx("Start") And data(i, 1) <= ctx("End") And (ctx("Crime") = "All crime types" Or ctx("Crime") = data(i, 2)) Then
                records = records + 1
                If Not categories.Exists(data(i, 2)) Then categories.Add data(i, 2), True
                If data(i, 5) Then
                    eligible = eligible + 1
                    If Not seen.Exists(data(i, 3)) Then seen.Add data(i, 3), True
                Else
                    missingIDs = missingIDs + 1
                End If
            End If
        End If
    Next
    For Each row In rows
        key = data(row, 4)
        If Len(key) = 0 Then
            missingLocations = missingLocations + 1
        Else
            If Not locations.Exists(key) Then locations.Add key, 0&
            locations(key) = locations(key) + 1
        End If
    Next
    Set r = ctx("Source")
    If reportTarget Is Nothing Then
    Set r = ctx("Source"): Set ws = NewOutput(r.Parent.Parent, "Crime_Profile", r, "")
    AnnotateAnalysisScope ws, r
    Else
        Set ws = reportTarget: ClearReportingResults ws
    End If
    WriteHeaders ws, 6, 1, Array("Metric", "Value")
    ws.Range("A7").Value2 = "Selected measure total": ws.Range("B7").Value2 = rows.Count
    ws.Range("A8").Value2 = "Dated source/offense rows": ws.Range("B8").Value2 = records
    ws.Range("A9").Value2 = "Valid unique incidents": ws.Range("B9").Value2 = IIf(ctx("IDMapped"), seen.Count, "IncidentID not mapped")
    ws.Range("A10").Value2 = "Missing/conflicting ID rows": ws.Range("B10").Value2 = IIf(ctx("IDMapped"), missingIDs, "IncidentID not mapped")
    ws.Range("A11").Value2 = "Crime categories (incl missing/error)": ws.Range("B11").Value2 = categories.Count
    ws.Range("A12").Value2 = "Eligible events missing location": ws.Range("B12").Value2 = IIf(ctx("LocationMapped"), missingLocations, "Location not mapped")
    ws.Range("A14:H14").Merge: ws.Range("A14").Value2 = "Crime-type incident counts can overlap: one incident may have multiple offenses. They are not additive shares of mutually exclusive incident categories."
    ws.Range("A14").WrapText = True: ws.Rows(14).RowHeight = 32
    WriteHeaders ws, 15, 1, Array("Crime type", "Event count", "Percent of measure total")
    If categories.Count > 0 Then
        ReDim output(1 To categories.Count, 1 To 3): i = 0
        For Each key In categories.Keys
            Set typeRows = ReportRows(ctx, 0, 0, CStr(key)): i = i + 1
            output(i, 1) = key: output(i, 2) = typeRows.Count
            If rows.Count > 0 Then output(i, 3) = typeRows.Count / rows.Count Else output(i, 3) = "N/A"
        Next
        SortConcentration output, 1, categories.Count
        ws.Range("A16").Resize(categories.Count, 3).NumberFormat = "@": ws.Range("A16").Resize(categories.Count, 3).Value2 = output
        ws.Range("B16:B" & categories.Count + 15).NumberFormat = "0": ws.Range("C16:C" & categories.Count + 15).NumberFormat = "0.0%"
    End If
    WriteHeaders ws, 15, 5, Array("Top 20 locations", "Event count", "Additional events")
    If locations.Count > 0 Then
        ReDim detail(1 To locations.Count, 1 To 3): i = 0
        For Each key In locations.Keys: i = i + 1: detail(i, 1) = key: detail(i, 2) = locations(key): detail(i, 3) = locations(key) - 1: Next
        SortConcentration detail, 1, locations.Count
        ws.Range("E16").Resize(Application.Min(20, locations.Count), 3).NumberFormat = "@"
        For i = 1 To Application.Min(20, locations.Count)
            ws.Cells(i + 15, 5).Value2 = detail(i, 1): ws.Cells(i + 15, 6).Value2 = detail(i, 2): ws.Cells(i + 15, 7).Value2 = detail(i, 3)
        Next
    Else
        ws.Range("E16").Value2 = "No mapped eligible locations"
    End If
    FinishOutput ws, Application.Max(16, categories.Count + 15, Application.Min(20, locations.Count) + 15), 8
    ConfigureReport ws, ctx, ws.Range("A1:H" & Application.Min(35, ws.UsedRange.Rows.Count))
    ws.Range("A15:C15").Font.Bold = True: ws.Range("E15:G15").Font.Bold = True
    ws.Columns("A").ColumnWidth = 36: ws.Columns("E").ColumnWidth = 36
    If categories.Count > 0 Then CreateReportChart ws, "CAPReport_CrimeMix", "Crime-type counts | " & ctx("Measure"), ws.Range("A16:A" & Application.Min(35, categories.Count + 15)), ws.Range("B16:B" & Application.Min(35, categories.Count + 15)), "J6", xlBarClustered
    AttachReportingFilter ws, ctx, "Profile", ""
    If showSummary Then MsgBox "Output: " & ws.Name & ". Missing/invalid date rows: " & ctx("MissingDates") & ". Crime profile created. Incident counts by crime type can overlap when an incident contains multiple offenses.", vbInformation, "CAP Index"
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    If Not source Is Nothing Then Err.Raise errNo, "Crime Profile", description
    MsgBox description, vbExclamation, "CAP Index"
End Sub

Public Sub CreateReportChart(ByVal ws As Worksheet, ByVal name As String, ByVal title As String, ByVal labels As Range, ByVal values As Range, ByVal anchor As String, ByVal chartType As Long)
    Dim chart As ChartObject, series As Series
    Set chart = ws.ChartObjects.Add(ws.Range(anchor).Left, ws.Range(anchor).Top, 720, 380)
    chart.Name = name: chart.Chart.ChartType = chartType: chart.Chart.HasTitle = True: chart.Chart.ChartTitle.Text = title
    Set series = chart.Chart.SeriesCollection.NewSeries: series.XValues = labels: series.Values = values
    series.Name = "Recorded events": chart.Chart.HasLegend = False: chart.Chart.ChartArea.Font.Name = "Calibri"
    chart.Chart.Axes(xlValue).MinimumScale = 0
    chart.Chart.Axes(xlValue).MajorUnit = Application.Max(1, -Int(-Application.Max(values) / 5))
    If chartType = xlBarClustered Then chart.Chart.Axes(xlCategory).ReversePlotOrder = True
End Sub

Private Function PeriodGroupCounts(ByVal ctx As Object, ByVal firstDate As Double, ByVal lastDate As Double) As Object
    Dim data As Variant, counts As Object, seen As Object, i As Long, key As String, identity As String, eligible As Boolean
    Set counts = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary"): data = ctx("Data")
    For i = 2 To UBound(data, 1)
        eligible = False
        If Not IsEmpty(data(i, 1)) Then eligible = data(i, 1) >= firstDate And data(i, 1) <= lastDate And (ctx("Crime") = "All crime types" Or ctx("Crime") = data(i, 2))
        If eligible Then
            If ctx("Measure") = "Unique incidents" And Not data(i, 5) Then eligible = False
        End If
        If eligible Then
            key = IIf(ctx("LocationMapped"), data(i, 4), data(i, 2)): If Len(key) = 0 Then key = "[Missing location]"
            identity = KeyPart(key) & KeyPart(data(i, 3))
            If ctx("Measure") = "Source records" Or Not seen.Exists(identity) Then
                If Not counts.Exists(key) Then counts.Add key, 0&
                counts(key) = counts(key) + 1: seen(identity) = True
            End If
        End If
    Next
    Set PeriodGroupCounts = counts
End Function
