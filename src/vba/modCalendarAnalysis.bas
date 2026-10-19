Attribute VB_Name = "modCalendarAnalysis"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
Option Explicit

Public Sub BuildYearOverYear()
    BuildCalendarReport kind:="YoY"
End Sub
Public Sub BuildSeasonality()
    BuildCalendarReport kind:="Seasonality"
End Sub

Public Sub BuildCalendarReport(Optional ByVal source As Range, Optional ByVal kind As String = "YoY", Optional ByVal dateHeader As String = "Date", _
    Optional ByVal crimeHeader As String = "CrimeType", Optional ByVal incidentHeader As String = "IncidentID", Optional ByVal measure As String = "", _
    Optional ByVal firstDate As Double = 0, Optional ByVal lastDate As Double = 0, Optional ByVal unit As String = "Months", Optional ByVal showSummary As Boolean = True, _
    Optional ByVal reportTarget As Worksheet, Optional ByVal reportContext As Object)
    Dim ctx As Object, ws As Worksheet, r As Range, choice As Variant, description As String, errNo As Long
    On Error GoTo failed
    If source Is Nothing And reportContext Is Nothing Then OpenToolSetup IIf(kind = "YoY", "capYoY", "capSeasonality"): Exit Sub
    If reportContext Is Nothing Then Set ctx = ReadReportInput(source, dateHeader, crimeHeader, incidentHeader, "", "All crime types", measure, firstDate, lastDate, False) Else Set ctx = reportContext
    If ctx Is Nothing Then Exit Sub
    If kind <> "YoY" And kind <> "Seasonality" Then Err.Raise vbObjectError + 985, , "Invalid calendar report kind."
    If kind = "YoY" And source Is Nothing Then
        choice = Application.InputBox("Year-over-Year: 1 = monthly comparisons; 2 = annual comparisons.", "CAP Index", "1", Type:=1)
        If VarType(choice) = vbBoolean Then Exit Sub
        If choice <> 1 And choice <> 2 Then Err.Raise vbObjectError + 985, , "Choose 1 or 2."
        unit = IIf(choice = 1, "Months", "Years")
    End If
    Set r = ctx("Source")
    If reportTarget Is Nothing Then
        Set ws = NewOutput(r.Parent.Parent, IIf(kind = "YoY", "Year_Over_Year", "Seasonality"), r, "")
        AnnotateAnalysisScope ws, r
    Else
        Set ws = reportTarget: ClearReportingResults ws
        ws.Range("A6:AD1006").Clear
    End If
    If kind = "YoY" Then RenderYearOverYear ws, ctx, unit Else RenderSeasonality ws, ctx
    AttachReportingFilter ws, ctx, kind, unit
    If kind = "Seasonality" Then FormatSeasonality ws, ctx
    If showSummary Then MsgBox "Output: " & ws.Name & ". Missing/invalid dates: " & ctx("MissingDates") & ". Partial or observed-edge calendar bins are flagged. Reporting coverage remains unverified.", vbInformation, "CAP Index"
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    If Not source Is Nothing Then Err.Raise errNo, "Calendar analysis", description
    MsgBox description, vbExclamation, "CAP Index | Calendar analysis"
End Sub

Private Function FullCalendarBin(ByVal ctx As Object, ByVal firstDate As Double, ByVal lastDate As Double, Optional ByVal requested As Boolean = True) As Boolean
    FullCalendarBin = firstDate >= ctx("SourceStart") And lastDate <= ctx("SourceEnd")
    If requested Then FullCalendarBin = FullCalendarBin And firstDate >= ctx("Start") And lastDate <= ctx("End")
End Function

Private Sub RenderYearOverYear(ByVal ws As Worksheet, ByVal ctx As Object, ByVal unit As String)
    Dim origin As Date, boundary As Date, nextBoundary As Date, previous As Date, nextPrevious As Date, p As Long, i As Long
    Dim count As Long, priorCount As Long, days As Long, priorDays As Long, rows As Collection, output() As Variant, offset As Double, full As Boolean
    If unit <> "Months" And unit <> "Years" Then Err.Raise vbObjectError + 985, , "Invalid Year-over-Year interval."
    If unit = "Months" Then origin = DateSerial(Year(CDate(ctx("Start"))), Month(CDate(ctx("Start"))), 1) Else origin = DateSerial(Year(CDate(ctx("Start"))), 1, 1)
    p = DateDiff(IIf(unit = "Months", "m", "yyyy"), origin, CDate(ctx("End"))) + 1
    If p > 1000 Then Err.Raise vbObjectError + 985, , "Select at most 1000 calendar periods."
    ReDim output(1 To p, 1 To 8)
    If ws.Parent.Date1904 Then offset = 1462
    For i = 1 To p
        boundary = DateAdd(IIf(unit = "Months", "m", "yyyy"), i - 1, origin): nextBoundary = DateAdd(IIf(unit = "Months", "m", "yyyy"), 1, boundary)
        previous = DateAdd("yyyy", -1, boundary): nextPrevious = DateAdd("yyyy", -1, nextBoundary)
        Set rows = ReportRows(ctx, Application.Max(CDbl(boundary), ctx("Start")), Application.Min(CDbl(nextBoundary) - 1, ctx("End"))): count = rows.Count
        Set rows = ReportRows(ctx, CDbl(previous), CDbl(nextPrevious) - 1): priorCount = rows.Count
        days = Application.Min(CDbl(nextBoundary) - 1, ctx("End")) - Application.Max(CDbl(boundary), ctx("Start")) + 1: priorDays = DateDiff("d", previous, nextPrevious)
        full = FullCalendarBin(ctx, CDbl(boundary), CDbl(nextBoundary) - 1) And FullCalendarBin(ctx, CDbl(previous), CDbl(nextPrevious) - 1, False)
        output(i, 1) = CDbl(boundary) - offset: output(i, 2) = count: output(i, 3) = priorCount
        output(i, 4) = "N/A - partial/edge comparison": output(i, 8) = output(i, 4)
        If full Then output(i, 4) = ReportChange(priorCount, count): output(i, 8) = ReportChange(priorCount / priorDays, count / days)
        output(i, 5) = count / days: output(i, 6) = priorCount / priorDays
        output(i, 7) = IIf(full, "Full bins; unverified", "Partial/observed edge")
        If days <> priorDays Then output(i, 7) = output(i, 7) & "; unequal days"
    Next
    WriteHeaders ws, 6, 1, Array("Period", "Event count", "Prior-year count", "Count change", "Events/day", "Prior events/day", "Coverage", "Daily-rate change")
    ws.Range("A7").Resize(p, 8).NumberFormat = "@": ws.Range("A7").Resize(p, 8).Value2 = output
    ws.Range("A7:A" & p + 6).NumberFormat = IIf(unit = "Months", "mmm yyyy", "yyyy")
    ws.Range("B7:C" & p + 6).NumberFormat = "0": ws.Range("D7:D" & p + 6).NumberFormat = "0.0%"
    ws.Range("E7:F" & p + 6).NumberFormat = "0.000": ws.Range("H7:H" & p + 6).NumberFormat = "0.0%"
    FinishOutput ws, p + 6, 8: ConfigureReport ws, ctx, ws.Range("A1:H" & Application.Min(30, p + 6))
    ws.Columns("G").ColumnWidth = 26: ws.Range("A6:H" & p + 6).WrapText = True: ws.Range("A6:H" & p + 6).Rows.AutoFit
    ws.Range("A7:H" & p + 6).VerticalAlignment = xlCenter
    CreateReportChart ws, "CAPReport_YoY", "Year-over-Year | " & ctx("Crime") & " | " & ctx("Measure"), ws.Range("A7:A" & p + 6), ws.Range("B7:B" & p + 6), "J6", xlLineMarkers
    With ws.ChartObjects("CAPReport_YoY").Chart
        .SeriesCollection(1).Name = "Current period"
        Dim second As Series
        Set second = .SeriesCollection.NewSeries: second.Name = "Prior year": second.XValues = ws.Range("A7:A" & p + 6): second.Values = ws.Range("C7:C" & p + 6)
        .HasLegend = True: .Legend.Position = xlLegendPositionBottom
    End With
End Sub

Private Sub RenderSeasonality(ByVal ws As Worksheet, ByVal ctx As Object)
    Dim firstYear As Long, lastYear As Long, years As Long, yr As Long, m As Long, col As Long, cols As Long
    Dim boundary As Date, nextBoundary As Date, rows As Collection, count As Long, comparable As Long, total As Long, days As Long, output() As Variant
    firstYear = Year(CDate(ctx("Start"))): lastYear = Year(CDate(ctx("End"))): years = lastYear - firstYear + 1: cols = years + 5
    If years > 20 Then Err.Raise vbObjectError + 985, , "Select at most 20 years for the seasonality matrix."
    ReDim output(1 To 12, 1 To cols)
    ws.Cells(6, 1).Value2 = "Month"
    For yr = firstYear To lastYear: ws.Cells(6, yr - firstYear + 2).Value2 = yr: Next
    WriteHeaders ws, 6, years + 2, Array("Comparable years", "Avg/full month", "Events/day", "Coverage")
    For m = 1 To 12
        output(m, 1) = MonthName(m): comparable = 0: total = 0: days = 0
        For yr = firstYear To lastYear
            boundary = DateSerial(yr, m, 1): nextBoundary = DateAdd("m", 1, boundary): col = yr - firstYear + 2
            If CDbl(nextBoundary) - 1 < ctx("Start") Or CDbl(boundary) > ctx("End") Then
                output(m, col) = "Outside request"
            Else
                Set rows = ReportRows(ctx, Application.Max(CDbl(boundary), ctx("Start")), Application.Min(CDbl(nextBoundary) - 1, ctx("End"))): count = rows.Count
                If FullCalendarBin(ctx, CDbl(boundary), CDbl(nextBoundary) - 1) Then
                    output(m, col) = count: comparable = comparable + 1: total = total + count: days = days + DateDiff("d", boundary, nextBoundary)
                Else
                    output(m, col) = "Partial/edge: " & count
                End If
            End If
        Next
        output(m, years + 2) = comparable: output(m, years + 3) = "N/A": output(m, years + 4) = "N/A"
        If comparable > 0 Then output(m, years + 3) = total / comparable: output(m, years + 4) = total / days
        output(m, years + 5) = "Full bins only; unverified"
    Next
    ws.Range("A7").Resize(12, cols).NumberFormat = "@": ws.Range("A7").Resize(12, cols).Value2 = output
    ws.Range(ws.Cells(7, 2), ws.Cells(18, years + 2)).NumberFormat = "0"
    ws.Range(ws.Cells(7, years + 3), ws.Cells(18, years + 3)).NumberFormat = "0.0": ws.Range(ws.Cells(7, years + 4), ws.Cells(18, years + 4)).NumberFormat = "0.000"
    FinishOutput ws, 18, cols: ConfigureReport ws, ctx, ws.Range(ws.Cells(1, 1), ws.Cells(18, cols))
    Dim heatColors As Object
    Set heatColors = ws.Range(ws.Cells(7, 2), ws.Cells(18, years + 1)).FormatConditions.AddColorScale(3)
    heatColors.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 250, 240)
    heatColors.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 198, 109)
    heatColors.ColorScaleCriteria(3).FormatColor.Color = RGB(210, 43, 43)
    ws.Names.Add Name:="_CAPReportBody", RefersTo:="='" & ws.Name & "'!" & ws.Range(ws.Cells(1, 1), ws.Cells(18, cols)).Address
    CreateReportChart ws, "CAPReport_Seasonality", "Seasonality: pooled daily rate | " & ctx("Crime"), ws.Range("A7:A18"), ws.Range(ws.Cells(7, years + 4), ws.Cells(18, years + 4)), ws.Cells(6, Application.Max(10, cols + 2)).Address, xlColumnClustered
    With ws.ChartObjects("CAPReport_Seasonality").Chart.Axes(xlValue)
        .MajorUnitIsAuto = True: .MaximumScaleIsAuto = True: .TickLabels.NumberFormat = "0.000"
    End With
End Sub

Private Sub FormatSeasonality(ByVal ws As Worksheet, ByVal ctx As Object)
    Dim years As Long, cols As Long, rowNo As Variant
    years = Year(CDate(ctx("End"))) - Year(CDate(ctx("Start"))) + 1: cols = years + 5
    For Each rowNo In Array(1, 2, 4, 5)
        ws.Range(ws.Cells(rowNo, 1), ws.Cells(rowNo, Application.Max(8, cols))).UnMerge
        ws.Range(ws.Cells(rowNo, 1), ws.Cells(rowNo, cols)).Merge
    Next
    ws.Range("E3:H3").UnMerge: ws.Range(ws.Cells(3, 5), ws.Cells(3, cols)).Merge
    ws.Columns("A").ColumnWidth = 18
    ws.Range(ws.Cells(6, 2), ws.Cells(18, years + 1)).EntireColumn.ColumnWidth = 12
    ws.Range(ws.Cells(6, years + 2), ws.Cells(18, years + 4)).EntireColumn.ColumnWidth = 17
    ws.Columns(cols).ColumnWidth = 28
    ws.Range(ws.Cells(6, 2), ws.Cells(6, years + 1)).NumberFormat = "0"
    ws.Range(ws.Cells(6, 2), ws.Cells(6, years + 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(6, 1), ws.Cells(18, cols)).WrapText = True
    ws.Range(ws.Cells(7, 1), ws.Cells(18, cols)).VerticalAlignment = xlCenter
    ws.Rows("6:18").AutoFit: ws.Rows(2).RowHeight = 42: ws.Rows(5).RowHeight = 42
    ws.Range("A4").Value2 = ws.Range("A4").Value2 & "; complete calendar bins use observed/requested span only, not verified reporting coverage."
    ws.Rows(4).RowHeight = 42
End Sub
