Attribute VB_Name = "modReportingTests"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit
Private checks As Long
Private Function ReportingTestSheet(ByVal wb As Workbook, ByVal index As Long) As Worksheet
    Dim ws As Worksheet, n As Long
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            n = n + 1
            If n = index Then Set ReportingTestSheet = ws: Exit Function
        End If
    Next
    Err.Raise vbObjectError + 879, , "Report fixture sheet missing"
End Function
Private Sub CheckReporting(ByVal condition As Boolean, ByVal description As String)
    checks = checks + 1
    If Not condition Then Err.Raise vbObjectError + 870, , "FAIL: " & description
End Sub

Public Function RunReportingTests(ByVal previewDir As String) As String
    Dim wb As Workbook, source As Worksheet, trend As Worksheet, comparison As Worksheet, profile As Worksheet, ws As Worksheet
    Dim a(1 To 12, 1 To 5) As Variant, i As Long, january As Double, february As Double, marchEnd As Double
    Dim rows As Collection, ctx As Object, sheets As Collection, ppt As Object, deck As Object, ownApp As Boolean
    Dim oldEvents As Boolean, oldScreen As Boolean, beforeCharts As Long, caught As Long, oldArea As String, xml As String
    Dim errNo As Long, description As String, path As String
    On Error GoTo failed
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating: Application.EnableEvents = False
    checks = 0: Set wb = Workbooks.Add(xlWBATWorksheet): Set source = wb.Worksheets(1)
    january = CDbl(DateSerial(2026, 1, 1)): february = CDbl(DateSerial(2026, 2, 1)): marchEnd = CDbl(DateSerial(2026, 3, 31))
    a(1, 1) = "Date": a(1, 2) = "CrimeType": a(1, 3) = "IncidentID": a(1, 4) = "Location": a(1, 5) = "PrivateNotes"
    For i = 2 To 12: a(i, 2) = "Theft": a(i, 3) = "I" & i: a(i, 4) = "100 MAIN, CITY, IL": a(i, 5) = "PRIVATE SENTINEL 8724": Next
    a(2, 1) = "2026-01-01": a(2, 3) = "I1"
    a(3, 1) = "2026-01-01": a(3, 3) = "I1": a(3, 2) = "Burglary"
    a(4, 1) = "2026-01-31": a(4, 3) = "I2"
    a(5, 1) = "2026-02-01": a(5, 3) = "I3": a(5, 4) = "200 OAK, CITY, IL"
    a(6, 1) = "2026-02-01": a(6, 3) = "I3": a(6, 4) = "200 OAK, CITY, IL"
    a(7, 1) = "2026-02-28": a(7, 3) = "I4": a(7, 2) = "Robbery": a(7, 4) = "200 OAK, CITY, IL"
    a(8, 1) = "2026-03-01": a(8, 3) = "I5": a(8, 4) = "300 ELM, CITY, IL"
    a(9, 1) = "2026-03-31": a(9, 3) = "I6": a(9, 4) = "300 ELM, CITY, IL"
    a(10, 1) = "invalid": a(10, 3) = "I7"
    a(11, 1) = "2026-03-02": a(11, 3) = "I8": a(12, 1) = "2026-03-03": a(12, 3) = "I8"
    source.Range("A1:E12").NumberFormat = "@": source.Range("A1:E12").Value2 = a
    Set ctx = ReadReportInput(source.Range("A1:E12"), "Date", "CrimeType", "IncidentID", "", "All crime types", "Unique incidents", january, marchEnd)
    Set rows = ReportRows(ctx): CheckReporting rows.Count = 6 And ctx("MissingDates") = 1, "Incident deduplication/conflicting dates and date exclusions"
    Set rows = ReportRows(ctx, 0, 0, "Type: THEFT"): CheckReporting rows.Count = 5, "Crime-specific unique incident count"
    CheckReporting ReportChange(0, 4) = "N/A - zero baseline" And ReportChange(2, 4) = 1, "Zero baseline and percentage arithmetic"
    BuildTrendReport source.Range("A1:E12"), "Date", "CrimeType", "IncidentID", "All crime types", "Unique incidents", "Months", january, marchEnd, False
    Set trend = ReportingTestSheet(wb, 2)
    CheckReporting trend.Range("B7").Value2 = 2 And trend.Range("B8").Value2 = 2 And trend.Range("B9").Value2 = 2, "Monthly unique incident totals"
    CheckReporting trend.Range("F8").Value2 = 0 And trend.Range("C8").Value2 = 28, "Full-period percentage and February denominator"
    CheckReporting trend.ChartObjects.Count = 1 And Not CAPSummaryRange(trend) Is Nothing, "Trend chart and summary export marker"
    BuildTrendReport source.Range("A1:E12"), "Date", "CrimeType", "IncidentID", "All crime types", "Source records", "Months", january, marchEnd, False
    Set ws = ReportingTestSheet(wb, 3): CheckReporting ws.Range("B7").Value2 = 3 And ws.Range("B8").Value2 = 3 And ws.Range("B9").Value2 = 4, "Offense-record counts preserved"
    BuildTrendReport source.Range("A1:E12"), "Date", "CrimeType", "IncidentID", "All crime types", "Unique incidents", "Months", january + 14, marchEnd, False
    Set ws = ReportingTestSheet(wb, 4): CheckReporting ws.Range("B7").Value2 = 1 And InStr(ws.Range("F8").Value2, "partial") > 0, "Partial-month comparison suppressed"
    BuildTrendReport source.Range("A1:E12"), "Date", "CrimeType", "IncidentID", "All crime types", "Unique incidents", "Months", january, CDbl(DateSerial(2026, 4, 30)), False
    Set ws = ReportingTestSheet(wb, 5): CheckReporting ws.Range("B10").Value2 = 0 And InStr(ws.Range("F10").Value2, "partial") > 0, "Empty periods retained with observed-edge warning"
    BuildTrendReport source.Range("A1:E12"), "Date", "CrimeType", "IncidentID", "All crime types", "Unique incidents", "Years", january, marchEnd, False
    Set ws = ReportingTestSheet(wb, 6): CheckReporting ws.Range("B7").Value2 = 6, "Annual report"
    BuildPeriodComparison source.Range("A1:E12"), "Date", "CrimeType", "IncidentID", "Location", "All crime types", "Unique incidents", january, CDbl(DateSerial(2026, 1, 31)), february, CDbl(DateSerial(2026, 2, 28)), False
    Set comparison = ReportingTestSheet(wb, 7)
    CheckReporting comparison.ChartObjects(1).Chart.SeriesCollection.Count = 2, "Two-period comparison chart"
    CheckReporting comparison.Range("B7").Value2 = 2 And comparison.Range("C7").Value2 = 0, "Period comparison locations"
    CheckReporting InStr(comparison.Range("G7").Value2, "unequal") > 0 And Abs(comparison.Range("F7").Value2 - 2 / 28) < 0.000001, "Unequal-period daily rates"
    BuildPeriodComparison source.Range("A1:E12"), "Date", "CrimeType", "IncidentID", "", "All crime types", "Unique incidents", january, january + 27, february, february + 27, False
    Set ws = ReportingTestSheet(wb, 8): CheckReporting Application.Sum(ws.Range("C7:C9")) = 2 And InStr(ws.Range("A5").Value2, "overlap") > 0, "Multi-offense incident belongs to each crime category"
    BuildCrimeProfile source.Range("A1:E12"), "Date", "CrimeType", "IncidentID", "Location", "All crime types", "Unique incidents", january, marchEnd, False
    Set profile = ReportingTestSheet(wb, 9)
    CheckReporting profile.Range("B7").Value2 = 6 And profile.Range("B8").Value2 = 10 And profile.Range("B9").Value2 = 6, "Profile separates incidents/offense rows"
    CheckReporting profile.Range("B10").Value2 = 2 And profile.Range("B16").Value2 = 5, "Profile excluded IDs and theft incidence"
    CheckReporting Application.Sum(profile.Range("B16:B18")) = 7 And InStr(profile.Range("A14").Value2, "overlap") > 0, "Crime counts overlap without false exclusive shares"
    CheckReporting profile.Range("F16").Value2 = 2 And profile.Range("G16").Value2 = 1, "Profile repeat-location metrics"
    CheckReporting source.Range("E12").Value2 = "PRIVATE SENTINEL 8724" And source.UsedRange.Columns.Count = 5, "Original data unchanged"
    BuildCrimeProfile source.Range("A1:E12"), "Date", "CrimeType", "", "Location", "All crime types", "Source records", january, marchEnd, False
    Set ws = ReportingTestSheet(wb, 10)
    CheckReporting ws.Range("B7").Value2 = 10 And ws.Range("B9").Value2 = "IncidentID not mapped", "Optional missing ID is disclosed rather than reported as zero incidents"
    CheckReporting InStr(DefaultExportFilename("https://example.invalid/workbook", "png"), "://") = 0, "Cloud workbook URL never becomes an export filename"
    On Error Resume Next: RequireNewExportFile "https://example.invalid/test.png": caught = Err.Number: Err.Clear: On Error GoTo failed
    CheckReporting caught = vbObjectError + 840, "URL destination produces actionable validation error rather than bad filename"
    CheckReporting trend.Range("B3").Value2 = "All crime types" And trend.Range("B3").Validation.Type = xlValidateList, "Trend starts with all-crime dropdown"
    trend.Range("B3").Value2 = "Type: BURGLARY": RefreshReportingFilter trend, True
    CheckReporting trend.Range("B7").Value2 = 1 And trend.Range("B8").Value2 = 0 And trend.Range("B9").Value2 = 0, "Dropdown narrows monthly counts and retains zero periods"
    source.Range("B3").Value2 = "SOURCE CHANGED"
    trend.Range("B3").Value2 = "All crime types": RefreshReportingFilter trend, True
    CheckReporting trend.Range("B7").Value2 = 2 And trend.Range("B8").Value2 = 2 And trend.Range("B9").Value2 = 2, "Saved snapshot retains source values and full span on reset"
    source.Range("B3").Value2 = "Burglary"
    comparison.Range("B3").Value2 = "Type: BURGLARY": RefreshReportingFilter comparison, True
    CheckReporting comparison.Range("B7").Value2 = 0 And comparison.Range("C7").Value2 = 1 And comparison.ChartObjects.Count = 1, "Comparison filter updates counts and replaces chart in place"
    comparison.Range("B3").Value2 = "All crime types": RefreshReportingFilter comparison, True
    Application.EnableEvents = True
    trend.Range("B3").Value2 = "Type: ROBBERY"
    CheckReporting trend.Range("B8").Value2 = 1 And trend.Range("B7").Value2 = 0, "Live dropdown event refreshes report"
    trend.Range("B3").Value2 = "All crime types": Application.EnableEvents = False
    path = previewDir & "\monthly-report.png": beforeCharts = trend.ChartObjects.Count
    ExportRangePNG CAPSummaryRange(trend), path
    CheckReporting FileLen(path) > 0 And trend.ChartObjects.Count = beforeCharts, "PNG export and temporary chart cleanup"
    On Error Resume Next: ExportRangePNG CAPSummaryRange(trend), path: caught = Err.Number: Err.Clear: On Error GoTo failed
    CheckReporting caught <> 0, "Existing image not overwritten"
    ExportChartPNG trend.ChartObjects(1).Chart, previewDir & "\trend-chart-caption.png"
    CheckReporting FileLen(previewDir & "\trend-chart-caption.png") > 0 And trend.ChartObjects.Count = beforeCharts, "Captioned chart PNG and cleanup"
    oldArea = profile.PageSetup.PrintArea
    ExportReportPDF profile, previewDir & "\crime-profile.pdf", False
    CheckReporting FileLen(previewDir & "\crime-profile.pdf") > 0 And profile.PageSetup.PrintArea = oldArea, "PDF export and print settings restored"
    Set sheets = New Collection: sheets.Add trend: sheets.Add profile
    ExportBriefingPowerPoint sheets, previewDir & "\reporting-briefing.pptx", False
    CheckReporting FileLen(previewDir & "\reporting-briefing.pptx") > 0, "PowerPoint briefing saved"
    On Error Resume Next: Set ppt = GetObject(, "PowerPoint.Application"): Err.Clear: On Error GoTo failed
    If ppt Is Nothing Then Set ppt = CreateObject("PowerPoint.Application"): ownApp = True
    Set deck = ppt.Presentations.Open(previewDir & "\reporting-briefing.pptx", -1, 0, 0)
    CheckReporting deck.Slides.Count = 4, "Summary and chart slides for both reports"
    CheckReporting deck.Slides(2).Shapes(3).Type = 13, "Charts embedded as pictures rather than workbook/OLE objects"
    deck.Slides(1).Export previewDir & "\briefing-summary.png", "PNG", 1280, 720
    deck.Slides(2).Export previewDir & "\briefing-chart.png", "PNG", 1280, 720
    deck.Close: Set deck = Nothing: If ownApp Then ppt.Quit
    Set ppt = Nothing
    Set ws = wb.Worksheets.Add
    ws.Range("A1:B3").Value2 = "Capture failure fixture": ws.Protect
    caught = 0
    On Error Resume Next
    ExportRangePNG ws.Range("A1:B3"), previewDir & "\protected-capture.png"
    caught = Err.Number: Err.Clear: On Error GoTo failed
    CheckReporting caught <> 0 And ws.ChartObjects.Count = 0, "Failed capture retries clean temporary charts"
    CheckReporting Len(Dir$(previewDir & "\protected-capture.png")) = 0, "Failed capture leaves no export file"
    ws.Unprotect
    wb.Close False: Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    RunReportingTests = "PASS: " & checks & " reporting/PNG/PDF/PowerPoint checks": Exit Function
failed:
    errNo = Err.Number: description = Err.Description
    On Error Resume Next
    If Not deck Is Nothing Then deck.Close
    If Not ppt Is Nothing And ownApp Then ppt.Quit
    If Not wb Is Nothing Then wb.Close False
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    On Error GoTo 0: Err.Raise errNo, "Reporting tests", description
End Function
