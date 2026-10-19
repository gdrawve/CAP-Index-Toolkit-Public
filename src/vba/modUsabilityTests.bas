Attribute VB_Name = "modUsabilityTests"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
Option Explicit
Private checks As Long
Private Sub CheckUX(ByVal condition As Boolean, ByVal description As String)
    checks = checks + 1
    If Not condition Then Err.Raise vbObjectError + 970, , description
End Sub
Public Function RunUsabilityTests(ByVal previewDir As String) As String
    Dim wb As Workbook, source As Worksheet, report As Worksheet, nearby As Worksheet, profile As Worksheet
    Dim a(1 To 13, 1 To 8) As Variant, i As Long, lat As Double, lon As Double, mileOffset As Double, pi As Double, caught As Long
    Dim oldEvents As Boolean, errNo As Long, description As String, before As Long
    Dim recordReport As Worksheet, recordName As String
    oldEvents = Application.EnableEvents
    On Error GoTo failed
    Application.EnableEvents = False: checks = 0
    pi = 4 * Atn(1)
    CheckUX DistanceMiles(35, -92, 35, -92) = 0, "Coincident coordinates"
    CheckUX Abs(DistanceMiles(0, 0, 1, 0) - 69.09341899) < 0.0001, "One degree on the reference sphere"
    CheckUX DistanceMiles(0, 179.99, 0, -179.99) < 2, "Dateline shortest distance"
    CheckUX DistanceMiles(0, 0, 0, 180) > 12000, "Antipodal calculation is finite"
    CheckUX DistanceBand(0.25) = 1 And DistanceBand(0.5) = 2 And DistanceBand(1) = 3, "Inclusive radius boundaries"
    CheckUX DistanceBand(0.25005) = 2 And DistanceBand(1.00005) = 0, "Counts use unrounded distance"
    CheckUX Not CoordinatePair(92, 35, lat, lon) And Not CoordinatePair(0, 0, lat, lon), "Reversed latitude and placeholders rejected"
    Set wb = Workbooks.Add(xlWBATWorksheet): Set source = wb.Worksheets(1)
    source.Range("A1:C5").Value2 = "TEST"
    source.Range("A1").Value2 = "Address": source.Range("A2:A5").Value2 = "1200 MAIN ST"
    CheckUX FullWorkingSource(source.Range("A1:A3")).Address = source.Range("A1:A3").Address, "Plain range preserves selected rows and fields"
    source.ListObjects.Add xlSrcRange, source.Range("A1:C5"), , xlYes
    CheckUX FullWorkingSource(source.Range("A1:A3")).Address = source.Range("A1:A3").Address, "Table selection does not expand"
    source.Protect
    CheckUX FullWorkingSource(source.Range("A1:A3")).Address = source.Range("A1:A3").Address, "Protected table selection does not expand"
    source.Unprotect: source.ListObjects(1).Unlist
    source.Protect
    BuildHundredBlocks source.Range("A1:A3"), showSummary:=False
    CheckUX wb.Worksheets(wb.Worksheets.Count).ListObjects(1).ListRows.Count = 2 And wb.Worksheets(wb.Worksheets.Count).ListObjects(1).ListColumns.Count = 3, "Protected plain source copies only selected rows and fields"
    Application.DisplayAlerts = False: wb.Worksheets(wb.Worksheets.Count).Delete: Application.DisplayAlerts = True
    source.Unprotect: source.Cells.Clear
    RememberField wb, "Occurrence DATE or timestamp.", "IncidentDate"
    CheckUX SavedField(wb, "Date or timestamp", "Date") = "IncidentDate", "Semantic field mapping reused across tools"
    RememberField wb, "Optional stable IncidentID", ""
    CheckUX SavedField(wb, "IncidentID", "IncidentID") = "", "Optional omitted mapping remembered"
    a(1, 1) = "Date": a(1, 2) = "CrimeType": a(1, 3) = "IncidentID": a(1, 4) = "Latitude": a(1, 5) = "Longitude": a(1, 6) = "Review Flag": a(1, 7) = "PrivateNotes": a(1, 8) = "OffenseID"
    For i = 2 To 13
        a(i, 1) = "01/01/2026": a(i, 2) = "Theft": a(i, 3) = "I" & i - 1: a(i, 5) = -92: a(i, 7) = "PRIVATE FIXTURE": a(i, 8) = "O" & i
        mileOffset = 0.1
        If i = 4 Then mileOffset = 0.4
        If i = 5 Or i = 12 Then mileOffset = 0.75
        If i = 6 Then mileOffset = 2
        a(i, 4) = 35 + mileOffset / 3958.7613 * 180 / pi
    Next
    a(3, 3) = "I1": a(3, 2) = "Robbery"
    a(7, 4) = 999: a(8, 4) = 0: a(8, 5) = 0: a(9, 3) = "": a(10, 6) = "Synthetic block estimate"
    a(11, 3) = "CONFLICT": a(12, 3) = "CONFLICT": a(13, 1) = "bad date"
    source.Range("A1:H13").NumberFormat = "@": source.Range("A1:H13").Value2 = a
    BuildSingleSite source.Range("A1:H13"), 35, -92, measure:="Unique incidents", firstDate:=CDbl(DateSerial(2026, 1, 1)), lastDate:=CDbl(DateSerial(2026, 1, 31)), showSummary:=False
    Set report = wb.Worksheets(2): Set nearby = wb.Worksheets(3)
    CheckUX report.Range("B7").Value2 = 1 And report.Range("B8").Value2 = 2 And report.Range("B9").Value2 = 3, "Unique cumulative radius counts"
    CheckUX report.Range("D7").Value2 = 1 And report.Range("D8").Value2 = 1 And report.Range("D9").Value2 = 1, "Nonoverlapping bands"
    CheckUX nearby.ListObjects(1).ListRows.Count = 4 And nearby.ListObjects(1).ListColumns.Count = 11, "Nearby retains offense rows and original fields plus distance/band/precision"
    CheckUX InStr(report.Range("A10").Value2, "positions in period=3") > 0 And InStr(report.Range("A10").Value2, "ID rows in period=3") > 0, "Invalid/synthetic coordinates and missing/conflicting ID exclusions"
    CheckUX Application.Sum(report.Range("D14:D15")) = 4, "Multi-offense crime counts overlap without multiplying incident total"
    CheckUX Not CAPNamedRange(report, "_CAPRunStatus") Is Nothing And report.Range("B3").Validation.Type = xlValidateList, "Status and visible common crime selector"
    before = wb.Worksheets.Count
    report.Range("B3").Value2 = "Type: ROBBERY": UnifiedRefresh report, False
    CheckUX report.Range("B9").Value2 = 1 And nearby.ListObjects(1).ListRows.Count = 1 And wb.Worksheets.Count = before, "Unified refresh updates report and nearby table in place"
    ResetReportFilters report, False
    CheckUX report.Range("B3").Value2 = "All crime types" And report.Range("B9").Value2 = 3, "Reset retains other settings"
    source.Range("A2").Value2 = "SOURCE EDITED": UnifiedRefresh report, False
    CheckUX report.Range("B9").Value2 = 3, "Saved data refresh does not silently reload source edits"
    source.Range("A2").Value2 = "01/01/2026"
    report.Range("B3").Value2 = "Type: ROBBERY"
    On Error Resume Next: RequireFreshReport report: caught = Err.Number: Err.Clear: On Error GoTo failed
    CheckUX caught <> 0, "Changed filters block stale exports"
    ResetReportFilters report, False
    ExportRangePNG CAPSummaryRange(report), previewDir & "\single-site.png"
    ExportRangePNG nearby.Range("A1:K10"), previewDir & "\site-nearby.png"
    report.Range("B4").Value2 = 45: UnifiedRefresh report, False
    CheckUX report.Range("B9").Value2 = 0 And nearby.ListObjects(1).ListRows.Count = 0, "Empty radius result has no fabricated nearby records"
    report.Range("B4").Value2 = 35: UnifiedRefresh report, False
    Application.EnableEvents = True: report.Range("B3").Value2 = "Type: ROBBERY"
    CheckUX report.Range("B9").Value2 = 1, "Live Single Site filter event"
    Application.EnableEvents = False: ResetReportFilters report, False
    BuildCrimeProfile source.Range("A1:H13"), "Date", "CrimeType", "IncidentID", "", showSummary:=False
    For Each profile In wb.Worksheets
        If profile.Visible = xlSheetVisible And profile.Name Like "Crime_Profile*" Then Exit For
    Next
    profile.Range("B3").Value2 = "Type: ROBBERY": UnifiedRefresh profile, False
    CheckUX profile.Range("B7").Value2 = 1, "Crime Profile also supports unified refresh/filter"
    ResetReportFilters profile, False
    CheckUX profile.Range("B3").Value2 = "All crime types", "Crime Profile reset"
    On Error Resume Next: UnifiedRefresh source, False: caught = Err.Number: Err.Clear: On Error GoTo failed
    CheckUX caught <> 0, "Unsupported/static worksheet gets explicit rebuild guidance"
    CheckUX source.UsedRange.Columns.Count = 8 And source.Range("G2").Value2 = "PRIVATE FIXTURE", "Source records preserved"
    wb.Date1904 = True
    BuildSingleSite source.Range("A1:H13"), 35, -92, measure:="Source records", firstDate:=CDbl(DateSerial(2026, 1, 1)), lastDate:=CDbl(DateSerial(2026, 1, 31)), showSummary:=False
    Set recordReport = ActiveSheet: recordName = recordReport.Name
    CheckUX recordReport.Range("B7").Value2 = 4 And recordReport.Range("B8").Value2 = 5 And recordReport.Range("B9").Value2 = 7, "Record measure counts multiple offenses and coordinate-conflicted IDs as separate source rows"
    CheckUX SiteReportDate(recordReport.Range("F4").Value2, True) = CDbl(DateSerial(2026, 1, 1)), "1904 workbook date boundaries"
    wb.SaveAs previewDir & "\site-snapshot-fixture.xlsx", xlOpenXMLWorkbook
    wb.Close False: Set wb = Workbooks.Open(previewDir & "\site-snapshot-fixture.xlsx")
    Set recordReport = wb.Worksheets(recordName)
    recordReport.Range("B3").Value2 = "Type: ROBBERY": UnifiedRefresh recordReport, False
    CheckUX recordReport.Range("B9").Value2 = 1 And SavedField(wb, "Date", "") = "IncidentDate", "Snapshot and remembered mappings persist after save/reopen"
    Set source = wb.Worksheets(1)
    BuildSingleSite source.Range("A1:H13"), 35, -92, dateHeader:="", measure:="Unique incidents", showSummary:=False
    Set recordReport = ActiveSheet
    CheckUX recordReport.Range("B9").Value2 = 4 And recordReport.Range("F4").Value2 = "Not mapped", "Undated datasets count valid incidents without inventing dates"
    CheckUX InStr(SitePeriodLabel(recordReport), "not mapped") > 0, "Undated report/export period is disclosed"
    wb.Close False: Application.EnableEvents = oldEvents
    RunUsabilityTests = "PASS: " & checks & " distance/single-site/refresh/reset/preferences/status checks": Exit Function
failed:
    errNo = Err.Number: description = Err.Description
    On Error Resume Next: If Not wb Is Nothing Then wb.Close False
    Application.EnableEvents = oldEvents: Application.StatusBar = False
    On Error GoTo 0: Err.Raise errNo, "Usability tests", description
End Function
