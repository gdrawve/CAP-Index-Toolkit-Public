Attribute VB_Name = "modRelease015Tests"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
Option Explicit
Private checks As Long
Private Sub Check015(ByVal condition As Boolean, ByVal description As String)
    checks = checks + 1
    If Not condition Then Err.Raise vbObjectError + 999, , description
End Sub
Public Function RunRelease015Tests(ByVal previewDir As String) As String
    Dim wb As Workbook, source As Worksheet, yoy As Worksheet, seasons As Worksheet, plan As Worksheet, crosswalk As Worksheet, classified As Worksheet
    Dim r As Range, a(1 To 10, 1 To 4) As Variant, data As Variant, catalogue As Object, modes As Object, titles As Object, selected As Collection
    Dim planName As String, i As Long, j As Long, caught As Long, oldEvents As Boolean, oldAlerts As Boolean, errNo As Long, description As String, ppt As Object, deck As Object, ownApp As Boolean
    oldEvents = Application.EnableEvents: oldAlerts = Application.DisplayAlerts
    On Error GoTo failed
    Application.EnableEvents = False: checks = 0
    Set catalogue = NIBRSCatalog()
    Check015 catalogue.Count = 81 And catalogue.Exists("11D") And Not catalogue.Exists("90A"), "2025 catalogue with retired codes excluded"
    Set wb = Workbooks.Add(xlWBATWorksheet): Set source = wb.Worksheets(1)
    a(1, 1) = "Date": a(1, 2) = "CrimeType": a(1, 3) = "IncidentID": a(1, 4) = "Agency"
    For i = 2 To 10: a(i, 1) = "02/15/2025": a(i, 2) = "Robbery": a(i, 3) = "I" & i: a(i, 4) = "Agency A": Next
    a(2, 1) = "01/01/2024": a(3, 1) = "02/15/2024": a(4, 1) = "02/29/2024": a(5, 1) = "03/01/2024"
    a(6, 1) = "02/01/2025": a(7, 1) = "02/28/2025": a(8, 1) = "12/31/2025"
    a(9, 3) = a(6, 3): a(10, 2) = "Theft": a(10, 4) = "Agency B"
    Set r = source.Range("A1:D10"): r.NumberFormat = "@": r.Value2 = a
    source.Rows(3).Hidden = True: SetAnalysisScope wb, "Visible rows only", False
    data = ScopedValues(r)
    Check015 UBound(data, 1) = 9 And ScopeSourceRow(r, 3) = 4, "Visible scope excludes hidden and retains original row references"
    ReviewCases r, "IncidentID", "CrimeType"
    Check015 ActiveSheet.Cells(8, 1).Value2 = "02/29/2024" And ActiveSheet.Cells(8, 5).Value2 = 4, "Case review data and row references stay aligned after scoping"
    Check015 InStr(ActiveSheet.Range("A2").Value2, "excluded=1") > 0, "Output records included/excluded counts"
    SetAnalysisScope wb, "All rows", False
    BuildCalendarReport source:=r, kind:="YoY", measure:="Source records", firstDate:=CDbl(DateSerial(2025, 2, 1)), lastDate:=CDbl(DateSerial(2025, 2, 28)), showSummary:=False
    Set yoy = ActiveSheet
    Check015 yoy.Range("B7").Value2 = 4 And yoy.Range("C7").Value2 = 2, "YoY compares current and prior calendar-month records"
    Check015 Abs(yoy.Range("D7").Value2 - 1) < 0.0001 And Abs(yoy.Range("H7").Value2 - ((4 / 28) / (2 / 29) - 1)) < 0.0001, "Leap-year comparison distinguishes count change from daily rate change"
    yoy.Range("B3").Value2 = "Type: ROBBERY": RefreshReportingFilter yoy, True
    Check015 yoy.Range("B7").Value2 = 3, "YoY crime dropdown refreshes saved snapshot"
    ResetReportFilters yoy, False
    Check015 yoy.Range("B7").Value2 = 4, "Unified reset supports YoY"
    BuildCalendarReport source:=r, kind:="YoY", measure:="Unique incidents", firstDate:=CDbl(DateSerial(2025, 2, 1)), lastDate:=CDbl(DateSerial(2025, 2, 28)), showSummary:=False
    Check015 ActiveSheet.Range("B7").Value2 = 2, "Unique incidents exclude conflicting multi-offense dates and count remaining incident IDs once"
    BuildCalendarReport source:=r, kind:="YoY", measure:="Source records", firstDate:=CDbl(DateSerial(2025, 2, 15)), lastDate:=CDbl(DateSerial(2025, 2, 28)), showSummary:=False
    Check015 InStr(ActiveSheet.Range("D7").Value2, "partial") > 0, "Partial requested comparisons suppress percent change"
    BuildCalendarReport source:=r, kind:="YoY", measure:="Source records", firstDate:=CDbl(DateSerial(2025, 4, 1)), lastDate:=CDbl(DateSerial(2025, 4, 30)), showSummary:=False
    Check015 ActiveSheet.Range("D7").Value2 = "N/A - zero baseline", "Zero prior baseline never divides by zero"
    wb.Date1904 = True
    BuildCalendarReport source:=r, kind:="YoY", measure:="Source records", firstDate:=CDbl(DateSerial(2025, 1, 1)), lastDate:=CDbl(DateSerial(2025, 12, 31)), unit:="Years", showSummary:=False
    Check015 ActiveSheet.Range("B7").Value2 = 5 And ActiveSheet.Range("C7").Value2 = 4 And ActiveSheet.Range("A7").Value2 = CDbl(DateSerial(2025, 1, 1)) - 1462, "Annual comparisons and 1904 display dates"
    wb.Date1904 = False
    BuildCalendarReport source:=r, kind:="Seasonality", measure:="Source records", showSummary:=False
    Set seasons = ActiveSheet
    Check015 seasons.Range("B8").Value2 = 2 And seasons.Range("C8").Value2 = 4 And seasons.Range("D8").Value2 = 2, "Seasonality matrix and comparable years"
    Check015 Abs(seasons.Range("F8").Value2 - 6 / 57) < 0.0001, "Seasonality daily rate pools calendar days including leap year"
    seasons.Range("B3").Value2 = "Type: ROBBERY": RefreshReportingFilter seasons, True
    Check015 seasons.Range("C8").Value2 = 3, "Seasonality crime filter"
    ExportRangePNG yoy.Range("A1:H7"), previewDir & "\year-over-year.png"
    ExportRangePNG seasons.Range("A1:G18"), previewDir & "\seasonality.png"
    PrepareNIBRSCrosswalk r, "CrimeType", "Agency", False, False: Set crosswalk = ActiveSheet
    For i = 7 To crosswalk.Cells(crosswalk.Rows.Count, 2).End(xlUp).Row
        If crosswalk.Cells(i, 2).Value2 = "ROBBERY" Then
            Check015 crosswalk.Cells(i, 3).Value2 = "120" And crosswalk.Cells(i, 5).Value2 = "No", "Exact official label proposed but not automatically approved"
            crosswalk.Cells(i, 5).Value2 = "Yes"
        Else
            Check015 Len(crosswalk.Cells(i, 3).Value2) = 0, "Broad Theft label not inferred"
        End If
    Next
    ApplyNIBRSCrosswalk crosswalk, False: Set classified = ActiveSheet
    Check015 classified.Range("E7").Value2 = "120" And Len(classified.Range("E15").Value2) = 0, "Reviewed classification preserves unresolved records"
    Check015 classified.ListObjects(1).DataBodyRange.Rows.Count = 9 And source.Range("B10").Value2 = "Theft", "Classification retains source fields/records without editing source"
    source.Range("B10").Value2 = "Robbery": ApplyNIBRSCrosswalk crosswalk, False
    Check015 Len(ActiveSheet.Range("E15").Value2) = 0, "Same label from unmapped agency does not inherit another agency's mapping"
    source.Range("B10").Value2 = "Theft"
    crosswalk.Range("A9:E9").Value2 = crosswalk.Range("A7:E7").Value2: ApplyNIBRSCrosswalk crosswalk, False
    Check015 Len(ActiveSheet.Range("E7").Value2) = 0 And InStr(ActiveSheet.Range("I7").Value2, "Duplicate") > 0, "Duplicate crosswalk keys block classification"
    crosswalk.Range("A9:F9").ClearContents
    crosswalk.Range("C7").Value2 = "90Z": ApplyNIBRSCrosswalk crosswalk, False
    Check015 ActiveSheet.Range("G7").Value2 = "B" And InStr(ActiveSheet.Range("I7").Value2, "arrest-only") > 0, "Group B arrest-only classification remains explicit"
    crosswalk.Range("C7").Value2 = "11B": ApplyNIBRSCrosswalk crosswalk, False
    Check015 Len(ActiveSheet.Range("E7").Value2) = 0 And InStr(ActiveSheet.Range("I7").Value2, "Legacy") > 0, "Legacy code needs conversion review"
    crosswalk.Range("C7").Value2 = "360": ApplyNIBRSCrosswalk crosswalk, False
    Check015 Len(ActiveSheet.Range("E7").Value2) = 0 And InStr(ActiveSheet.Range("I7").Value2, "Federal/tribal") > 0, "Restricted code blocked for local/state crosswalk"
    crosswalk.Range("C7").Value2 = "90A": ApplyNIBRSCrosswalk crosswalk, False
    Check015 Len(ActiveSheet.Range("E7").Value2) = 0 And InStr(ActiveSheet.Range("I7").Value2, "obsolete") > 0, "Obsolete code cannot be approved into outputs"
    source.Range("A1:D10").AutoFilter Field:=2, Criteria1:="Robbery": source.Rows(3).Hidden = True
    SetAnalysisScope wb, "Visible rows only", False: data = ScopedValues(r)
    Check015 UBound(data, 1) = 8, "Visible scope respects AutoFilter and manually hidden rows together"
    source.ShowAllData: source.Rows("2:10").Hidden = True
    On Error Resume Next: data = ScopedValues(r): caught = Err.Number: Err.Clear: On Error GoTo failed
    Check015 caught <> 0, "Empty visible scope fails with clear error"
    source.Rows("2:10").Hidden = False: SetAnalysisScope wb, "All rows", False
    CreateReportBuilder wb, False: Set plan = ActiveSheet
    For i = 7 To CAPNamedRange(plan, "_CAPReportPlan").Rows.Count + 5
        If plan.Cells(i, 3).Value2 = yoy.Name Then plan.Cells(i, 1).Value2 = "Yes": plan.Cells(i, 2).Value2 = 2: plan.Cells(i, 4).Value2 = "Year-over-Year": plan.Cells(i, 5).Value2 = "Tables"
        If plan.Cells(i, 3).Value2 = seasons.Name Then plan.Cells(i, 1).Value2 = "Yes": plan.Cells(i, 2).Value2 = 1: plan.Cells(i, 4).Value2 = "Seasonality": plan.Cells(i, 5).Value2 = "Charts"
    Next
    Set selected = SelectedReportSections(plan, modes, titles)
    Check015 selected.Count = 2 And selected(1).Name = seasons.Name And modes(yoy.Name) = "Tables", "Report selection honors order and content"
    For i = 7 To CAPNamedRange(plan, "_CAPReportPlan").Rows.Count + 5
        If plan.Cells(i, 3).Value2 = yoy.Name Then plan.Cells(i, 2).Value2 = 1
    Next
    caught = 0: On Error Resume Next: Set selected = SelectedReportSections(plan, modes, titles): caught = Err.Number: Err.Clear: On Error GoTo failed
    Check015 caught <> 0, "Duplicate selected order rejected"
    For i = 7 To CAPNamedRange(plan, "_CAPReportPlan").Rows.Count + 5
        If plan.Cells(i, 3).Value2 = yoy.Name Then plan.Cells(i, 2).Value2 = 2
    Next
    ExportSelectedReport plan, "pdf", previewDir & "\selected-report.pdf", False
    Check015 FileLen(previewDir & "\selected-report.pdf") > 2000 And Application.DisplayAlerts = oldAlerts, "Combined PDF and alert restoration"
    ExportSelectedReport plan, "pptx", previewDir & "\selected-report.pptx", False
    Set ppt = CreateObject("PowerPoint.Application"): ownApp = True: Set deck = ppt.Presentations.Open(previewDir & "\selected-report.pptx", -1, 0, 0)
    Check015 deck.Slides.Count = 2 And InStr(deck.Slides(1).Shapes(2).TextFrame.TextRange.Text, "Seasonality") > 0, "Selected slides omit unchecked content and honor custom titles/order"
    Check015 InStr(deck.Slides(1).Shapes(4).TextFrame.TextRange.Text, "Scope: All rows") > 0 And deck.Slides(1).Shapes(3).Type = 13, "Chart slides retain saved scope and static pictures"
    deck.Slides(1).Export previewDir & "\selected-slide.png", "PNG", 1280, 720
    deck.Close: Set deck = Nothing: ppt.Quit: Set ppt = Nothing
    For i = 7 To CAPNamedRange(plan, "_CAPReportPlan").Rows.Count + 5
        If plan.Cells(i, 3).Value2 = seasons.Name Then plan.Cells(i, 1).Value2 = "No"
    Next
    ExportSelectedReport plan, "pdf", previewDir & "\selected-tables-only.pdf", False
    Check015 FileLen(previewDir & "\selected-tables-only.pdf") > 2000, "Tables-only PDF completes without temporary chart files"
    caught = 0: On Error Resume Next: ExportSelectedReport plan, "pdf", previewDir & "\selected-tables-only.pdf", False: caught = Err.Number: Err.Clear: On Error GoTo failed
    Check015 caught <> 0, "Selected export protects existing files"
    SetAnalysisScope wb, "Visible rows only", False
    planName = plan.Name
    wb.SaveAs previewDir & "\release015-test.xlsx", xlOpenXMLWorkbook
    wb.Close False: Set wb = Workbooks.Open(previewDir & "\release015-test.xlsx")
    Check015 AnalysisScope(wb) = "Visible rows only" And Not CAPNamedRange(wb.Worksheets(planName), "_CAPReportPlan") Is Nothing, "Nondefault scope and report selector persist through save/reopen"
    wb.Close False: Application.EnableEvents = oldEvents: Application.DisplayAlerts = oldAlerts
    RunRelease015Tests = "PASS: " & checks & " scope/NIBRS/calendar/selected PDF/PowerPoint checks": Exit Function
failed:
    errNo = Err.Number: description = Err.Description
    On Error Resume Next
    If Not deck Is Nothing Then deck.Close
    If Not ppt Is Nothing And ownApp Then ppt.Quit
    If Not wb Is Nothing Then wb.Close False
    Application.EnableEvents = oldEvents: Application.DisplayAlerts = oldAlerts
    On Error GoTo 0: Err.Raise errNo, "Release 0.15 tests", description
End Function
