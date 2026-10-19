Attribute VB_Name = "modToolGuide"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
Option Explicit

Public Sub BuildToolGuide(Optional ByVal wb As Workbook, Optional ByVal showSummary As Boolean = True)
    Dim ws As Worksheet, r As Range, rowNo As Long, errNo As Long, description As String
    On Error GoTo failed
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise vbObjectError + 910, , "Open a writable data workbook first."
    If wb Is ThisWorkbook Or wb.ReadOnly Or wb.ProtectStructure Then Err.Raise vbObjectError + 910, , "Open a writable data workbook that permits new sheets."
    Set r = wb.Worksheets(1).Range("A1")
    Set ws = NewOutput(wb, "CAP_Tool_Guide", r, "Use the setup window: select headers/data, choose fields/options, then Run. Timestamp and Hour remain single-column tools. Source data is never saved automatically.")
    ws.Range("A4:D4").Merge: ws.Range("A4").Value2 = "Records may be offenses, not incidents. Map a stable IncidentID for incident counts. IDs and addresses must be qualified when combining agencies."
    ws.Range("A5:D5").Merge: ws.Range("A5").Value2 = "Analysis reports use saved snapshots. Rebuild after source edits. Analyses follow Analysis Scope; preparation tools include all rows. Review privacy and coverage before sharing."
    ws.Columns("A").ColumnWidth = 22: ws.Columns("B:D").ColumnWidth = 28
    ws.Range("A1:D1").Merge: ws.Range("A1").Value2 = "CAP Index | Help / Tool Guide | " & CAP_VERSION
    ws.Range("A1:D1").Interior.Color = RGB(26, 55, 82): ws.Range("A1:D1").Font.Color = vbWhite
    ws.Range("A1").Font.Size = 20: ws.Rows(1).RowHeight = 34
    ws.Range("A2:D2").Merge: ws.Range("A2").WrapText = True: ws.Rows(2).RowHeight = 38
    ws.Range("A3:D3").Merge: ws.Range("A3").WrapText = True: ws.Rows(3).RowHeight = 32
    ws.Range("A4:D5").WrapText = True: ws.Rows("4:5").RowHeight = 34
    rowNo = 7
    GuideRow ws, rowNo, "Dataset Check + Coverage", "Setup window: select headers/data; optional Date, Crime Type, IncidentID, Address, Latitude and Longitude", "New Dataset_Check sheet with checks, months, crime mix and field completeness", "Zero-record months can be gaps. Coordinates are checked for ranges, not jurisdiction or accuracy. No source rows removed."
    GuideRow ws, rowNo, "Duplicate / Case Review", "Headers/data; IncidentID and Crime Type", "New review sheet, original records plus review flags and incident crime-type list", "Repeated IDs can be valid multiple offenses. Review before removing anything."
    GuideRow ws, rowNo, "Hundred-Block Builder", "Address in the working dataset; include City/State for later analysis", "New selected-data working table plus CAP_HundredBlock and Review Status", "Uses address bins, not validated street segments; review aliases and jurisdiction."
    GuideRow ws, rowNo, "Randomize Block Addresses", "Agency-masked block Address; optional IncidentID/City/State and seed", "New selected-data working table plus synthetic address, block and status", "1200 or 12XX generates 1200-1299. Synthetic addresses do not recover true locations; never present them as observed incident points."
    GuideRow ws, rowNo, "Census Geocoder", "Address plus City/State or ZIP; headers/data", "New selected-data working table with matched address, coordinates and match/review fields", "Sends address components to Census. Batch limit 10,000 per request; larger sources are split. Short/numeric ZIP repairs, including inferred ZIP+4, require verification in Review Flag. Match status is not numeric confidence or rooftop accuracy."
    GuideRow ws, rowNo, "Timestamp Fields", "One Date/timestamp column, including its header", "Adjacent six date fields; Time Only and Hour when valid times exist", "US month/day/year or ISO text. Date-only rows remain untimed. Requires an unprotected source sheet."
    GuideRow ws, rowNo, "Hour from Time", "One Time column, including its header", "Adjacent Hour (0-23)", "Use with separate Date and Time fields. Accepts Excel times and time text; invalid/blank values stay blank."
    GuideRow ws, rowNo, "Single Site", "Latitude, Longitude, Crime Type; optional Date and IncidentID. Enter site coordinates and reporting dates.", "Radius summary for 0.25, 0.5 and 1 mile, separate bands, crime breakdown and nearby source records", "Blue B3 selects crime; B4/D4 set site coordinates; F4/H4 set dates. Uses approximate straight-line miles, not driving distance. Synthetic locations are excluded; geocoded precision is labelled. Unique incident counts exclude missing IDs and conflicting dates/coordinates."
    GuideRow ws, rowNo, "Crime Concentration", "Hundred Block, Crime Type; optional IncidentID/City/State", "Filtered concentration report and charts", "Use blue B3 for crime and the count-measure control. Denominator includes only blocks with selected crime. Percentages are not directly comparable with published estimates that include zero-crime places."
    GuideRow ws, rowNo, "Temporal Heat Map", "Crime Type plus Date/Time, Date + Time, Day of Week + Hour, or Day of Week + Time", "Monday-Sunday by hours 0-23", "Choose Time input in setup; unused fields are hidden. Missing hours are excluded; date-only events cannot be assigned an hour."
    GuideRow ws, rowNo, "Repeat Locations", "Address, Crime Type; optional IncidentID/City/State", "Address counts and additional events", "No date/time required. Match addresses and qualify locations across jurisdictions."
    GuideRow ws, rowNo, "Hot Streets", "Hundred Block/street ID, Crime Type, Date; optional IncidentID", "Week/month interval profiles across the source span", "Descriptive trajectories; partial intervals and coverage affect interpretation. Time is not required. B9 sets the minimum events across early and late comparison windows combined, default 2; Refresh updates labels. Below the minimum no trend label is assigned."
    GuideRow ws, rowNo, "Monthly / Annual Trends", "Date, Crime Type; optional IncidentID", "New report with crime dropdown at B3, counts, rates and change", "Starts with All crime types. Change the blue dropdown to refresh counts/chart. Partial comparisons and zero baselines are flagged."
    GuideRow ws, rowNo, "Period Comparison", "Date, Crime Type; optional IncidentID and qualified location", "Period A/B comparisons and crime dropdown at B3", "Period A ends before B begins. Unequal lengths use daily rates; crime groups can overlap for multi-offense incidents."
    GuideRow ws, rowNo, "Crime Profile", "Date, Crime Type; optional IncidentID and qualified location", "Incident/record totals, crime mix and top locations", "Incident crime categories are not mutually exclusive; percentages may sum above 100%."
    GuideRow ws, rowNo, "NIBRS Classification", "Crime label/code; optional Agency", "Reference catalogue, editable crosswalk and classified working dataset", "Prepare crosswalk, review codes and mark Approved = Yes, then Apply. Broad labels are unresolved. Group B is arrest-only. Analytical mapping is not submission validation."
    GuideRow ws, rowNo, "Analysis Scope", "All Rows or Visible Rows Only; selected source range", "Scope is saved in workbook and labelled in new analyses", "Visible excludes filtered/manually hidden rows. Existing saved snapshots retain their scope; rebuild to change. Geography and date preparation include all selected rows."
    GuideRow ws, rowNo, "Year over Year", "Date, Crime Type; optional IncidentID", "Monthly or annual current/prior counts and daily rates; crime filter", "Partial bins, observed edges and zero baselines are flagged. Coverage is not inferred from dates."
    GuideRow ws, rowNo, "Seasonality", "Date, Crime Type; optional IncidentID", "Month-by-year matrix and pooled daily rate chart; crime filter", "Only complete bins inside the observed/requested span contribute to averages. Zero counts do not establish reporting coverage."
    GuideRow ws, rowNo, "Select Report Sections / Export Selected Report", "Existing CAP reports in the active workbook", "Selector sheet; combined PDF or PowerPoint", "Mark Include = Yes, choose unique Order values, edit titles, select Tables/Charts/Both, then Export Selected Report. Review output before sharing; no source workbook embedded."
    GuideRow ws, rowNo, "Refresh Report", "Active supported CAP analysis report", "Updates counts/charts using saved data and current selectors; shows last refresh status", "Supports concentration, temporal/location reports, Trends, Period Comparison, Crime Profile and Single Site. Does not reload source edits. Rebuild static Dataset Check/geography outputs or older reports with their original tool."
    GuideRow ws, rowNo, "Reset to All", "Active supported CAP report", "Restores All crime types and refreshes results", "Retains dates, site coordinates, measure and interval settings. It resets the crime selector only."
    GuideRow ws, rowNo, "Excel Essentials", "Current Excel selection", "Native Filter, Sort, Freeze Panes, PivotTable, Table, Clear Filters, Text to Columns, Find and Conditional Formatting commands", "These shortcuts also remain on their original Excel tabs. They operate on the current selection using Excel's usual behavior."
    GuideRow ws, rowNo, "Save Image", "Active report, selected chart, or selected formatted range", "PNG image", "Range limit 100 rows x 80 columns. Include labels and context. Choose an accessible local or synced folder."
    GuideRow ws, rowNo, "PowerPoint Briefing", "Choose CAP report sheets in slide order", "Static summary and chart pictures; source workbook not embedded", "Desktop PowerPoint required. Summaries can be truncated; inspect notes and slides before sharing."
    GuideRow ws, rowNo, "Save Report PDF", "Active CAP report sheet", "Full report body plus chart pages", "Temporary print copy; original print settings retained. Review pagination and sensitive content before sharing."
    GuideRow ws, rowNo, "Maps / 3D Maps", "Prepared location fields; use Latitude and Longitude for 3D Maps", "Excel mapping interface", "Existing Excel commands. Automated point/density mapping is deferred; geocoded and synthetic precision must remain visible."
    GuideRow ws, rowNo, "About CAP Index", "No selection required", "Ownership and attribution notice", "CAP toolkit attribution does not claim ownership of agency data or crime facts."
    ws.Range("A1:D" & rowNo - 1).Font.Name = "Calibri"
    ws.Range("A1:D" & rowNo - 1).VerticalAlignment = xlTop
    ws.Activate: ActiveWindow.DisplayGridlines = False: ws.Range("A1").Select
    If showSummary Then MsgBox "Created " & ws.Name & ". Keep this guide with the working workbook.", vbInformation, "CAP Index"
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    If Not showSummary Then Err.Raise errNo, "Tool Guide", description
    MsgBox description, vbExclamation, "CAP Index | Tool Guide"
End Sub

Private Sub GuideRow(ByVal ws As Worksheet, ByRef rowNo As Long, ByVal tool As String, ByVal inputs As String, ByVal output As String, ByVal notes As String)
    With ws.Range(ws.Cells(rowNo, 1), ws.Cells(rowNo, 4))
        .Merge: .Value2 = tool: .Interior.Color = RGB(26, 55, 82): .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 14
    End With
    ws.Rows(rowNo).RowHeight = 28: rowNo = rowNo + 1
    GuideDetail ws, rowNo, "Select / map", Replace(inputs, "; ", vbCrLf)
    GuideDetail ws, rowNo, "Creates", Replace(output, "; ", vbCrLf)
    GuideDetail ws, rowNo, "Use / limits", Replace(notes, ". ", "." & vbCrLf)
    ws.Rows(rowNo).RowHeight = 14: rowNo = rowNo + 1
End Sub

Private Sub GuideDetail(ByVal ws As Worksheet, ByRef rowNo As Long, ByVal label As String, ByVal content As String)
    Dim piece As Variant, lines As Long
    ws.Cells(rowNo, 1).Value2 = label: ws.Cells(rowNo, 1).Font.Bold = True
    ws.Range(ws.Cells(rowNo, 2), ws.Cells(rowNo, 4)).Merge
    ws.Cells(rowNo, 2).Value2 = content: ws.Range(ws.Cells(rowNo, 1), ws.Cells(rowNo, 4)).WrapText = True
    ws.Range(ws.Cells(rowNo, 1), ws.Cells(rowNo, 4)).Interior.Color = RGB(242, 246, 250)
    For Each piece In Split(content, vbCrLf): lines = lines + Application.Max(1, Application.RoundUp(Len(piece) / 70, 0)): Next
    ws.Rows(rowNo).RowHeight = lines * 16 + 10: rowNo = rowNo + 1
End Sub
