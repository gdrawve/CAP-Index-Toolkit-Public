Attribute VB_Name = "modConcentration"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit
Public CAPEvents As clsCAPEvents
Public Const REPORT_DATA_NAME As String = "_CAPConcentrationData"
Public Const REPORT_TYPES_NAME As String = "_CAPConcentrationTypes"

Public Sub InitializeCAPEvents()
    If CAPEvents Is Nothing Then
        Set CAPEvents = New clsCAPEvents: Set CAPEvents.ExcelApp = Application
    End If
End Sub

Public Function ConcentrationCache(ByVal ws As Worksheet) As Range
    Dim named As Name, r As Range
    On Error Resume Next
    For Each named In ws.Names
        If Right$(named.Name, Len(REPORT_DATA_NAME)) = REPORT_DATA_NAME Then
            Set r = named.RefersToRange
            If Not r Is Nothing Then
                If r.Parent.Parent Is ws.Parent Then
                    If r.Columns.Count = 6 And r.Rows.Count >= 2 Then Set ConcentrationCache = r
                End If
            End If
            Exit Function
        End If
    Next
    On Error GoTo 0
End Function

Private Function UniqueSheet(ByVal wb As Workbook, ByVal prefix As String) As Worksheet
    Dim i As Long, name As String, found As Boolean, ws As Worksheet
    For i = 1 To 9999
        name = Left$(prefix, 25) & Format$(i, "0000"): found = False
        For Each ws In wb.Worksheets
            If ws.Name = name Then found = True: Exit For
        Next
        If Not found Then Exit For
    Next
    If found Then Err.Raise vbObjectError + 500, , "No available concentration sheet name."
    Set UniqueSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    UniqueSheet.Name = name
End Function

Public Function CrimeCategory(ByVal value As Variant) As String
    If IsError(value) Then
        CrimeCategory = "[Crime field error]"
    ElseIf Len(SafeText(value)) = 0 Then
        CrimeCategory = "[Missing crime type]"
    Else
        CrimeCategory = "Type: " & NormalText(value)
    End If
End Function

Public Sub BuildConcentrationReport(Optional ByVal source As Range, Optional ByVal blockHeader As String = "CAP_HundredBlock", Optional ByVal crimeHeader As String = "CrimeType", Optional ByVal incidentHeader As String = "IncidentID", Optional ByVal labelHeader As String = "CAP_HundredBlock", Optional ByVal showSummary As Boolean = True, Optional ByVal cityHeader As String = "", Optional ByVal stateHeader As String = "")
    Dim r As Range, wb As Workbook, report As Worksheet, cache As Worksheet, a As Variant, data() As Variant
    Dim bc As Long, cc As Long, ic As Long, lc As Long, i As Long, j As Long, n As Long, id As String, block As String
    Dim ids As Object, categories As Object, category As String, k As Variant, labels() As Variant, value As String
    Dim cityCol As Long, stateCol As Long
    Dim oldEvents As Boolean, changing As Boolean, marker As Name, description As String, errNo As Long
    On Error GoTo failed
    If source Is Nothing Then OpenToolSetup "capPareto": Exit Sub
    Set r = source: If r Is Nothing Then Set r = PickSource("Include Hundred Block, Crime Type and optional IncidentID, City and State. Dates and times are not required for concentration.")
    If r Is Nothing Then Exit Sub
    Set wb = r.Parent.Parent
    If wb Is ThisWorkbook Or wb.ReadOnly Or wb.ProtectStructure Then Err.Raise vbObjectError + 501, , "Select a writable data workbook with unprotected structure."
    If r.Areas.Count <> 1 Or r.Rows.Count < 2 Or r.Rows.Count > 200001 Or r.Columns.Count > 200 Then Err.Raise vbObjectError + 502, , "Select a header/data rectangle, up to 200,000 data rows and 200 columns."
    a = ScopedValues(r): n = UBound(a, 1) - 1
    If source Is Nothing Then
        bc = AskColumn(a, "Hundred-block address field. City/state are combined internally when supplied.", "CAP_HundredBlock", False, "HUNDRED-BLOCK")
        cityCol = AskColumn(a, "Optional City; include when combining jurisdictions.", "", True, "CITY")
        stateCol = AskColumn(a, "Optional State.", "", True, "STATE")
        lc = AskColumn(a, "Optional readable hundred-block label for charts.", "", True, "LOCATION LABEL")
        cc = AskColumn(a, "Crime-type field for the report filter.", "CrimeType", False, "CRIME")
        ic = AskColumn(a, "Optional stable IncidentID. Leave blank to count source records; incident mode requires a valid ID.", "", True, "INCIDENTID")
    Else
        cityCol = HeaderColumn(a, cityHeader): stateCol = HeaderColumn(a, stateHeader)
        bc = HeaderColumn(a, blockHeader): lc = HeaderColumn(a, labelHeader)
        cc = HeaderColumn(a, crimeHeader): ic = HeaderColumn(a, incidentHeader)
    End If
    Set ids = CreateObject("Scripting.Dictionary"): Set categories = CreateObject("Scripting.Dictionary")
    ReDim data(1 To n + 1, 1 To 6)
    data(1, 1) = "IncidentID": data(1, 2) = "BlockKey": data(1, 3) = "BlockLabel"
    data(1, 4) = "CrimeCategory": data(1, 5) = "IncidentValidity": data(1, 6) = "SourceRow"
    For i = 2 To n + 1
        id = "": block = ""
        If ic > 0 Then
            If Not IsError(a(i, ic)) Then id = SafeText(a(i, ic))
        End If
        If Not IsError(a(i, bc)) Then block = SafeText(a(i, bc))
        data(i, 3) = block
        If Len(block) > 0 Then
            If cityCol > 0 Then block = KeyPart(NormalText(a(i, cityCol))) & KeyPart(block)
            If stateCol > 0 Then block = KeyPart(NormalText(a(i, stateCol))) & KeyPart(block)
        End If
        If cityCol > 0 Then
            If IsError(a(i, cityCol)) Or Len(SafeText(a(i, cityCol))) = 0 Then block = ""
        End If
        If stateCol > 0 Then
            If IsError(a(i, stateCol)) Or Len(SafeText(a(i, stateCol))) = 0 Then block = ""
        End If
        data(i, 1) = id: data(i, 2) = block
        If lc > 0 Then
            If Not IsError(a(i, lc)) And Len(SafeText(a(i, lc))) > 0 Then data(i, 3) = SafeText(a(i, lc))
        End If
        category = CrimeCategory(a(i, cc)): data(i, 4) = category
        If Not categories.Exists(category) Then categories.Add category, True
        data(i, 6) = ScopeSourceRow(r, i)
        If Len(id) > 0 Then
            If Not ids.Exists(id) Then
                ids.Add id, block
            ElseIf ids(id) <> block Then
                ids(id) = ""
            End If
        End If
    Next
    For i = 2 To n + 1
        id = data(i, 1): block = data(i, 2)
        If Len(id) = 0 Then
            data(i, 5) = "Missing ID"
        ElseIf Len(ids(id)) = 0 Then
            data(i, 5) = "Missing/conflicting incident block"
        Else
            data(i, 5) = "Eligible"
        End If
    Next
    oldEvents = Application.EnableEvents: Application.EnableEvents = False: changing = True
    Set report = UniqueSheet(wb, "CrimeConcentration_")
    Set cache = UniqueSheet(wb, "_CAPConcentration_")
    cache.Range("A1").Resize(n + 1, 6).NumberFormat = "@"
    cache.Range("A1").Resize(n + 1, 6).Value2 = data
    ReDim labels(1 To categories.Count + 1, 1 To 1)
    labels(1, 1) = "All crime types": i = 1
    For Each k In categories.Keys: i = i + 1: labels(i, 1) = k: Next
    cache.Range("J1").Resize(UBound(labels, 1), 1).NumberFormat = "@"
    cache.Range("J1").Resize(UBound(labels, 1), 1).Value2 = labels
    If categories.Count > 1 Then cache.Range("J2:J" & categories.Count + 1).Sort Key1:=cache.Range("J2"), Order1:=xlAscending, Header:=xlNo
    Set marker = report.Names.Add(Name:=REPORT_DATA_NAME, RefersTo:="='" & cache.Name & "'!" & cache.Range("A1:F" & n + 1).Address): marker.Visible = False
    Set marker = report.Names.Add(Name:=REPORT_TYPES_NAME, RefersTo:="='" & cache.Name & "'!" & cache.Range("J1:J" & categories.Count + 1).Address): marker.Visible = False
    report.Range("A1").Value2 = "CAP Index | Crime Concentration"
    report.Range("A2").Value2 = "Source snapshot: " & r.Parent.Name & "!" & r.Address(False, False) & " | " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | Toolkit " & CAP_VERSION & " | " & CAP_ATTRIBUTION
    AnnotateAnalysisScope report, r
    report.Range("A4").Value2 = "Crime type": report.Range("B4").Value2 = "All crime types"
    report.Range("A5").Value2 = "Count measure": report.Range("B5").Value2 = IIf(ic > 0, "Unique incidents", "Source records")
    report.Range("D4").Value2 = "Displayed crime": report.Range("D5").Value2 = "Displayed measure"
    report.Range("A6").Value2 = "Block denominator": report.Range("B6").Value2 = "Only blocks with selected crime and eligible activity"
    With report.Range("B4").Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & REPORT_TYPES_NAME
        .InCellDropdown = True: .ErrorTitle = "Select a listed crime type": .ErrorMessage = "Choose a value from this report's crime-type list.": .ShowError = True
    End With
    With report.Range("B5").Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Unique incidents" & Application.International(xlListSeparator) & "Source records"
        .InCellDropdown = True: .ShowError = True
    End With
    report.Range("A27").Value2 = "Geography: hundred-block address bins, not validated street segments. Zero-crime streets outside this source are not included."
    report.Range("A28").Value2 = "Incident counts are distinct within each filter and are not additive across offense categories. Conflicting/missing incident locations are excluded across the entire source snapshot."
    report.Range("A29").Value2 = "Counts use the saved Analysis Scope shown above; suspected duplicates remain included. Rebuild this report after source edits. The crime filter updates this saved snapshot."
    report.Range("A30").Value2 = "The workbook contains a hidden local snapshot with incident IDs. Share an approved summary/export rather than the complete workbook when those IDs are sensitive."
    report.Range("A31").Value2 = "Block percentages use ceiling(target share * populated blocks). Actual percentages and boundary ties are reported; 5% does not assume 65% of crime."
    report.Range("A32").Value2 = "Research context: Weisburd (2015), Law of Crime Concentration. https://doi.org/10.1111/1745-9125.12070. This report counts only blocks with recorded crime; its percentages are not directly comparable with published estimates that include zero crime places."
    report.Columns("A").ColumnWidth = 27: report.Columns("B").ColumnWidth = 43
    report.Columns("C:G").ColumnWidth = 18: report.Columns("H").ColumnWidth = 24
    For j = 1 To 2
        report.Range(report.Cells(j, 1), report.Cells(j, 8)).Merge
    Next
    For j = 27 To 32
        report.Range(report.Cells(j, 1), report.Cells(j, 8)).Merge: report.Rows(j).WrapText = True: report.Rows(j).RowHeight = 30
    Next
    report.Range("A1:H32").Font.Name = "Calibri": report.Range("A1").Font.Size = 20
    report.Range("A1:H1").Interior.Color = RGB(26, 55, 82): report.Range("A1:H1").Font.Color = vbWhite: report.Rows(1).RowHeight = 34
    report.Range("B4:B5").Interior.Color = RGB(224, 239, 250): report.Range("B4:B5").Font.Bold = True
    report.Range("B6").WrapText = True: report.Rows(6).RowHeight = 30
    cache.Visible = xlSheetVeryHidden
    InitializeCAPEvents
    RefreshConcentration report, True
    report.Activate: report.Range("B4").Select
    Application.EnableEvents = oldEvents: changing = False
    If showSummary Then MsgBox "Output: " & report.Name & ". Review coverage/exclusions in the report. Concentration report created. Select Crime type and Count measure in the blue cells. Counts, rankings and charts update automatically while the add-in is enabled. Denominator: only populated blocks for the selected crime/measure.", vbInformation, "CAP Index"
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    If changing Then Application.EnableEvents = oldEvents
    If Not source Is Nothing Then Err.Raise errNo, "Crime Concentration", description
    MsgBox "Crime Concentration: " & description, vbExclamation, "CAP Index"
End Sub

Public Function ConcentrationCounts(ByVal a As Variant, ByVal selectedCrime As String, ByVal measure As String, ByRef coverage As Variant) As Object
    Dim blocks As Object, seen As Object, i As Long, key As String, id As String, value As Variant
    Dim selectedRows As Long, missingBlocks As Long, missingIDs As Long, excludedRows As Long
    Set blocks = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
    If measure <> "Unique incidents" And measure <> "Source records" Then Err.Raise vbObjectError + 503, , "Select Unique incidents or Source records."
    For i = 2 To UBound(a, 1)
        If selectedCrime = "All crime types" Or selectedCrime = a(i, 4) Then
            selectedRows = selectedRows + 1: key = a(i, 2): id = a(i, 1)
            If Len(key) = 0 Then missingBlocks = missingBlocks + 1
            If Len(id) = 0 Then missingIDs = missingIDs + 1
            If a(i, 5) = "Missing/conflicting incident block" Then excludedRows = excludedRows + 1
            If Len(key) > 0 Then
                If measure = "Source records" Or a(i, 5) = "Eligible" Then
                    If measure = "Source records" Or Not seen.Exists(id) Then
                        If Not blocks.Exists(key) Then blocks.Add key, Array(0&, CStr(a(i, 3)))
                        value = blocks(key): value(0) = value(0) + 1: blocks(key) = value
                        If Len(id) > 0 Then seen(id) = True
                    End If
                End If
            End If
        End If
    Next
    coverage = Array(selectedRows, missingBlocks, missingIDs, excludedRows)
    Set ConcentrationCounts = blocks
End Function

Public Sub RefreshConcentration(Optional ByVal report As Worksheet, Optional ByVal propagateErrors As Boolean = False)
    Dim cache As Range, a As Variant, blocks As Object, coverage As Variant, selectedCrime As String, measure As String
    Dim rows() As Variant, n As Long, total As Long, cumulative As Long, i As Long, j As Long, k As Variant, value As Variant
    Dim targets As Variant, cutoff As Long, share As Double, tieEnd As Long, reach50 As Long, reach80 As Long
    Dim oldEvents As Boolean, oldScreen As Boolean, changing As Boolean, description As String, errNo As Long, permitted As Boolean, lastRow As Long
    On Error GoTo failed
    If report Is Nothing Then Set report = ActiveSheet
    Set cache = ConcentrationCache(report)
    If cache Is Nothing Then Err.Raise vbObjectError + 504, , "Select a CAP Index concentration report."
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    changing = True: Application.EnableEvents = False: Application.ScreenUpdating = False
    report.Range("A8:H25").ClearContents
    If report.AutoFilterMode Then report.AutoFilterMode = False
    lastRow = Application.Max(36, report.Cells(report.Rows.Count, 1).End(xlUp).Row)
    report.Range("A36:G" & lastRow).ClearContents
    For i = report.ChartObjects.Count To 1 Step -1
        Select Case report.ChartObjects(i).Name
            Case "CAP_ConcentrationCurve", "CAP_TopBlocks", "CAP_BlockDistribution": report.ChartObjects(i).Delete
        End Select
    Next
    selectedCrime = SafeText(report.Range("B4").Value2): measure = SafeText(report.Range("B5").Value2)
    For Each value In cache.Parent.Range("J1:J" & cache.Parent.Cells(cache.Parent.Rows.Count, 10).End(xlUp).Row).Value2
        If selectedCrime = value Then permitted = True: Exit For
    Next
    If Not permitted Then Err.Raise vbObjectError + 505, , "The selected crime type is not in this report snapshot."
    a = cache.Value2: Set blocks = ConcentrationCounts(a, selectedCrime, measure, coverage)
    n = blocks.Count
    report.Range("A8").Value2 = "Eligible activity": report.Range("A9").Value2 = "Populated hundred blocks"
    report.Range("A10").Value2 = "Selected source rows": report.Range("B10").Value2 = coverage(0)
    report.Range("A11").Value2 = "Rows missing block key": report.Range("B11").Value2 = coverage(1)
    report.Range("A12").Value2 = "Rows missing IncidentID": report.Range("B12").Value2 = coverage(2)
    report.Range("A13").Value2 = "Rows with invalid incident location": report.Range("B13").Value2 = coverage(3)
    report.Range("A14").Value2 = "Status"
    report.Range("B9").Value2 = n
    If n = 0 Then
        report.Range("B8").Value2 = 0
        SetConcentrationStatus report, selectedCrime, measure, "No eligible activity for this crime/measure. Concentration unavailable."
        report.Range("B14:H14").Merge: report.Range("B14").WrapText = True: report.Rows(14).RowHeight = 34
        GoTo done
    End If
    ReDim rows(1 To n, 1 To 7)
    i = 0
    For Each k In blocks.Keys
        value = blocks(k): i = i + 1: rows(i, 1) = CStr(k): rows(i, 2) = value(0): rows(i, 3) = value(1): total = total + value(0)
    Next
    SortConcentration rows, 1, n
    For i = 1 To n
        cumulative = cumulative + rows(i, 2): rows(i, 4) = i: rows(i, 5) = rows(i, 2) / total
        rows(i, 6) = cumulative / total: rows(i, 7) = i / n
        If reach50 = 0 And rows(i, 6) >= 0.5 Then reach50 = i
        If reach80 = 0 And rows(i, 6) >= 0.8 Then reach80 = i
    Next
    report.Range("B8").Value2 = total
    SetConcentrationStatus report, selectedCrime, measure, "Calculated from " & measure & "; selected-crime populated blocks only."
    report.Range("B14:H14").Merge: report.Range("B14").WrapText = True: report.Rows(14).RowHeight = 34
    WriteHeaders report, 16, 1, Array("Target block share", "Blocks used", "Actual block share", "Crime share", "Boundary ties through rank")
    targets = Array(0.01, 0.05, 0.1, 0.2)
    For j = 0 To 3
        cutoff = -Int(-targets(j) * n): share = rows(cutoff, 6): tieEnd = cutoff
        Do While tieEnd < n
            If rows(tieEnd + 1, 2) <> rows(cutoff, 2) Then Exit Do
            tieEnd = tieEnd + 1
        Loop
        report.Cells(j + 17, 1).Value2 = targets(j): report.Cells(j + 17, 2).Value2 = cutoff
        report.Cells(j + 17, 3).Value2 = cutoff / n: report.Cells(j + 17, 4).Value2 = share: report.Cells(j + 17, 5).Value2 = tieEnd
    Next
    report.Range("A17:A20").NumberFormat = "0%": report.Range("C17:D20").NumberFormat = "0.0%"
    report.Range("A22").Value2 = "Blocks to reach 50%": report.Range("B22").Value2 = reach50: report.Range("C22").Value2 = reach50 / n
    report.Range("A23").Value2 = "Blocks to reach 80%": report.Range("B23").Value2 = reach80: report.Range("C23").Value2 = reach80 / n
    report.Range("C22:C23").NumberFormat = "0.0%"
    report.Range("A25:H25").Merge
    report.Range("A25").Value2 = "Among " & n & " hundred blocks with at least one eligible event, the top " & report.Range("B18").Value2 & " blocks (actual " & Format$(report.Range("C18").Value2, "0.0%") & ") account for " & Format$(report.Range("D18").Value2, "0.0%") & " of eligible " & LCase$(measure) & "."
    report.Range("A25").Font.Bold = True: report.Range("A25").WrapText = True: report.Rows(25).RowHeight = 36
    WriteHeaders report, 35, 1, Array("BlockKey", "Count", "Hundred-block label", "Rank", "Activity share", "Cumulative crime share", "Cumulative block share")
    report.Range("A36").Resize(n, 3).NumberFormat = "@"
    report.Range("A36").Resize(n, 7).Value2 = rows
    report.Range("B36:B" & n + 35).NumberFormat = "0": report.Range("E36:G" & n + 35).NumberFormat = "0.0%"
    report.Range("A35:G" & n + 35).AutoFilter
    report.Columns("A").ColumnWidth = 27: report.Columns("C").ColumnWidth = 27
    report.Range("A35:G35").WrapText = True: report.Rows(35).RowHeight = 34
    With report.Range("A16:E16")
        .Font.Bold = True: .Interior.Color = RGB(26, 55, 82): .Font.Color = vbWhite: .WrapText = True
    End With
    report.Rows(16).RowHeight = 45
    With report.Range("A35:G35")
        .Font.Bold = True: .Interior.Color = RGB(26, 55, 82): .Font.Color = vbWhite
    End With
    CreateConcentrationCharts report, rows, selectedCrime, measure
done:
    CommonCrimeControl report: StampReportStatus report
    If changing Then Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    On Error Resume Next
    If Not report Is Nothing Then report.Range("B14").Value2 = "Refresh failed: " & description
    If changing Then Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    If propagateErrors Then Err.Raise errNo, "Crime Concentration", description
End Sub

Private Sub SetConcentrationStatus(ByVal ws As Worksheet, ByVal crime As String, ByVal measure As String, ByVal status As String)
    ws.Range("E4").Value2 = crime: ws.Range("E5").Value2 = measure
    ws.Range("B14").Formula = "=IF(AND(B4=E4,B5=E5),""" & Replace(status, """", """""") & """,""Filter changed: enable CAP Index and use Refresh Report. Values/charts still describe Displayed crime and measure."")"
    ws.Range("B14").Calculate
End Sub

Private Sub ChartStyle(ByVal chart As Chart, ByVal title As String)
    chart.HasTitle = True: chart.ChartTitle.Text = title
    chart.ChartArea.Font.Name = "Calibri": chart.ChartArea.Font.Size = 11
    chart.ChartArea.Format.Fill.ForeColor.RGB = vbWhite
    chart.ChartArea.Format.Line.Visible = msoFalse
End Sub

Private Sub CreateConcentrationCharts(ByVal ws As Worksheet, ByVal rows As Variant, ByVal crime As String, ByVal measure As String)
    Dim n As Long, stride As Long, points As Long, i As Long, j As Long, x() As Variant, y() As Variant
    Dim chart As ChartObject, series As Series, top As Long, names() As Variant, values() As Variant
    Dim bins(0 To 6) As Variant, counts(0 To 6) As Variant, count As Long
    n = UBound(rows, 1): stride = -Int(-n / 1000#): points = 1
    For i = 1 To n Step stride: points = points + 1: Next
    If (n - 1) Mod stride <> 0 Then points = points + 1
    ReDim x(1 To points): ReDim y(1 To points): x(1) = 0: y(1) = 0: j = 1
    For i = 1 To n Step stride: j = j + 1: x(j) = rows(i, 7): y(j) = rows(i, 6): Next
    If j < points Then x(points) = 1: y(points) = 1
    Set chart = ws.ChartObjects.Add(ws.Range("J2").Left, ws.Range("J2").Top, 620, 290)
    chart.Name = "CAP_ConcentrationCurve": chart.Chart.ChartType = xlXYScatterLinesNoMarkers
    ChartStyle chart.Chart, "Crime concentration | " & crime
    Set series = chart.Chart.SeriesCollection.NewSeries
    series.Name = "Cumulative " & LCase$(measure): series.XValues = x: series.Values = y
    series.Format.Line.ForeColor.RGB = RGB(20, 95, 140): series.Format.Line.Weight = 2.5
    Set series = chart.Chart.SeriesCollection.NewSeries
    series.Name = "Equal distribution": series.XValues = Array(0, 1): series.Values = Array(0, 1)
    series.Format.Line.ForeColor.RGB = RGB(150, 150, 150): series.Format.Line.DashStyle = msoLineDash
    With chart.Chart.Axes(xlCategory)
        .MinimumScale = 0: .MaximumScale = 1: .TickLabels.NumberFormat = "0%": .HasTitle = True: .AxisTitle.Text = "Cumulative share of populated hundred blocks (highest crime first)"
    End With
    With chart.Chart.Axes(xlValue)
        .MinimumScale = 0: .MaximumScale = 1: .TickLabels.NumberFormat = "0%": .HasTitle = True: .AxisTitle.Text = "Cumulative share of eligible crime"
    End With
    chart.Chart.Legend.Position = xlLegendPositionBottom
    top = Application.Min(20, n): ReDim names(1 To top): ReDim values(1 To top)
    For i = 1 To top: names(i) = CStr(i) & ". " & rows(i, 3): values(i) = rows(i, 2): Next
    Set chart = ws.ChartObjects.Add(ws.Range("J22").Left, ws.Range("J22").Top, 620, 380)
    chart.Name = "CAP_TopBlocks": chart.Chart.ChartType = xlBarClustered
    ChartStyle chart.Chart, "Highest-crime hundred blocks | " & crime
    Set series = chart.Chart.SeriesCollection.NewSeries: series.Name = measure: series.XValues = names: series.Values = values
    series.Format.Fill.ForeColor.RGB = RGB(20, 95, 140): chart.Chart.HasLegend = False
    chart.Chart.Axes(xlCategory).ReversePlotOrder = True
    chart.Chart.Axes(xlValue).MinimumScale = 0
    chart.Chart.Axes(xlValue).HasTitle = True: chart.Chart.Axes(xlValue).AxisTitle.Text = measure
    bins(0) = "1": bins(1) = "2-5": bins(2) = "6-10": bins(3) = "11-25": bins(4) = "26-50": bins(5) = "51-100": bins(6) = "101+"
    For i = 0 To 6: counts(i) = 0: Next
    For i = 1 To n
        count = rows(i, 2)
        Select Case count
            Case 1: j = 0
            Case 2 To 5: j = 1
            Case 6 To 10: j = 2
            Case 11 To 25: j = 3
            Case 26 To 50: j = 4
            Case 51 To 100: j = 5
            Case Else: j = 6
        End Select
        counts(j) = counts(j) + 1
    Next
    Set chart = ws.ChartObjects.Add(ws.Range("J47").Left, ws.Range("J47").Top, 620, 290)
    chart.Name = "CAP_BlockDistribution": chart.Chart.ChartType = xlColumnClustered
    ChartStyle chart.Chart, "Crime-count distribution across hundred blocks"
    Set series = chart.Chart.SeriesCollection.NewSeries: series.Name = "Populated blocks": series.XValues = bins: series.Values = counts
    series.Format.Fill.ForeColor.RGB = RGB(47, 135, 116): chart.Chart.HasLegend = False
    chart.Chart.Axes(xlValue).MinimumScale = 0: chart.Chart.Axes(xlValue).HasTitle = True: chart.Chart.Axes(xlValue).AxisTitle.Text = "Number of populated hundred blocks"
    chart.Chart.Axes(xlCategory).HasTitle = True: chart.Chart.Axes(xlCategory).AxisTitle.Text = measure & " per block"
End Sub
