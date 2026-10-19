Attribute VB_Name = "modTimeAnalysis"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit
Public Const HOT_STREETS_MIN_EVENTS As Long = 2

Public Sub ApplyHeatColors(ByVal cells As Range, ByVal maximum As Double)
    Dim heatScale As ColorScale
    If maximum < 1 Then maximum = 1
    cells.FormatConditions.Delete
    Set heatScale = cells.FormatConditions.AddColorScale(3)
    heatScale.ColorScaleCriteria(1).Type = xlConditionValueNumber: heatScale.ColorScaleCriteria(1).Value = 0
    heatScale.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 255, 255)
    heatScale.ColorScaleCriteria(2).Type = xlConditionValueNumber: heatScale.ColorScaleCriteria(2).Value = maximum / 2
    heatScale.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 214, 102)
    heatScale.ColorScaleCriteria(3).Type = xlConditionValueNumber: heatScale.ColorScaleCriteria(3).Value = maximum
    heatScale.ColorScaleCriteria(3).FormatColor.Color = RGB(245, 100, 75)
End Sub

Public Sub RenderTemporalHeatmap(ByVal ws As Worksheet, ByVal events As Collection)
    Dim grid(1 To 7, 1 To 24) As Variant, days As Variant, observation As Variant, i As Long, j As Long, count As Long, peak As Long, peakText As String
    days = Array("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
    For i = 1 To 7: For j = 1 To 24: grid(i, j) = 0: Next: Next
    For Each observation In events
        i = Weekday(CDate(observation(1)), vbMonday): j = CLng(observation(2)) + 1: grid(i, j) = grid(i, j) + 1
    Next
    ws.Range("A17:Y17").Merge: ws.Range("A17").Value2 = "Day of week x hour (0-23). Counts, not exposure-adjusted rates. White = zero; hotter colors = more selected events."
    ws.Range("A18").Value2 = "Day / Hour"
    For j = 0 To 23: ws.Cells(18, j + 2).Value2 = j: Next
    For i = 1 To 7: ws.Cells(i + 18, 1).Value2 = days(i - 1): Next
    ws.Range("B19:Y25").Value2 = grid: ws.Range("B18:Y25").HorizontalAlignment = xlCenter
    ws.Range("B18:Y18").NumberFormat = "00": ws.Range("B19:Y25").NumberFormat = "0"
    For i = 1 To 7
        For j = 1 To 24
            count = grid(i, j)
            If count > peak Then peak = count: peakText = days(i - 1) & " " & Format$(j - 1, "00") & ":00-" & Format$(j - 1, "00") & ":59"
        Next
    Next
    ApplyHeatColors ws.Range("B19:Y25"), peak
    ws.Range("A18:Y18").Interior.Color = RGB(26, 55, 82): ws.Range("A18:Y18").Font.Color = vbWhite: ws.Range("A18:Y18").Font.Bold = True
    ws.Range("A18:Y25").Borders.Color = RGB(220, 225, 230)
    ws.Range("A19:A25").Font.Bold = True: ws.Range("A18:Y25").Font.Name = "Calibri"
    ws.Columns("B:Y").ColumnWidth = 5.2: ws.Rows("18:25").RowHeight = 26
    ws.Range("A27:Y27").Merge: ws.Range("A27").Value2 = IIf(peak > 0, "First peak cell: " & peakText & " | Count=" & peak & ". Equal peaks may exist; inspect the grid.", "No eligible timed observations. Missing times remain excluded, never assigned midnight.")
    ws.Range("A27").Font.Bold = True: ws.Rows(27).RowHeight = 26
    MarkTimeBody ws, ws.Range("A17:Y27")
End Sub

Public Sub SortDates(ByRef values As Variant, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long, pivot As Double, temp As Double
    i = first: j = last: pivot = values((first + last) \ 2)
    Do While i <= j
        Do While values(i) < pivot: i = i + 1: Loop
        Do While values(j) > pivot: j = j - 1: Loop
        If i <= j Then temp = values(i): values(i) = values(j): values(j) = temp: i = i + 1: j = j - 1
    Loop
    If first < j Then SortDates values, first, j
    If i < last Then SortDates values, i, last
End Sub

Private Function IntegerControl(ByVal ws As Worksheet, ByVal address As String, ByVal minimum As Long, ByVal maximum As Long) As Long
    Dim value As Variant: value = ws.Range(address).Value2
    If IsError(value) Or Len(SafeText(value)) = 0 Then Err.Raise vbObjectError + 610, , "Missing/invalid parameter in " & address
    If Not IsNumeric(value) Then Err.Raise vbObjectError + 611, , "Parameter must be numeric in " & address
    If CDbl(value) <> Fix(CDbl(value)) Or CDbl(value) < minimum Or CDbl(value) > maximum Then Err.Raise vbObjectError + 612, , "Parameter out of range in " & address
    IntegerControl = CLng(value)
End Function

Public Sub RenderRepeatLocations(ByVal ws As Worksheet, ByVal events As Collection)
    Dim groups As Object, labels As Object, observation As Variant, dates As Collection, key As Variant, values() As Variant, gaps() As Variant
    Dim output() As Variant, i As Long, j As Long, n As Long, shown As Long, minCount As Long, gap As Double, sumGap As Double, median As Double
    Dim windowA As Long, windowB As Long, windowC As Long, a As Long, b As Long, c As Long, offset As Double, maxCount As Long
    minCount = IntegerControl(ws, "B6", 2, 36500): windowA = IntegerControl(ws, "B7", 1, 36500)
    windowB = IntegerControl(ws, "B8", 1, 36500): windowC = IntegerControl(ws, "B9", 1, 36500)
    If windowA > windowB Or windowB > windowC Then Err.Raise vbObjectError + 613, , "Repeat windows must be in increasing order."
    Set groups = CreateObject("Scripting.Dictionary"): Set labels = CreateObject("Scripting.Dictionary")
    For Each observation In events
        key = observation(3)
        If Not groups.Exists(key) Then
            Set dates = New Collection: groups.Add key, dates: labels.Add key, observation(4)
        End If
        Set dates = groups(key): dates.Add observation(1)
    Next
    For Each key In groups.Keys
        Set dates = groups(key): If dates.Count >= minCount Then shown = shown + 1
    Next
    ws.Rows(17).RowHeight = 32: ws.Range("A17:L17").WrapText = True: ws.Range("A17:L17").Merge: ws.Range("A17").Value2 = "Selected-location keys with at least " & minCount & " events. Repeat gaps are calendar days between consecutive selected events, not all observation pairs."
    WriteHeaders ws, 18, 1, Array("Location key", "Event count", "Location label", "First date", "Last date", "Mean gap days", "Median gap days", "Minimum gap days", "Gaps <= " & windowA & " days", "Gaps <= " & windowB & " days", "Gaps <= " & windowC & " days", "Additional events")
    If shown = 0 Then
        ws.Range("A19").Value2 = "No locations meet the repeat-count threshold."
        MarkTimeBody ws, ws.Range("A17:L19"): Exit Sub
    End If
    ReDim output(1 To shown, 1 To 12): i = 0
    If ws.Parent.Date1904 Then offset = 1462
    For Each key In groups.Keys
        Set dates = groups(key): n = dates.Count
        If n >= minCount Then
            i = i + 1: ReDim values(1 To n)
            For j = 1 To n: values(j) = CDbl(dates(j)): Next
            SortDates values, 1, n: ReDim gaps(1 To n - 1): a = 0: b = 0: c = 0: sumGap = 0
            For j = 1 To n - 1
                gap = values(j + 1) - values(j): gaps(j) = gap: sumGap = sumGap + gap
                If gap <= windowA Then a = a + 1
                If gap <= windowB Then b = b + 1
                If gap <= windowC Then c = c + 1
            Next
            SortDates gaps, 1, n - 1
            If (n - 1) Mod 2 = 0 Then median = (gaps((n - 1) \ 2) + gaps((n - 1) \ 2 + 1)) / 2 Else median = gaps(n \ 2)
            output(i, 1) = key: output(i, 2) = n: output(i, 3) = labels(key)
            output(i, 4) = values(1) - offset: output(i, 5) = values(n) - offset
            output(i, 6) = sumGap / (n - 1): output(i, 7) = median: output(i, 8) = gaps(1)
            output(i, 9) = a: output(i, 10) = b: output(i, 11) = c: output(i, 12) = n - 1
            If n > maxCount Then maxCount = n
        End If
    Next
    SortConcentration output, 1, shown
    ws.Range("A19").Resize(shown, 12).NumberFormat = "@": ws.Range("A19").Resize(shown, 12).Value2 = output
    ws.Range("B19:B" & shown + 18).NumberFormat = "0": ws.Range("D19:E" & shown + 18).NumberFormat = "mm/dd/yyyy"
    ws.Range("F19:H" & shown + 18).NumberFormat = "0.0": ws.Range("I19:L" & shown + 18).NumberFormat = "0"
    ws.Columns("A").ColumnWidth = 30: ws.Columns("B").ColumnWidth = 18: ws.Columns("C").ColumnWidth = 35: ws.Columns("D:L").ColumnWidth = 16
    ws.Range("A19:A" & shown + 18).WrapText = True: ws.Range("C19:C" & shown + 18).WrapText = True
    TimeTableStyle ws, shown + 18, 12: MarkTimeBody ws, ws.Range("A17:L" & shown + 18)
    CreateRepeatChart ws, output
End Sub

Private Sub CreateRepeatChart(ByVal ws As Worksheet, ByVal output As Variant)
    Dim chart As ChartObject, series As Series, names() As Variant, values() As Variant, i As Long, top As Long
    top = Application.Min(20, UBound(output, 1)): ReDim names(1 To top): ReDim values(1 To top)
    For i = 1 To top: names(i) = CStr(i) & ". " & output(i, 3): values(i) = output(i, 12): Next
    Set chart = ws.ChartObjects.Add(ws.Range("N18").Left, ws.Range("N18").Top, 650, 360)
    chart.Name = "CAPTime_Repeats": chart.Chart.ChartType = xlBarClustered
    chart.Chart.HasTitle = True: chart.Chart.ChartTitle.Text = "Additional selected events at repeat locations"
    Set series = chart.Chart.SeriesCollection.NewSeries: series.Name = "Events after first": series.XValues = names: series.Values = values
    series.Format.Fill.ForeColor.RGB = RGB(20, 95, 140): chart.Chart.HasLegend = False: chart.Chart.Axes(xlCategory).ReversePlotOrder = True
    chart.Chart.ChartArea.Font.Name = "Calibri"
End Sub

Public Function StreetPeriodStart(ByVal origin As Double, ByVal index As Long, ByVal unit As String, ByVal binWidth As Long) As Double
    If unit = "Weeks" Then
        StreetPeriodStart = CDbl(DateAdd("d", (index - 1) * binWidth * 7, CDate(origin)))
    ElseIf unit = "Months" Then
        StreetPeriodStart = CDbl(DateAdd("m", (index - 1) * binWidth, CDate(origin)))
    Else
        Err.Raise vbObjectError + 614, , "Interval unit must be Weeks or Months."
    End If
End Function

Public Function StreetPeriodIndex(ByVal dateValue As Double, ByVal origin As Double, ByVal unit As String, ByVal binWidth As Long) As Long
    If unit = "Weeks" Then
        StreetPeriodIndex = Fix((dateValue - origin) / (7# * binWidth)) + 1
    ElseIf unit = "Months" Then
        StreetPeriodIndex = DateDiff("m", CDate(origin), CDate(dateValue)) \ binWidth + 1
    Else
        Err.Raise vbObjectError + 614, , "Invalid interval unit."
    End If
End Function

Public Function StreetProfile(ByVal counts As Variant, ByVal starts As Variant, ByVal minDate As Double, ByVal maxDate As Double, ByVal threshold As Double, ByVal highVolume As Boolean, Optional ByVal minEvents As Long = HOT_STREETS_MIN_EVENTS) As Variant
    Dim comparisonEvents As Double
    Dim complete As Collection, i As Long, binWidth As Long, early As Double, late As Double, active As Long, ratio As Variant, label As String, k As Long, days As Double
    Set complete = New Collection
    For i = LBound(counts) To UBound(counts)
        If starts(i) >= minDate And starts(i + 1) - 1 <= maxDate Then complete.Add i
    Next
    If complete.Count < 3 Then StreetProfile = Array("Insufficient complete intervals", Empty, Empty, Empty, Empty, complete.Count): Exit Function
    binWidth = Application.Max(1, complete.Count \ 3)
    For i = 1 To complete.Count
        k = complete(i): days = starts(k + 1) - starts(k)
        If counts(k) > 0 Then active = active + 1
        If i <= binWidth Then early = early + counts(k) / days: comparisonEvents = comparisonEvents + counts(k)
        If i > complete.Count - binWidth Then late = late + counts(k) / days: comparisonEvents = comparisonEvents + counts(k)
    Next
    early = early / binWidth: late = late / binWidth
    If comparisonEvents < minEvents Then
        StreetProfile = Array("Low volume; no trend label", early, late, "Below minimum events", active / complete.Count, complete.Count): Exit Function
    End If
    If early = 0 And late > 0 Then
        label = "Emerging from zero baseline": ratio = "No prior activity"
    ElseIf early > 0 Then
        ratio = (late - early) / early
        If ratio >= threshold Then
            label = "Rising (descriptive)"
        ElseIf ratio <= -threshold Then
            label = "Falling (descriptive)"
        End If
    End If
    If Len(label) = 0 Then
        If highVolume And active / complete.Count >= 0.75 Then
            label = "Persistent higher-volume"
        ElseIf active / complete.Count < 0.5 Then
            label = "Intermittent"
        Else
            label = "Stable/mixed"
        End If
    End If
    StreetProfile = Array(label, early, late, ratio, active / complete.Count, complete.Count)
End Function

Public Sub RenderHotStreets(ByVal ws As Worksheet, ByVal events As Collection, ByVal minDate As Double, ByVal maxDate As Double)
    Dim minEvents As Long, unit As String, binWidth As Long, threshold As Variant, origin As Double, p As Long, index As Long, starts() As Variant
    Dim groups As Object, labels As Object, observation As Variant, key As Variant, counts As Variant, output() As Variant
    Dim i As Long, j As Long, total As Long, n As Long, maximum As Long, profile As Variant, highThreshold As Long, offset As Double
    unit = SafeText(ws.Range("B6").Value2): binWidth = IntegerControl(ws, "B7", 1, 120): threshold = ws.Range("B8").Value2
    If ws.Range("A9").Value2 <> "Minimum comparison events" Then
        ws.Range("D9").Value2 = "Grouping key": ws.Range("E9").Value2 = ws.Range("B9").Value2
        ws.Range("A9").Value2 = "Minimum comparison events": ws.Range("B9").Value2 = HOT_STREETS_MIN_EVENTS
        ws.Range("B9").NumberFormat = "0": ws.Range("B9").Interior.Color = RGB(224, 239, 250)
        ws.Range("B9").Validation.Delete
        ws.Range("B9").Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="200000"
    End If
    minEvents = IntegerControl(ws, "B9", 1, 200000)
    If unit <> "Weeks" And unit <> "Months" Then Err.Raise vbObjectError + 615, , "Select Weeks or Months."
    If IsError(threshold) Or Not IsNumeric(threshold) Then Err.Raise vbObjectError + 616, , "Invalid descriptive rise/fall threshold."
    If CDbl(threshold) < 0.01 Or CDbl(threshold) > 10 Then Err.Raise vbObjectError + 616, , "Threshold must be 1% to 1000%."
    If minDate = 0 Or maxDate = 0 Or events.Count = 0 Then
        ws.Range("A17").Value2 = "No eligible dated activity for this selection."
        MarkTimeBody ws, ws.Range("A17"): Exit Sub
    End If
    If unit = "Weeks" Then origin = minDate - Weekday(CDate(minDate), vbMonday) + 1 Else origin = CDbl(DateSerial(Year(CDate(minDate)), Month(CDate(minDate)), 1))
    p = StreetPeriodIndex(maxDate, origin, unit, binWidth)
    If p > 2000 Then Err.Raise vbObjectError + 617, , "More than 2000 intervals: increase units per interval."
    ReDim starts(1 To p + 1)
    For i = 1 To p + 1: starts(i) = StreetPeriodStart(origin, i, unit, binWidth): Next
    Set groups = CreateObject("Scripting.Dictionary"): Set labels = CreateObject("Scripting.Dictionary")
    For Each observation In events
        key = observation(3)
        If Not groups.Exists(key) Then
            If (groups.Count + 1) * CDbl(p + 9) > 400000 Then Err.Raise vbObjectError + 618, , "Timeline exceeds 400,000 cells: increase interval binWidth or narrow the source geography."
            ReDim counts(1 To p): groups.Add key, counts: labels.Add key, observation(4)
        End If
        counts = groups(key): index = StreetPeriodIndex(observation(1), origin, unit, binWidth)
        If IsEmpty(counts(index)) Then counts(index) = 0
        counts(index) = counts(index) + 1: groups(key) = counts
    Next
    n = groups.Count: ReDim output(1 To n, 1 To p + 9): i = 0
    For Each key In groups.Keys
        i = i + 1: counts = groups(key): total = 0
        output(i, 1) = key: output(i, 3) = labels(key)
        For j = 1 To p
            If IsEmpty(counts(j)) Then counts(j) = 0
            output(i, j + 9) = counts(j): total = total + counts(j)
            If counts(j) > maximum Then maximum = counts(j)
        Next
        output(i, 2) = total: groups(key) = counts
    Next
    SortConcentration output, 1, n: highThreshold = output(-Int(-0.25 * n), 2)
    For i = 1 To n
        counts = groups(output(i, 1))
        profile = StreetProfile(counts, starts, minDate, maxDate, CDbl(threshold), output(i, 2) >= highThreshold, minEvents)
        For j = 0 To 5: output(i, j + 4) = profile(j): Next
    Next
    ws.Rows(17).RowHeight = 60: ws.Range("A17:I17").WrapText = True: ws.Range("A17:I17").Merge: ws.Range("A17").Value2 = "Descriptive profiles: complete calendar bins inside observed date span only. Matrix retains partial edge bins and empty intervening periods; reporting completeness is unverified. Labels require at least " & minEvents & " events across the early and late windows combined. Minimum is user-controlled in B9."
    WriteHeaders ws, 18, 1, Array("Location key", "Total events", "Location label", "Descriptive profile", "Early mean/day", "Late mean/day", "Relative change", "Active bin share", "Complete bins")
    If ws.Parent.Date1904 Then offset = 1462
    For j = 1 To p: ws.Cells(18, j + 9).Value2 = starts(j) - offset: Next
    ws.Range("J18").Resize(1, p).NumberFormat = IIf(unit = "Weeks", "mm/dd/yy", "mmm yyyy")
    ws.Range("A19").Resize(n, p + 9).NumberFormat = "@": ws.Range("A19").Resize(n, p + 9).Value2 = output
    ws.Range("B19:B" & n + 18).NumberFormat = "0": ws.Range("E19:F" & n + 18).NumberFormat = "0.000"
    ws.Range("G19:H" & n + 18).NumberFormat = "0.0%": ws.Range("I19").Resize(n, p + 1).NumberFormat = "0"
    ApplyHeatColors ws.Range("J19").Resize(n, p), maximum
    ws.Columns("A").ColumnWidth = 28: ws.Columns("B").ColumnWidth = 18: ws.Columns("C:D").ColumnWidth = 32
    ws.Columns("E:I").ColumnWidth = 16: ws.Range("J18").Resize(n + 1, p).Columns.ColumnWidth = 12
    TimeTableStyle ws, n + 18, p + 9: MarkTimeBody ws, ws.Range(ws.Cells(17, 1), ws.Cells(n + 18, p + 9))
    CreateStreetProfileChart ws, output, starts, minDate, maxDate, p
End Sub

Private Sub CreateStreetProfileChart(ByVal ws As Worksheet, ByVal output As Variant, ByVal starts As Variant, ByVal minDate As Double, ByVal maxDate As Double, ByVal p As Long)
    Dim sums As Object, i As Long, j As Long, label As String, group As Variant, key As Variant, x() As Variant, y() As Variant, days As Double
    Dim chart As ChartObject, series As Series, offset As Double
    Set sums = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(output, 1)
        label = output(i, 4)
        If Not sums.Exists(label) Then
            ReDim group(0 To p)
            For j = 0 To p: group(j) = 0: Next
            sums.Add label, group
        End If
        group = sums(label): group(0) = group(0) + 1
        For j = 1 To p
            days = Application.Min(maxDate, starts(j + 1) - 1) - Application.Max(minDate, starts(j)) + 1
            group(j) = group(j) + output(i, j + 9) / days
        Next
        sums(label) = group
    Next
    If ws.Parent.Date1904 Then offset = 1462
    ReDim x(1 To p): For j = 1 To p: x(j) = starts(j) - offset: Next
    Set chart = ws.ChartObjects.Add(ws.Range("J2").Left, ws.Range("J2").Top, 720, 300)
    chart.Name = "CAPTime_StreetProfiles": chart.Chart.ChartType = xlLineMarkers
    chart.Chart.HasTitle = True: chart.Chart.ChartTitle.Text = "Descriptive street profiles | mean events per day of observed span"
    For Each key In sums.Keys
        group = sums(key): ReDim y(1 To p)
        For j = 1 To p: y(j) = group(j) / group(0): Next
        Set series = chart.Chart.SeriesCollection.NewSeries
        series.Name = CStr(key): series.XValues = x: series.Values = y
    Next
    chart.Chart.Axes(xlCategory).TickLabels.NumberFormat = "mmm yy": chart.Chart.Legend.Position = xlLegendPositionBottom
    chart.Chart.ChartArea.Font.Name = "Calibri"
End Sub
