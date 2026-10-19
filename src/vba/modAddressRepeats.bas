Attribute VB_Name = "modAddressRepeats"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit
Public Sub RenderAddressRepeats(ByVal ws As Worksheet, ByVal observations As Collection)
    Dim groups As Object, labels As Object, observation As Variant, key As Variant, minimum As Variant
    Dim output() As Variant, i As Long, n As Long, chart As ChartObject, series As Series, names() As Variant, values() As Variant, top As Long
    minimum = ws.Range("B6").Value2
    If IsError(minimum) Or Not IsNumeric(minimum) Then Err.Raise vbObjectError + 740, , "Minimum events must be an integer of at least two."
    If minimum < 2 Or minimum > 36500 Or minimum <> Fix(minimum) Then Err.Raise vbObjectError + 740, , "Minimum events must be 2-36500."
    Set groups = CreateObject("Scripting.Dictionary"): Set labels = CreateObject("Scripting.Dictionary")
    For Each observation In observations
        key = observation(3)
        If Not groups.Exists(key) Then groups.Add key, 0: labels.Add key, observation(4)
        groups(key) = groups(key) + 1
    Next
    For Each key In groups.Keys: If groups(key) >= minimum Then n = n + 1
    Next
    ws.Range("A17:C17").Merge: ws.Range("A17").Value2 = "Address-specific counts; no date/time required. Additional events = selected events minus first event."
    ws.Range("A17").WrapText = True: ws.Rows(17).RowHeight = 32
    WriteHeaders ws, 18, 1, Array("Address", "Event count", "Additional events")
    If n = 0 Then
        ws.Range("A19").Value2 = "No addresses meet the repeat-count threshold."
        MarkTimeBody ws, ws.Range("A17:C19"): Exit Sub
    End If
    ReDim output(1 To n, 1 To 4)
    For Each key In groups.Keys
        If groups(key) >= minimum Then
            i = i + 1: output(i, 1) = key: output(i, 2) = groups(key): output(i, 3) = labels(key): output(i, 4) = groups(key) - 1
        End If
    Next
    SortConcentration output, 1, n
    Dim visible() As Variant
    ReDim visible(1 To n, 1 To 3)
    For i = 1 To n: visible(i, 1) = output(i, 3): visible(i, 2) = output(i, 2): visible(i, 3) = output(i, 4): Next
    ws.Range("A19").Resize(n, 3).NumberFormat = "@": ws.Range("A19").Resize(n, 3).Value2 = visible
    ws.Range("B19:B" & n + 18).NumberFormat = "0": ws.Range("C19:C" & n + 18).NumberFormat = "0"
    ws.Columns("A").ColumnWidth = 38: ws.Columns("B").ColumnWidth = 18: ws.Columns("C").ColumnWidth = 35: ws.Columns("D").ColumnWidth = 20
    TimeTableStyle ws, n + 18, 3: MarkTimeBody ws, ws.Range("A17:C" & n + 18)
    top = Application.Min(20, n): ReDim names(1 To top): ReDim values(1 To top)
    For i = 1 To top: names(i) = output(i, 3): values(i) = output(i, 4): Next
    Set chart = ws.ChartObjects.Add(ws.Range("F18").Left, ws.Range("F18").Top, 650, 360)
    chart.Name = "CAPTime_Repeats": chart.Chart.ChartType = xlBarClustered
    chart.Chart.HasTitle = True: chart.Chart.ChartTitle.Text = "Additional selected events at repeat addresses"
    Set series = chart.Chart.SeriesCollection.NewSeries: series.XValues = names: series.Values = values
    chart.Chart.HasLegend = False: chart.Chart.Axes(xlCategory).ReversePlotOrder = True
End Sub
