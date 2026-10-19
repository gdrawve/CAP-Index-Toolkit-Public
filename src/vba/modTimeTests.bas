Attribute VB_Name = "modTimeTests"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit
Private checks As Long
Private Sub CheckTime(ByVal condition As Boolean, ByVal description As String)
    checks = checks + 1
    If Not condition Then Err.Raise vbObjectError + 650, , "FAIL: " & description
End Sub

Public Function RunTimeAnalysisTests(Optional ByVal previewDir As String = "") As String
    Dim wb As Workbook, source As Worksheet, heat As Worksheet, repeats As Worksheet, streets As Worksheet
    Dim a(1 To 9, 1 To 5) As Variant, i As Long, j As Long, oldEvents As Boolean, oldScreen As Boolean
    Dim starts(1 To 7) As Variant, counts(1 To 6) As Variant, profile As Variant, caught As Long
    Dim cache As Range, body As Range, picture As ChartObject, description As String, errNo As Long
    On Error GoTo failed
    checks = 0: oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Set wb = Workbooks.Add(xlWBATWorksheet): Set source = wb.Worksheets(1)
    a(1, 1) = "IncidentID": a(1, 2) = "OccurredAt": a(1, 3) = "CrimeType": a(1, 4) = "CAP_BlockKey": a(1, 5) = "CAP_HundredBlock"
    For i = 2 To 9
        a(i, 1) = "I" & i: a(i, 3) = "Theft": a(i, 4) = "IL:TEST:A": a(i, 5) = "1200 BLOCK MAIN ST"
    Next
    a(2, 2) = "01/05/2026 14:10": a(3, 2) = "01/05/2026 14:10": a(3, 1) = a(2, 1)
    a(4, 2) = "01/12/2026 00:00": a(5, 2) = "02/15/2026": a(6, 2) = "invalid"
    a(7, 2) = "04/20/2026 23:00": a(7, 3) = "Robbery": a(7, 4) = "IL:TEST:B": a(7, 5) = "1300 BLOCK OAK ST"
    a(8, 2) = "06/30/2026 08:00": a(9, 2) = "06/30/2026 08:00": a(9, 1) = ""
    source.Range("A1:E9").NumberFormat = "@": source.Range("A1:E9").Value2 = a
    BuildTemporalHeatmap source.Range("A1:E9"), "OccurredAt", "CrimeType", "IncidentID", "", False
    Set heat = wb.Worksheets(2)
    CheckTime heat.Range("B10").Value2 = 4, "Heatmap unique timed incidents"
    CheckTime Application.Sum(heat.Range("B19:Y25")) = 4, "Heatmap reconciles eligible count"
    CheckTime heat.Range("P19").Value2 = 1 And heat.Range("B19").Value2 = 1, "Monday 14 and explicit midnight"
    CheckTime heat.Range("A25").Value2 = "Sunday" And heat.Range("Y18").Value2 = 23, "Seven days and 24 hours"
    CheckTime heat.Range("B19:Y25").FormatConditions.Count = 1, "Native conditional color scale"
    CheckTime heat.Range("B19").FormatConditions(1).ColorScaleCriteria(1).Value = 0, "Zero color anchor"
    CheckTime heat.Range("B12").Value2 = 1 And heat.Range("D12").Value2 = "Unknown hours: 1", "Invalid date and unknown hour reported"
    Application.EnableEvents = True
    heat.Range("B4").Value2 = "Type: ROBBERY"
    CheckTime heat.Range("B10").Value2 = 1 And heat.Range("Y19").Value2 = 1, "Heat crime filter event"
    heat.Range("B4").Value2 = "All crime types": heat.Range("B5").Value2 = "Source records"
    CheckTime heat.Range("B10").Value2 = 6 And heat.Range("P19").Value2 = 2, "Offense rows remain distinct"
    heat.Range("B5").Value2 = "Unique incidents"
    Application.EnableEvents = False
    If Len(previewDir) > 0 Then
        ExportRangePNG heat.Range("A17:Y27"), previewDir & "\temporal-heatmap.png"
    End If
    BuildRepeatLocationsWithGaps source.Range("A1:E9"), "CAP_BlockKey", "OccurredAt", "CrimeType", "IncidentID", "CAP_HundredBlock", False
    Set repeats = wb.Worksheets(4)
    CheckTime repeats.Range("B19").Value2 = 4 And repeats.Range("L19").Value2 = 3, "Repeat unique count and additional events"
    CheckTime repeats.Range("I19").Value2 = 1 And repeats.Range("J19").Value2 = 1 And repeats.Range("K19").Value2 = 2, "Consecutive gaps within 7/30/90 days"
    CheckTime repeats.Range("H19").Value2 = 7 And repeats.Range("G19").Value2 = 34, "Gap minimum and median"
    CheckTime repeats.ChartObjects.Count = 1, "Repeat chart"
    If Len(previewDir) > 0 Then repeats.Activate: ExportTimeChart repeats.ChartObjects(1), previewDir & "\repeat-locations.png"
    Application.EnableEvents = True: repeats.Range("B6").Value2 = 5
    CheckTime repeats.ChartObjects.Count = 0 And InStr(repeats.Range("A19").Value2, "No locations") > 0, "Repeat minimum event refresh"
    repeats.Range("B6").Value2 = 2
    Application.EnableEvents = False
    BuildHotStreets source.Range("A1:E9"), "CAP_BlockKey", "OccurredAt", "CrimeType", "IncidentID", "CAP_HundredBlock", False
    Set streets = wb.Worksheets(6)
    CheckTime streets.Range("B19").Value2 = 4 And streets.Range("B20").Value2 = 1, "Street unique totals"
    CheckTime streets.Range("J18").Value2 = CDbl(DateSerial(2026, 1, 1)) And streets.Range("O18").Value2 = CDbl(DateSerial(2026, 6, 1)), "Entire source span monthly axis"
    CheckTime streets.Range("L19").Value2 = 0 And Application.Sum(streets.Range("J19:O20")) = 5, "Empty periods retained and matrix reconciles"
    CheckTime streets.Range("I19").Value2 = 5, "Partial January excluded from profiles"
    If Len(previewDir) > 0 Then streets.Activate: ExportTimeChart streets.ChartObjects(1), previewDir & "\hot-street-profiles.png"
    Application.EnableEvents = True: streets.Range("B4").Value2 = "Type: ROBBERY"
    CheckTime streets.Range("J18").Value2 = CDbl(DateSerial(2026, 1, 1)) And streets.Range("O18").Value2 = CDbl(DateSerial(2026, 6, 1)), "Crime filter preserves full timeline"
    streets.Range("B7").Value2 = 2
    CheckTime streets.Range("L18").Value2 = CDbl(DateSerial(2026, 5, 1)) And streets.Range("D19").Value2 = "Insufficient complete intervals", "Multi-month parameter and insufficient history"
    streets.Range("B6").Value2 = "Weeks": streets.Range("B7").Value2 = 1
    CheckTime streets.Range("J18").Value2 = CDbl(DateSerial(2026, 1, 5)), "Week starts Monday"
    Application.EnableEvents = False: streets.Range("B7").Value2 = 2: streets.Range("B14").Calculate
    CheckTime InStr(streets.Range("B14").Value2, "Controls changed") = 1, "Disabled events show stale interval parameters"
    streets.Range("B7").Value2 = 0
    On Error Resume Next: RefreshTimeReport streets, True: caught = Err.Number: Err.Clear: On Error GoTo failed
    CheckTime caught <> 0 And streets.ChartObjects.Count = 0 And InStr(streets.Range("B14").Value2, "Refresh failed") = 1, "Invalid interval clears stale output"
    For i = 1 To 7: starts(i) = CDbl(DateSerial(2026, 1, 1)) + (i - 1) * 7: Next
    For i = 1 To 6: counts(i) = IIf(i <= 2, 1, 4): Next
    profile = StreetProfile(counts, starts, starts(1), starts(7) - 1, 0.25, False)
    CheckTime profile(0) = "Rising (descriptive)" And profile(3) = 3, "Known rising descriptive profile"
    For i = 1 To 6: counts(i) = IIf(i <= 2, 4, 1): Next
    profile = StreetProfile(counts, starts, starts(1), starts(7) - 1, 0.25, False)
    CheckTime profile(0) = "Falling (descriptive)" And profile(3) = -0.75, "Known falling profile"
    For i = 1 To 6: counts(i) = IIf(i <= 2, 0, 4): Next
    profile = StreetProfile(counts, starts, starts(1), starts(7) - 1, 0.25, False)
    CheckTime profile(0) = "Emerging from zero baseline", "Zero baseline handled without divide by zero"
    For i = 1 To 6: counts(i) = 4: Next
    profile = StreetProfile(counts, starts, starts(1), starts(7) - 1, 0.25, True)
    CheckTime profile(0) = "Persistent higher-volume", "Persistent high volume"
    For i = 1 To 6: counts(i) = 0: Next
    counts(1) = 1: counts(5) = 2
    profile = StreetProfile(counts, starts, starts(1), starts(7) - 1, 0.25, False, 10)
    CheckTime profile(0) = "Low volume; no trend label" And profile(3) = "Below minimum events" And profile(1) > 0 And profile(2) > 0, "User minimum suppresses sparse labels and retains rates"
    counts(1) = 0: counts(5) = 3
    profile = StreetProfile(counts, starts, starts(1), starts(7) - 1, 0.25, False, 10)
    CheckTime profile(0) = "Low volume; no trend label", "Sparse emerging suppressed"
    counts(5) = 12
    profile = StreetProfile(counts, starts, starts(1), starts(7) - 1, 0.25, False, 10)
    CheckTime profile(0) = "Emerging from zero baseline", "Emerging above chosen minimum"
    counts(1) = 6
    profile = StreetProfile(counts, starts, starts(1), starts(7) - 1, 0.25, False, 10)
    CheckTime profile(0) = "Rising (descriptive)", "Rising above chosen minimum"
    counts(1) = 12: counts(5) = 6
    profile = StreetProfile(counts, starts, starts(1), starts(7) - 1, 0.25, False, 10)
    CheckTime profile(0) = "Falling (descriptive)", "Falling above chosen minimum"
    counts(1) = 0: counts(5) = 1: counts(3) = 100
    profile = StreetProfile(counts, starts, starts(1), starts(7) - 1, 0.25, False)
    CheckTime profile(0) = "Low volume; no trend label", "Default minimum two excludes middle-window counts"
    counts(5) = 2
    profile = StreetProfile(counts, starts, starts(1), starts(7) - 1, 0.25, False)
    CheckTime profile(0) = "Emerging from zero baseline", "Default minimum two includes boundary"
    profile = StreetProfile(counts, starts, starts(1), starts(3) - 1, 0.25, False)
    CheckTime profile(0) = "Insufficient complete intervals", "Insufficient intervals takes priority"
    streets.Range("B7").Value2 = 1: streets.Range("B6").Value2 = "Months"
    streets.Range("B4").Value2 = "All crime types": streets.Range("B9").Value2 = 200000
    RefreshTimeReport streets, True
    CheckTime streets.Range("D19").Value2 = "Low volume; no trend label" And streets.Range("B9").Validation.Type = xlValidateWholeNumber, "Sheet minimum controls rendered labels"
    Application.EnableEvents = True: streets.Range("B9").Value2 = 2
    CheckTime InStr(streets.Range("A17").Value2, "at least 2 events") > 0, "Minimum change refreshes report"
    Application.EnableEvents = False: streets.Range("B9").Value2 = 3: streets.Range("B14").Calculate
    CheckTime InStr(streets.Range("B14").Value2, "Controls changed") = 1, "Minimum included in stale-report detection"
    CheckTime source.Range("A3").Value2 = "I2" And source.UsedRange.Columns.Count = 5, "Source unchanged"
    wb.Close False: Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    RunTimeAnalysisTests = checks & " temporal/repeat/street checks passed": Exit Function
failed:
    errNo = Err.Number: description = Err.Description
    On Error Resume Next: If Not wb Is Nothing Then wb.Close False
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    On Error GoTo 0: Err.Raise errNo, "Time analysis tests", description
End Function

Private Sub ExportTimeChart(ByVal picture As ChartObject, ByVal path As String)
    picture.Activate: picture.Chart.Refresh: DoEvents
    If Not picture.Chart.Export(path, "PNG") Then Err.Raise vbObjectError + 651, , "Chart export failed"
    If FileLen(path) = 0 Then Err.Raise vbObjectError + 652, , "Empty chart preview"
End Sub
