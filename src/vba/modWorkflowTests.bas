Attribute VB_Name = "modWorkflowTests"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit
Private checks As Long
Private Sub CheckFlow(ByVal condition As Boolean, ByVal description As String)
    checks = checks + 1
    If Not condition Then Err.Raise vbObjectError + 760, , "FAIL: " & description
End Sub

Public Function RunWorkflowTests(Optional ByVal previewDir As String = "") As String
    Dim wb As Workbook, ws As Worksheet, blocks As Worksheet, randomized As Worksheet, geocoded As Worksheet, report As Worksheet
    Dim a(1 To 7, 1 To 8) As Variant, i As Long, h As Long, fake As clsCensusFixture, oldEvents As Boolean
    Dim caught As Long, description As String, errNo As Long, picture As ChartObject, oldScreen As Boolean
    On Error GoTo failed
    checks = 0: oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating: Application.EnableEvents = False
    CheckFlow HourFromTime(0#, False, h) And h = 0, "Numeric time midnight"
    CheckFlow HourFromTime("12:00 AM", False, h) And h = 0, "AM midnight"
    CheckFlow HourFromTime("1:25 PM", False, h) And h = 13, "PM time"
    CheckFlow HourFromTime("23:59:59", False, h) And h = 23, "24-hour time"
    CheckFlow Not HourFromTime("24:00", False, h), "Invalid clock hour"
    CheckFlow Not HourFromTime("", False, h), "Blank time"
    CheckFlow Not HourFromTime("01/05/2026", False, h), "Date only does not supply hour"
    CheckFlow DayOfWeekNumber("Sun") = 7 And DayOfWeekNumber("Monday") = 1 And DayOfWeekNumber(1) = 0, "Unambiguous day text"
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    a(1, 1) = "Address": a(1, 2) = "City": a(1, 3) = "State": a(1, 4) = "IncidentID"
    a(1, 5) = "CrimeType": a(1, 6) = "Date": a(1, 7) = "Time": a(1, 8) = "Day of Week"
    For i = 2 To 7
        a(i, 1) = "1200 MAIN ST": a(i, 2) = "TESTCITY": a(i, 3) = "IL": a(i, 4) = "I" & i
        a(i, 5) = "Theft": a(i, 6) = "01/05/2026": a(i, 7) = "2:15 PM": a(i, 8) = "Monday"
    Next
    a(3, 4) = a(2, 4): a(4, 1) = "12XX OAK ST": a(4, 7) = "12:00 AM": a(4, 8) = "Tuesday": a(4, 6) = "01/06/2026"
    a(5, 1) = "1234 ELM ST APT 2": a(5, 7) = "": a(6, 1) = "MAIN & OAK": a(6, 7) = "bad"
    a(7, 1) = "1200 MAIN ST": a(7, 2) = "OTHER CITY": a(7, 5) = "Robbery": a(7, 8) = "bad": a(7, 6) = "06/30/2026"
    ws.Range("A1:H7").NumberFormat = "@": ws.Range("A1:H7").Value2 = a
    ws.Protect
    BuildHundredBlocks ws.Range("A1:H7"), showSummary:=False
    Set blocks = wb.Worksheets(2)
    CheckFlow blocks.Name Like "Hundred_Blocks_*" And blocks.UsedRange.Columns.Count = 10, "Named compact hundred-block sheet"
    CheckFlow blocks.Range("I6").Value2 = "CAP_HundredBlock" And blocks.Range("J6").Value2 = "Review Status", "Only two generated block fields"
    CheckFlow blocks.Range("I10").Value2 = "1200 BLOCK ELM ST" And Len(blocks.Range("I11").Value2) = 0, "Block conversion and review"
    CheckFlow ws.ProtectContents And ws.UsedRange.Columns.Count = 8, "Protected source is untouched"
    BuildHundredBlocks ws.Range("A1:H7"), "Address", "City", "State", "IncidentID", True, 42, False
    Set randomized = wb.Worksheets(3)
    CheckFlow randomized.Name Like "Randomized_Addresses_*" And randomized.UsedRange.Columns.Count = 11, "Named compact randomized sheet"
    CheckFlow randomized.Range("J7").Value2 = randomized.Range("J8").Value2, "Same incident stable synthetic number"
    CheckFlow randomized.Range("J10").Value2 = "1234 ELM ST APT 2" And InStr(randomized.Range("K7").Value2, "Synthetic") > 0, "Original exact address retained and synthetic clearly labeled"
    CheckFlow InStr(randomized.Range("A3").Value2, "Seed=42") > 0, "Seed recorded once"
    Set fake = New clsCensusFixture
    GeocodeCensus randomized.Range("A6:K12"), "CAP_SyntheticAddress", "City", "State", "", fake, 10000, False
    Set geocoded = wb.Worksheets(4)
    CheckFlow geocoded.Name Like "Geocoded_Addresses_*" And geocoded.UsedRange.Columns.Count = 18, "Named geocode sheet adds only seven fields"
    CheckFlow geocoded.Range("L6").Value2 = "Full Address" And geocoded.Range("R6").Value2 = "Review Flag", "Minimal geocoder field set"
    CheckFlow InStr(geocoded.Range("R7").Value2, "Synthetic") > 0 And geocoded.Range("N7").Value2 > 38, "Synthetic warning survives coordinate response"
    CheckFlow InStr(fake.Captured, "Theft") = 0 And InStr(fake.Captured, "I2") = 0, "No crime or incident IDs submitted"
    CheckFlow randomized.UsedRange.Columns.Count = 11 And ws.UsedRange.Columns.Count = 8, "Geocoding source preservation"
    ws.Unprotect
    BuildHourFields ws.Range("G1:G7"), False
    CheckFlow ws.Range("H1").Value2 = "Hour (0-23)" And ws.Range("H2").Value2 = 14 And ws.Range("H4").Value2 = 0, "Hour inserts only one adjacent field"
    CheckFlow IsEmpty(ws.Range("H5").Value2) And IsEmpty(ws.Range("H6").Value2), "Missing/invalid times stay blank"
    BuildHourFields ws.Range("G1:G7"), False
    CheckFlow ws.UsedRange.Columns.Count = 9, "Hour rerun refreshes instead of inserting again"
    BuildTemporalHeatmap ws.Range("A1:I7"), "Date", "CrimeType", "IncidentID", "", False, "", "Time"
    Set report = wb.Worksheets(5)
    CheckFlow report.Range("P19").Value2 = 1 And report.Range("P20").Value2 = 1 And report.Range("B20").Value2 = 1 And report.Range("B10").Value2 = 3, "Separate Date/Time heat map"
    BuildTemporalHeatmap ws.Range("A1:I7"), "", "CrimeType", "IncidentID", "Hour (0-23)", False, "Day of Week"
    Set report = wb.Worksheets(7)
    CheckFlow report.Range("P19").Value2 = 1 And report.Range("B20").Value2 = 1 And report.Range("B10").Value2 = 2, "Direct Day of Week/Hour heat map"
    CheckFlow report.Range("B12").Value2 = 1, "Invalid day reported separately"
    If Len(previewDir) > 0 Then
        ExportRangePNG report.Range("A1:Y27"), previewDir & "\heatmap-day-hour.png"
    End If
    ws.Range("A6").Value2 = "1200 MAIN ST"
    BuildRepeatLocations ws.Range("A1:E7"), "Address", "", "CrimeType", "IncidentID", "", False, "City", "State"
    Set report = wb.Worksheets(9)
    CheckFlow report.Range("B10").Value2 = 5 And report.Range("B19").Value2 = 2 And report.Range("C19").Value2 = 1, "Address repeats need no date/time and distinguish cities"
    CheckFlow report.Range("C18").Value2 = "Additional events" And IsEmpty(report.Range("D18").Value2), "Three repeat result fields"
    If Len(previewDir) > 0 Then
        Application.ScreenUpdating = True: report.Activate: report.ChartObjects(1).Activate
        report.ChartObjects(1).Chart.Refresh: DoEvents
        If Not report.ChartObjects(1).Chart.Export(previewDir & "\address-repeats.png", "PNG") Then Err.Raise vbObjectError + 762, , "Address preview failed"
    End If
    Application.EnableEvents = True: report.Range("B5").Value2 = "Source records"
    CheckFlow report.Range("B19").Value2 = 3, "Address offense record selector"
    report.Range("B4").Value2 = "Type: ROBBERY"
    CheckFlow InStr(report.Range("A19").Value2, "No addresses") > 0, "Address crime filter"
    Application.EnableEvents = False
    BuildHotStreets blocks.Range("A6:J12"), "CAP_HundredBlock", "Date", "CrimeType", "IncidentID", "CAP_HundredBlock", False, "City", "State"
    Set report = wb.Worksheets(11)
    CheckFlow report.Range("B10").Value2 = 4 And report.Range("O18").Value2 = CDbl(DateSerial(2026, 6, 1)), "Hot streets uses hundred block/crime/date with no time requirement"
    BuildConcentrationReport blocks.Range("A6:J12"), "CAP_HundredBlock", "CrimeType", "IncidentID", "CAP_HundredBlock", False, "City", "State"
    Set report = wb.Worksheets(13)
    CheckFlow report.Range("B8").Value2 = 4 And report.Range("B9").Value2 = 4, "Concentration constructs qualified block keys internally"
    ws.Range("F2").Value2 = CDbl(DateSerial(2026, 1, 5)): ws.Range("F2:F7").NumberFormat = "mm/dd/yyyy"
    BuildHundredBlocks ws.Range("A1:E7"), showSummary:=False
    Set report = wb.Worksheets(wb.Worksheets.Count)
    CheckFlow report.UsedRange.Columns.Count = 7 And report.ListObjects.Count = 1, "Only selected source fields are copied"
    BuildHundredBlocks ws.Range("A1:I7"), showSummary:=False
    Set report = wb.Worksheets(wb.Worksheets.Count)
    CheckFlow report.Range("F7").Text = "01/05/2026", "Source date formatting is retained"
    Set fake = New clsCensusFixture: fake.FailAt = 2
    On Error Resume Next: GeocodeCensus ws.Range("A1:E7"), "Address", "City", "State", "", fake, 1, False: caught = Err.Number: Err.Clear: On Error GoTo failed
    CheckFlow caught <> 0 And wb.Worksheets(wb.Worksheets.Count).Range("J7").Value2 = "Match", "New geocoder sheet retains completed batches on failure"
    BuildHundredBlocks blocks.Range("A6:J12"), "Address", "City", "State", "IncidentID", True, 42, False
    Set report = wb.Worksheets(wb.Worksheets.Count)
    CheckFlow report.UsedRange.Columns.Count = 11 And report.Range("I6").Value2 = "CAP_HundredBlock", "Chained block/randomizer output reuses existing generated headers"
    CheckFlow report.Range("K7").Value2 = randomized.Range("J7").Value2, "Chained synthetic numbers match original seeded run"
    Set fake = New clsCensusFixture
    GeocodeCensus geocoded.Range("A6:R12"), "CAP_SyntheticAddress", "City", "State", "", fake, 10000, False
    Set report = wb.Worksheets(wb.Worksheets.Count)
    CheckFlow report.UsedRange.Columns.Count = 18, "Geocode rerun does not duplicate existing result columns"
    blocks.Range("B12").Value2 = ""
    BuildHotStreets blocks.Range("A6:J12"), "CAP_HundredBlock", "Date", "CrimeType", "IncidentID", "CAP_HundredBlock", False, "City", "State"
    Set report = wb.Worksheets(wb.Worksheets.Count - 1)
    CheckFlow report.Range("B10").Value2 = 3, "Mapped missing geography is excluded from Hot Streets"
    BuildConcentrationReport blocks.Range("A6:J12"), "CAP_HundredBlock", "CrimeType", "IncidentID", "CAP_HundredBlock", False, "City", "State"
    Set report = wb.Worksheets(wb.Worksheets.Count - 1)
    CheckFlow report.Range("B8").Value2 = 3 And report.Range("B9").Value2 = 3, "Mapped missing geography is excluded from concentration"
    ws.Range("K1").Value2 = "Second Time": ws.Range("K2").Value2 = "15:45"
    BuildHourFields ws.Range("K1:K2"), False
    CheckFlow ws.Range("L1").Value2 = "Hour from Time (0-23)" And ws.Range("L2").Value2 = 15, "Existing timestamp Hour field does not create duplicate header"
    BuildHourFields ws.Range("K1:K2"), False
    CheckFlow ws.Range("L2").Value2 = 15 And ws.UsedRange.Columns.Count = 12, "Alternate Hour header refreshes safely"
    wb.Close False: Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    RunWorkflowTests = "PASS: " & checks & " revised workflow checks": Exit Function
failed:
    errNo = Err.Number: description = Err.Description
    On Error Resume Next: If Not wb Is Nothing Then wb.Close False
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    On Error GoTo 0: Err.Raise errNo, "Revised workflow tests", description
End Function
