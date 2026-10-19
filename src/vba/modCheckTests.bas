Attribute VB_Name = "modCheckTests"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
Option Explicit
Private checks As Long
Private Sub CheckQuality(ByVal condition As Boolean, ByVal label As String)
    checks = checks + 1
    If Not condition Then Err.Raise vbObjectError + 920, , label
End Sub
Private Function Metric(ByVal ws As Worksheet, ByVal label As String) As Variant
    Dim found As Range
    Set found = ws.Columns(1).Find(What:=label, LookIn:=xlValues, LookAt:=xlWhole)
    If found Is Nothing Then Err.Raise vbObjectError + 921, , "Missing metric: " & label
    Metric = found.Offset(0, 1).Value2
End Function
Public Function RunDatasetCheckTests(ByVal previewDir As String) As String
    Dim wb As Workbook, source As Worksheet, report As Worksheet, guide As Worksheet, a(1 To 9, 1 To 7) As Variant, i As Long, j As Long, found As Range
    Dim oldEvents As Boolean, errNo As Long, description As String
    oldEvents = Application.EnableEvents
    On Error GoTo failed
    Application.EnableEvents = False: checks = 0
    Set wb = Workbooks.Add(xlWBATWorksheet): Set source = wb.Worksheets(1)
    a(1, 1) = "Date": a(1, 2) = "CrimeType": a(1, 3) = "IncidentID": a(1, 4) = "Address": a(1, 5) = "Latitude": a(1, 6) = "Longitude": a(1, 7) = "OffenseID"
    For i = 2 To 8
        a(i, 1) = "01/05/2026": a(i, 2) = "Theft": a(i, 3) = "I1": a(i, 4) = "100 MAIN ST": a(i, 5) = 34: a(i, 6) = -92: a(i, 7) = "O" & i
    Next
    a(4, 1) = "03/05/2026": a(4, 3) = "I2": a(4, 5) = -92: a(4, 6) = 34
    a(5, 1) = "bad date": a(5, 3) = "I3": a(5, 2) = "Robbery": a(5, 5) = Empty: a(5, 6) = Empty
    a(6, 3) = "I4": a(6, 5) = 0: a(6, 6) = 0: a(7, 3) = "I4": a(7, 1) = "03/06/2026"
    For j = 1 To 7: a(8, j) = a(2, j): Next
    source.Range("A1:G9").NumberFormat = "@": source.Range("A1:G9").Value2 = a
    BuildDatasetCheck source.Range("A1:G9"), showSummary:=False: Set report = wb.Worksheets(2)
    CheckQuality Metric(report, "Source records") = 8 And Metric(report, "Blank rows") = 1, "Selected records and blank rows distinguished"
    CheckQuality Metric(report, "Exact duplicate extra rows") = 1, "Exact duplicate extra row detected"
    CheckQuality Metric(report, "Recorded unique incident IDs") = 4 And Metric(report, "IDs with conflicting dates") = 1, "Multi-offense IDs and conflicting dates distinguished"
    CheckQuality Metric(report, "Missing dates") = 1 And Metric(report, "Invalid dates") = 1, "Missing and invalid dates distinguished"
    CheckQuality Metric(report, "Missing coordinate pairs") = 2 And Metric(report, "Invalid coordinate pairs") = 1, "Missing and invalid coordinates distinguished"
    CheckQuality Metric(report, "Possible reversed pairs") = 1 And Metric(report, "Zero-zero coordinate pairs") = 1, "Coordinate swap and zero placeholders flagged"
    Set found = report.Columns(1).Find(What:="2026-02", LookIn:=xlValues, LookAt:=xlWhole)
    CheckQuality Not found Is Nothing, "Empty month retained"
    CheckQuality found.Offset(0, 1).Value2 = 0 And InStr(found.Offset(0, 3).Value2, "possible gap") > 0, "Zero month disclosed as possible coverage gap"
    Set found = report.Columns(1).Find(What:="2026-01", LookIn:=xlValues, LookAt:=xlWhole)
    CheckQuality found.Offset(0, 1).Value2 = 4 And found.Offset(0, 2).Value2 = 1, "Monthly records retain offenses/duplicates; incidents deduplicate and exclude date conflicts"
    CheckQuality source.Range("A5").Value2 = "bad date" And source.UsedRange.Columns.Count = 7, "Source data unchanged"
    ExportRangePNG report.Range("A1:E23"), previewDir & "\dataset-check.png"
    BuildToolGuide wb, False: Set guide = wb.Worksheets(3)
    Set found = guide.Columns(1).Find(What:="About CAP Index", LookIn:=xlValues, LookAt:=xlWhole)
    CheckQuality guide.Range("A7").Value2 = "Dataset Check + Coverage" And Not found Is Nothing, "Guide includes separate current-tool sections and ownership"
    CheckQuality guide.Range("A8").Value2 = "Select / map" And InStr(guide.Range("B10").Value2, vbCrLf) > 0, "Guide uses labelled steps and hard line breaks"
    ExportRangePNG guide.Range("A1:D16"), previewDir & "\tool-guide.png"
    BuildDatasetCheck source.Range("A1:G9"), "", "", "", "", "", "", False: Set report = wb.Worksheets(4)
    CheckQuality Metric(report, "Recorded unique incident IDs") = "Not mapped" And Metric(report, "Invalid dates") = "Not mapped", "Absent optional fields disclosed rather than zero counts"
    source.Cells.Clear
    source.Range("A1").Value2 = "Label": source.Range("B1").Value2 = "Label"
    source.Range("A2:A5").NumberFormat = "@"
    source.Range("A2").Value2 = " X": source.Range("A3").Value2 = "X"
    source.Range("A4").Value2 = CVErr(xlErrNA): source.Range("A5").Value2 = CVErr(xlErrValue)
    BuildDatasetCheck source.Range("A1:B5"), "", "Label", "", "", "", "", False: Set report = wb.Worksheets(5)
    CheckQuality Metric(report, "Header issues") = 1 And Metric(report, "Missing/error crime types") = "Not mapped", "Ambiguous mapped headers flagged without selecting an arbitrary column"
    CheckQuality Metric(report, "Cell errors") = 2 And Metric(report, "Exact duplicate extra rows") = 0, "Exact duplicate check preserves whitespace and distinct Excel errors"
    wb.Close False: Application.EnableEvents = oldEvents
    RunDatasetCheckTests = "PASS: " & checks & " dataset quality/coverage/tool guide checks": Exit Function
failed:
    errNo = Err.Number: description = Err.Description
    On Error Resume Next: If Not wb Is Nothing Then wb.Close False
    Application.EnableEvents = oldEvents
    On Error GoTo 0: Err.Raise errNo, "Dataset check tests", description
End Function
