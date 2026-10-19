Attribute VB_Name = "modDatasetTests"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit
Private checks As Long
Private Sub CheckDataset(ByVal condition As Boolean, ByVal description As String)
    checks = checks + 1
    If Not condition Then Err.Raise vbObjectError + 780, , "FAIL: " & description
End Sub

Public Function RunDatasetTests(Optional ByVal previewDir As String = "") As String
    Dim oldScreen As Boolean
    Dim wb As Workbook, ws As Worksheet, report As Worksheet, r As Range, lo As ListObject
    Dim a(1 To 7, 1 To 3) As Variant, fake As clsCensusFixture, caught As Long, errNo As Long, description As String
    On Error GoTo failed
    checks = 0: oldScreen = Application.ScreenUpdating: Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    a(1, 1) = "IncidentID": a(1, 2) = "CrimeType": a(1, 3) = "OffenseID"
    a(2, 1) = "I1": a(2, 2) = "Theft": a(2, 3) = "O1"
    a(3, 1) = "I1": a(3, 2) = "Burglary": a(3, 3) = "O2"
    a(4, 1) = "I1": a(4, 2) = " theft ": a(4, 3) = "O3"
    a(5, 1) = "I2": a(5, 2) = "Robbery": a(5, 3) = "O4"
    a(6, 1) = "": a(6, 2) = "Assault": a(6, 3) = "O5"
    a(7, 1) = "I3": a(7, 2) = "": a(7, 3) = "O6"
    ws.Range("A1:C7").Value2 = a
    ReviewCases ws.Range("A1:C7"), "IncidentID", "CrimeType"
    Set report = wb.Worksheets(2)
    CheckDataset report.Range("I6").Value2 = "CAP_IncidentCrimeTypes", "Incident crime-type list added"
    CheckDataset report.Range("I7").Value2 = "Burglary; Theft" And report.Range("I8").Value2 = "Burglary; Theft" And report.Range("I9").Value2 = "Burglary; Theft", "Several offense rows share sorted distinct incident list"
    CheckDataset report.Range("G7").Value2 = 3 And report.Range("E7").Value2 = 1, "Repeated incident differs from an exact duplicate"
    CheckDataset report.Range("I10").Value2 = "Robbery" And report.Range("I11").Value2 = "[No valid incident ID]", "Single incident and missing ID"
    CheckDataset report.Range("I12").Value2 = "[No recorded crime type]" And ws.Range("C7").Value2 = "O6", "No invented offense and all source rows retained"
    If Len(previewDir) > 0 Then ExportDatasetPreview report, "A6:I12", previewDir & "\incident-crime-types.png"
    Set ws = wb.Worksheets.Add
    ws.Range("A2:A3").NumberFormat = "@"
    ws.Range("A1").Value2 = "Date": ws.Range("B1").Value2 = "CrimeType"
    ws.Range("A2").Value2 = "01/05/2026": ws.Range("A3").Value2 = "01/06/2026": ws.Range("B2").Value2 = "KEEP"
    BuildDateTime ws.Range("A1:A3"), False
    CheckDataset ws.Range("G1").Value2 = "Date Only" And ws.Range("H1").Value2 = "CrimeType" And ws.UsedRange.Columns.Count = 8, "Six date fields and no empty time columns"
    BuildDateTime ws.Range("A1:A3"), False
    CheckDataset ws.Range("H2").Value2 = "KEEP" And ws.UsedRange.Columns.Count = 8, "Date-only rerun does not insert extra columns"
    ws.Range("A2").Value2 = "01/05/2026 00:00"
    BuildDateTime ws.Range("A1:A3"), False
    CheckDataset ws.Range("H1").Value2 = "Time Only" And ws.Range("I2").Value2 = 0 And ws.Range("J2").Value2 = "KEEP", "Explicit midnight expands six fields to eight"
    CheckDataset IsEmpty(ws.Range("H3").Value2), "Mixed date-only row remains untimed"
    ws.Range("A2").Value2 = "01/05/2026"
    BuildDateTime ws.Range("A1:A3"), False
    CheckDataset ws.Range("H2").Value2 = "KEEP" And ws.Range("H1").Value2 = "CrimeType", "All-time-blank rerun removes only owned time fields"
    Set ws = wb.Worksheets.Add
    ws.Range("A1").Value2 = "Date": ws.Range("B1").Value2 = "CrimeType"
    ws.Range("A2").NumberFormat = "@"
    ws.Range("A2").Value2 = "01/05/2026": ws.Range("B2").Value2 = "KEEP TABLE"
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B2"), , xlYes): lo.ShowTotals = True
    BuildDateTime lo.HeaderRowRange.Cells(1).Resize(2, 1), False
    CheckDataset lo.ListColumns.Count = 8 And lo.ShowTotals, "Date-only table adds six fields and retains totals"
    ws.Range("A2").Value2 = "01/05/2026 14:15"
    BuildDateTime lo.HeaderRowRange.Cells(1).Resize(2, 1), False
    CheckDataset lo.ListColumns.Count = 10 And lo.ListColumns(10).DataBodyRange.Value2 = "KEEP TABLE", "Date-only table expands safely"
    ws.Range("A2").Value2 = "01/05/2026"
    BuildDateTime lo.HeaderRowRange.Cells(1).Resize(2, 1), False
    CheckDataset lo.ListColumns.Count = 8 And lo.ListColumns(8).DataBodyRange.Value2 = "KEEP TABLE" And lo.ShowTotals, "Table shrinks safely when time disappears"
    Set ws = wb.Worksheets.Add
    WriteHeaders ws, 1, 1, Array("Address", "City", "State", "Date", "CrimeType", "IncidentID")
    ws.Range("A2").Value2 = "1200 MAIN ST": ws.Range("B2").Value2 = "TESTCITY": ws.Range("C2").Value2 = "IL"
    ws.Range("D2").Value2 = CDbl(DateSerial(2026, 1, 5)): ws.Range("D2").NumberFormat = "mm/dd/yyyy"
    ws.Range("E2").Value2 = "Theft": ws.Range("F2").Value2 = "I1"
    BuildHundredBlocks ws.Range("A1:F2"), showSummary:=False
    Set report = wb.Worksheets(wb.Worksheets.Count)
    CheckDataset report.ListObjects.Count = 1 And report.ListObjects(1).ListColumns.Count = 8, "Selected six source fields are copied into working table"
    CheckDataset report.Range("E7").Value2 = "Theft" And report.Range("F7").Value2 = "I1" And report.Range("D7").Text = "01/05/2026", "Crime/ID/date fields preserved"
    report.Range("J7").Formula = "=ISNUMBER(D7)": report.Range("J7").Calculate
    CheckDataset report.Range("J7").Value2 = True, "Copied date remains a numeric Excel date"
    report.Range("J7").Clear
    CheckDataset report.Range("G7").Value2 = "1200 BLOCK MAIN ST" And ws.UsedRange.Columns.Count = 6, "Hundred block added without modifying source"
    If Len(previewDir) > 0 Then ExportDatasetPreview report, "A6:H7", previewDir & "\working-hundred-blocks.png"
    BuildHundredBlocks report.ListObjects(1).HeaderRowRange.Resize(2, report.ListObjects(1).ListColumns.Count), "Address", "City", "State", "IncidentID", True, 42, False
    Set report = wb.Worksheets(wb.Worksheets.Count)
    CheckDataset report.ListObjects(1).ListColumns.Count = 9 And report.Range("F7").Value2 = "I1", "Randomizer copies whole working table and adds synthetic address"
    Set fake = New clsCensusFixture
    GeocodeCensus report.ListObjects(1).HeaderRowRange.Resize(2, report.ListObjects(1).ListColumns.Count), "CAP_SyntheticAddress", "City", "State", "", fake, 10000, False
    Set report = wb.Worksheets(wb.Worksheets.Count)
    CheckDataset report.ListObjects(1).ListColumns.Count = 16 And report.Range("E7").Value2 = "Theft" And report.Range("F7").Value2 = "I1", "Geocoder copies entire working dataset with seven result fields"
    CheckDataset InStr(report.Range("P7").Value2, "Synthetic") > 0, "Working geocoded dataset retains synthetic warning"
    wb.Close False: Application.ScreenUpdating = oldScreen: RunDatasetTests = "PASS: " & checks & " incident-list/date-only/full-dataset checks": Exit Function
failed:
    errNo = Err.Number: description = Err.Description
    On Error Resume Next: If Not wb Is Nothing Then wb.Close False
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0: Err.Raise errNo, "Dataset tests", description
End Function

Private Sub ExportDatasetPreview(ByVal ws As Worksheet, ByVal address As String, ByVal path As String)
    ExportRangePNG ws.Range(address), path
End Sub
