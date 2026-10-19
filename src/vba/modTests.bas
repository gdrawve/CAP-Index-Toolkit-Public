Attribute VB_Name = "modTests"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit
Private checks As Long
Public Function VerifyRibbonIcon(ByVal imageID As String) As Boolean
    Dim picture As stdole.IPictureDisp
    Set picture = Application.CommandBars.GetImageMso(imageID, 32, 32)
    VerifyRibbonIcon = picture.Width > 0 And picture.Height > 0
End Function
Private Sub Check(ByVal result As Boolean, ByVal name As String)
    checks = checks + 1
    If Not result Then Err.Raise vbObjectError + 100, , "FAIL: " & name
End Sub

Public Function RunSelfTests() As String
    Dim d As Date, ht As Boolean, status As String, ok As Boolean, v As String, t As Double
    Dim a(1 To 8, 1 To 2) As Variant, blocks As Object, missing As Long, excluded As Long, ids As Long
    Dim rows(1 To 3, 1 To 2) As Variant
    checks = 0
    ok = ParseTimestamp("2024-02-29 00:00:00", False, d, ht, status)
    Check ok And ht And Hour(d) = 0, "Explicit midnight and leap day"
    ok = ParseTimestamp("2023-02-29", False, d, ht, status): Check Not ok, "Reject non-leap date"
    ok = ParseTimestamp("2024-13-01", False, d, ht, status): Check Not ok, "Reject month overflow"
    ok = ParseTimestamp("2024-04-31", False, d, ht, status): Check Not ok, "Reject day overflow"
    ok = ParseTimestamp("2024-01-01 24:00", False, d, ht, status): Check Not ok, "Reject hour overflow"
    ok = ParseTimestamp("2024-01-01T12:34:56Z", False, d, ht, status): Check Not ok, "Reject unhandled timezone"
    ok = ParseTimestamp("01/02/2024", False, d, ht, status): Check ok And Month(d) = 1 And Day(d) = 2, "US month/day/year text"
    ok = ParseTimestamp("3/4/2026 7:05 PM", False, d, ht, status): Check ok And Month(d) = 3 And Day(d) = 4 And Hour(d) = 19, "US text with PM"
    ok = ParseTimestamp("3/4/2026 12:00 AM", False, d, ht, status): Check ok And ht And Hour(d) = 0, "US explicit midnight"
    ok = ParseTimestamp("3/4/2026 12:00 PM", False, d, ht, status): Check ok And Hour(d) = 12, "US noon"
    ok = ParseTimestamp("3/4/2026 00:00 AM", False, d, ht, status): Check Not ok, "Invalid AM hour"
    ok = ParseTimestamp("2024-01-01", False, d, ht, status): Check ok And Not ht, "Date-only missing time"
    ok = ParseTimestamp(45292#, False, d, ht, status): Check ok And Year(d) = 2024 And Month(d) = 1 And Day(d) = 1, "1900 serial"
    ok = ParseTimestamp(43830#, True, d, ht, status): Check ok And Year(d) = 2024 And Day(d) = 1, "1904 serial"
    ok = ParseTimestamp(60#, False, d, ht, status): Check Not ok, "Reject fictitious date"
    ok = ParseTimestamp(Empty, False, d, ht, status): Check Not ok, "Missing date"
    ok = ParseTimestamp(CVErr(xlErrNA), False, d, ht, status): Check Not ok, "Date cell error"
    Check ParseTimeOnly("00:00", t) And t = 0, "Separate explicit midnight"
    Check Not ParseTimeOnly("23:60", t), "Invalid minute"
    Check Not ParseTimeOnly(2400, t), "No implicit HHMM"
    v = HundredBlock("123 N Main St Apt 2", status): Check v = "100 BLOCK N MAIN ST", "Unit stripped"
    v = HundredBlock("123 E APTITUDE AVE", status): Check v = "100 BLOCK E APTITUDE AVE", "Street prefix not confused with apartment marker"
    v = HundredBlock("99 MAIN ST", status): Check v = "0 BLOCK MAIN ST", "Zero hundred block"
    v = HundredBlock("100 MAIN ST", status): Check v = "100 BLOCK MAIN ST", "Boundary"
    v = HundredBlock("120-124 MAIN ST", status): Check v = "", "Reject range"
    v = HundredBlock("123 1/2 MAIN ST", status): Check v = "", "Reject fractional address"
    v = HundredBlock("MAIN & OAK", status): Check v = "", "Reject intersection"
    v = HundredBlock("100 BLOCK MAIN ST", status): Check v = "", "Reject already masked block"
    v = HundredBlock("PO BOX 123", status): Check v = "", "Reject PO box"
    a(1, 1) = "ID": a(1, 2) = "Block"
    a(2, 1) = "I1": a(2, 2) = "A"
    a(3, 1) = "I1": a(3, 2) = "A"
    a(4, 1) = "I2": a(4, 2) = "A"
    a(5, 1) = "I2": a(5, 2) = "B"
    a(6, 1) = "I3": a(6, 2) = "B"
    a(7, 1) = "": a(7, 2) = "B"
    a(8, 1) = "I4": a(8, 2) = ""
    Set blocks = AggregateBlocks(a, 1, 2, missing, excluded, ids)
    Check blocks("A")(0) = 3 And blocks("B")(0) = 3, "Offense rows retained"
    Check blocks("A")(1) = 1 And blocks("B")(1) = 1, "Repeated IDs count once; conflict excluded globally"
    Check ids = 4 And excluded = 2 And missing = 1, "Coverage reconciliation"
    Check RowSignature(a, 2) = RowSignature(a, 3), "Exact repeated rows"
    Check RowSignature(a, 2) <> RowSignature(a, 4), "Different rows"
    Check ExactRowEqual(a, 2, 3), "Exact row equality"
    a(2, 2) = 1.000000000000001: a(3, 2) = 1.000000000000002
    Check Not ExactRowEqual(a, 2, 3), "Floating point values not merged by text rounding"
    rows(1, 1) = "B": rows(1, 2) = 5
    rows(2, 1) = "C": rows(2, 2) = 10
    rows(3, 1) = "A": rows(3, 2) = 5
    SortConcentration rows, 1, 3
    Check rows(1, 1) = "C" And rows(2, 1) = "A" And rows(3, 1) = "B", "Deterministic ranking and ties"
    RunSelfTests = "PASS: " & checks & " VBA assertions"
End Function

Public Function RunCommandTests() As String
    Dim wb As Workbook, ws As Worksheet, r As Range, blockRange As Range
    Dim a(1 To 8, 1 To 7) As Variant, i As Long, j As Long, before As Long
    Dim idValues As Variant, addresses As Variant, after As Variant
    On Error GoTo failed
    before = checks
    Set wb = Workbooks.Add(xlWBATWorksheet)
    WriteHeaders wb.Worksheets(1), 1, 1, Array("IncidentID", "OffenseID", "OccurredAt", "Address", "City", "State", "ZIP")
    idValues = Array("I1", "I1", "I2", "I2", "I3", "", "I4")
    addresses = Array("123 MAIN ST", "123 MAIN ST", "123 MAIN ST", "234 OAK ST", "234 OAK ST", "234 OAK ST", "MAIN & OAK")
    For j = 1 To 7: a(1, j) = wb.Worksheets(1).Cells(1, j).Value2: Next
    For i = 2 To 8
        a(i, 1) = idValues(i - 2): a(i, 2) = "O" & i
        a(i, 3) = "2024-02-29 00:00:00": a(i, 4) = addresses(i - 2)
        a(i, 5) = "TEST CITY": a(i, 6) = "IL": a(i, 7) = "01234"
    Next
    a(4, 3) = "2023-02-29": a(5, 3) = "2024-03-01"
    Set r = wb.Worksheets(1).Range("A1:G8")
    r.NumberFormat = "@": r.Value2 = a
    BuildHundredBlocksLegacy r
    Set ws = wb.Worksheets(wb.Worksheets.Count)
    Check ws.Cells(7, 8).Value2 = "100 BLOCK MAIN ST", "Block output label"
    Check Len(ws.Cells(13, 9).Value2) = 0, "Unlocated address excluded from block key"
    Set blockRange = ws.Range("A6:J13")
    ReviewCases r
    Set ws = wb.Worksheets(wb.Worksheets.Count)
    Check ws.Cells(7, 11).Value2 = 2 And ws.Cells(8, 11).Value2 = 2, "Case output repeats not removed"
    Check InStr(ws.Cells(12, 12).Value2, "Missing") > 0, "Case output missing ID"
    PrepareCensus r, "Address", "City", "State", "ZIP"
    Set ws = wb.Worksheets(wb.Worksheets.Count)
    Check ws.Cells(7, 5).Value2 = "01234", "Census leading-zero ZIP preserved"
    Check ws.Cells(7, 6).Value2 = 1 And ws.Cells(12, 7).Value2 = 6, "Census batch allocation"
    Check ws.Cells(13, 9).Value2 = "Unsupported address; review" And IsEmpty(ws.Cells(13, 6)), "Census review row not allocated"
    Check ws.Cells(7, 10).Value2 = "Not submitted", "Census no submission"
    BuildPareto blockRange, "IncidentID", "CAP_BlockKey", "INCIDENTS"
    Set ws = wb.Worksheets(wb.Worksheets.Count)
    Check ws.Cells(7, 2).Value2 = 1 And ws.Cells(8, 2).Value2 = 1, "Incident concentration reconciles"
    Check ws.Cells(8, 5).Value2 = 1, "Incident cumulative share reconciles"
    BuildPareto blockRange, "IncidentID", "CAP_BlockKey", "RECORDS"
    Set ws = wb.Worksheets(wb.Worksheets.Count)
    Check ws.Cells(7, 2).Value2 = 3 And ws.Cells(8, 2).Value2 = 3, "Record concentration reconciles"
    after = r.Value2
    For i = 1 To 8
        For j = 1 To 7
            Check SafeText(after(i, j)) = SafeText(a(i, j)), "Source preservation " & i & "/" & j
        Next
    Next
    Check wb.Worksheets.Count = 6, "Other commands create separate output"
    wb.Close False
    RunCommandTests = "PASS: " & checks - before & " end-to-end assertions across other commands and both concentration measures"
    Exit Function
failed:
    Dim description As String: description = Err.Description
    If Not wb Is Nothing Then wb.Close False
    Err.Raise vbObjectError + 102, , description
End Function

Public Function RunTimestampTests() As String
    Dim wb As Workbook, ws As Worksheet, r As Range, lo As ListObject
    Dim a(1 To 10, 1 To 1) As Variant, before As Long, i As Long, j As Long
    Dim headers As Variant, saved As Variant, after As Variant, endCol As Long, caught As Long
    On Error GoTo failed
    before = checks: headers = TimestampHeaders()
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    a(1, 1) = "OccurredAt"
    a(2, 1) = CDbl(DateSerial(2026, 10, 13) + TimeSerial(14, 35, 0))
    a(3, 1) = CDbl(DateSerial(2026, 10, 13))
    a(4, 1) = "10/13/2026 12:00 AM"
    a(5, 1) = "03/04/2026 7:05 PM"
    a(6, 1) = Empty
    a(7, 1) = "2023-02-29"
    a(8, 1) = "2026-10-13"
    a(9, 1) = "=1+1"
    a(10, 1) = "2026-10-13 23:59:59"
    Set r = ws.Range("B1:B10"): r.NumberFormat = "@": r.Value2 = a
    saved = r.Value2
    ws.Range("A1").Value2 = "RecordID": ws.Range("A2").Value2 = "KEEP"
    ws.Range("C1").Value2 = "Existing field": ws.Range("C2").Value2 = "Original neighbor"
    ws.Range("D2").Formula = "=B2"
    BuildDateTime r, False
    Check wb.Worksheets.Count = 1, "Timestamp stays on original worksheet"
    For j = 0 To 7
        Check ws.Cells(1, j + 3).Value2 = headers(j), "Timestamp field order " & j + 1
    Next
    Check ws.Range("C2").Value2 = 10 And ws.Range("D2").Value2 = "October", "Number and full month name"
    Check ws.Range("E2").Value2 = 13 And ws.Range("F2").Value2 = "Tuesday", "Day number and weekday text"
    Check ws.Range("G2").Value2 = 2026, "Year number"
    Check ws.Range("H2").Value2 = CDbl(DateSerial(2026, 10, 13)) And VarType(ws.Range("H2").Value2) = vbDouble, "Typed date-only serial"
    Check Abs(ws.Range("I2").Value2 - CDbl(TimeSerial(14, 35, 0))) < 0.000000001 And ws.Range("J2").Value2 = 14, "Typed time and 24-hour hour"
    Check IsEmpty(ws.Range("I3")) And IsEmpty(ws.Range("J3")), "Numeric date-only time fields blank"
    Check ws.Range("I4").Value2 = 0 And ws.Range("J4").Value2 = 0, "Explicit midnight remains zero"
    Check ws.Range("C5").Value2 = 3 And ws.Range("E5").Value2 = 4 And ws.Range("J5").Value2 = 19, "US date/time output"
    Check Application.CountA(ws.Range("C6:J6")) = 0, "Blank source yields eight blanks"
    Check Application.CountA(ws.Range("C7:J7")) = 0 And ws.Range("C7").Interior.Color = RGB(255, 235, 238), "Invalid source highlighted with blank results"
    Check IsEmpty(ws.Range("I8")) And IsEmpty(ws.Range("J8")), "ISO date-only time fields blank"
    Check ws.Range("B9").Value2 = "=1+1" And Not ws.Range("B9").HasFormula, "Formula-like source text preserved"
    Check ws.Range("J10").Value2 = 23, "Hour upper bound 23"
    Check ws.Range("K2").Value2 = "Original neighbor" And ws.Range("L2").Formula = "=B2", "Adjacent values and formulas shift safely"
    after = r.Value2
    For i = 1 To 10: Check SafeText(saved(i, 1)) = SafeText(after(i, 1)), "Timestamp source unchanged row " & i: Next
    endCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    ws.Range("B2").Value2 = CDbl(DateSerial(2026, 11, 1) + TimeSerial(0, 30, 0))
    BuildDateTime r, False
    Check ws.Range("C2").Value2 = 11 And ws.Range("D2").Value2 = "November", "Rerun refreshes values"
    Check ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1 = endCol, "Rerun does not insert duplicates"
    Set ws = wb.Worksheets.Add
    ws.Range("A1").Value2 = "OccurredAt": ws.Range("B1").Value2 = "Neighbor"
    ws.Range("A2").Value2 = CDbl(DateSerial(2026, 10, 13) + TimeSerial(23, 59, 0))
    ws.Range("A3").Value2 = CDbl(DateSerial(2026, 10, 13)): ws.Range("B2").Value2 = "KEEP TABLE"
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B3"), , xlYes)
    lo.ShowTotals = True
    Set r = lo.ListColumns(1).Range.Resize(lo.ListRows.Count + 1)
    BuildDateTime r, False
    Check lo.ListColumns.Count = 10 And lo.ListRows.Count = 2, "Table inserts eight fields and retains rows"
    Check lo.ListColumns(10).Name = "Neighbor" And ws.Range("J2").Value2 = "KEEP TABLE", "Table adjacent field retained"
    Check ws.Range("I2").Value2 = 23 And IsEmpty(ws.Range("I3")), "Table mixed dates and timestamps"
    Check lo.ShowTotals, "Table total row retained"
    BuildDateTime r, False
    Check lo.ListColumns.Count = 10, "Table rerun refreshes existing fields"
    wb.Date1904 = True
    Set ws = wb.Worksheets.Add
    ws.Range("A1").Value2 = "OccurredAt": ws.Range("A2").Value2 = CDbl(DateSerial(2026, 10, 13) + TimeSerial(12, 0, 0)) - 1462
    Set r = ws.Range("A1:A2")
    BuildDateTime r, False
    Check ws.Range("G2").Value2 = CDbl(DateSerial(2026, 10, 13)) - 1462 And ws.Range("I2").Value2 = 12, "1904 date system preserved"
    Set ws = wb.Worksheets.Add
    ws.Range("A1").Value2 = "OccurredAt": ws.Range("A2").Value2 = "not a date"
    ws.Range("B1").Value2 = "Neighbor": ws.Range("B2").Value2 = "KEEP INVALID"
    On Error Resume Next
    BuildDateTime ws.Range("A1:A2"), False
    caught = Err.Number: Err.Clear
    On Error GoTo failed
    Check caught <> 0 And ws.Range("B2").Value2 = "KEEP INVALID", "All-invalid input rejects before insertion"
    wb.Close False
    RunTimestampTests = "PASS: " & checks - before & " timestamp assertions: mixed values, eight fields, adjacent columns, tables, refresh, invalid rows, 1904 system"
    Exit Function
failed:
    Dim description As String: description = Err.Description
    If Not wb Is Nothing Then wb.Close False
    Err.Raise vbObjectError + 103, , description
End Function

Public Function RunOutputSmokeTest() As String
    Dim wb As Workbook, source As Range, ws As Worksheet, a(1 To 3, 1 To 2) As Variant
    Dim outName As String
    On Error GoTo failed
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    a(1, 1) = "IncidentID": a(1, 2) = "Address"
    a(2, 1) = "I01": a(2, 2) = "123 Main St"
    a(3, 1) = "=1+1": a(3, 2) = "100 Main St"
    Set source = wb.Worksheets(1).Range("A1:B3")
    source.NumberFormat = "@": source.Value2 = a
    Set ws = NewOutput(wb, "SmokeTest", source, "Synthetic QA")
    CopySourceValues source, ws
    Check ws.Range("A8").Value2 = "=1+1" And Not ws.Range("A8").HasFormula, "Source text not executed as a formula"
    WriteHeaders ws, 6, 3, Array("TestField")
    ws.Range("C7").Value2 = 1
    FinishOutput ws, 8, 3
    Check ws.AutoFilterMode And wb.Worksheets.Count = 2, "New sheet and filter"
    Check source.Cells(2, 1).Value2 = "I01", "Source remains unchanged"
    wb.Close False
    RunOutputSmokeTest = "PASS: output sheet, filters, source preservation, formula-like text"
    Exit Function
failed:
    Dim description As String: description = Err.Description
    If Not wb Is Nothing Then wb.Close False
    Err.Raise vbObjectError + 101, , description
End Function
