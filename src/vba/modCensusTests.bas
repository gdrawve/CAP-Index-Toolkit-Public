Attribute VB_Name = "modCensusTests"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit
Private checks As Long
Private Sub AssertCensus(ByVal condition As Boolean, ByVal label As String)
    checks = checks + 1
    If Not condition Then Err.Raise vbObjectError + 351, , "FAIL: " & label
End Sub

Public Function RunCensusTests() As String
    Dim wb As Workbook, ws As Worksheet, r As Range, lo As ListObject, fake As clsCensusFixture
    Dim data(1 To 8, 1 To 5) As Variant, streets As Variant, i As Long, j As Long, caught As Long
    Dim records As Collection, expected As Object, results As Object, csv As String, saved As Variant, headers As Variant
    Dim zipInputs As Variant, zipExpected As Variant, zipFlags As Variant, wasRepaired As Boolean, zipValue As String
    On Error GoTo failed
    checks = 0
    AssertCensus FullAddress("123 Main St", "City", "IL", "") = "123 Main St, City, IL", "Optional ZIP concatenation"
    AssertCensus FullAddress("123 Main St", "", "", "01234") = "123 Main St 01234", "Optional city/state with ZIP"
    AssertCensus Len(AddressProblem("123 Main St", "City", "IL", "")) = 0, "No ZIP required with city/state"
    AssertCensus Len(AddressProblem("PO BOX 1", "City", "IL", "")) > 0, "PO box review"
    AssertCensus Len(AddressProblem("100 BLOCK MAIN ST", "City", "IL", "")) > 0, "Masked block review"
    AssertCensus Len(AddressProblem("123 Main St", "", "", "")) > 0, "Insufficient locality review"
    zipInputs = Array(72701, 2134, "2134", 601, 727011234, 21341234, "72701-1234", "12", "ABCDE", "123456", 12.5, CVErr(xlErrValue), "00123", "1E3")
    zipExpected = Array("72701", "02134", "02134", "00601", "72701-1234", "02134-1234", "72701-1234", "12", "ABCDE", "00012-3456", "12.5", "[CELL ERROR]", "00123", "1E3")
    zipFlags = Array(False, True, True, True, True, True, False, False, False, True, False, False, False, False)
    For i = 0 To UBound(zipInputs)
        wasRepaired = True: zipValue = RepairZip(zipInputs(i), wasRepaired)
        AssertCensus zipValue = zipExpected(i) And wasRepaired = zipFlags(i), "ZIP repair case " & i
    Next
    Set records = ReadCSV("A,""B,C"",""D""""E""" & vbCrLf & "F,""line" & vbLf & "two"",G")
    AssertCensus records.Count = 2 And records(1)(1) = "B,C" And records(1)(2) = "D""E" And InStr(records(2)(1), vbLf) > 0, "Quoted CSV commas quotes embedded newlines"
    AssertCensus UTF8Size("abc") = 3 And UTF8Size(ChrW$(233)) = 2, "UTF-8 byte accounting without BOM"
    Set expected = CreateObject("Scripting.Dictionary"): expected.Add "A000001", True
    Set results = CensusResponse("A000001,input,No_Match" & vbCrLf, expected)
    AssertCensus results("A000001")(3) = "No Match" And IsEmpty(results("A000001")(1)), "Actual no-match spelling and blank coordinates"
    On Error Resume Next
    Set results = CensusResponse("unknown,input,No_Match", expected): caught = Err.Number: Err.Clear
    On Error GoTo failed
    AssertCensus caught <> 0, "Unknown response ID rejects batch"
    On Error Resume Next
    Set results = CensusResponse("A000001,input,No_Match" & vbCrLf & "A000001,input,No_Match", expected): caught = Err.Number: Err.Clear
    On Error GoTo failed
    AssertCensus caught <> 0, "Duplicate response ID rejects batch"
    On Error Resume Next
    Set results = CensusResponse("", expected): caught = Err.Number: Err.Clear
    On Error GoTo failed
    AssertCensus caught <> 0, "Missing response ID rejects batch"
    On Error Resume Next
    Set results = CensusResponse("A000001,input,Match,Exact,matched,""-181,38""", expected): caught = Err.Number: Err.Clear
    On Error GoTo failed
    AssertCensus caught <> 0, "Out-of-range coordinate rejects batch"
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    headers = Array("Address", "City", "State", "ZIP", "IncidentID")
    For j = 0 To 4: data(1, j + 1) = headers(j): Next
    streets = Array("4600 Silver Hill Rd", "4600 SILVER HILL RD", "123 NONEXACT RD", "123 TIE ST", "123 MISSING ST", "PO BOX 1", "")
    For i = 2 To 8
        data(i, 1) = streets(i - 2): data(i, 2) = "Washington": data(i, 3) = "DC": data(i, 4) = "20233": data(i, 5) = "PRIVATE-CASE-" & i
    Next
    Set r = ws.Range("A1:E8"): r.NumberFormat = "@": r.Value2 = data
    Set fake = New clsCensusFixture
    GeocodeCensusInPlace r, "Address", "City", "State", "ZIP", fake, 2, False
    AssertCensus fake.CallCount = 2 And fake.MaxSeen = 2, "Distinct addresses and automatic batching"
    AssertCensus InStr(fake.Captured, "PRIVATE-CASE") = 0, "Incident fields never submitted"
    AssertCensus ws.Range("E2").Value2 = "4600 Silver Hill Rd, Washington, DC 20233", "Full Address inserted after address fields"
    AssertCensus ws.Range("N2").Value2 = "PRIVATE-CASE-2", "Existing neighbor shifts intact"
    AssertCensus ws.Range("I2").Value2 = "Match" And ws.Range("I3").Value2 = "Match", "One result returned to repeated address rows"
    AssertCensus Abs(ws.Range("G2").Value2 - 38.846016) < 0.000001 And Abs(ws.Range("H2").Value2 + 76.927487) < 0.000001, "Latitude longitude correctly separated"
    AssertCensus ws.Range("J4").Value2 = "Non-Exact" And Len(ws.Range("K4").Value2) > 0, "Non-exact match review"
    AssertCensus ws.Range("I5").Value2 = "Tie" And IsEmpty(ws.Range("G5").Value2), "Tie remains unresolved"
    AssertCensus ws.Range("I6").Value2 = "No Match" And IsEmpty(ws.Range("H6").Value2), "No match remains blank"
    AssertCensus ws.Range("I7").Value2 = "Invalid Input" And ws.Range("I8").Value2 = "Invalid Input", "Invalid rows never submitted"
    AssertCensus ws.Range("L2").Value2 = CENSUS_BENCHMARK And IsNumeric(ws.Range("M2").Value2), "Benchmark and timestamp provenance"
    For i = 1 To 8
        For j = 1 To 4: AssertCensus SafeText(ws.Cells(i, j).Value2) = SafeText(data(i, j)), "Source preservation " & i & "/" & j: Next
    Next
    Set fake = New clsCensusFixture
    GeocodeCensusInPlace r, "Address", "City", "State", "ZIP", fake, 2, False
    AssertCensus ws.Range("N2").Value2 = "PRIVATE-CASE-2" And ws.UsedRange.Columns.Count = 14, "Rerun refreshes same columns"
    Set ws = wb.Worksheets.Add: ws.Range("A1:E8").NumberFormat = "@": ws.Range("A1:E8").Value2 = data
    Set fake = New clsCensusFixture: fake.FailAt = 2
    On Error Resume Next
    GeocodeCensusInPlace ws.Range("A1:E8"), "Address", "City", "State", "ZIP", fake, 1, False
    caught = Err.Number: Err.Clear
    On Error GoTo failed
    AssertCensus caught <> 0 And ws.Range("I2").Value2 = "Match" And ws.Range("I3").Value2 = "Match", "Completed results retained after later batch failure"
    AssertCensus ws.Range("I4").Value2 = "Batch Failed" And IsEmpty(ws.Range("G4").Value2), "Unfinished rows flagged not false no-match"
    Set ws = wb.Worksheets.Add: ws.Range("A1:E8").NumberFormat = "@": ws.Range("A1:E8").Value2 = data
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E8"), , xlYes): lo.ShowTotals = True
    Set fake = New clsCensusFixture
    GeocodeCensusInPlace lo.HeaderRowRange.Resize(8, 5), "Address", "City", "State", "ZIP", fake, 2, False
    AssertCensus lo.ListColumns.Count = 14 And lo.ListRows.Count = 7 And lo.ShowTotals, "Table columns rows and totals retained"
    Set fake = New clsCensusFixture
    GeocodeCensusInPlace lo.HeaderRowRange.Resize(8, 14), "Address", "City", "State", "ZIP", fake, 2, False
    AssertCensus lo.ListColumns.Count = 14, "Table rerun refreshes instead of duplicate columns"
    Set ws = wb.Worksheets.Add
    ws.Range("A1").Value2 = "Address": ws.Range("B1").Value2 = "City": ws.Range("C1").Value2 = "State"
    For i = 2 To 4
        ws.Cells(i, 1).Value2 = i & " " & String$(140, "X") & " ST"
        ws.Cells(i, 2).Value2 = "Test City": ws.Cells(i, 3).Value2 = "IL"
    Next
    Set fake = New clsCensusFixture
    GeocodeCensusInPlace ws.Range("A1:C4"), "Address", "City", "State", "", fake, 10000, False, 256
    AssertCensus fake.CallCount = 3 And fake.MaxSeen = 1, "Byte limit splits batches before row limit"
    AssertCensus ws.Range("H4").Value2 = "Match", "Missing ZIP submits and matches"
    Set ws = wb.Worksheets.Add
    WriteHeaders ws, 1, 1, Array("Address", "City", "State", "ZIP")
    ws.Range("A2").Value2 = "4600 Silver Hill Rd": ws.Range("B2").Value2 = "Test City": ws.Range("C2").Value2 = "MA": ws.Range("D2").Value2 = 2134
    Set fake = New clsCensusFixture
    GeocodeCensusInPlace ws.Range("A1:D2"), "Address", "City", "State", "ZIP", fake, 10000, False
    AssertCensus InStr(fake.Captured, "02134") > 0 And ws.Range("I2").Value2 = "Match", "Numeric ZIP repaired and submitted"
    AssertCensus InStr(ws.Range("K2").Value2, "ZIP restored") > 0 And InStr(ws.Range("E2").Value2, "02134") > 0, "Repair warning and full address survive response"
    Set ws = wb.Worksheets.Add
    WriteHeaders ws, 1, 1, Array("Address", "City", "State", "ZIP")
    ws.Range("A2").Value2 = "4600 Silver Hill Rd": ws.Range("B2").Value2 = "Test City": ws.Range("C2").Value2 = "MA": ws.Range("D2").Value2 = 123456
    Set fake = New clsCensusFixture
    GeocodeCensus ws.Range("A1:D2"), "Address", "City", "State", "ZIP", fake, 10000, False
    AssertCensus InStr(fake.Captured, "00012-3456") > 0 And InStr(wb.Worksheets(wb.Worksheets.Count).Range("K7").Value2, "inferred") > 0, "Working-dataset wrapper preserves inferred ZIP+4 warning"
    wb.Close False
    RunCensusTests = "PASS: " & checks & " geocoder assertions: deduplication, response joins, CSV, statuses, retention, privacy, insertion, tables and refresh"
    Exit Function
failed:
    Dim description As String: description = Err.Description
    If Not wb Is Nothing Then wb.Close False
    Err.Raise vbObjectError + 352, , description
End Function

Public Function RunCensusBatchLimitTest() As String
    Dim wb As Workbook, ws As Worksheet, fake As clsCensusFixture, a() As Variant, i As Long
    On Error GoTo failed
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    ReDim a(1 To 10002, 1 To 4)
    a(1, 1) = "Address": a(1, 2) = "City": a(1, 3) = "State": a(1, 4) = "ZIP"
    For i = 2 To 10002
        a(i, 1) = i & " TEST ST": a(i, 2) = "TEST CITY": a(i, 3) = "IL": a(i, 4) = "01234"
    Next
    ws.Range("A1:D10002").NumberFormat = "@": ws.Range("A1:D10002").Value2 = a
    Set fake = New clsCensusFixture
    GeocodeCensusInPlace ws.Range("A1:D10002"), "Address", "City", "State", "ZIP", fake, 10000, False
    If fake.CallCount <> 2 Or fake.MaxSeen <> 10000 Or ws.Range("I10002").Value2 <> "Match" Then Err.Raise vbObjectError + 353, , "10,001-address split failed."
    Set ws = wb.Worksheets.Add
    WriteHeaders ws, 1, 1, Array("Address", "City", "State", "ZIP")
    ws.Range("A2").Value2 = "4600 Silver Hill Rd": ws.Range("B2").Value2 = "Test City": ws.Range("C2").Value2 = "MA": ws.Range("D2").Value2 = 2134
    Set fake = New clsCensusFixture
    GeocodeCensusInPlace ws.Range("A1:D2"), "Address", "City", "State", "ZIP", fake, 10000, False
    AssertCensus InStr(fake.Captured, "02134") > 0 And ws.Range("I2").Value2 = "Match", "Numeric ZIP repaired and submitted"
    AssertCensus InStr(ws.Range("K2").Value2, "ZIP restored") > 0 And InStr(ws.Range("E2").Value2, "02134") > 0, "Repair warning and full address survive response"
    Set ws = wb.Worksheets.Add
    WriteHeaders ws, 1, 1, Array("Address", "City", "State", "ZIP")
    ws.Range("A2").Value2 = "4600 Silver Hill Rd": ws.Range("B2").Value2 = "Test City": ws.Range("C2").Value2 = "MA": ws.Range("D2").Value2 = 123456
    Set fake = New clsCensusFixture
    GeocodeCensus ws.Range("A1:D2"), "Address", "City", "State", "ZIP", fake, 10000, False
    AssertCensus InStr(fake.Captured, "00012-3456") > 0 And InStr(wb.Worksheets(wb.Worksheets.Count).Range("K7").Value2, "inferred") > 0, "Working-dataset wrapper preserves inferred ZIP+4 warning"
    wb.Close False
    RunCensusBatchLimitTest = "PASS: 10001 distinct addresses split 10000 + 1; final result applied"
    Exit Function
failed:
    Dim description As String: description = Err.Description
    If Not wb Is Nothing Then wb.Close False
    Err.Raise vbObjectError + 354, , description
End Function

Public Function RunCensusLiveSmoke() As String
    Dim expected As Object, results As Object, csv As String
    Set expected = CreateObject("Scripting.Dictionary")
    expected.Add "PUBLIC1", True: expected.Add "PUBLIC2", True
    csv = "PUBLIC1,4600 Silver Hill Rd,Washington,DC,20233" & vbCrLf & "PUBLIC2,1600 Pennsylvania Ave NW,Washington,DC,20500" & vbCrLf
    Set results = CensusResponse(CensusPost(csv), expected)
    If results("PUBLIC1")(3) <> "Match" Then Err.Raise vbObjectError + 355, , "Public Census example did not match."
    RunCensusLiveSmoke = "PASS: live Census batch, two public demonstration addresses, IDs reconciled, address/match type/coordinates validated"
End Function
