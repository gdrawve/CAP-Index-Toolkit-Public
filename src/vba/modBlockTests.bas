Attribute VB_Name = "modBlockTests"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit
Private checks As Long
Private Sub AssertBlock(ByVal condition As Boolean, ByVal label As String)
    checks = checks + 1
    If Not condition Then Err.Raise vbObjectError + 450, , "FAIL: " & label
End Sub

Public Function RunBlockTests() As String
    Dim base As Long, street As String, precision As String, status As String, masked As Boolean, ok As Boolean
    Dim i As Long, j As Long, suffix As Long, foundZero As Boolean, found99 As Boolean
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, r As Range, data(1 To 8, 1 To 5) As Variant, addresses As Variant
    Dim fake As clsCensusFixture, oldValue As String
    On Error GoTo failed
    checks = 0
    ok = ParseBlockAddress("1234 MAIN ST", base, street, precision, status, masked)
    AssertBlock ok And base = 1200 And Not masked, "Exact number floors to hundred block"
    ok = ParseBlockAddress("1299 MAIN ST", base, street, precision, status, masked)
    AssertBlock ok And base = 1200, "Upper block boundary"
    ok = ParseBlockAddress("1200 MAIN ST", base, street, precision, status, masked)
    AssertBlock ok And base = 1200 And masked And precision = "Hundred block (00)", "User-selected 00 means block input"
    ok = ParseBlockAddress("12xx MAIN ST", base, street, precision, status, masked)
    AssertBlock ok And base = 1200 And masked, "XX normalization"
    ok = ParseBlockAddress("1200 BLOCK OF MAIN ST", base, street, precision, status, masked)
    AssertBlock ok And base = 1200 And street = "MAIN ST" And masked, "Explicit block-of label"
    ok = ParseBlockAddress("1200 BLK MAIN ST", base, street, precision, status, masked)
    AssertBlock ok And masked And street = "MAIN ST", "BLK label"
    ok = ParseBlockAddress("1234 BLOCK MAIN ST", base, street, precision, status, masked)
    AssertBlock Not ok, "Non-boundary explicit block rejected"
    ok = ParseBlockAddress("123 1ST ST", base, street, precision, status, masked)
    AssertBlock ok And street = "1ST ST", "Ordinal street names"
    ok = ParseBlockAddress("123 60 ST", base, street, precision, status, masked)
    AssertBlock ok And street = "60 ST", "Numeric street names"
    ok = ParseBlockAddress("123 E APTITUDE AVE APT 2", base, street, precision, status, masked)
    AssertBlock ok And street = "E APTITUDE AVE", "Unit removed without corrupting street"
    ok = ParseBlockAddress("123 1/2 MAIN ST", base, street, precision, status, masked): AssertBlock Not ok, "Fractional address review"
    ok = ParseBlockAddress("123-129 MAIN ST", base, street, precision, status, masked): AssertBlock Not ok, "Address range review"
    ok = ParseBlockAddress("MAIN & OAK", base, street, precision, status, masked): AssertBlock Not ok, "Intersection review"
    ok = ParseBlockAddress("PO BOX 1", base, street, precision, status, masked): AssertBlock Not ok, "PO box review"
    ok = ParseBlockAddress("XX MAIN ST", base, street, precision, status, masked): AssertBlock ok And base = 0 And masked, "Zero hundred block XX"
    AssertBlock BlockRandomSuffix(42, "INCIDENT:I1") = BlockRandomSuffix(42, "INCIDENT:I1"), "Stable seeded result"
    For i = 1 To 1000
        suffix = BlockRandomSuffix(42, "KEY:" & i)
        If suffix < 0 Or suffix > 99 Then Err.Raise vbObjectError + 451, , "Random suffix outside 00-99."
        If suffix = 0 Then foundZero = True
        If suffix = 99 Then found99 = True
    Next
    AssertBlock foundZero And found99, "Random range includes 00 and 99"
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    data(1, 1) = "Address": data(1, 2) = "City": data(1, 3) = "State": data(1, 4) = "IncidentID": data(1, 5) = "OffenseID"
    addresses = Array("1234 MAIN ST", "1200 MAIN ST", "12XX MAIN ST", "1200 BLOCK OF MAIN ST", "99 OAK ST", "MAIN & OAK", "")
    For i = 2 To 8
        data(i, 1) = addresses(i - 2): data(i, 2) = "TEST CITY": data(i, 3) = "IL"
        data(i, 4) = "I" & i: data(i, 5) = "O" & i
    Next
    data(4, 4) = data(3, 4): data(5, 4) = data(3, 4)
    Set r = ws.Range("A1:E8"): r.NumberFormat = "@": r.Value2 = data
    BuildHundredBlocksInPlace r, "Address", "City", "State", "IncidentID", False, 42, False
    AssertBlock wb.Worksheets.Count = 1 And ws.Range("B2").Value2 = 1200, "In-sheet block number insertion"
    AssertBlock ws.Range("C2").Value2 = "1200 BLOCK MAIN ST", "Hundred-block label"
    AssertBlock ws.Range("F3").Value2 = "" And ws.Range("G3").Value2 = False, "Builder does not generate fictional addresses"
    AssertBlock ws.Range("J2").Value2 = "TEST CITY" And ws.Range("M2").Value2 = "O2", "Adjacent fields shift safely"
    BuildHundredBlocksInPlace r, "Address", "City", "State", "IncidentID", True, 42, False
    AssertBlock ws.Range("F2").Value2 = "" And ws.Range("G2").Value2 = False, "Exact address never randomized"
    AssertBlock ws.Range("F3").Value2 = ws.Range("F4").Value2 And ws.Range("F4").Value2 = ws.Range("F5").Value2, "Same incident and block share synthetic address across offense rows"
    suffix = CLng(Split(ws.Range("F3").Value2, " ")(0))
    AssertBlock suffix >= 1200 And suffix <= 1299, "Synthetic house number within requested block"
    AssertBlock ws.Range("G3").Value2 = True And ws.Range("H3").Value2 = 42 And InStr(ws.Range("I3").Value2, "not an observed") > 0, "Synthetic provenance retained"
    AssertBlock ws.Range("D3").Value2 = ws.Range("D4").Value2, "00 and XX share geographic block key"
    AssertBlock Len(ws.Range("D7").Value2) = 0 And Len(ws.Range("F7").Value2) = 0, "Unsupported addresses not invented"
    oldValue = ws.Range("F3").Value2
    BuildHundredBlocksInPlace r, "Address", "City", "State", "IncidentID", True, 42, False
    AssertBlock ws.UsedRange.Columns.Count = 13 And ws.Range("F3").Value2 = oldValue, "Repeat run stable without duplicate columns"
    For i = 1 To 8
        AssertBlock SafeText(ws.Cells(i, 1).Value2) = SafeText(data(i, 1)), "Original address preserved row " & i
    Next
    Set fake = New clsCensusFixture
    GeocodeCensusInPlace ws.Range("A1:M8"), "CAP_SyntheticAddress", "City", "State", "", fake, 10000, False
    AssertBlock InStr(ws.Range("R3").Value2, "Synthetic hundred-block estimate") > 0, "Census preserves synthetic location warning even on exact match"
    Set ws = wb.Worksheets.Add: ws.Range("A1:E8").NumberFormat = "@": ws.Range("A1:E8").Value2 = data
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E8"), , xlYes): lo.ShowTotals = True
    BuildHundredBlocksInPlace lo.HeaderRowRange.Resize(8, 5), "Address", "City", "State", "IncidentID", True, 42, False
    AssertBlock lo.ListColumns.Count = 13 And lo.ListRows.Count = 7 And lo.ShowTotals, "Table inserts fields and preserves rows/totals"
    BuildHundredBlocksInPlace lo.HeaderRowRange.Resize(8, 13), "Address", "City", "State", "IncidentID", True, 42, False
    AssertBlock lo.ListColumns.Count = 13, "Table refresh avoids duplicate columns"
    wb.Close False
    RunBlockTests = "PASS: " & checks & " block/randomizer assertions plus 1000 seeded range checks"
    Exit Function
failed:
    Dim description As String: description = Err.Description
    If Not wb Is Nothing Then wb.Close False
    Err.Raise vbObjectError + 452, , description
End Function
