Attribute VB_Name = "modConcentrationTests"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit
Private checks As Long
Private Sub AssertConcentration(ByVal condition As Boolean, ByVal label As String)
    checks = checks + 1
    If Not condition Then Err.Raise vbObjectError + 550, , "FAIL: " & label
End Sub

Public Function RunConcentrationTests(Optional ByVal previewDir As String = "") As String
    Dim wb As Workbook, source As Worksheet, report As Worksheet, cache As Range, r As Range, a() As Variant
    Dim i As Long, block As Long, pos As Long, record As Long, value As Variant, oldEvents As Boolean, caught As Long
    Dim curve As Variant, histogram As Variant, totalBins As Long, reportName As String, examplePath As String
    On Error GoTo failed
    checks = 0: oldEvents = Application.EnableEvents
    Set wb = Workbooks.Add(xlWBATWorksheet): Set source = wb.Worksheets(1)
    ReDim a(1 To 102, 1 To 5)
    a(1, 1) = "IncidentID": a(1, 2) = "CAP_BlockKey": a(1, 3) = "CAP_HundredBlock": a(1, 4) = "CrimeType": a(1, 5) = "OffenseID"
    pos = 1
    For block = 1 To 20
        If block = 1 Then
            record = 65
        ElseIf block <= 17 Then
            record = 2
        Else
            record = 1
        End If
        For i = 1 To record
            pos = pos + 1: a(pos, 1) = "I" & Format$(pos - 1, "0000")
            a(pos, 2) = "IL:TESTCITY:B" & Format$(block, "00")
            a(pos, 3) = block * 100 & " BLOCK TEST ST " & block
            If block = 1 Then
                a(pos, 4) = IIf(i <= 45, "Theft", "Robbery")
            Else
                a(pos, 4) = IIf(block <= 17, "Theft", "Robbery")
            End If
            a(pos, 5) = "O" & pos - 1
        Next
    Next
    For i = 1 To 4: a(102, i) = a(2, i): Next
    a(102, 5) = "SECOND OFFENSE RECORD"
    Set r = source.Range("A1:E102"): r.NumberFormat = "@": r.Value2 = a
    BuildConcentrationReport r, "CAP_BlockKey", "CrimeType", "IncidentID", "CAP_HundredBlock", False
    Set report = wb.Worksheets(2): Set cache = ConcentrationCache(report)
    AssertConcentration Not cache Is Nothing And cache.Parent.Visible = xlSheetVeryHidden, "Local snapshot exists and is hidden"
    AssertConcentration source.Range("E102").Value2 = "SECOND OFFENSE RECORD" And source.UsedRange.Columns.Count = 5, "Source unchanged"
    AssertConcentration report.Range("B8").Value2 = 100 And report.Range("B9").Value2 = 20, "Unique incidents and populated blocks"
    AssertConcentration Abs(report.Range("D18").Value2 - 0.65) < 0.00000001 And report.Range("C18").Value2 = 0.05, "Top 5 percent equals calculated 65 percent in synthetic fixture"
    AssertConcentration report.Range("B18").Value2 = 1 And report.Range("B22").Value2 = 1, "Rounded cutoff and 50 percent threshold"
    AssertConcentration report.ChartObjects.Count = 3, "Three distribution charts"
    RefreshConcentration report
    AssertConcentration Left$(report.Range("A25").Value2, 8) = "Among 20", "Headline names eligible populated block denominator"
    AssertConcentration InStr(report.Range("A32").Value2, "zero crime") > 0, "Research context distinguishes zero crime places"
    curve = report.ChartObjects("CAP_ConcentrationCurve").Chart.SeriesCollection(1).Values
    AssertConcentration curve(1) = 0 And curve(UBound(curve)) = 1, "Curve origin and full activity reconciliation"
    histogram = report.ChartObjects("CAP_BlockDistribution").Chart.SeriesCollection(1).Values
    totalBins = 0: For i = LBound(histogram) To UBound(histogram): totalBins = totalBins + histogram(i): Next
    AssertConcentration totalBins = 20, "Distribution bins sum to block denominator"
    AssertConcentration report.Range("B3").Validation.Type = xlValidateList, "Crime dropdown exists"
    If Len(previewDir) > 0 Then
        Application.ScreenUpdating = True: report.Activate
        ExportConcentrationPreview report.ChartObjects("CAP_ConcentrationCurve"), previewDir & "\concentration-curve.png"
        ExportConcentrationPreview report.ChartObjects("CAP_TopBlocks"), previewDir & "\highest-crime-blocks.png"
        ExportConcentrationPreview report.ChartObjects("CAP_BlockDistribution"), previewDir & "\block-distribution.png"
    End If
    Application.EnableEvents = True
    report.Range("B4").Value2 = "Type: ROBBERY"
    AssertConcentration report.Range("B8").Value2 = 23 And report.Range("B9").Value2 = 4, "Crime selector triggers automatic recalculation"
    AssertConcentration Abs(report.Range("D18").Value2 - 20 / 23) < 0.00000001 And report.Range("C18").Value2 = 0.25, "Selected-crime block universe and honest rounding"
    AssertConcentration InStr(report.Range("A25").Value2, "25.0%") > 0, "Small-sample actual block share visible"
    AssertConcentration report.AutoFilterMode, "Rank filter remains enabled after refresh"
    report.Range("B4").Value2 = "Type: THEFT"
    AssertConcentration report.Range("B8").Value2 = 77 And report.Range("B9").Value2 = 17, "Theft filter counts incidents once"
    report.Range("B5").Value2 = "Source records"
    AssertConcentration report.Range("B8").Value2 = 78, "Record measure retains extra offense row"
    report.Range("B4").Value2 = "All crime types"
    AssertConcentration report.Range("B8").Value2 = 101 And report.Range("B9").Value2 = 20, "All source records reconcile"
    report.Range("B4").Value2 = "Invalid manual paste"
    AssertConcentration report.ChartObjects.Count = 0 And IsEmpty(report.Range("D18").Value2) And InStr(report.Range("B14").Value2, "failed") > 0, "Invalid pasted filter clears stale metrics and charts"
    report.Range("B4").Value2 = "All crime types": report.Range("B5").Value2 = "Unique incidents"
    AssertConcentration report.Range("B8").Value2 = 100 And report.ChartObjects.Count = 3, "Valid selector recovers report"
    Application.EnableEvents = False
    report.Range("B4").Value2 = "Type: ROBBERY": report.Range("B14").Calculate
    AssertConcentration InStr(report.Range("B14").Value2, "Filter changed") > 0 And report.Range("E4").Value2 = "All crime types", "Disabled events expose stale filter state"
    ' Deliberately invalidate one incident across the entire snapshot, then filter to a crime.
    cache.Cells(2, 5).Value2 = "Missing/conflicting incident block"
    cache.Cells(102, 5).Value2 = "Missing/conflicting incident block"
    report.Range("B4").Value2 = "Type: THEFT": RefreshConcentration report, True
    AssertConcentration report.Range("B8").Value2 = 76 And report.Range("B13").Value2 = 2, "Global incident validity and exclusion coverage survive filters"
    For i = 2 To cache.Rows.Count: cache.Cells(i, 5).Value2 = "Missing ID": cache.Cells(i, 1).Value2 = "": Next
    RefreshConcentration report, True
    AssertConcentration report.Range("B8").Value2 = 0 And report.ChartObjects.Count = 0 And IsEmpty(report.Range("D18").Value2), "No eligible activity is unavailable rather than stale concentration"
    Application.EnableEvents = oldEvents
    wb.Close False
    RunConcentrationTests = "PASS: " & checks & " concentration assertions: 5%/65% fixture, live crime/measure filters, charts, rounding, exclusions, source preservation and empty/error states"
    Exit Function
failed:
    Dim description As String: description = Err.Description
    Application.EnableEvents = oldEvents
    If Not wb Is Nothing Then wb.Close False
    Err.Raise vbObjectError + 551, , description
End Function

Private Sub ExportConcentrationPreview(ByVal picture As ChartObject, ByVal path As String)
    picture.Activate: picture.Chart.Refresh: DoEvents
    If Not picture.Chart.Export(path, "PNG") Then Err.Raise vbObjectError + 552, , "Native chart export failed."
    If FileLen(path) = 0 Then Err.Raise vbObjectError + 553, , "Native chart export was empty."
End Sub
