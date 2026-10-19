Attribute VB_Name = "modReportInput"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit

Public Function ReadReportInput(ByVal source As Range, ByVal dateHeader As String, ByVal crimeHeader As String, ByVal incidentHeader As String, ByVal locationHeader As String, ByVal crimeFilter As String, ByVal measure As String, ByVal firstDate As Double, ByVal lastDate As Double, Optional ByVal promptCrime As Boolean = True) As Object
    Dim ctx As Object, r As Range, a As Variant, data() As Variant, dc As Long, cc As Long, ic As Long, lc As Long
    Dim i As Long, j As Long, d As Date, ht As Boolean, status As String, id As String, loc As String, signature As String
    Dim ids As Object, categories As Object, choice As Variant, minimum As Double, maximum As Double, missing As Long
    Dim list As String, category As Variant, response As Variant, parsed As Date, validDate As Boolean
    Set r = source
    If r Is Nothing Then Set r = PickSource("Reporting inputs: occurrence DATE, Crime Type, optional IncidentID and qualified Address/Hundred Block for location summaries. Include headers/data; Analysis Scope applies.")
    If r Is Nothing Then Exit Function
    ValidateOutputSource r: a = ScopedValues(r)
    If source Is Nothing Then
        dc = AskColumn(a, "Occurrence DATE or timestamp. Separate Time is not required.", "Date", False, "DATE")
        cc = AskColumn(a, "Crime type or offense description.", "CrimeType", False, "CRIME")
        ic = AskColumn(a, "Optional stable IncidentID. Missing/conflicting IDs are excluded in incident mode.", "", True, "INCIDENTID")
        lc = AskColumn(a, "Optional qualified Address or Hundred Block for location rankings. Include jurisdiction in the key when combining agencies.", "", True, "LOCATION")
    Else
        dc = HeaderColumn(a, dateHeader): cc = HeaderColumn(a, crimeHeader)
        ic = HeaderColumn(a, incidentHeader): lc = HeaderColumn(a, locationHeader)
    End If
    Set ids = CreateObject("Scripting.Dictionary"): Set categories = CreateObject("Scripting.Dictionary")
    ReDim data(1 To UBound(a, 1), 1 To 6)
    For i = 2 To UBound(a, 1)
        id = "": loc = ""
        If ic > 0 Then
            If Not IsError(a(i, ic)) Then id = SafeText(a(i, ic))
        End If
        If lc > 0 Then
            If Not IsError(a(i, lc)) Then loc = NormalText(a(i, lc))
        End If
        validDate = ParseTimestamp(a(i, dc), r.Parent.Parent.Date1904, d, ht, status)
        If validDate Then
            data(i, 1) = CDbl(DateValue(d))
            If minimum = 0 Or data(i, 1) < minimum Then minimum = data(i, 1)
            If data(i, 1) > maximum Then maximum = data(i, 1)
        Else
            missing = missing + 1
        End If
        data(i, 2) = CrimeCategory(a(i, cc)): data(i, 3) = id: data(i, 4) = loc: data(i, 6) = ScopeSourceRow(r, i)
        If Not categories.Exists(data(i, 2)) Then categories.Add data(i, 2), True
        If Len(id) > 0 Then
            signature = "": If validDate Then signature = KeyPart(CStr(data(i, 1))) & KeyPart(loc)
            If Not ids.Exists(id) Then
                ids.Add id, signature
            ElseIf ids(id) <> signature Then
                ids(id) = ""
            End If
        End If
    Next
    For i = 2 To UBound(a, 1)
        id = data(i, 3): data(i, 5) = False
        If Len(id) > 0 Then data(i, 5) = Len(ids(id)) > 0
    Next
    If minimum = 0 Then Err.Raise vbObjectError + 800, , "No valid occurrence dates were found."
    If source Is Nothing Then
        If promptCrime Then
        list = "0 = All crime types": i = 0
        For Each category In categories.Keys: i = i + 1: list = list & vbCrLf & i & " = " & category: Next
        If Len(list) > 900 Then
            response = Application.InputBox("Enter the exact crime label, or All crime types. Crime labels use Type: plus the uppercase source description.", "CAP Index | Crime filter", "All crime types", Type:=2)
            If VarType(response) = vbBoolean Then Exit Function
            crimeFilter = SafeText(response)
        Else
            MsgBox list, vbInformation, "CAP Index | Available crime types"
            choice = Application.InputBox("Enter a crime-type number from the list (0 means all).", "CAP Index | Crime filter", "0", Type:=1)
            If VarType(choice) = vbBoolean Then Exit Function
            If choice < 0 Or choice > categories.Count Or choice <> Fix(choice) Then Err.Raise vbObjectError + 801, , "Invalid crime-type selection."
            crimeFilter = "All crime types"
            If choice > 0 Then response = categories.Keys: crimeFilter = response(CLng(choice) - 1)
        End If
        Else
            crimeFilter = "All crime types"
        End If
        choice = Application.InputBox("Count measure: 1 = Unique incidents; 2 = Source/offense records. Incident mode requires stable IDs.", "CAP Index | Measure", IIf(ic > 0, "1", "2"), Type:=1)
        If VarType(choice) = vbBoolean Then Exit Function
        If choice <> 1 And choice <> 2 Then Err.Raise vbObjectError + 802, , "Choose measure 1 or 2."
        measure = IIf(choice = 1, "Unique incidents", "Source records")
        firstDate = AskReportDate("Start date (inclusive). Observed source coverage is not proof of reporting completeness.", minimum)
        lastDate = AskReportDate("End date (inclusive).", maximum)
    End If
    If Len(crimeFilter) = 0 Then crimeFilter = "All crime types"
    If crimeFilter <> "All crime types" And Not categories.Exists(crimeFilter) Then Err.Raise vbObjectError + 803, , "Crime filter is not present in the source."
    If Len(measure) = 0 Then measure = IIf(ic > 0, "Unique incidents", "Source records")
    If measure <> "Unique incidents" And measure <> "Source records" Then Err.Raise vbObjectError + 804, , "Invalid count measure."
    If measure = "Unique incidents" And ic = 0 Then Err.Raise vbObjectError + 805, , "Map an IncidentID field or choose Source records."
    If firstDate = 0 Then firstDate = minimum
    If lastDate = 0 Then lastDate = maximum
    If firstDate > lastDate Or firstDate <> Fix(firstDate) Or lastDate <> Fix(lastDate) Then Err.Raise vbObjectError + 806, , "Choose ordered date-only period boundaries."
    Set ctx = CreateObject("Scripting.Dictionary")
    ctx.Add "Source", r: ctx.Add "Data", data: ctx.Add "Crime", crimeFilter: ctx.Add "Measure", measure
    ctx.Add "Start", firstDate: ctx.Add "End", lastDate: ctx.Add "SourceStart", minimum: ctx.Add "SourceEnd", maximum
    ctx.Add "IDMapped", ic > 0
    ctx.Add "MissingDates", missing: ctx.Add "LocationMapped", lc > 0
    Set ReadReportInput = ctx
End Function

Public Function AskReportDate(ByVal prompt As String, ByVal defaultDate As Double) As Double
    Dim answer As Variant, parsed As Date, hasTime As Boolean, status As String
    answer = Application.InputBox(prompt & " Use month/day/year or yyyy-mm-dd.", "CAP Index | Period", Format$(CDate(defaultDate), "mm/dd/yyyy"), Type:=2)
    If VarType(answer) = vbBoolean Then Err.Raise vbObjectError + 807, , "Operation cancelled."
    If Not ParseTimestamp(answer, False, parsed, hasTime, status) Then Err.Raise vbObjectError + 808, , "Invalid period date."
    If hasTime Then Err.Raise vbObjectError + 808, , "Enter a date without a time."
    AskReportDate = CDbl(parsed)
End Function

Public Function ReportRows(ByVal ctx As Object, Optional ByVal firstDate As Double = 0, Optional ByVal lastDate As Double = 0, Optional ByVal category As String = "") As Collection
    Dim data As Variant, rows As New Collection, seen As Object, i As Long, id As String
    Set seen = CreateObject("Scripting.Dictionary"): data = ctx("Data")
    If firstDate = 0 Then firstDate = ctx("Start")
    If lastDate = 0 Then lastDate = ctx("End")
    If Len(category) = 0 Then category = ctx("Crime")
    For i = 2 To UBound(data, 1)
        If Not IsEmpty(data(i, 1)) Then
            If data(i, 1) >= firstDate And data(i, 1) <= lastDate And (category = "All crime types" Or category = data(i, 2)) Then
                id = data(i, 3)
                If ctx("Measure") = "Source records" Then
                    rows.Add i
                ElseIf data(i, 5) And Not seen.Exists(id) Then
                    seen.Add id, True: rows.Add i
                End If
            End If
        End If
    Next
    Set ReportRows = rows
End Function

Public Sub ConfigureReport(ByVal ws As Worksheet, ByVal ctx As Object, ByVal summary As Range)
    Dim named As Name, r As Range
    Set r = ctx("Source")
    ws.Range("A1").Value2 = Replace(ws.Range("A1").Value2, "_", " ")
    ws.Range("A1:H1").Merge: ws.Range("A1").Font.Size = 20: ws.Rows(1).RowHeight = 34
    ws.Range("A1:H1").Interior.Color = RGB(26, 55, 82): ws.Range("A1:H1").Font.Color = vbWhite
    ws.Range("A2:H2").Merge: ws.Range("A2").WrapText = True: ws.Rows(2).RowHeight = 30
    ws.Range("A3:H3").Merge: ws.Range("A3").WrapText = True: ws.Rows(3).RowHeight = 42
    ws.Range("A3").Value2 = ctx("Crime") & " | " & ctx("Measure") & " | " & Format$(CDate(ctx("Start")), "mm/dd/yyyy") & " through " & Format$(CDate(ctx("End")), "mm/dd/yyyy")
    ws.Range("A4:H4").Merge: ws.Range("A4").Value2 = "Observed source span: " & Format$(CDate(ctx("SourceStart")), "mm/dd/yyyy") & " - " & Format$(CDate(ctx("SourceEnd")), "mm/dd/yyyy") & "; missing/invalid date rows: " & ctx("MissingDates")
    ws.Range("A4").WrapText = True: ws.Rows(4).RowHeight = 30
    ws.Range("A5:H5").Merge: ws.Range("A5").Value2 = "Saved report snapshot. Rebuild after source/filter changes. Zero-record periods do not verify zero crime. Missing/conflicting IDs excluded in incident mode."
    If ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1 > summary.Row + summary.Rows.Count - 1 Then
        ws.Range("A5").Value2 = "Snapshot. Image/slide summaries show the first " & summary.Rows.Count - 6 & " table rows; full results remain in this sheet. Rebuild after source/filter changes. Reporting coverage is unverified."
    End If
    ws.Range("A5").WrapText = True: ws.Rows(5).RowHeight = 30
    ws.Columns("A").ColumnWidth = 32: ws.Columns("B:H").ColumnWidth = 16
    Set named = ws.Names.Add(Name:="_CAPExportSummary", RefersTo:="='" & ws.Name & "'!" & summary.Address): named.Visible = False
    Set named = ws.Names.Add(Name:="_CAPReportBody", RefersTo:="='" & ws.Name & "'!" & ws.Range("A1:H" & ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1).Address): named.Visible = False
    ' PDF export configures a temporary workbook; dropdown refresh does not query the printer.
    ws.Range("A6:H" & ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1).Rows.AutoFit
    ws.Activate: ActiveWindow.DisplayGridlines = False
End Sub
