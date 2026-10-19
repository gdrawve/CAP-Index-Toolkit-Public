Attribute VB_Name = "modTimeReports"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit
Public Const TIME_DATA_NAME As String = "_CAPTimeData"
Public Const TIME_BODY_NAME As String = "_CAPTimeBody"

Public Function TimeReportCache(ByVal ws As Worksheet) As Range
    Dim named As Name, r As Range
    On Error Resume Next
    For Each named In ws.Names
        If Right$(named.Name, Len(TIME_DATA_NAME)) = TIME_DATA_NAME Then
            Set r = named.RefersToRange
            If Not r Is Nothing Then
                If r.Parent.Parent Is ws.Parent Then
                    If r.Columns.Count = 10 And r.Rows.Count >= 2 Then Set TimeReportCache = r
                End If
            End If
            Exit Function
        End If
    Next
    On Error GoTo 0
End Function

Private Function NewTimeSheet(ByVal wb As Workbook, ByVal prefix As String) As Worksheet
    Dim ws As Worksheet, i As Long, candidate As String, exists As Boolean
    For i = 1 To 9999
        candidate = Left$(prefix, 25) & Format$(i, "0000"): exists = False
        For Each ws In wb.Worksheets
            If ws.Name = candidate Then exists = True: Exit For
        Next
        If Not exists Then Exit For
    Next
    If exists Then Err.Raise vbObjectError + 600, , "No available analysis sheet name."
    Set NewTimeSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    NewTimeSheet.Name = candidate
End Function

Public Sub BuildTemporalHeatmap(Optional ByVal source As Range, Optional ByVal dateHeader As String = "OccurredAt", Optional ByVal crimeHeader As String = "CrimeType", Optional ByVal incidentHeader As String = "IncidentID", Optional ByVal hourHeader As String = "", Optional ByVal showSummary As Boolean = True, Optional ByVal dayHeader As String = "", Optional ByVal timeHeader As String = "")
    If source Is Nothing Then OpenToolSetup "capHeatmap": Exit Sub
    BuildTimeReport "Heatmap", source, dateHeader, crimeHeader, incidentHeader, "", "", hourHeader, showSummary, dayHeader, timeHeader
End Sub

Public Sub BuildRepeatLocations(Optional ByVal source As Range, Optional ByVal locationHeader As String = "Address", Optional ByVal dateHeader As String = "OccurredAt", Optional ByVal crimeHeader As String = "CrimeType", Optional ByVal incidentHeader As String = "IncidentID", Optional ByVal labelHeader As String = "", Optional ByVal showSummary As Boolean = True, Optional ByVal cityHeader As String = "", Optional ByVal stateHeader As String = "")
    If source Is Nothing Then OpenToolSetup "capRepeat": Exit Sub
    BuildTimeReport "RepeatLocations", source, "", crimeHeader, incidentHeader, locationHeader, labelHeader, "", showSummary, "", "", cityHeader, stateHeader
End Sub

Public Sub BuildRepeatLocationsWithGaps(Optional ByVal source As Range, Optional ByVal locationHeader As String = "Full Address", Optional ByVal dateHeader As String = "OccurredAt", Optional ByVal crimeHeader As String = "CrimeType", Optional ByVal incidentHeader As String = "IncidentID", Optional ByVal labelHeader As String = "", Optional ByVal showSummary As Boolean = True)
    BuildTimeReport "RepeatGaps", source, dateHeader, crimeHeader, incidentHeader, locationHeader, labelHeader, "", showSummary
End Sub

Public Sub BuildHotStreets(Optional ByVal source As Range, Optional ByVal locationHeader As String = "CAP_HundredBlock", Optional ByVal dateHeader As String = "OccurredAt", Optional ByVal crimeHeader As String = "CrimeType", Optional ByVal incidentHeader As String = "IncidentID", Optional ByVal labelHeader As String = "CAP_HundredBlock", Optional ByVal showSummary As Boolean = True, Optional ByVal cityHeader As String = "", Optional ByVal stateHeader As String = "")
    If source Is Nothing Then OpenToolSetup "capHotStreets": Exit Sub
    BuildTimeReport "HotStreets", source, dateHeader, crimeHeader, incidentHeader, locationHeader, labelHeader, "", showSummary, "", "", cityHeader, stateHeader
End Sub

Private Sub BuildTimeReport(ByVal kind As String, ByVal source As Range, ByVal dateHeader As String, ByVal crimeHeader As String, ByVal incidentHeader As String, ByVal locationHeader As String, ByVal labelHeader As String, ByVal hourHeader As String, ByVal showSummary As Boolean, Optional ByVal dayHeader As String = "", Optional ByVal timeHeader As String = "", Optional ByVal cityHeader As String = "", Optional ByVal stateHeader As String = "")
    Dim r As Range, a As Variant, data() As Variant, wb As Workbook, report As Worksheet, cache As Worksheet
    Dim dc As Long, cc As Long, ic As Long, lc As Long, labelCol As Long, hc As Long, n As Long, i As Long, j As Long
    Dim parsed As Date, hasTime As Boolean, valid As Boolean, status As String, id As String, location As String, signature As String
    Dim ids As Object, types As Object, category As String, item As Variant, labels() As Variant, h As Variant
    Dim mode As Variant, dayCol As Long, timeCol As Long, dayNum As Long, timeHour As Long, cityCol As Long, stateCol As Long
    Dim requirements As String, timeProblem As String
    Dim minDate As Double, maxDate As Double, named As Name, oldEvents As Boolean, changing As Boolean, offset As Double
    Dim description As String, errNo As Long, headers As Variant
    On Error GoTo failed
    Set r = source
    If r Is Nothing Then
        If kind = "Heatmap" Then
            mode = Application.InputBox("Choose inputs, then select headers/data:" & vbCrLf & "1 Date/timestamp (optional Hour)" & vbCrLf & "2 Separate Date + Time" & vbCrLf & "3 Day of Week text + Hour 0-23" & vbCrLf & "4 Day of Week text + Time" & vbCrLf & "Include Crime Type; IncidentID optional.", "CAP Index | Temporal Heat Map", "2", Type:=1)
            If VarType(mode) = vbBoolean Then Exit Sub
            If mode < 1 Or mode > 4 Or mode <> Fix(mode) Then Err.Raise vbObjectError + 607, , "Choose input mode 1, 2, 3 or 4."
            requirements = "Include Crime Type, optional IncidentID, and " & Choose(CLng(mode), "Date/timestamp (optional Hour).", "separate Date and Time.", "Day of Week and Hour (0-23).", "Day of Week and Time.")
        ElseIf kind = "RepeatLocations" Then
            requirements = "Include Address (or Full Address), Crime Type and optional IncidentID, City and State. Dates and times are NOT required. Addresses are grouped exactly after case/space normalization; apartment/unit text is retained."
        Else
            requirements = "Include Hundred Block (or stable street ID), Crime Type and Date. Date-only values work; time is NOT required. Date assigns events to weeks/months so change can be measured. Include optional IncidentID, City and State."
        End If
        Set r = PickSource(requirements)
    End If
    If r Is Nothing Then Exit Sub
    Set wb = r.Parent.Parent
    If wb Is ThisWorkbook Or wb.ReadOnly Or wb.ProtectStructure Then Err.Raise vbObjectError + 601, , "Select a writable workbook with unprotected structure."
    If r.Areas.Count <> 1 Or r.Rows.Count < 2 Or r.Rows.Count > 200001 Or r.Columns.Count > 200 Then Err.Raise vbObjectError + 602, , "Select source headers and up to 200,000 data rows."
    a = ScopedValues(r): n = UBound(a, 1) - 1
    If source Is Nothing Then
        If kind <> "RepeatLocations" Then
            If kind = "Heatmap" And mode >= 3 Then
                dayCol = AskColumn(a, "Day of Week text: Monday-Sunday or Mon-Sun.", "Day of Week", False, "DAY OF WEEK")
            Else
                dc = AskColumn(a, "Occurrence DATE. Date-only is sufficient for Hot Streets; text dates use month/day/year.", "Date", False, "DATE")
            End If
        End If
        cc = AskColumn(a, "Crime-type field for filtering.", "CrimeType", False, "CRIME")
        ic = AskColumn(a, "Optional stable IncidentID; leave blank to count source records.", "", True, "INCIDENTID")
        If kind = "Heatmap" Then
            If mode = 2 Or mode = 4 Then
                timeCol = AskColumn(a, "TIME column. Accepts Excel times, 24-hour text or AM/PM. Missing time stays excluded.", "Time", False, "TIME")
            ElseIf mode = 3 Then
                hc = AskColumn(a, "Hour field: integer 0-23. Blank never means midnight.", IIf(OptionalHeader(a, "Hour from Time (0-23)") > 0, "Hour from Time (0-23)", "Hour (0-23)"), False, "HOUR")
            Else
                hc = AskColumn(a, "Optional Hour field: integer 0-23.", "", True, "HOUR")
            End If
        End If
        If kind <> "Heatmap" Then
            lc = AskColumn(a, IIf(kind = "RepeatLocations", "ADDRESS field. Use the full address; apartment/unit text is retained.", "HUNDRED-BLOCK field or stable street ID. City/state are combined internally."), locationHeader, False, IIf(kind = "HotStreets", "HUNDRED-BLOCK", "ADDRESS"))
            labelCol = AskColumn(a, "Optional readable location label.", "", True, "LOCATION LABEL")
            cityCol = AskColumn(a, "Optional City to prevent combining identical addresses/blocks across cities.", "", True, "CITY")
            stateCol = AskColumn(a, "Optional State.", "", True, "STATE")
        End If
    Else
        If kind <> "RepeatLocations" And Len(dayHeader) = 0 Then dc = HeaderColumn(a, dateHeader)
        cc = HeaderColumn(a, crimeHeader): ic = HeaderColumn(a, incidentHeader)
        cityCol = HeaderColumn(a, cityHeader): stateCol = HeaderColumn(a, stateHeader)
        dayCol = HeaderColumn(a, dayHeader): timeCol = HeaderColumn(a, timeHeader)
        hc = HeaderColumn(a, hourHeader): lc = HeaderColumn(a, locationHeader): labelCol = HeaderColumn(a, labelHeader)
    End If
    Set ids = CreateObject("Scripting.Dictionary"): Set types = CreateObject("Scripting.Dictionary")
    ReDim data(1 To n + 1, 1 To 10)
    headers = Array("IncidentID", "TimestampVBA", "DateVBA", "Hour", "LocationKey", "LocationLabel", "CrimeCategory", "IncidentValidity", "DateStatus", "SourceRow")
    For j = 0 To 9: data(1, j + 1) = headers(j): Next
    For i = 2 To n + 1
        id = "": location = ""
        If ic > 0 Then
            If Not IsError(a(i, ic)) Then id = SafeText(a(i, ic))
        End If
        If lc > 0 Then
            If Not IsError(a(i, lc)) Then location = NormalText(a(i, lc))
            If Len(location) > 0 Then
                If cityCol > 0 Then location = KeyPart(NormalText(a(i, cityCol))) & KeyPart(location)
                If stateCol > 0 Then location = KeyPart(NormalText(a(i, stateCol))) & KeyPart(location)
            End If
            If cityCol > 0 Then
                If IsError(a(i, cityCol)) Or Len(SafeText(a(i, cityCol))) = 0 Then location = ""
            End If
            If stateCol > 0 Then
                If IsError(a(i, stateCol)) Or Len(SafeText(a(i, stateCol))) = 0 Then location = ""
            End If
        End If
        data(i, 1) = id: data(i, 5) = location: data(i, 6) = location
        If lc > 0 Then data(i, 6) = NormalText(a(i, lc))
        If labelCol > 0 Then
            If Not IsError(a(i, labelCol)) And Len(SafeText(a(i, labelCol))) > 0 Then data(i, 6) = SafeText(a(i, labelCol))
        End If
        If cityCol > 0 Then data(i, 6) = data(i, 6) & ", " & SafeText(a(i, cityCol))
        If stateCol > 0 Then data(i, 6) = data(i, 6) & ", " & SafeText(a(i, stateCol))
        category = CrimeCategory(a(i, cc)): data(i, 7) = category
        If Not types.Exists(category) Then types.Add category, True
        valid = False: hasTime = False
        If kind = "RepeatLocations" Then
            parsed = DateSerial(2024, 1, 1): valid = True: status = "Date not required"
        ElseIf dayCol > 0 Then
            dayNum = DayOfWeekNumber(a(i, dayCol))
            If dayNum > 0 Then parsed = DateSerial(2024, 1, 1) + dayNum - 1: valid = True: status = "Day-of-week input" Else status = "Invalid/missing day of week"
        ElseIf dc > 0 Then
            valid = ParseTimestamp(a(i, dc), wb.Date1904, parsed, hasTime, status)
        End If
        If valid Then
            data(i, 2) = CDbl(parsed): data(i, 3) = CDbl(DateValue(parsed))
            If hasTime Then data(i, 4) = Hour(parsed)
            If timeCol > 0 Then
                If HourFromTime(a(i, timeCol), wb.Date1904, timeHour) Then
                    If hasTime And timeHour <> Hour(parsed) Then data(i, 4) = Empty: status = "Time contradicts timestamp" Else data(i, 4) = timeHour
                Else
                    data(i, 4) = Empty: status = "Invalid/missing separate time"
                End If
            End If
            If hc > 0 Then
                h = a(i, hc)
                If IsError(h) Then
                    data(i, 4) = Empty: status = "Hour cell error"
                ElseIf Len(SafeText(h)) = 0 Then
                    data(i, 4) = Empty: status = "Missing explicit hour"
                ElseIf Not IsNumeric(h) Then
                    data(i, 4) = Empty: status = "Invalid explicit hour"
                ElseIf CDbl(h) < 0 Or CDbl(h) > 23 Or CDbl(h) <> Fix(CDbl(h)) Then
                    data(i, 4) = Empty: status = "Hour must be an integer 0-23"
                ElseIf hasTime And CLng(h) <> Hour(parsed) Then
                    data(i, 4) = Empty: status = "Hour contradicts timestamp"
                Else
                    data(i, 4) = CLng(h)
                End If
            End If
            If minDate = 0 Or data(i, 3) < minDate Then minDate = data(i, 3)
            If data(i, 3) > maxDate Then maxDate = data(i, 3)
        End If
        data(i, 9) = status: data(i, 10) = ScopeSourceRow(r, i)
        If Len(id) > 0 Then
            signature = KeyPart(location) & KeyPart(SafeText(data(i, 3))) & KeyPart(SafeText(data(i, 4)))
            If valid Then signature = signature & KeyPart(Format$(parsed, "yyyy-mm-dd hh:nn:ss"))
            If Not valid Or (kind <> "Heatmap" And Len(location) = 0) Then signature = ""
            If Not ids.Exists(id) Then
                ids.Add id, signature
            ElseIf ids(id) <> signature Then
                ids(id) = ""
            End If
        End If
    Next
    For i = 2 To n + 1
        id = data(i, 1)
        If Len(id) = 0 Then
            data(i, 8) = "Missing ID"
        ElseIf Len(ids(id)) = 0 Then
            data(i, 8) = "Invalid/conflicting incident date/time/location"
        Else
            data(i, 8) = "Eligible"
        End If
    Next
    oldEvents = Application.EnableEvents: Application.EnableEvents = False: changing = True
    Set report = NewTimeSheet(wb, kind & "_"): Set cache = NewTimeSheet(wb, "_CAPTime_")
    cache.Range("A1").Resize(n + 1, 10).NumberFormat = "@": cache.Range("A1").Resize(n + 1, 10).Value2 = data
    cache.Range("N1").Value2 = kind: cache.Range("N2").Value2 = minDate: cache.Range("N3").Value2 = maxDate
    cache.Range("N5").Value2 = dayHeader: If dayCol > 0 Then cache.Range("N5").Value2 = "Day field"
    cache.Range("N4").Value2 = IIf(lc > 0, "Mapped location key", "No geography required")
    If lc > 0 Then cache.Range("N4").Value2 = SafeText(a(1, lc))
    ReDim labels(1 To types.Count + 1, 1 To 1): labels(1, 1) = "All crime types": i = 1
    For Each item In types.Keys: i = i + 1: labels(i, 1) = item: Next
    cache.Range("L1").Resize(UBound(labels, 1)).NumberFormat = "@": cache.Range("L1").Resize(UBound(labels, 1)).Value2 = labels
    If types.Count > 1 Then cache.Range("L2:L" & types.Count + 1).Sort Key1:=cache.Range("L2"), Order1:=xlAscending, Header:=xlNo
    Set named = report.Names.Add(Name:=TIME_DATA_NAME, RefersTo:="='" & cache.Name & "'!" & cache.Range("A1:J" & n + 1).Address): named.Visible = False
    Set named = report.Names.Add(Name:="_CAPTimeTypes", RefersTo:="='" & cache.Name & "'!" & cache.Range("L1:L" & types.Count + 1).Address): named.Visible = False
    report.Range("A1:Y1").Merge: report.Range("A1").Value2 = "CAP Index | " & IIf(kind = "Heatmap", "Temporal Heat Map", IIf(kind = "RepeatLocations", "Repeat Locations", "Hot Streets"))
    report.Range("A2:Y2").Merge: report.Range("A2").Value2 = "Source snapshot: " & r.Parent.Name & "!" & r.Address(False, False) & " | " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & CAP_VERSION & " | " & CAP_ATTRIBUTION
    AnnotateAnalysisScope report, r
    report.Range("A4").Value2 = "Crime type": report.Range("B4").Value2 = "All crime types"
    report.Range("A5").Value2 = "Count measure": report.Range("B5").Value2 = IIf(ic > 0, "Unique incidents", "Source records")
    report.Range("D4").Value2 = "Displayed crime": report.Range("D5").Value2 = "Displayed measure"
    With report.Range("B4").Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=_CAPTimeTypes": .InCellDropdown = True: .ShowError = True
    End With
    With report.Range("B5").Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Unique incidents" & Application.International(xlListSeparator) & "Source records": .InCellDropdown = True: .ShowError = True
    End With
    If kind = "RepeatLocations" Or kind = "RepeatGaps" Then
        report.Range("A6").Value2 = "Minimum events": report.Range("B6").Value2 = 2
        If kind = "RepeatGaps" Then
        report.Range("A7").Value2 = "Gap window A (days)": report.Range("B7").Value2 = 7
        report.Range("A8").Value2 = "Gap window B (days)": report.Range("B8").Value2 = 30
        report.Range("A9").Value2 = "Gap window C (days)": report.Range("B9").Value2 = 90
        End If
        For i = 6 To IIf(kind = "RepeatGaps", 9, 6)
            report.Cells(i, 2).Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=IIf(i = 6, "2", "1"), Formula2:="36500"
        Next
    ElseIf kind = "HotStreets" Then
        report.Range("A6").Value2 = "Interval unit": report.Range("B6").Value2 = "Months"
        report.Range("B6").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Weeks" & Application.International(xlListSeparator) & "Months"
        report.Range("A7").Value2 = "Units per interval": report.Range("B7").Value2 = 1
        report.Range("B7").Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="120"
        report.Range("A8").Value2 = "Rise/fall threshold": report.Range("B8").Value2 = 0.25: report.Range("B8").NumberFormat = "0%"
        report.Range("B8").Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0.01", Formula2:="10"
        report.Range("D9").Value2 = "Grouping key": report.Range("E9").Value2 = cache.Range("N4").Value2
        report.Range("A9").Value2 = "Minimum comparison events": report.Range("B9").Value2 = HOT_STREETS_MIN_EVENTS
        report.Range("B9").Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="200000"
        report.Range("B9").Interior.Color = RGB(224, 239, 250)
    End If
    report.Columns("A").ColumnWidth = 28: report.Columns("B").ColumnWidth = 38: report.Columns("C:Y").ColumnWidth = 12
    report.Range("A1:Y16").Font.Name = "Calibri": report.Range("A1").Font.Size = 20: report.Rows(1).RowHeight = 34
    report.Range("A1:Y1").Interior.Color = RGB(26, 55, 82): report.Range("A1:Y1").Font.Color = vbWhite
    report.Range("B4:B8").Interior.Color = RGB(224, 239, 250)
    report.Range("B14:Y14").Merge: report.Range("B14").WrapText = True: report.Rows(14).RowHeight = 32
    report.Range("A15:Y15").Merge: report.Range("A15").Value2 = "Saved snapshot. Rebuild after source changes. The saved scope is shown in the source line. Hidden data contains IDs and location keys; share an approved summary/export."
    report.Range("A16:Y16").Merge: report.Range("A16").Value2 = "Location analyses use the supplied key's granularity. Observed date span does not verify agency coverage; no-record periods are not proven zero-crime periods. Profiles are descriptive, not fitted trajectory models."
    report.Range("A15:Y16").WrapText = True: report.Rows("15:16").RowHeight = 30
    If kind = "Heatmap" Then
        report.Range("D4:E5").ClearContents
        report.Range("B4:H4").Merge: report.Range("B5:H5").Merge
        report.Range("B6:B8").Interior.Pattern = xlNone
        report.Range("D12:E13").ClearContents
    Else
        For i = 4 To 9: report.Range(report.Cells(i, 2), report.Cells(i, 3)).Merge: Next
        report.Range("E4:H4").Merge: report.Range("E5:H5").Merge
    End If
    If kind = "RepeatLocations" Then report.Range("A16").Value2 = "Address-specific counts; no date/time requirement. Address spelling aliases are not automatically resolved. Units/apartments are retained. Hidden snapshots contain IDs and location keys."
    If kind = "Heatmap" Then report.Range("A16").Value2 = "Day/hour counts. Missing or invalid time/hour stays excluded. Day-of-week fields need Monday-Sunday text; numeric days are ambiguous and rejected. Hidden snapshots contain IDs."
    cache.Visible = xlSheetVeryHidden: InitializeCAPEvents: RefreshTimeReport report, True
    report.Activate: report.Range("B4").Select
    Application.EnableEvents = oldEvents: changing = False
    If showSummary Then MsgBox "Output: " & report.Name & ". Eligible events: " & report.Range("B10").Value2 & ". Exclusions: see rows 11-14. Analysis report created. Use the blue crime/measure cells" & IIf(kind = "HotStreets", " and interval controls", "") & " to update the saved snapshot. Review coverage and scope notes.", vbInformation, "CAP Index"
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    If changing Then Application.EnableEvents = oldEvents
    If Not source Is Nothing Then Err.Raise errNo, kind, description
    MsgBox kind & ": " & description, vbExclamation, "CAP Index"
End Sub

Public Function TimeEvents(ByVal data As Variant, ByVal category As String, ByVal measure As String, ByVal kind As String, ByRef coverage As Variant) As Collection
    Dim events As New Collection, seen As Object, i As Long, id As String, missingDates As Long, missingHours As Long, missingLocations As Long, excluded As Long, selected As Long
    Set seen = CreateObject("Scripting.Dictionary")
    If measure <> "Unique incidents" And measure <> "Source records" Then Err.Raise vbObjectError + 603, , "Select a valid count measure."
    For i = 2 To UBound(data, 1)
        If category = "All crime types" Or category = data(i, 7) Then
            selected = selected + 1: id = SafeText(data(i, 1))
            If Not IsNumeric(data(i, 3)) Or IsEmpty(data(i, 3)) Then
                missingDates = missingDates + 1
            ElseIf kind = "Heatmap" And IsEmpty(data(i, 4)) Then
                missingHours = missingHours + 1
            ElseIf kind <> "Heatmap" And Len(SafeText(data(i, 5))) = 0 Then
                missingLocations = missingLocations + 1
            ElseIf measure = "Unique incidents" And data(i, 8) <> "Eligible" Then
                excluded = excluded + 1
            ElseIf measure = "Source records" Or Not seen.Exists(id) Then
                events.Add Array(CDbl(data(i, 2)), CDbl(data(i, 3)), data(i, 4), CStr(data(i, 5)), CStr(data(i, 6)), id)
                If Len(id) > 0 Then seen(id) = True
            End If
        End If
    Next
    coverage = Array(selected, missingDates, missingHours, missingLocations, excluded)
    Set TimeEvents = events
End Function

Public Sub RefreshTimeReport(Optional ByVal report As Worksheet, Optional ByVal propagateErrors As Boolean = False)
    Dim cache As Range, data As Variant, events As Collection, coverage As Variant, category As String, measure As String, kind As String
    Dim oldEvents As Boolean, oldScreen As Boolean, changing As Boolean, named As Name, body As Range, i As Long, permitted As Boolean, v As Variant
    Dim errNo As Long, description As String
    On Error GoTo failed
    If report Is Nothing Then Set report = ActiveSheet
    If Not ReportingCache(report) Is Nothing Then
        RefreshReportingFilter report, propagateErrors: Exit Sub
    End If
    Set cache = TimeReportCache(report): If cache Is Nothing Then Err.Raise vbObjectError + 604, , "Select a CAP Index temporal/location report."
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    changing = True: Application.EnableEvents = False: Application.ScreenUpdating = False
    report.Range("B10:Y13").ClearContents
    For Each named In report.Names
        If Right$(named.Name, Len(TIME_BODY_NAME)) = TIME_BODY_NAME Then Set body = named.RefersToRange: body.Clear: Exit For
    Next
    If report.AutoFilterMode Then report.AutoFilterMode = False
    For i = report.ChartObjects.Count To 1 Step -1
        If Left$(report.ChartObjects(i).Name, 8) = "CAPTime_" Then report.ChartObjects(i).Delete
    Next
    category = SafeText(report.Range("B4").Value2): measure = SafeText(report.Range("B5").Value2): kind = cache.Parent.Range("N1").Value2
    For Each v In cache.Parent.Range("L1:L" & cache.Parent.Cells(cache.Parent.Rows.Count, 12).End(xlUp).Row).Value2
        If category = v Then permitted = True: Exit For
    Next
    If Not permitted Then Err.Raise vbObjectError + 605, , "Choose a crime type from the report dropdown."
    data = cache.Value2: Set events = TimeEvents(data, category, measure, kind, coverage)
    report.Range("A10").Value2 = "Eligible events": report.Range("B10").Value2 = events.Count
    report.Range("A11").Value2 = "Selected source rows": report.Range("B11").Value2 = coverage(0)
    report.Range("A12").Value2 = IIf(kind = "Heatmap", "Invalid/missing date or day", "Missing/invalid dates"): report.Range("B12").Value2 = coverage(1)
    report.Range("A13").Value2 = "Other excluded rows": report.Range("B13").Value2 = coverage(2) + coverage(3) + coverage(4)
    If kind = "Heatmap" Then
        report.Range("D12:H12").Merge: report.Range("D13:H13").Merge
        report.Range("D12").Value2 = "Unknown hours: " & coverage(2)
        report.Range("D13").Value2 = "Invalid IDs: " & coverage(4)
    Else
        report.Range("D12").Value2 = "Unknown hours": report.Range("E12").Value2 = coverage(2)
        report.Range("D13").Value2 = "Invalid ID/location": report.Range("E13").Value2 = coverage(3) + coverage(4)
    End If
    If kind <> "Heatmap" Then
        report.Range("E4").Value2 = category: report.Range("E5").Value2 = measure
    End If
    If kind = "RepeatLocations" Then
        report.Range("A12").Value2 = "Missing addresses": report.Range("B12").Value2 = coverage(3)
        report.Range("A13").Value2 = "Missing/conflicting IDs": report.Range("B13").Value2 = coverage(4)
        report.Range("D12:E13").ClearContents
    ElseIf kind = "HotStreets" Then
        report.Range("D12:E12").ClearContents
    End If
    Select Case kind
        Case "Heatmap": RenderTemporalHeatmap report, events
        Case "RepeatLocations": RenderAddressRepeats report, events
        Case "RepeatGaps": RenderRepeatLocations report, events
        Case "HotStreets": RenderHotStreets report, events, CDbl(cache.Parent.Range("N2").Value2), CDbl(cache.Parent.Range("N3").Value2)
        Case Else: Err.Raise vbObjectError + 606, , "Unknown report kind."
    End Select
    cache.Parent.Range("P4:P9").Value2 = report.Range("B4:B9").Value2
    description = "AND("
    For i = 4 To 9
        If i > 4 Then description = description & ","
        description = description & "B" & i & "='" & cache.Parent.Name & "'!P" & i
    Next
    description = description & ")"
    report.Range("B14").Formula = "=IF(" & description & ",""Calculated from saved snapshot; review coverage and scope notes."",""Controls changed: enable CAP Index and refresh. Existing results use previous controls."")"
    report.Range("B14").Calculate
    CommonCrimeControl report: StampReportStatus report
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    On Error Resume Next
    If Not report Is Nothing Then report.Range("B14").Value2 = "Refresh failed: " & description
    If changing Then Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    If propagateErrors Then Err.Raise errNo, "Temporal/location analysis", description
End Sub

Public Sub MarkTimeBody(ByVal ws As Worksheet, ByVal body As Range)
    Dim named As Name
    Set named = ws.Names.Add(Name:=TIME_BODY_NAME, RefersTo:="='" & Replace(ws.Name, "'", "''") & "'!" & body.Address)
    named.Visible = False
End Sub

Public Sub TimeTableStyle(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    With ws.Range(ws.Cells(18, 1), ws.Cells(18, lastCol))
        .Font.Bold = True: .Interior.Color = RGB(26, 55, 82): .Font.Color = vbWhite: .WrapText = True
    End With
    ws.Rows(18).RowHeight = 42
    ws.Range(ws.Cells(18, 1), ws.Cells(lastRow, lastCol)).Font.Name = "Calibri"
    ws.Range(ws.Cells(18, 1), ws.Cells(lastRow, lastCol)).AutoFilter
End Sub
