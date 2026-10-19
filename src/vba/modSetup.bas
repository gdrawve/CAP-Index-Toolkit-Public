Attribute VB_Name = "modSetup"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit

Public Function SetupValue(ByVal values As Object, ByVal key As String) As String
    If values.Exists(key) Then SetupValue = SafeText(values(key))
End Function
Public Function SetupRoleActive(ByVal toolId As String, ByVal key As String, ByVal selections As Object) As Boolean
    Dim mode As String
    SetupRoleActive = True
    If toolId <> "capHeatmap" Then Exit Function
    mode = SetupValue(selections, "TIME INPUT")
    Select Case key
        Case "DATE": SetupRoleActive = (mode = "Date or timestamp" Or mode = "Date plus separate Time")
        Case "TIME": SetupRoleActive = (mode = "Date plus separate Time" Or mode = "Day of Week plus Time")
        Case "HOUR": SetupRoleActive = (mode = "Date or timestamp" Or mode = "Day of Week plus Hour")
        Case "DAY OF WEEK": SetupRoleActive = (mode = "Day of Week plus Hour" Or mode = "Day of Week plus Time")
    End Select
End Function
Public Function SetupRoleRequired(ByVal spec As Object, ByVal role As Object, ByVal selections As Object) As Boolean
    Dim key As String, mode As String
    key = role("Key"): SetupRoleRequired = role("Required")
    If Not SetupRoleActive(spec("Id"), key, selections) Then SetupRoleRequired = False: Exit Function
    If spec("Id") <> "capHeatmap" Then Exit Function
    mode = SetupValue(selections, "TIME INPUT")
    If key = "DATE" Or key = "TIME" Or key = "DAY OF WEEK" Then SetupRoleRequired = True
    If key = "HOUR" Then SetupRoleRequired = (mode = "Day of Week plus Hour")
End Function
Public Function SetupProblem(ByVal spec As Object, ByVal headers As Variant, ByVal selections As Object, ByVal source As Range) As String
    Dim role As Object, item As Object, seen As Object, key As String, value As String, found As Long, j As Long
    Dim first As Date, last As Date, hasTime As Boolean, status As String, number As Double
    On Error GoTo invalid
    If source Is Nothing Then SetupProblem = "Select headers and at least one data row.": Exit Function
    ValidateOutputSource source
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    For Each role In spec("Roles")
        key = role("Key")
        If SetupRoleActive(spec("Id"), key, selections) Then
            value = SetupValue(selections, key)
            If SetupRoleRequired(spec, role, selections) And Len(value) = 0 Then SetupProblem = "Choose a " & role("Label") & " field.": Exit Function
            If Len(value) > 0 Then
                found = 0
                For j = 1 To UBound(headers, 2)
                    If StrComp(SafeText(headers(1, j)), value, vbTextCompare) = 0 Then found = found + 1
                Next
                If found <> 1 Then SetupProblem = "Choose a unique source header for " & role("Label") & ".": Exit Function
                If seen.Exists(value) Then SetupProblem = "Use different fields for " & seen(value) & " and " & role("Label") & ".": Exit Function
                seen.Add value, role("Label")
            End If
        End If
    Next
    If spec("Id") = "capCensus" Then
        If Len(SetupValue(selections, "ZIP")) = 0 And (Len(SetupValue(selections, "CITY")) = 0 Or Len(SetupValue(selections, "STATE")) = 0) Then
            SetupProblem = "Choose City and State, or choose ZIP.": Exit Function
        End If
    End If
    If SetupValue(selections, "MEASURE") = "Unique incidents" And Len(SetupValue(selections, "INCIDENTID")) = 0 Then SetupProblem = "Map Incident ID or choose Source records.": Exit Function
    For Each item In spec("Options")
        key = item("Key"): value = SetupValue(selections, key)
        If item("Kind") = "list" And key <> "CRIME FILTER" Then
            If InStr(1, "|" & item("Choices") & "|", "|" & value & "|", vbBinaryCompare) = 0 Or Len(value) = 0 Then SetupProblem = "Choose " & item("Label") & ".": Exit Function
        End If
        If key = "SEED" Then
            If Not IsNumeric(value) Then SetupProblem = "Enter a whole-number seed from 1 to 2147483646.": Exit Function
            number = CDbl(value)
            If number < 1 Or number > 2147483646# Or number <> Fix(number) Then SetupProblem = "Enter a whole-number seed from 1 to 2147483646.": Exit Function
        End If
        If key = "SITE LAT" Or key = "SITE LON" Then
            If Not IsNumeric(value) Then SetupProblem = "Enter " & item("Label") & " in decimal degrees.": Exit Function
            number = CDbl(value)
            If Abs(number) > IIf(key = "SITE LAT", 90, 180) Then SetupProblem = "Check the range of " & item("Label") & ".": Exit Function
        End If
    Next
    If spec("Id") = "capSingleSite" Then
        If Val(SetupValue(selections, "SITE LAT")) = 0 And Val(SetupValue(selections, "SITE LON")) = 0 Then SetupProblem = "Enter site coordinates; 0,0 is a placeholder.": Exit Function
    End If
    If selections.Exists("START") And (spec("Id") <> "capSingleSite" Or Len(SetupValue(selections, "DATE")) > 0) Then
        If Not SetupDate(SetupValue(selections, "START"), first) Or Not SetupDate(SetupValue(selections, "END"), last) Then SetupProblem = "Enter valid date-only boundaries (month/day/year or ISO).": Exit Function
        If first > last Then SetupProblem = "Start date must be on or before End date.": Exit Function
        If selections.Exists("A START") Then
            If Not SetupDate(SetupValue(selections, "A START"), first) Or Not SetupDate(SetupValue(selections, "A END"), last) Then SetupProblem = "Enter valid Period A dates.": Exit Function
            If first > last Then SetupProblem = "Period A start must be on or before its end.": Exit Function
            SetupDate SetupValue(selections, "START"), first
            If last >= first Then SetupProblem = "Period A must end before Period B begins.": Exit Function
        End If
    End If
    Exit Function
invalid:
    SetupProblem = Err.Description
End Function
Public Function SetupDate(ByVal text As String, ByRef value As Date) As Boolean
    Dim hasTime As Boolean, status As String
    If ParseTimestamp(text, False, value, hasTime, status) Then SetupDate = Not hasTime
End Function
Public Function SuggestHeader(ByVal role As Object, ByVal headers As Variant, ByVal wb As Workbook) As String
    Dim saved As String, candidate As Variant, j As Long, count As Long, matched As String
    saved = SavedField(wb, CStr(role("Key")), "[CAP NO MAPPING]", CStr(role("Key")))
    If Len(saved) = 0 And Not role("Required") Then Exit Function
    If Len(saved) > 0 Then
        For j = 1 To UBound(headers, 2)
            If StrComp(SafeText(headers(1, j)), saved, vbTextCompare) = 0 Then count = count + 1: matched = SafeText(headers(1, j))
        Next
        If count = 1 Then SuggestHeader = matched: Exit Function
    End If
    For Each candidate In Split(role("Suggested"), "|")
        count = 0
        If Len(candidate) > 0 Then
            For j = 1 To UBound(headers, 2)
                If StrComp(SafeText(headers(1, j)), candidate, vbTextCompare) = 0 Then count = count + 1: matched = SafeText(headers(1, j))
            Next
            If count = 1 Then SuggestHeader = matched: Exit Function
        End If
    Next
End Function
Public Function RunSetup(ByVal toolId As String) As Object
    Load frmCAPSetup
    If frmCAPSetup.ToolId <> toolId Then frmCAPSetup.Configure toolId
    frmCAPSetup.Show vbModal
    Set RunSetup = frmCAPSetup.Result
End Function
Public Sub OpenToolSetup(ByVal toolId As String)
    Dim inputs As Object, problem As String, role As Object, spec As Object, source As Range
    On Error GoTo failed
    Do
        Set inputs = RunSetup(toolId)
        If inputs Is Nothing Then Exit Do
        frmCAPSetup.Controls("lblStatus").Caption = "Running...": DoEvents
        problem = ExecuteSetupSafely(toolId, inputs)
        If Len(problem) = 0 Then
            Set spec = ToolSpec(toolId): Set source = inputs("Source")
            For Each role In spec("Roles")
                If SetupRoleActive(toolId, role("Key"), inputs) Then RememberField source.Parent.Parent, CStr(role("Key")), SetupValue(inputs, role("Key")), CStr(role("Key"))
            Next
            Exit Do
        End If
        frmCAPSetup.Controls("lblStatus").Caption = "Could not run: " & problem
    Loop
    Unload frmCAPSetup
    Exit Sub
failed:
    problem = Err.Description
    On Error Resume Next: Unload frmCAPSetup: Application.StatusBar = False
    MsgBox problem, vbExclamation, "CAP Index | Tool setup"
End Sub
Public Function ExecuteSetupSafely(ByVal toolId As String, ByVal inputs As Object) As String
    On Error GoTo failed
    ExecuteSetup toolId, inputs
    Exit Function
failed:
    ExecuteSetupSafely = Err.Description
End Function
Public Sub ExecuteSetup(ByVal toolId As String, ByVal s As Object)
    Dim r As Range, ctx As Object, first As Date, last As Date, aFirst As Date, aLast As Date, crime As String, measure As String
    Set r = s("Source"): measure = SetupValue(s, "MEASURE"): crime = SetupValue(s, "CRIME FILTER")
    If Len(crime) = 0 Then crime = "All crime types"
    If s.Exists("START") And Len(SetupValue(s, "DATE")) > 0 Then
        SetupDate SetupValue(s, "START"), first: SetupDate SetupValue(s, "END"), last
    End If
    If toolId = "capTrends" Or toolId = "capCompare" Or toolId = "capProfile" Then
        Set ctx = ReadReportInput(r, SetupValue(s, "DATE"), SetupValue(s, "CRIME"), SetupValue(s, "INCIDENTID"), SetupValue(s, "LOCATION"), crime, measure, CDbl(first), CDbl(last), False)
    End If
    Select Case toolId
        Case "capBlocks", "capBlockRandom"
            If toolId = "capBlockRandom" Then
                BuildHundredBlocks r, SetupValue(s, "ADDRESS"), SetupValue(s, "CITY"), SetupValue(s, "STATE"), SetupValue(s, "INCIDENTID"), True, CLng(SetupValue(s, "SEED")), False
            Else
                BuildHundredBlocks r, SetupValue(s, "ADDRESS"), showSummary:=False
            End If
        Case "capCensus"
            GeocodeCensus r, SetupValue(s, "ADDRESS"), SetupValue(s, "CITY"), SetupValue(s, "STATE"), SetupValue(s, "ZIP"), showSummary:=False
        Case "capPareto"
            BuildConcentrationReport r, SetupValue(s, "HUNDRED-BLOCK"), SetupValue(s, "CRIME"), SetupValue(s, "INCIDENTID"), SetupValue(s, "LOCATION LABEL"), False, SetupValue(s, "CITY"), SetupValue(s, "STATE")
        Case "capHeatmap"
            BuildTemporalHeatmap r, SetupValue(s, "DATE"), SetupValue(s, "CRIME"), SetupValue(s, "INCIDENTID"), SetupValue(s, "HOUR"), False, SetupValue(s, "DAY OF WEEK"), SetupValue(s, "TIME")
        Case "capRepeat"
            BuildRepeatLocations r, SetupValue(s, "ADDRESS"), "", SetupValue(s, "CRIME"), SetupValue(s, "INCIDENTID"), SetupValue(s, "LOCATION LABEL"), False, SetupValue(s, "CITY"), SetupValue(s, "STATE")
        Case "capHotStreets"
            BuildHotStreets r, SetupValue(s, "HUNDRED-BLOCK"), SetupValue(s, "DATE"), SetupValue(s, "CRIME"), SetupValue(s, "INCIDENTID"), SetupValue(s, "LOCATION LABEL"), False, SetupValue(s, "CITY"), SetupValue(s, "STATE")
        Case "capTrends"
            BuildTrendReport r, SetupValue(s, "DATE"), SetupValue(s, "CRIME"), SetupValue(s, "INCIDENTID"), crime, measure, SetupValue(s, "UNIT"), CDbl(first), CDbl(last), False, reportContext:=ctx
        Case "capCompare"
            SetupDate SetupValue(s, "A START"), aFirst: SetupDate SetupValue(s, "A END"), aLast
            BuildPeriodComparison r, SetupValue(s, "DATE"), SetupValue(s, "CRIME"), SetupValue(s, "INCIDENTID"), SetupValue(s, "LOCATION"), crime, measure, CDbl(aFirst), CDbl(aLast), CDbl(first), CDbl(last), False, reportContext:=ctx
        Case "capProfile"
            BuildCrimeProfile r, SetupValue(s, "DATE"), SetupValue(s, "CRIME"), SetupValue(s, "INCIDENTID"), SetupValue(s, "LOCATION"), crime, measure, CDbl(first), CDbl(last), False, reportContext:=ctx
        Case "capCrosswalk"
            PrepareNIBRSCrosswalk r, SetupValue(s, "CRIME"), SetupValue(s, "AGENCY"), SetupValue(s, "FEDERAL") = "Yes", False
        Case "capSingleSite"
            BuildSingleSite r, CDbl(SetupValue(s, "SITE LAT")), CDbl(SetupValue(s, "SITE LON")), SetupValue(s, "DATE"), SetupValue(s, "CRIME"), SetupValue(s, "INCIDENTID"), SetupValue(s, "LATITUDE"), SetupValue(s, "LONGITUDE"), measure, CDbl(first), CDbl(last), False
        Case "capYoY", "capSeasonality"
            BuildCalendarReport r, IIf(toolId = "capYoY", "YoY", "Seasonality"), SetupValue(s, "DATE"), SetupValue(s, "CRIME"), SetupValue(s, "INCIDENTID"), measure, CDbl(first), CDbl(last), IIf(toolId = "capYoY", SetupValue(s, "UNIT"), "Months"), False
        Case "capCheck"
            BuildDatasetCheck r, SetupValue(s, "DATE"), SetupValue(s, "CRIME"), SetupValue(s, "INCIDENTID"), SetupValue(s, "ADDRESS"), SetupValue(s, "LATITUDE"), SetupValue(s, "LONGITUDE"), False
        Case "capReview"
            ReviewCases r, SetupValue(s, "INCIDENTID"), SetupValue(s, "CRIME")
        Case Else: Err.Raise vbObjectError + 1201, , "Unknown setup tool."
    End Select
End Sub
Public Function SetupHelp(ByVal spec As Object) As String
    Dim role As Object, item As Object, text As String
    text = spec("Title") & vbCrLf & vbCrLf & spec("Description") & vbCrLf & vbCrLf & _
        "1. Select a rectangular range including headers and data. Only selected rows and fields are used." & vbCrLf & _
        "2. Choose fields from the dropdowns. An asterisk marks required fields. Use (none) for an absent optional field." & vbCrLf & _
        "3. Review options and the scope line. Run becomes available when the setup is valid." & vbCrLf & vbCrLf
    For Each role In spec("Roles"): text = text & role("Label") & ": " & role("Help") & vbCrLf: Next
    If spec("Id") = "capCensus" Then text = text & vbCrLf & "Address components and request IDs are sent to the US Census Bureau. Match information does not establish rooftop accuracy. Inferred ZIP repairs require verification."
    If spec("Id") = "capHotStreets" Then text = text & vbCrLf & "Change weeks/months, interval width, rise/fall threshold and minimum comparison events on the output sheet. Labels are descriptive, not statistical significance tests."
    If spec("Id") = "capPareto" Then text = text & vbCrLf & "The denominator contains only blocks with the selected recorded crime, excluding unknown zero-crime places."
    text = text & vbCrLf & vbCrLf & "Incident counts require stable, qualified IDs. Offense records may contain several rows per incident. Observed dates do not prove complete reporting. Source data is not saved automatically. Help and Cancel do not create sheets. A failed run may retain completed output; review it before retrying."
    SetupHelp = text
End Function
