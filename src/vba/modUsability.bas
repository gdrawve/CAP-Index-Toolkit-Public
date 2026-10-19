Attribute VB_Name = "modUsability"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
Option Explicit

Public Function CAPNamedRange(ByVal ws As Worksheet, ByVal label As String) As Range
    Dim named As Name
    On Error Resume Next
    For Each named In ws.Names
        If Right$(named.Name, Len(label)) = label Then Set CAPNamedRange = named.RefersToRange: Exit For
    Next
    On Error GoTo 0
End Function

Public Function SavedField(ByVal wb As Workbook, ByVal prompt As String, ByVal fallback As String, Optional ByVal roleKey As String = "") As String
    Dim role As String, value As String, found As Boolean
    role = ExplicitFieldRole(prompt, roleKey): SavedField = fallback
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    Err.Clear: value = wb.CustomDocumentProperties("CAP Role " & role).Value: found = (Err.Number = 0)
    Err.Clear
    If Not found Then
        Select Case role
            Case "STATE", "CRIME", "INCIDENTID", "DATE", "TIME", "HOUR", "DAY OF WEEK", "LATITUDE", "LONGITUDE", "LOCATION LABEL"
                value = wb.CustomDocumentProperties("CAP Field " & role).Value: found = (Err.Number = 0)
        End Select
    End If
    On Error GoTo 0
    If found Then SavedField = value
End Function
Public Sub RememberField(ByVal wb As Workbook, ByVal prompt As String, ByVal header As String, Optional ByVal roleKey As String = "")
    Dim key As String
    If wb Is Nothing Then Exit Sub
    If wb Is ThisWorkbook Or wb.ReadOnly Then Exit Sub
    key = "CAP Role " & ExplicitFieldRole(prompt, roleKey)
    On Error Resume Next
    wb.CustomDocumentProperties(key).Delete
    wb.CustomDocumentProperties.Add key, False, 4, header
    On Error GoTo 0
End Sub
Private Function ExplicitFieldRole(ByVal prompt As String, ByVal roleKey As String) As String
    If Len(roleKey) > 0 Then ExplicitFieldRole = UCase$(roleKey): Exit Function
    Select Case UCase$(prompt)
        Case "ADDRESS", "CITY", "STATE", "ZIP", "HUNDRED-BLOCK", "LOCATION", "LOCATION LABEL", "CRIME", "INCIDENTID", "DATE", "TIME", "HOUR", "DAY OF WEEK", "AGENCY", "LATITUDE", "LONGITUDE"
            ExplicitFieldRole = UCase$(prompt)
        Case Else: ExplicitFieldRole = FieldRole(prompt)
    End Select
End Function

Private Function FieldRole(ByVal prompt As String) As String
    Dim token As Variant
    If InStr(1, prompt, "readable", vbTextCompare) > 0 Or InStr(1, prompt, "label", vbTextCompare) > 0 Then FieldRole = "LOCATION LABEL": Exit Function
    For Each token In Array("Latitude", "Longitude", "IncidentID", "Hour", "Day of Week", "Hundred-block", "ADDRESS", "City", "State", "Crime", "DATE", "timestamp", "TIME")
        If InStr(1, prompt, token, vbTextCompare) > 0 Then
            FieldRole = UCase$(token): If FieldRole = "TIMESTAMP" Then FieldRole = "DATE"
            Exit Function
        End If
    Next
    FieldRole = Left$(prompt, 180)
End Function

Public Sub CommonCrimeControl(ByVal ws As Worksheet)
    Dim types As Range, oldEvents As Boolean, errNo As Long, description As String, typeName As String
    oldEvents = Application.EnableEvents
    On Error GoTo failed
    Application.EnableEvents = False
    If Not TimeReportCache(ws) Is Nothing Then
        Set types = CAPNamedRange(ws, "_CAPTimeTypes")
        typeName = "_CAPTimeTypes"
    Else
        Set types = CAPNamedRange(ws, REPORT_TYPES_NAME)
        typeName = REPORT_TYPES_NAME
    End If
    If types Is Nothing Then GoTo finished
    ws.Range("A3:D3").UnMerge: ws.Range("A3").Value2 = "Crime type / blue field"
    ws.Range("B3:D3").Merge: ws.Range("B3").Value2 = ws.Range("B4").Value2
    ws.Range("B3:D3").Interior.Color = RGB(224, 239, 250): ws.Range("A3:D3").WrapText = True: ws.Rows(3).RowHeight = 36
    With ws.Range("B3").Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & typeName
        .ShowError = True: .InCellDropdown = True
    End With
    ws.Range("B4").Validation.Delete: ws.Range("A4").Value2 = "Selected crime"
    ws.Range("B4").Interior.Color = RGB(242, 246, 250)
    ws.Names.Add Name:="_CAPCommonCrime", RefersTo:="='" & Replace(ws.Name, "'", "''") & "'!$B$3"
finished:
    Application.EnableEvents = oldEvents: Exit Sub
failed:
    errNo = Err.Number: description = Err.Description: Application.EnableEvents = oldEvents
    Err.Raise errNo, "Crime filter", description
End Sub

Public Sub StampReportStatus(ByVal ws As Worksheet)
    Dim named As Name, cell As Range, re As Object, version As String, note As String
    Set cell = CAPNamedRange(ws, "_CAPRunStatus")
    If cell Is Nothing Then
        If Not TimeReportCache(ws) Is Nothing Or Not ConcentrationCache(ws) Is Nothing Then
            Set cell = ws.Range("E3:H3")
        Else
            Set cell = ws.Range("J2:Q3")
        End If
        cell.Merge
        Set named = ws.Names.Add(Name:="_CAPRunStatus", RefersTo:="='" & Replace(ws.Name, "'", "''") & "'!" & cell.Address): named.Visible = False
    End If
    Set re = CreateObject("VBScript.RegExp"): re.Pattern = "\b\d+\.\d+\.\d+\b"
    version = "Unrecorded"
    If re.Test(SafeText(ws.Range("A1").Value2)) Then
        version = re.Execute(SafeText(ws.Range("A1").Value2))(0).Value
    ElseIf re.Test(SafeText(ws.Range("A2").Value2)) Then
        re.Global = True: version = re.Execute(SafeText(ws.Range("A2").Value2))(re.Execute(SafeText(ws.Range("A2").Value2)).Count - 1).Value
    End If
    If version <> CAP_VERSION Then note = "Older report: rebuild for current layout. "
    If cell.Row = 3 Then ws.Rows(3).RowHeight = Application.Max(54, ws.Rows(3).RowHeight)
    cell.Value2 = "Report v" & version & " | Tool v" & CAP_VERSION & vbCrLf & "Refreshed: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & note & "Saved data; rebuild for source edits."
    cell.WrapText = True: cell.Font.Name = "Calibri": cell.Font.Size = 10
    cell.Interior.Color = RGB(242, 246, 250)
End Sub

Public Sub UnifiedRefresh(Optional ByVal report As Worksheet, Optional ByVal showSummary As Boolean = True)
    Dim status As String
    On Error GoTo failed
    If report Is Nothing Then
        If TypeName(ActiveSheet) <> "Worksheet" Then Err.Raise vbObjectError + 950, , "Select a CAP report worksheet."
        Set report = ActiveSheet
    End If
    Application.StatusBar = "CAP Index: refreshing " & report.Name
    If Not CAPNamedRange(report, "_CAPCommonCrime") Is Nothing Then
        Dim oldEvents As Boolean
        oldEvents = Application.EnableEvents: Application.EnableEvents = False
        report.Range("B4").Value2 = report.Range("B3").Value2
        Application.EnableEvents = oldEvents
    End If
    If Not SingleSiteCache(report) Is Nothing Then
        RefreshSingleSite report, True
    ElseIf Not ReportingCache(report) Is Nothing Then
        RefreshReportingFilter report, True
    ElseIf Not TimeReportCache(report) Is Nothing Then
        RefreshTimeReport report, True
    ElseIf Not ConcentrationCache(report) Is Nothing Then
        RefreshConcentration report, True
    Else
        Err.Raise vbObjectError + 951, , "This sheet has no refreshable saved analysis. Rebuild Dataset Check, geography outputs or older reports using their source data and original tool."
    End If
    StampReportStatus report
    Application.StatusBar = False
    If showSummary Then MsgBox "Updated " & report.Name & " using saved data and current filters. Source edits require rebuilding the report.", vbInformation, "CAP Index | Refresh Report"
    Exit Sub
failed:
    status = Err.Description: Application.StatusBar = False
    If Not showSummary Then Err.Raise vbObjectError + 952, "Refresh Report", status
    MsgBox status, vbExclamation, "CAP Index | Refresh Report"
End Sub

Public Sub ResetReportFilters(Optional ByVal report As Worksheet, Optional ByVal showSummary As Boolean = True)
    Dim oldEvents As Boolean, description As String
    oldEvents = Application.EnableEvents
    On Error GoTo failed
    If report Is Nothing Then
        If TypeName(ActiveSheet) <> "Worksheet" Then Err.Raise vbObjectError + 953, , "Select a CAP report worksheet."
        Set report = ActiveSheet
    End If
    Application.EnableEvents = False
    If Not SingleSiteCache(report) Is Nothing Or Not ReportingCache(report) Is Nothing Then
        report.Range("B3").Value2 = "All crime types"
    ElseIf Not TimeReportCache(report) Is Nothing Or Not ConcentrationCache(report) Is Nothing Then
        report.Range("B4").Value2 = "All crime types"
        If Not CAPNamedRange(report, "_CAPCommonCrime") Is Nothing Then report.Range("B3").Value2 = "All crime types"
    Else
        Err.Raise vbObjectError + 953, , "This sheet has no supported crime filter. Rebuild an older report with its original tool."
    End If
    UnifiedRefresh report, False
    Application.EnableEvents = oldEvents
    If showSummary Then MsgBox "Crime filter reset to All crime types. Date ranges, count measure and other settings are retained.", vbInformation, "CAP Index"
    Exit Sub
failed:
    description = Err.Description: Application.EnableEvents = oldEvents
    If Not showSummary Then Err.Raise vbObjectError + 953, "Reset Filters", description
    MsgBox description, vbExclamation, "CAP Index | Reset Filters"
End Sub

Public Sub CompleteCAPCommand(ByVal workbookBefore As Workbook, ByVal countBefore As Long)
    Dim ws As Worksheet
    If workbookBefore Is Nothing Then Exit Sub
    If ActiveWorkbook Is Nothing Then Exit Sub
    If Not ActiveWorkbook Is workbookBefore Then Exit Sub
    If workbookBefore.Worksheets.Count <= countBefore Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet
    MsgBox "Output: " & ws.Name & vbCrLf & "Review exclusions, quality flags and coverage notes in the output before interpreting or sharing it.", vbInformation, "CAP Index | Completed"
End Sub
