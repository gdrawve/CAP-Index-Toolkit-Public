Attribute VB_Name = "modHour"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit

Public Function HourFromTime(ByVal raw As Variant, ByVal date1904 As Boolean, ByRef hourValue As Long) As Boolean
    Dim value As Double, parsed As Date, hasTime As Boolean, status As String, re As Object, matches As Object
    Dim h As Long, m As Long, sec As Long, ampm As String
    If IsError(raw) Or Len(SafeText(raw)) = 0 Then Exit Function
    If VarType(raw) <> vbString And IsNumeric(raw) Then
        value = CDbl(raw)
        If value >= 0 And value < 1 Then
            hourValue = Hour(CDate(value)): HourFromTime = True: Exit Function
        End If
    End If
    If ParseTimestamp(raw, date1904, parsed, hasTime, status) Then
        If hasTime Then hourValue = Hour(parsed): HourFromTime = True
        Exit Function
    End If
    Set re = CreateObject("VBScript.RegExp"): re.IgnoreCase = True
    re.Pattern = "^(\d{1,2}):(\d{2})(:(\d{2}))?\s*(AM|PM)?$"
    If Not re.Test(SafeText(raw)) Then Exit Function
    Set matches = re.Execute(SafeText(raw))(0)
    h = CLng(matches.SubMatches(0)): m = CLng(matches.SubMatches(1))
    If Len(matches.SubMatches(3)) > 0 Then sec = CLng(matches.SubMatches(3))
    ampm = UCase$(matches.SubMatches(4))
    If Len(ampm) > 0 Then
        If h < 1 Or h > 12 Then Exit Function
        If h = 12 Then h = 0
        If ampm = "PM" Then h = h + 12
    End If
    If h > 23 Or m > 59 Or sec > 59 Then Exit Function
    hourValue = h: HourFromTime = True
End Function

Public Function DayOfWeekNumber(ByVal raw As Variant) As Long
    Dim names As Variant, i As Long, s As String
    If IsError(raw) Then Exit Function
    names = Array("MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY")
    s = NormalText(raw)
    For i = 0 To 6
        If s = names(i) Or s = Left$(names(i), 3) Then DayOfWeekNumber = i + 1: Exit Function
    Next
End Function

Public Sub BuildHourFields(Optional ByVal source As Range, Optional ByVal showSummary As Boolean = True)
    Dim r As Range, ws As Worksheet, lo As ListObject, marked As Name, previous As Range, target As Range
    Dim a As Variant, output() As Variant, i As Long, h As Long, bad As Long, inserted As Boolean, pos As Long
    Dim targetHeader As String, duplicate As Range
    Dim oldEvents As Boolean, changing As Boolean, errNo As Long, description As String
    On Error GoTo failed
    Set r = source
    If r Is Nothing Then
        On Error Resume Next
        Set r = Application.InputBox("Select ONE Time column including its header and data. Accepts Excel times, 24-hour times and AM/PM times. Inserts only Hour (0-23) beside it. Missing/invalid values remain blank.", "CAP Index | Hour", Type:=8)
        On Error GoTo failed
    End If
    If r Is Nothing Then Exit Sub
    If r.Columns.Count <> 1 Then Err.Raise vbObjectError + 720, , "Select one Time column including the header."
    ValidateOutputSource r: Set ws = r.Parent
    If ws.ProtectContents Then Err.Raise vbObjectError + 721, , "Unprotect the worksheet to insert Hour beside Time."
    If IsNull(r.MergeCells) Then Err.Raise vbObjectError + 722, , "Unmerge the selected Time cells."
    If r.MergeCells Then Err.Raise vbObjectError + 722, , "Unmerge the selected Time cells."
    targetHeader = "Hour (0-23)"
    a = r.Value2: ReDim output(1 To UBound(a, 1) - 1, 1 To 1)
    For i = 2 To UBound(a, 1)
        If HourFromTime(a(i, 1), ws.Parent.Date1904, h) Then output(i - 1, 1) = h Else bad = bad + 1
    Next
    For Each marked In ws.Names
        If InStr(marked.Name, "_CAPHour_") > 0 Then
            Set previous = marked.RefersToRange
            If previous.Row = r.Row And previous.Column = r.Column Then
                targetHeader = SafeText(r.Cells(1).Offset(0, 1).Value2)
                If targetHeader <> "Hour (0-23)" And targetHeader <> "Hour from Time (0-23)" Then Err.Raise vbObjectError + 723, , "Restore the generated Hour header before refreshing."
                Set target = r.Cells(1).Offset(0, 1): Exit For
            End If
        End If
    Next
    If target Is Nothing Then
        Set duplicate = ws.Rows(r.Row).Find(What:="Hour (0-23)", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
        If Not duplicate Is Nothing Then targetHeader = "Hour from Time (0-23)"
        Set duplicate = ws.Rows(r.Row).Find(What:=targetHeader, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
        If Not duplicate Is Nothing Then Err.Raise vbObjectError + 725, , "The sheet already has " & targetHeader & ". Select the Time column previously used by this tool to refresh its marked Hour output."
    End If
    On Error Resume Next: Set lo = r.Cells(1).ListObject: On Error GoTo failed
    If Not lo Is Nothing Then
        pos = r.Column - lo.Range.Column + 1
        If r.Row <> lo.HeaderRowRange.Row Or r.Rows.Count <> lo.ListRows.Count + 1 Then Err.Raise vbObjectError + 724, , "Select the complete Time table column including its header."
        If target Is Nothing Then
            For i = 1 To lo.ListColumns.Count
                If StrComp(lo.ListColumns(i).Name, targetHeader, vbTextCompare) = 0 Then Err.Raise vbObjectError + 725, , "The table already has Hour (0-23). Rename it before creating another Hour field."
            Next
        End If
    End If
    oldEvents = Application.EnableEvents: changing = True: Application.EnableEvents = False
    If target Is Nothing Then
        If lo Is Nothing Then ws.Columns(r.Column + 1).Insert Shift:=xlToRight Else lo.ListColumns.Add Position:=pos + 1
        inserted = True: Set target = ws.Cells(r.Row, r.Column + 1)
        Set marked = ws.Names.Add(Name:="_CAPHour_" & r.Column & "_" & r.Row, RefersTo:="='" & Replace(ws.Name, "'", "''") & "'!" & r.Cells(1).Resize(1, 2).Address): marked.Visible = False
    End If
    target.Value2 = targetHeader: target.Offset(1).Resize(UBound(output, 1)).NumberFormat = "0"
    target.Offset(1).Resize(UBound(output, 1)).Value2 = output
    target.Font.Bold = True: target.Font.Color = vbWhite: target.Interior.Color = RGB(26, 55, 82): target.EntireColumn.AutoFit
    Application.EnableEvents = oldEvents: changing = False
    If showSummary Then MsgBox "Hour created/refreshed. Blank or invalid time rows: " & bad & ". Only numeric Excel time zero or explicit text midnight produces Hour 0.", vbInformation, "CAP Index"
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    On Error Resume Next
    If inserted Then
        If lo Is Nothing Then ws.Columns(r.Column + 1).Delete Else lo.ListColumns(pos + 1).Delete
        If Not marked Is Nothing Then marked.Delete
    End If
    If changing Then Application.EnableEvents = oldEvents
    On Error GoTo 0
    If Not source Is Nothing Then Err.Raise errNo, "Hour", description
    MsgBox description, vbExclamation, "CAP Index"
End Sub
