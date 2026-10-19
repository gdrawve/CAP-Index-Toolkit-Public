Attribute VB_Name = "modDateTime"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit

Public Function ParseTimestamp(ByVal raw As Variant, ByVal date1904 As Boolean, ByRef parsed As Date, ByRef hasTime As Boolean, ByRef status As String) As Boolean
    Dim serial As Double, s As String, re As Object, m As Object
    Dim y As Long, mo As Long, dy As Long, hh As Long, nn As Long, ss As Long, usDate As Boolean, ampm As String
    status = "Invalid date/time": hasTime = False
    If IsError(raw) Then status = "Source cell error": Exit Function
    If Len(SafeText(raw)) = 0 Then status = "Missing date/time": Exit Function
    On Error GoTo invalid
    If VarType(raw) <> vbString And IsNumeric(raw) Then
        serial = CDbl(raw)
        If date1904 Then serial = serial + 1462
        ' Reject the fictitious 1900-02-29 and pre-March-1900 serial ambiguity.
        If serial < 61 Or serial >= 2958466# Then GoTo invalid
        parsed = CDate(serial)
        hasTime = (serial <> Fix(serial))
        status = IIf(hasTime, "Valid", "Date only; time missing")
        ParseTimestamp = True: Exit Function
    End If
    s = SafeText(raw)
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "^(\d{4})-(\d{2})-(\d{2})([ T](\d{1,2}):(\d{2})(:(\d{2}))?( ?(AM|PM))?)?$"
    If Not re.Test(s) Then
        re.Pattern = "^(\d{1,2})/(\d{1,2})/(\d{4})([ T](\d{1,2}):(\d{2})(:(\d{2}))?( ?(AM|PM))?)?$"
        If Not re.Test(s) Then status = "Use Excel dates, ISO dates or US month/day/year; timezone not accepted": Exit Function
        usDate = True
    End If
    Set m = re.Execute(s)(0)
    If usDate Then
        mo = CLng(m.SubMatches(0)): dy = CLng(m.SubMatches(1)): y = CLng(m.SubMatches(2))
    Else
        y = CLng(m.SubMatches(0)): mo = CLng(m.SubMatches(1)): dy = CLng(m.SubMatches(2))
    End If
    If y < 1900 Or y > 9999 Or mo < 1 Or mo > 12 Or dy < 1 Or dy > 31 Then GoTo invalid
    parsed = DateSerial(y, mo, dy)
    If Year(parsed) <> y Or Month(parsed) <> mo Or Day(parsed) <> dy Then GoTo invalid
    If parsed < DateSerial(1900, 3, 1) Then GoTo invalid
    hasTime = Len(m.SubMatches(3)) > 0
    If hasTime Then
        hh = CLng(m.SubMatches(4)): nn = CLng(m.SubMatches(5))
        If Len(m.SubMatches(7)) > 0 Then ss = CLng(m.SubMatches(7))
        ampm = UCase$(m.SubMatches(9))
        If Len(ampm) > 0 Then
            If hh < 1 Or hh > 12 Then GoTo invalid
            If hh = 12 Then hh = 0
            If ampm = "PM" Then hh = hh + 12
        End If
        If hh > 23 Or nn > 59 Or ss > 59 Then GoTo invalid
        parsed = parsed + TimeSerial(hh, nn, ss)
    End If
    status = IIf(hasTime, "Valid", "Date only; time missing")
    ParseTimestamp = True: Exit Function
invalid:
    status = "Invalid date/time"
End Function

Public Function ParseTimeOnly(ByVal raw As Variant, ByRef fraction As Double) As Boolean
    Dim s As String, re As Object, m As Object, h As Long, n As Long, sec As Long
    If IsError(raw) Or Len(SafeText(raw)) = 0 Then Exit Function
    If VarType(raw) <> vbString And IsNumeric(raw) Then
        fraction = CDbl(raw)
        ParseTimeOnly = fraction >= 0 And fraction < 1
        Exit Function
    End If
    s = SafeText(raw): Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^(\d{2}):(\d{2})(:(\d{2}))?$"
    If Not re.Test(s) Then Exit Function
    Set m = re.Execute(s)(0)
    h = CLng(m.SubMatches(0)): n = CLng(m.SubMatches(1))
    If Len(m.SubMatches(3)) > 0 Then sec = CLng(m.SubMatches(3))
    If h > 23 Or n > 59 Or sec > 59 Then Exit Function
    fraction = CDbl(TimeSerial(h, n, sec)): ParseTimeOnly = True
End Function

Public Function TimestampHeaders() As Variant
    TimestampHeaders = Array("Month Number", "Month Text", "Day Number", "Day of Week", "Year", "Date Only", "Time Only", "Hour (0-23)")
End Function

Private Function SelectedTimestampColumn() As Range
    Dim selected As Range, ws As Worksheet, lo As ListObject, index As Long, lastRow As Long
    If TypeName(Selection) <> "Range" Then Err.Raise vbObjectError + 20, , "Select a cell in the timestamp column."
    Set selected = Selection: Set ws = selected.Parent
    If selected.Areas.Count <> 1 Or selected.Columns.Count <> 1 Then Err.Raise vbObjectError + 21, , "Select one timestamp column or one cell within it."
    On Error Resume Next
    Set lo = selected.Cells(1, 1).ListObject
    On Error GoTo 0
    If Not lo Is Nothing Then
        If lo.DataBodyRange Is Nothing Then Err.Raise vbObjectError + 22, , "The table has no data rows."
        index = selected.Column - lo.Range.Column + 1
        Set SelectedTimestampColumn = lo.ListColumns(index).Range.Resize(lo.ListRows.Count + 1)
    ElseIf selected.Rows.Count > 1 And selected.Rows.Count < ws.Rows.Count Then
        Set SelectedTimestampColumn = selected
    Else
        lastRow = ws.Cells(ws.Rows.Count, selected.Column).End(xlUp).Row
        Set SelectedTimestampColumn = ws.Cells(1, selected.Column).Resize(lastRow, 1)
    End If
End Function

Private Function ExistingTimestampFieldCount(ByVal source As Range, ByRef owner As Name) As Long
    Dim named As Name, marked As Range, headers As Variant, j As Long
    headers = TimestampHeaders()
    For Each named In source.Parent.Names
        If InStr(1, named.Name, "_CAPTimestamp_", vbTextCompare) > 0 Then
            Set marked = Nothing
            On Error Resume Next
            Set marked = named.RefersToRange
            On Error GoTo 0
            If Not marked Is Nothing Then
                If marked.Row = source.Row And marked.Column = source.Column Then
                    If marked.Columns.Count <> 7 And marked.Columns.Count <> 9 Then Err.Raise vbObjectError + 23, , "Generated timestamp fields were interrupted by another column. Restore their layout before refreshing."
                    For j = 0 To marked.Columns.Count - 2
                        If SafeText(source.Cells(1, 1).Offset(0, j + 1).Value2) <> headers(j) Then Err.Raise vbObjectError + 23, , "The previously created timestamp headers were changed. Restore them before refreshing this tool."
                    Next
                    ExistingTimestampFieldCount = marked.Columns.Count - 1: Set owner = named: Exit Function
                End If
            End If
        End If
    Next
End Function

Public Sub BuildDateTime(Optional ByVal source As Range, Optional ByVal showSummary As Boolean = True)
    On Error GoTo failed
    Dim r As Range, a As Variant, ws As Worksheet, wb As Workbook, out() As Variant, target As Range
    Dim i As Long, j As Long, d As Date, ht As Boolean, status As String, ok As Boolean, offset As Double
    Dim monthNames As Variant, dayNames As Variant, headers As Variant, bad() As Boolean
    Dim timedCount As Long, fieldCount As Long, previousCount As Long, compact() As Variant
    Dim validCount As Long, invalidCount As Long, blankCount As Long, inserted As Long, reuse As Boolean
    Dim lo As ListObject, tableIndex As Long, sourceCol As Long, totalRows As Long, marked As Name
    Dim oldScreen As Boolean, oldEvents As Boolean, changing As Boolean, errNo As Long, errText As String
    Set r = source
    If r Is Nothing Then Set r = SelectedTimestampColumn()
    If r.Areas.Count <> 1 Or r.Columns.Count <> 1 Or r.Rows.Count < 2 Or r.Rows.Count > 200001 Then Err.Raise vbObjectError + 24, , "Select one column with its header plus 1 to 200,000 data rows."
    Set ws = r.Parent: Set wb = ws.Parent
    If wb Is ThisWorkbook Then Err.Raise vbObjectError + 25, , "Select a data workbook, not the add-in."
    If wb.ReadOnly Or ws.ProtectContents Then Err.Raise vbObjectError + 26, , "The worksheet must be unprotected and its workbook writable."
    If IsNull(r.MergeCells) Then Err.Raise vbObjectError + 27, , "The source column contains merged cells."
    If r.MergeCells Then Err.Raise vbObjectError + 27, , "The source column contains merged cells."
    If Len(SafeText(r.Cells(1, 1).Value2)) = 0 Then Err.Raise vbObjectError + 28, , "The selected range must start with a column header."
    sourceCol = r.Column: totalRows = r.Rows.Count - 1
    On Error Resume Next
    Set lo = r.Cells(1, 1).ListObject
    On Error GoTo failed
    If Not lo Is Nothing Then
        tableIndex = sourceCol - lo.Range.Column + 1
        If r.Row <> lo.HeaderRowRange.Row Or totalRows <> lo.ListRows.Count Then Err.Raise vbObjectError + 30, , "Select the entire table column including its header."
    End If
    previousCount = ExistingTimestampFieldCount(r, marked): reuse = previousCount > 0
    headers = TimestampHeaders()
    a = r.Value2
    ReDim out(1 To totalRows, 1 To 8): ReDim bad(1 To totalRows)
    monthNames = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    dayNames = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
    If wb.Date1904 Then offset = 1462
    For i = 2 To UBound(a, 1)
        ok = ParseTimestamp(a(i, 1), wb.Date1904, d, ht, status)
        If ok Then
            validCount = validCount + 1
            out(i - 1, 1) = Month(d): out(i - 1, 2) = monthNames(Month(d) - 1)
            out(i - 1, 3) = Day(d): out(i - 1, 4) = dayNames(Weekday(d) - 1)
            out(i - 1, 5) = Year(d): out(i - 1, 6) = CDbl(DateValue(d)) - offset
            If ht Then timedCount = timedCount + 1: out(i - 1, 7) = CDbl(d) - Fix(CDbl(d)): out(i - 1, 8) = Hour(d)
        ElseIf Len(SafeText(a(i, 1))) = 0 Then
            blankCount = blankCount + 1
        Else
            invalidCount = invalidCount + 1: bad(i - 1) = True
        End If
    Next
    If validCount = 0 Then Err.Raise vbObjectError + 32, , "No valid dates were found. No columns were inserted. Text dates use month/day/year or ISO format."
    fieldCount = IIf(timedCount > 0, 8, 6)
    If sourceCol > ws.Columns.Count - fieldCount Then Err.Raise vbObjectError + 29, , "No room for generated timestamp columns."
    ReDim compact(1 To totalRows, 1 To fieldCount)
    For i = 1 To totalRows: For j = 1 To fieldCount: compact(i, j) = out(i, j): Next: Next
    ReDim Preserve headers(0 To fieldCount - 1)
    If Not lo Is Nothing Then
        For j = previousCount To fieldCount - 1
            For i = 1 To lo.ListColumns.Count
                If StrComp(lo.ListColumns(i).Name, headers(j), vbTextCompare) = 0 Then Err.Raise vbObjectError + 31, , "Table already contains '" & headers(j) & "'. Rename that field before expanding Timestamp Fields."
            Next
        Next
    End If
    If previousCount > fieldCount And lo Is Nothing Then
        If r.Row > 1 Then
            If Application.CountA(ws.Range(ws.Cells(1, sourceCol + 7), ws.Cells(r.Row - 1, sourceCol + 8))) > 0 Then Err.Raise vbObjectError + 33, , "Time columns contain data above the generated range; no columns removed."
        End If
        If Application.CountA(ws.Range(ws.Cells(r.Row + totalRows + 1, sourceCol + 7), ws.Cells(ws.Rows.Count, sourceCol + 8))) > 0 Then Err.Raise vbObjectError + 33, , "Time columns contain data below the generated range; no columns removed."
    End If
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    changing = True: Application.ScreenUpdating = False: Application.EnableEvents = False
    If fieldCount > previousCount Then
        If lo Is Nothing Then
            ws.Columns(sourceCol + previousCount + 1).Resize(, fieldCount - previousCount).Insert Shift:=xlToRight
            inserted = fieldCount - previousCount
        Else
            For j = previousCount + 1 To fieldCount
                lo.ListColumns.Add Position:=tableIndex + j
                inserted = inserted + 1
            Next
        End If
    ElseIf fieldCount < previousCount Then
        If lo Is Nothing Then
            ws.Columns(sourceCol + fieldCount + 1).Resize(, previousCount - fieldCount).Delete
        Else
            For j = 1 To previousCount - fieldCount: lo.ListColumns(tableIndex + fieldCount + 1).Delete: Next
        End If
    End If
    Set target = ws.Cells(r.Row, sourceCol + 1).Resize(totalRows + 1, fieldCount)
    WriteHeaders ws, r.Row, sourceCol + 1, headers
    With target.Offset(1, 0).Resize(totalRows, fieldCount)
        .NumberFormat = "General": .Value2 = compact: .Interior.Pattern = xlNone
    End With
    target.Offset(1, 0).Resize(totalRows, 1).NumberFormat = "0"
    target.Offset(1, 2).Resize(totalRows, 1).NumberFormat = "0"
    target.Offset(1, 4).Resize(totalRows, 1).NumberFormat = "0"
    target.Offset(1, 5).Resize(totalRows, 1).NumberFormat = "mm/dd/yyyy"
    If fieldCount = 8 Then
        target.Offset(1, 6).Resize(totalRows, 1).NumberFormat = "hh:mm:ss"
        target.Offset(1, 7).Resize(totalRows, 1).NumberFormat = "0"
    End If
    With target.Rows(1)
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(26, 55, 82)
        .WrapText = True
    End With
    For i = 1 To totalRows
        If bad(i) Then target.Rows(i + 1).Interior.Color = RGB(255, 235, 238)
    Next
    target.Columns.AutoFit: target.Rows(1).EntireRow.AutoFit
    If Not reuse Then
        Set marked = ws.Names.Add(Name:="_CAPTimestamp_" & sourceCol & "_" & r.Row, RefersTo:="='" & Replace(ws.Name, "'", "''") & "'!" & ws.Cells(r.Row, sourceCol).Resize(1, fieldCount + 1).Address)
        marked.Visible = False
    Else
        marked.RefersTo = "='" & Replace(ws.Name, "'", "''") & "'!" & ws.Cells(r.Row, sourceCol).Resize(1, fieldCount + 1).Address
    End If
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents: changing = False
    If showSummary Then MsgBox "Timestamp fields " & IIf(reuse, "refreshed", "inserted") & "." & vbCrLf & "Valid dates: " & validCount & vbCrLf & "Blank source rows: " & blankCount & vbCrLf & "Invalid source rows: " & invalidCount & " (output cells highlighted; left blank)." & vbCrLf & "Generated fields: " & fieldCount & ". Text dates use month/day/year. " & IIf(fieldCount = 6, "No valid times: Time Only and Hour omitted.", "Rows with no time leave Time Only and Hour blank."), vbInformation, "CAP Index | Timestamp Fields"
    Exit Sub
failed:
    errNo = Err.Number: errText = Err.Description
    On Error Resume Next
    If inserted > 0 Then
        If lo Is Nothing Then
            ws.Columns(sourceCol + previousCount + 1).Resize(, inserted).Delete
        Else
            For j = 1 To inserted: lo.ListColumns(tableIndex + previousCount + 1).Delete: Next
        End If
    End If
    If changing Then Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents
    On Error GoTo 0
    If Not source Is Nothing Then Err.Raise errNo, "Timestamp Fields", errText
    MsgBox "Timestamp Fields: " & errText, vbExclamation, "CAP Index"
End Sub
