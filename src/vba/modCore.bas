Attribute VB_Name = "modCore"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit
Public Const CAP_VERSION As String = "1.2.3"

Public Function PickSource(Optional ByVal requirements As String = "") As Range
    Dim r As Range
    On Error Resume Next
    Set r = Application.InputBox(IIf(Len(requirements) > 0, Left$(requirements, 190) & vbCrLf & "Select headers + data; Analysis Scope applies to analyses.", "Select one rectangular range including column headers. Analyses use Analysis Scope; preparation tools process all selected rows."), "CAP Index source", Type:=8)
    On Error GoTo 0
    If r Is Nothing Then Exit Function
    If r.Areas.Count <> 1 Or r.Rows.Count < 2 Or r.Columns.Count > 200 Or r.Rows.Count > 200001 Then
        Err.Raise vbObjectError + 1, , "Select headers plus 1 to 200,000 data rows and at most 200 columns."
    End If
    If r.Parent.Parent Is ThisWorkbook Then Err.Raise vbObjectError + 2, , "Select a data workbook, not the add-in."
    If r.Parent.Parent.ReadOnly Or r.Parent.Parent.ProtectStructure Then Err.Raise vbObjectError + 3, , "The data workbook must be writable and its structure unprotected."
    If MsgBox("Source: " & r.Parent.Parent.Name & "!" & r.Parent.Name & vbCrLf & _
        "Selected data rows: " & r.Rows.Count - 1 & "; fields: " & r.Columns.Count & vbCrLf & _
        "Confirm headers. Analyses use the workbook Analysis Scope; preparation tools include all selected rows.", vbOKCancel + vbInformation, "CAP Index | Confirm selection") = vbCancel Then Exit Function
    Set PickSource = r
End Function

Public Function HeaderColumn(ByVal data As Variant, ByVal header As String) As Long
    Dim j As Long, found As Long
    If Len(header) = 0 Then Exit Function
    For j = 1 To UBound(data, 2)
        If StrComp(SafeText(data(1, j)), header, vbTextCompare) = 0 Then
            If found <> 0 Then Err.Raise vbObjectError + 4, , "Duplicate source header: " & header
            found = j
        End If
    Next
    If found = 0 Then Err.Raise vbObjectError + 5, , "Missing source header: " & header
    HeaderColumn = found
End Function

Public Function AskColumn(ByVal data As Variant, ByVal caption As String, ByVal suggested As String, Optional ByVal optionalField As Boolean = False, Optional ByVal roleKey As String = "") As Long
    Dim response As Variant, wb As Workbook, saved As String
    Set wb = ActiveWorkbook
    If Not wb Is Nothing Then
        saved = SavedField(wb, caption, suggested, roleKey)
        If Len(saved) = 0 Or OptionalHeader(data, saved) > 0 Then suggested = saved
    End If
    response = Application.InputBox(caption & " Enter the exact source column header." & IIf(optionalField, " Leave blank if absent.", ""), "CAP Index field", suggested, Type:=2)
    If VarType(response) = vbBoolean Then Err.Raise vbObjectError + 6, , "Operation cancelled."
    If Len(Trim$(CStr(response))) = 0 And Not optionalField Then Err.Raise vbObjectError + 7, , "A source field is required."
    AskColumn = HeaderColumn(data, Trim$(CStr(response)))
    RememberField wb, caption, Trim$(CStr(response)), roleKey
End Function

Public Function SafeText(ByVal v As Variant) As String
    If IsError(v) Then
        SafeText = "[CELL ERROR]"
    ElseIf IsNull(v) Or IsEmpty(v) Then
        SafeText = ""
    Else
        SafeText = Trim$(CStr(v))
    End If
End Function

Public Function NormalText(ByVal v As Variant) As String
    Dim s As String
    s = UCase$(SafeText(v))
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    NormalText = Trim$(s)
End Function

Public Function KeyPart(ByVal s As String) As String
    KeyPart = CStr(Len(s)) & ":" & s
End Function

Public Function NewOutput(ByVal wb As Workbook, ByVal prefix As String, ByVal source As Range, ByVal notes As String) As Worksheet
    Dim ws As Worksheet, candidate As String, i As Long, exists As Boolean, other As Worksheet
    For i = 1 To 9999
        candidate = Left$(prefix, 24) & "_" & Format$(i, "0000")
        exists = False
        For Each other In wb.Worksheets
            If other.Name = candidate Then exists = True: Exit For
        Next
        If Not exists Then Exit For
    Next
    If exists Then Err.Raise vbObjectError + 8, , "No available output sheet name."
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = candidate
    ws.Cells(1, 1).Value2 = "CAP Index | " & prefix & " | " & CAP_VERSION
    ws.Cells(2, 1).Value2 = "Source: " & source.Parent.Name & "!" & source.Address(False, False) & " | Run: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & CAP_ATTRIBUTION
    ws.Cells(3, 1).Value2 = notes
    Set NewOutput = ws
End Function

Public Sub WriteHeaders(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal firstCol As Long, ByVal headers As Variant)
    Dim j As Long
    For j = LBound(headers) To UBound(headers)
        ws.Cells(rowNo, firstCol + j - LBound(headers)).Value2 = headers(j)
    Next
End Sub

Public Sub FinishOutput(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    With ws.Range(ws.Cells(6, 1), ws.Cells(lastRow, lastCol))
        .AutoFilter
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Columns.AutoFit
    End With
    With ws.Range(ws.Cells(6, 1), ws.Cells(6, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(26, 55, 82)
        .Font.Color = RGB(255, 255, 255)
    End With
    Dim j As Long
    For j = 1 To lastCol
        If ws.Columns(j).ColumnWidth > 40 Then
            ws.Columns(j).ColumnWidth = 40
            ws.Range(ws.Cells(6, j), ws.Cells(lastRow, j)).WrapText = True
        End If
    Next
    ws.Range(ws.Cells(6, 1), ws.Cells(lastRow, lastCol)).Rows.AutoFit
    ws.Rows("1:3").Font.Bold = True
    ws.Activate
    ws.Range("A7").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
End Sub

Public Sub CopySourceValues(ByVal r As Range, ByVal ws As Worksheet)
    ' Text-first prevents source strings beginning with =,+,-,@ becoming formulas.
    With ws.Cells(6, 1).Resize(r.Rows.Count, r.Columns.Count)
        .NumberFormat = "@"
        .Value2 = r.Value2
    End With
End Sub

Public Sub ToolkitError(ByVal tool As String)
    MsgBox tool & ": " & Err.Description, vbExclamation, "CAP Index"
End Sub
