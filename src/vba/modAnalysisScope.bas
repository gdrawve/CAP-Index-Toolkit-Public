Attribute VB_Name = "modAnalysisScope"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
Option Explicit
Public CAPRibbon As Office.IRibbonUI
Private lastKey As String
Private sourceRows() As Long
Private lastIncluded As Long
Private lastExcluded As Long
Private lastScope As String

Public Function AnalysisScope(ByVal wb As Workbook) As String
    AnalysisScope = "All rows"
    On Error Resume Next
    If wb.CustomDocumentProperties("CAP Analysis Scope").Value = "Visible rows only" Then AnalysisScope = "Visible rows only"
    On Error GoTo 0
End Function

Public Sub SetAnalysisScope(ByVal wb As Workbook, ByVal scope As String, Optional ByVal showSummary As Boolean = True)
    If scope <> "All rows" And scope <> "Visible rows only" Then Err.Raise vbObjectError + 980, , "Invalid analysis scope."
    If wb Is Nothing Then Err.Raise vbObjectError + 980, , "Open a writable data workbook."
    If wb Is ThisWorkbook Or wb.ReadOnly Then Err.Raise vbObjectError + 980, , "Select a writable data workbook."
    On Error Resume Next: wb.CustomDocumentProperties("CAP Analysis Scope").Delete: On Error GoTo 0
    wb.CustomDocumentProperties.Add "CAP Analysis Scope", False, 4, scope
    If Not CAPRibbon Is Nothing Then CAPRibbon.Invalidate
    If showSummary Then MsgBox "Analysis scope: " & scope & "." & vbCrLf & "New analyses, Dataset Check, Case Review and NIBRS use this scope. Rebuild existing reports to change their saved scope. Geography and Date/Time preparation continue to process all selected rows.", vbInformation, "CAP Index | Analysis Scope"
End Sub

Public Function SavedAnalysisScope(ByVal ws As Worksheet) As String
    Dim property As Object
    SavedAnalysisScope = "Scope: not recorded; rebuild to record scope"
    On Error Resume Next
    For Each property In ws.CustomProperties
        If property.Name = "CAP Scope" Then SavedAnalysisScope = "Scope: " & property.Value: Exit For
    Next
    On Error GoTo 0
End Function

Public Sub CAP100_ScopePressed(ByVal control As Office.IRibbonControl, ByRef returnedValue)
    returnedValue = False
    If ActiveWorkbook Is Nothing Then Exit Sub
    If control.Id = "capScopeAll" Then returnedValue = AnalysisScope(ActiveWorkbook) = "All rows"
    If control.Id = "capScopeVisible" Then returnedValue = AnalysisScope(ActiveWorkbook) = "Visible rows only"
End Sub
Public Sub CAP100_ScopeClick(ByVal control As Office.IRibbonControl, ByVal pressed As Boolean)
    On Error GoTo failed
    SetAnalysisScope ActiveWorkbook, IIf(control.Id = "capScopeVisible", "Visible rows only", "All rows")
    Exit Sub
failed: MsgBox Err.Description, vbExclamation, "CAP Index | Analysis Scope"
End Sub

Private Function ScopeKey(ByVal r As Range) As String
    ScopeKey = r.Parent.Parent.FullName & "|" & r.Parent.CodeName & "|" & r.Address
End Function

Public Function ScopedValues(ByVal r As Range) As Variant
    Dim original As Variant, result() As Variant, i As Long, j As Long, count As Long, rows() As Long
    original = r.Value2: ReDim rows(1 To UBound(original, 1)): count = 1: rows(1) = 1
    lastScope = AnalysisScope(r.Parent.Parent)
    If lastScope = "All rows" Then
        count = UBound(original, 1)
        For i = 2 To count: rows(i) = i: Next
    Else
        For i = 2 To UBound(original, 1)
            If Not r.Rows(i).EntireRow.Hidden Then count = count + 1: rows(count) = i
        Next
    End If
    If count = 1 Then Err.Raise vbObjectError + 981, , "No data rows remain in the selected analysis scope. Clear filters or choose All rows."
    ReDim result(1 To count, 1 To UBound(original, 2)): ReDim sourceRows(1 To count)
    For i = 1 To count
        sourceRows(i) = r.Row + rows(i) - 1
        For j = 1 To UBound(original, 2): result(i, j) = original(rows(i), j): Next
    Next
    lastKey = ScopeKey(r): lastIncluded = count - 1: lastExcluded = UBound(original, 1) - count
    ScopedValues = result
End Function

Public Function ScopeSourceRow(ByVal r As Range, ByVal index As Long) As Long
    ScopeSourceRow = r.Row + index - 1
    If ScopeKey(r) = lastKey Then ScopeSourceRow = sourceRows(index)
End Function

Public Sub AnnotateAnalysisScope(ByVal ws As Worksheet, ByVal r As Range)
    Dim i As Long
    If ScopeKey(r) <> lastKey Then Exit Sub
    ws.Range("A2").Value2 = ws.Range("A2").Value2 & " | Scope: " & lastScope & "; included=" & lastIncluded & "; excluded=" & lastExcluded
    For i = ws.CustomProperties.Count To 1 Step -1
        If ws.CustomProperties(i).Name = "CAP Scope" Then ws.CustomProperties(i).Delete
    Next
    ws.CustomProperties.Add "CAP Scope", lastScope & "; included=" & lastIncluded & "; excluded=" & lastExcluded
End Sub
