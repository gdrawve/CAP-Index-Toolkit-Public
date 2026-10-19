' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
Option Explicit
Public ToolId As String
Public Spec As Object
Public SetupSource As Range
Public Result As Object
Private headers As Variant, data As Variant, busy As Boolean, helpOpen As Boolean
Private dateField As String, crimeField As String, incidentField As String, dateIssue As String
Private crimeValues As Collection
Public Sub Configure(ByVal id As String)
    Dim i As Long, item As Object, role As Object, choice As Variant, r As Range
    busy = True: ToolId = id: Set Spec = ToolSpec(id): Set Result = Nothing
    Me.Caption = "CAP Index | " & Spec("Title") & " | " & CAP_VERSION
    Me.Controls("lblTitle").Caption = Spec("Title"): Me.Controls("lblDescription").Caption = Spec("Description")
    For i = 1 To 10
        Me.Controls("cboField" & i).Clear
        Me.Controls("lblField" & i).Visible = False: Me.Controls("cboField" & i).Visible = False: Me.Controls("lblHelp" & i).Visible = False
    Next
    For i = 1 To 6
        Me.Controls("lblOption" & i).Visible = False: Me.Controls("cboOption" & i).Visible = False: Me.Controls("txtOption" & i).Visible = False
        Me.Controls("cboOption" & i).Clear: Me.Controls("txtOption" & i).Text = ""
    Next
    i = 0
    For Each item In Spec("Options")
        i = i + 1: Me.Controls("lblOption" & i).Caption = item("Label")
        If item("Kind") = "list" Then
            For Each choice In Split(item("Choices"), "|"): Me.Controls("cboOption" & i).AddItem choice: Next
            Me.Controls("cboOption" & i).Value = item("Default")
        Else
            Me.Controls("txtOption" & i).Text = item("Default")
        End If
    Next
    busy = False
    On Error Resume Next
    If TypeName(Selection) = "Range" Then Set r = Selection
    On Error GoTo 0
    If Not r Is Nothing Then
        If r.Areas.Count <> 1 Or r.Rows.Count < 2 Or r.Rows.Count > 200001 Or r.Columns.Count > 200 Then Set r = Nothing
    End If
    If Not r Is Nothing Then
        If r.Parent.Parent Is ThisWorkbook Then Set r = Nothing
    End If
    SetSource r
End Sub
Public Sub SetSource(ByVal r As Range)
    Dim old As Object, role As Object, i As Long, j As Long, value As String, seen As Object
    On Error GoTo failed
    Set old = ReadChoices(): busy = True: Set SetupSource = r
    dateField = "[reload]": crimeField = "[reload]": incidentField = "": dateIssue = ""
    If r Is Nothing Then
        Me.Controls("lblRange").Caption = "No data selected. Choose Select range."
        Me.Controls("lblCounts").Caption = "Include a header row and all fields required by the tool."
    Else
        ValidateOutputSource r
        headers = r.Rows(1).Value2
        If r.Columns.Count = 1 Then
            ReDim headers(1 To 1, 1 To 1): headers(1, 1) = r.Cells(1, 1).Value2
        End If
        If Spec("Preparation") Then data = r.Value2 Else data = ScopedValues(r)
        Me.Controls("lblRange").Caption = r.Parent.Parent.Name & " | " & r.Parent.Name & "!" & r.Address
        Me.Controls("lblCounts").Caption = Format$(r.Rows.Count - 1, "#,##0") & " selected data rows | " & r.Columns.Count & " fields | " & IIf(Spec("Preparation"), "Processes all selected rows", AnalysisScope(r.Parent.Parent) & "; " & Format$(UBound(data, 1) - 1, "#,##0") & " included")
    End If
    i = 0
    For Each role In Spec("Roles")
        i = i + 1
        With Me.Controls("cboField" & i)
            .Clear
            If Not role("Required") Or ToolId = "capHeatmap" Then .AddItem "(none)"
            If Not r Is Nothing Then
                Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
                For j = 1 To UBound(headers, 2)
                    value = SafeText(headers(1, j))
                    If Len(value) > 0 And Not IsError(headers(1, j)) Then
                        If Not seen.Exists(value) Then .AddItem value: seen.Add value, True
                    End If
                Next
                value = SetupValue(old, role("Key"))
                If Len(value) = 0 Or Not seen.Exists(value) Then value = SuggestHeader(role, headers, r.Parent.Parent)
                If Len(value) > 0 Then
                    .Value = value
                ElseIf Not role("Required") Or ToolId = "capHeatmap" Then
                    .ListIndex = 0
                Else
                    .ListIndex = -1
                End If
            End If
        End With
    Next
    busy = False: Recheck: Exit Sub
failed:
    busy = False: Me.Controls("lblStatus").Caption = Err.Description: Me.Controls("cmdRun").Enabled = False
End Sub
Public Function ReadChoices() As Object
    Dim values As Object, role As Object, item As Object, i As Long, value As String
    Set values = CreateObject("Scripting.Dictionary")
    If Spec Is Nothing Then Set ReadChoices = values: Exit Function
    i = 0
    For Each item In Spec("Options")
        i = i + 1
        If item("Kind") = "list" Then value = Me.Controls("cboOption" & i).Value Else value = Me.Controls("txtOption" & i).Text
        If item("Key") = "CRIME FILTER" Then
            value = "All crime types"
            If Not crimeValues Is Nothing Then
                If Me.Controls("cboOption" & i).ListIndex >= 0 Then value = crimeValues(Me.Controls("cboOption" & i).ListIndex + 1)
            End If
        End If
        values.Add item("Key"), value
    Next
    i = 0
    For Each role In Spec("Roles")
        i = i + 1: value = SafeText(Me.Controls("cboField" & i).Value)
        If value = "(none)" Or Not SetupRoleActive(ToolId, role("Key"), values) Then value = ""
        values.Add role("Key"), value
    Next
    If Not SetupSource Is Nothing Then values.Add "Source", SetupSource
    Set ReadChoices = values
End Function
Private Function OptionIndex(ByVal key As String) As Long
    Dim item As Object, i As Long
    For Each item In Spec("Options")
        i = i + 1: If item("Key") = key Then OptionIndex = i: Exit Function
    Next
End Function
Private Sub FillCrime(ByVal values As Object)
    Dim index As Long, column As Long, i As Long, category As String, display As String, selected As String, seen As Object
    index = OptionIndex("CRIME FILTER"): If index = 0 Then Exit Sub
    selected = SetupValue(values, "CRIME FILTER")
    Set crimeValues = New Collection: crimeValues.Add "All crime types"
    Me.Controls("cboOption" & index).Clear: Me.Controls("cboOption" & index).AddItem "All crime types"
    If Len(SetupValue(values, "CRIME")) > 0 And Not SetupSource Is Nothing Then
        column = HeaderColumn(headers, SetupValue(values, "CRIME")): Set seen = CreateObject("Scripting.Dictionary")
        For i = 2 To UBound(data, 1)
            category = CrimeCategory(data(i, column))
            If Not seen.Exists(category) Then
                seen.Add category, True: crimeValues.Add category
                display = SafeText(data(i, column))
                If Len(display) = 0 Or IsError(data(i, column)) Then display = category
                If display = "All crime types" Then display = display & " (source value)"
                Me.Controls("cboOption" & index).AddItem display
            End If
        Next
    End If
    Me.Controls("cboOption" & index).ListIndex = 0
    For i = 1 To crimeValues.Count
        If crimeValues(i) = selected Then Me.Controls("cboOption" & index).ListIndex = i - 1: Exit For
    Next
End Sub
Private Sub FillDates(ByVal values As Object)
    Dim column As Long, i As Long, parsed As Date, ht As Boolean, status As String, minimum As Double, maximum As Double, index As Long
    dateIssue = "": Me.Controls("lblObserved").Caption = ""
    If SetupSource Is Nothing Or Len(SetupValue(values, "DATE")) = 0 Then Exit Sub
    column = HeaderColumn(headers, SetupValue(values, "DATE"))
    For i = 2 To UBound(data, 1)
        If ParseTimestamp(data(i, column), SetupSource.Parent.Parent.Date1904, parsed, ht, status) Then
            If minimum = 0 Or CDbl(DateValue(parsed)) < minimum Then minimum = CDbl(DateValue(parsed))
            If CDbl(DateValue(parsed)) > maximum Then maximum = CDbl(DateValue(parsed))
        End If
    Next
    If minimum = 0 Then
        If ToolId <> "capCheck" Then dateIssue = "No valid dates were found in the chosen Date field."
        Exit Sub
    End If
    Me.Controls("lblObserved").Caption = "Observed dates: " & Format$(CDate(minimum), "mm/dd/yyyy") & " to " & Format$(CDate(maximum), "mm/dd/yyyy")
    index = OptionIndex("START"): If index > 0 Then Me.Controls("txtOption" & index).Text = Format$(CDate(minimum), "mm/dd/yyyy")
    index = OptionIndex("END"): If index > 0 Then Me.Controls("txtOption" & index).Text = Format$(CDate(maximum), "mm/dd/yyyy")
    index = OptionIndex("A END"): If index > 0 Then Me.Controls("txtOption" & index).Text = Format$(CDate(minimum - 1), "mm/dd/yyyy")
    index = OptionIndex("A START"): If index > 0 Then Me.Controls("txtOption" & index).Text = Format$(CDate(minimum - 1 - (maximum - minimum)), "mm/dd/yyyy")
End Sub
Public Sub Recheck()
    Dim values As Object, role As Object, item As Object, i As Long, shown As Long, y As Double, bodyBottom As Double, index As Long, problem As String
    If busy Or Spec Is Nothing Then Exit Sub
    On Error GoTo failed
    busy = True: Set values = ReadChoices()
    If Not SetupSource Is Nothing Then
        If crimeField <> SetupValue(values, "CRIME") Then FillCrime values: crimeField = SetupValue(values, "CRIME")
        If dateField <> SetupValue(values, "DATE") Then FillDates values: dateField = SetupValue(values, "DATE")
    End If
    index = OptionIndex("MEASURE")
    If index > 0 Then
        If Len(SetupValue(values, "INCIDENTID")) = 0 Then
            Me.Controls("cboOption" & index).Value = "Source records": Me.Controls("cboOption" & index).Enabled = False
        Else
            Me.Controls("cboOption" & index).Enabled = True
            If Len(incidentField) = 0 Then Me.Controls("cboOption" & index).Value = "Unique incidents"
        End If
        incidentField = SetupValue(values, "INCIDENTID")
    End If
    Set values = ReadChoices(): i = 0
    For Each role In Spec("Roles")
        i = i + 1
        If SetupRoleActive(ToolId, role("Key"), values) Then
            y = 155 + shown * 30: shown = shown + 1
            Me.Controls("lblField" & i).Top = y + 2: Me.Controls("cboField" & i).Top = y: Me.Controls("lblHelp" & i).Top = y + 19
            Me.Controls("lblField" & i).Caption = role("Label") & IIf(SetupRoleRequired(Spec, role, values), " *", "")
            Me.Controls("lblHelp" & i).Caption = role("Help"): Me.Controls("cboField" & i).ControlTipText = role("Help")
        End If
        Me.Controls("lblField" & i).Visible = SetupRoleActive(ToolId, role("Key"), values) And Not helpOpen
        Me.Controls("cboField" & i).Visible = Me.Controls("lblField" & i).Visible
        Me.Controls("lblHelp" & i).Visible = Me.Controls("lblField" & i).Visible
    Next
    i = 0
    For Each item In Spec("Options")
        i = i + 1
        Me.Controls("lblOption" & i).Visible = Not helpOpen
        Me.Controls("cboOption" & i).Visible = (item("Kind") = "list") And Not helpOpen
        Me.Controls("txtOption" & i).Visible = (item("Kind") = "text") And Not helpOpen
        If ToolId = "capSingleSite" And (item("Key") = "START" Or item("Key") = "END") Then
            Me.Controls("txtOption" & i).Enabled = Len(SetupValue(values, "DATE")) > 0
        End If
    Next
    bodyBottom = 155 + Application.Max(shown * 30, Spec("Options").Count * 38, 120)
    Me.Controls("lblObserved").Top = bodyBottom + 2
    Me.Controls("lblNotice").Top = bodyBottom + 32
    Me.Controls("lblNotice").Caption = IIf(ToolId = "capCensus", "Address components and request IDs are sent to the US Census Bureau.", "Required fields are marked *. Original source data is not saved automatically.")
    Me.Controls("lblStatus").Top = bodyBottom + 52
    Me.Controls("cmdGuide").Top = bodyBottom + 86: Me.Controls("cmdCancel").Top = bodyBottom + 86: Me.Controls("cmdRun").Top = bodyBottom + 86
    Me.Height = bodyBottom + 146
    Me.Controls("txtHelp").Height = bodyBottom - 155 + 28
    Me.Controls("lblFields").Caption = IIf(helpOpen, "TOOL GUIDE", "FIELDS")
    Me.Controls("lblOptions").Visible = (Spec("Options").Count > 0 And Not helpOpen)
    Me.Controls("lblObserved").Visible = Not helpOpen
    Me.Controls("txtHelp").Visible = helpOpen
    If SetupSource Is Nothing Then
        problem = "Select headers and at least one data row."
    Else
        problem = SetupProblem(Spec, headers, values, SetupSource)
        If Len(problem) = 0 Then problem = dateIssue
    End If
    Me.Controls("lblStatus").Caption = problem: Me.Controls("cmdRun").Enabled = (Len(problem) = 0 And Not helpOpen)
    busy = False: Exit Sub
failed:
    busy = False: Me.Controls("lblStatus").Caption = Err.Description: Me.Controls("cmdRun").Enabled = False
End Sub
Private Sub cmdSelect_Click()
    Dim r As Range
    Me.Hide
    On Error Resume Next
    Set r = Application.InputBox("Select one rectangle including headers and data. Only selected rows and fields are used.", "CAP Index | Select data", Type:=8)
    On Error GoTo 0
    If Not r Is Nothing Then SetSource r
    Me.Show vbModal
End Sub
Private Sub cmdRun_Click()
    Recheck
    If Not Me.Controls("cmdRun").Enabled Then Exit Sub
    Set Result = ReadChoices(): Me.Hide
End Sub
Private Sub cmdCancel_Click()
    Set Result = Nothing: Me.Hide
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = True: Set Result = Nothing: Me.Hide
End Sub
Private Sub cmdGuide_Click()
    helpOpen = Not helpOpen
    Me.Controls("txtHelp").Text = SetupHelp(Spec)
    Me.Controls("cmdGuide").Caption = IIf(helpOpen, "Back to setup", "Tool Guide")
    Recheck
End Sub
Private Sub cboField1_Change()
    Recheck
End Sub
Private Sub cboField2_Change()
    Recheck
End Sub
Private Sub cboField3_Change()
    Recheck
End Sub
Private Sub cboField4_Change()
    Recheck
End Sub
Private Sub cboField5_Change()
    Recheck
End Sub
Private Sub cboField6_Change()
    Recheck
End Sub
Private Sub cboField7_Change()
    Recheck
End Sub
Private Sub cboField8_Change()
    Recheck
End Sub
Private Sub cboField9_Change()
    Recheck
End Sub
Private Sub cboField10_Change()
    Recheck
End Sub
Private Sub cboOption1_Change()
    Recheck
End Sub
Private Sub cboOption2_Change()
    Recheck
End Sub
Private Sub cboOption3_Change()
    Recheck
End Sub
Private Sub cboOption4_Change()
    Recheck
End Sub
Private Sub cboOption5_Change()
    Recheck
End Sub
Private Sub cboOption6_Change()
    Recheck
End Sub
Private Sub txtOption1_Change()
    Recheck
End Sub
Private Sub txtOption2_Change()
    Recheck
End Sub
Private Sub txtOption3_Change()
    Recheck
End Sub
Private Sub txtOption4_Change()
    Recheck
End Sub
Private Sub txtOption5_Change()
    Recheck
End Sub
Private Sub txtOption6_Change()
    Recheck
End Sub
