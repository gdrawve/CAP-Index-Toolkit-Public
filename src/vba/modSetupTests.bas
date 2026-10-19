Attribute VB_Name = "modSetupTests"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
Option Explicit
Private checks As Long
Private Sub CheckSetup(ByVal condition As Boolean, ByVal description As String)
    checks = checks + 1
    If Not condition Then Err.Raise vbObjectError + 1203, , "FAIL: " & description
End Sub
Public Function SetupDefaults(ByVal toolId As String) As Object
    Dim spec As Object, values As Object, role As Object, item As Object
    Set spec = ToolSpec(toolId): Set values = CreateObject("Scripting.Dictionary")
    For Each role In spec("Roles"): values.Add role("Key"), "": Next
    For Each item In spec("Options"): values.Add item("Key"), item("Default"): Next
    Set SetupDefaults = values
End Function
Public Function RunSetupTests() As String
    Dim wb As Workbook, ws As Worksheet, source As Range, spec As Object, values As Object, role As Object, item As Object
    Dim id As Variant, keys As Object, data As Variant, noRange As Range, form As Object, before As Long, errNo As Long, description As String
    Dim i As Long, fields As Variant, oldEvents As Boolean, oldAlerts As Boolean
    On Error GoTo failed
    checks = 0: oldEvents = Application.EnableEvents: oldAlerts = Application.DisplayAlerts
    Application.EnableEvents = False: Application.DisplayAlerts = False
    For Each id In ToolSpecIds()
        Set spec = ToolSpec(CStr(id)): Set keys = CreateObject("Scripting.Dictionary")
        CheckSetup Len(spec("Title")) > 0 And spec("Roles").Count > 0 And spec("Roles").Count <= 10 And spec("Options").Count <= 6, "Spec slots " & id
        For Each role In spec("Roles")
            CheckSetup Not keys.Exists(role("Key")), "Unique role " & id & "/" & role("Key")
            keys.Add role("Key"), True
        Next
        For Each item In spec("Options")
            If item("Kind") = "list" Then CheckSetup InStr("|" & item("Choices") & "|", "|" & item("Default") & "|") > 0, "Option default " & id & "/" & item("Key")
        Next
    Next
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    fields = Array("Address", "City", "State", "ZIP", "CAP_HundredBlock", "CrimeType", "IncidentID", "Date", "Time", "Hour (0-23)", "Day of Week", "Latitude", "Longitude", "Agency")
    For i = 0 To UBound(fields): ws.Cells(1, i + 1).Value2 = fields(i): Next
    For i = 2 To 5
        ws.Cells(i, 1).Value2 = "1200 MAIN ST": ws.Cells(i, 2).Value2 = "TESTCITY": ws.Cells(i, 3).Value2 = "IL": ws.Cells(i, 4).Value2 = "60000"
        ws.Cells(i, 5).Value2 = "1200 BLOCK MAIN ST": ws.Cells(i, 6).Value2 = "Robbery": ws.Cells(i, 7).Value2 = "I" & i
        ws.Cells(i, 8).Value2 = CDbl(DateSerial(2026, i - 1, 1)): ws.Cells(i, 9).Value2 = "14:30": ws.Cells(i, 10).Value2 = 14
        ws.Cells(i, 11).Value2 = "Monday": ws.Cells(i, 12).Value2 = 35: ws.Cells(i, 13).Value2 = -92: ws.Cells(i, 14).Value2 = "TEST"
    Next
    Set source = ws.Range("A1:N5"): data = source.Rows(1).Value2
    Set spec = ToolSpec("capPareto"): Set values = SetupDefaults("capPareto")
    CheckSetup Len(SetupProblem(spec, data, values, noRange)) > 0, "No range rejected"
    CheckSetup InStr(SetupProblem(spec, data, values, source), "Hundred Block") > 0, "Missing required mapping named"
    values("HUNDRED-BLOCK") = "CAP_HundredBlock": values("CRIME") = "CrimeType"
    CheckSetup SetupProblem(spec, data, values, source) = "", "Complete concentration setup accepted"
    values("LOCATION LABEL") = "CAP_HundredBlock"
    CheckSetup InStr(SetupProblem(spec, data, values, source), "different fields") > 0, "Duplicate role mapping rejected"
    Set spec = ToolSpec("capCensus"): Set values = SetupDefaults("capCensus"): values("ADDRESS") = "Address"
    CheckSetup InStr(SetupProblem(spec, data, values, source), "City and State") > 0, "Census locality alternatives required"
    values("ZIP") = "ZIP": CheckSetup SetupProblem(spec, data, values, source) = "", "Census ZIP alternative accepted"
    Set spec = ToolSpec("capTrends"): Set values = SetupDefaults("capTrends")
    values("DATE") = "Date": values("CRIME") = "CrimeType": values("START") = "04/01/2026": values("END") = "01/01/2026"
    CheckSetup InStr(SetupProblem(spec, data, values, source), "Start date") > 0, "Reversed dates rejected"
    values("END") = "05/01/2026": values("MEASURE") = "Unique incidents"
    CheckSetup InStr(SetupProblem(spec, data, values, source), "Incident ID") > 0, "Unique incidents requires ID"
    Set spec = ToolSpec("capCompare"): Set values = SetupDefaults("capCompare")
    values("DATE") = "Date": values("CRIME") = "CrimeType": values("START") = "03/01/2026": values("END") = "04/01/2026"
    values("A START") = "02/01/2026": values("A END") = "03/02/2026"
    CheckSetup InStr(SetupProblem(spec, data, values, source), "Period A must end") > 0, "Overlapping comparison periods rejected"
    Set spec = ToolSpec("capCheck"): Set values = SetupDefaults("capCheck")
    CheckSetup SetupProblem(spec, data, values, source) = "", "Dataset Check permits no optional mappings"
    wb.CustomDocumentProperties.Add "CAP Field ADDRESS", False, 4, "City"
    Set spec = ToolSpec("capBlocks"): Set role = spec("Roles")(1)
    CheckSetup SuggestHeader(role, data, wb) = "Address", "Legacy contaminated ADDRESS ignored"
    RememberField wb, "ADDRESS", "City", "ADDRESS"
    CheckSetup SuggestHeader(role, data, wb) = "City", "New explicit mapping preferred"
    RememberField wb, "ADDRESS", "Address", "ADDRESS"
    RememberField wb, "CITY", "City", "CITY": RememberField wb, "ZIP", "ZIP", "ZIP"
    CheckSetup SavedField(wb, "CITY", "", "CITY") = "City" And SavedField(wb, "ZIP", "", "ZIP") = "ZIP", "Independent City and ZIP role keys"
    before = wb.Worksheets.Count
    Load frmCAPSetup: Set form = frmCAPSetup
    For Each id In ToolSpecIds()
        form.Configure CStr(id): form.SetSource source
        CheckSetup form.Width <= 753 And form.Height <= 530, "Small-screen form bounds " & id
        CheckSetup InStr(SetupHelp(ToolSpec(CStr(id))), "Only selected rows") > 0, "Tool help uses selected data " & id
    Next
    form.Configure "capHeatmap": form.SetSource source
    form.Controls("cboOption1").Value = "Day of Week plus Hour"
    CheckSetup Not form.Controls("cboField2").Visible And form.Controls("cboField4").Visible And form.Controls("cboField5").Visible, "Heat map fields follow time mode"
    form.Configure "capTrends": form.SetSource source
    CheckSetup form.Controls("cboOption1").List(1) = "Robbery", "Crime dropdown shows plain source value"
    form.Controls("cboOption1").ListIndex = 1: Set values = form.ReadChoices()
    CheckSetup values("CRIME FILTER") = "Type: ROBBERY", "Plain crime maps to internal category"
    CheckSetup form.Controls("txtOption3").Text = "01/01/2026" And form.Controls("txtOption4").Text = "04/01/2026", "Date options use observed span"
    form.Controls("cboField3").Value = "(none)"
    CheckSetup form.Controls("cboOption2").Value = "Source records" And Not form.Controls("cboOption2").Enabled, "Measure fixed without Incident ID"
    CheckSetup form.Controls("cmdRun").Enabled, "Valid setup enables Run"
    form.Controls("cboField2").ListIndex = -1
    CheckSetup Not form.Controls("cmdRun").Enabled And InStr(form.Controls("lblStatus").Caption, "Crime Type") > 0, "Invalid setup disables Run with clear status"
    CheckSetup wb.Worksheets.Count = before, "Setup configuration and help create no sheets"
    Unload frmCAPSetup
    CheckSetup CheckCAPSetup() Like "PASS:*", "Packaged form structure"
    wb.Close False: Application.EnableEvents = oldEvents: Application.DisplayAlerts = oldAlerts
    RunSetupTests = "PASS: " & checks & " setup/spec/validation/mapping/form assertions": Exit Function
failed:
    errNo = Err.Number: description = Err.Description
    On Error Resume Next: Unload frmCAPSetup
    If Not wb Is Nothing Then wb.Close False
    Application.EnableEvents = oldEvents: Application.DisplayAlerts = oldAlerts
    On Error GoTo 0: Err.Raise errNo, "Setup tests", description
End Function
