Attribute VB_Name = "modReportBuilder"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
Option Explicit

Public Sub CreateReportBuilder(Optional ByVal wb As Workbook, Optional ByVal showSummary As Boolean = True)
    Dim plan As Worksheet, ws As Worksheet, entries As Collection, rowNo As Long
    On Error GoTo failed
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Err.Raise vbObjectError + 995, , "Open a workbook with CAP reports."
    If wb Is ThisWorkbook Or wb.ReadOnly Or wb.ProtectStructure Then Err.Raise vbObjectError + 995, , "Select a writable data workbook that permits new sheets."
    Set entries = New Collection
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            If Not CAPSummaryRange(ws) Is Nothing Then entries.Add ws
        End If
    Next
    If entries.Count = 0 Then Err.Raise vbObjectError + 995, , "Create the analyses you want first, then select their tables/charts with Create Report."
    Set plan = NewOutput(wb, "Report_Builder", wb.Worksheets(1).Range("A1"), "Choose report sections below. Set Include=Yes, order and Tables/Charts/Both, then click Export Selected Report. Create additional analyses first to make them available here.")
    plan.Range("A1:E1").Merge: plan.Range("A1").Font.Size = 20: plan.Range("A1:E1").Interior.Color = RGB(26, 55, 82): plan.Range("A1").Font.Color = vbWhite: plan.Rows(1).RowHeight = 34
    plan.Range("A2:E2").Merge: plan.Range("A3:E3").Merge: plan.Range("A2:E3").WrapText = True: plan.Rows("2:3").RowHeight = 42
    plan.Range("A4:E4").Merge: plan.Range("A4").Value2 = "PDF uses full report tables and captioned chart pages. PowerPoint uses summary tables and chart slides. Mixed measures, scopes and periods stay labelled; no totals are combined."
    plan.Range("A4").WrapText = True: plan.Rows(4).RowHeight = 36
    WriteHeaders plan, 6, 1, Array("Include", "Order", "Report sheet", "Section title", "Content")
    rowNo = 7
    For Each ws In entries
        plan.Cells(rowNo, 1).Value2 = "No": plan.Cells(rowNo, 2).Value2 = rowNo - 6
        plan.Cells(rowNo, 3).Value2 = ws.Name: plan.Cells(rowNo, 4).Value2 = ws.Range("A1").Value2: plan.Cells(rowNo, 5).Value2 = "Both": rowNo = rowNo + 1
    Next
    plan.Range("A7:A" & rowNo - 1).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes" & Application.International(xlListSeparator) & "No"
    plan.Range("E7:E" & rowNo - 1).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=Join(Array("Tables", "Charts", "Both"), Application.International(xlListSeparator))
    plan.Range("A7:B" & rowNo - 1 & ",D7:E" & rowNo - 1).Interior.Color = RGB(224, 239, 250)
    FinishOutput plan, rowNo - 1, 5
    plan.Columns("D").ColumnWidth = 50: plan.Range("D7:D" & rowNo - 1).WrapText = True: plan.Range("A6:E" & rowNo - 1).Rows.AutoFit
    plan.Names.Add Name:="_CAPReportPlan", RefersTo:="='" & plan.Name & "'!$A$6:$E$" & rowNo - 1
    If showSummary Then MsgBox "Output: " & plan.Name & ". Choose sections, content and order. The selection is saved with this workbook.", vbInformation, "CAP Index | Create Report"
    Exit Sub
failed:
    If Not showSummary Then Err.Raise Err.Number, "Create Report", Err.Description
    MsgBox Err.Description, vbExclamation, "CAP Index | Create Report"
End Sub

Public Function SelectedReportSections(ByVal plan As Worksheet, ByRef modes As Object, ByRef titles As Object) As Collection
    Dim data As Range, chosen As New Collection, orders As Object, selectedNames As Object, i As Long, j As Long, count As Long
    Dim numbers() As Long, names() As String, name As String, mode As String, title As String, ws As Worksheet, n As Double, temp As Long, s As String
    Set data = CAPNamedRange(plan, "_CAPReportPlan")
    If data Is Nothing Then Err.Raise vbObjectError + 995, , "Select the Report_Builder sheet first."
    Set orders = CreateObject("Scripting.Dictionary"): Set selectedNames = CreateObject("Scripting.Dictionary")
    Set modes = CreateObject("Scripting.Dictionary"): Set titles = CreateObject("Scripting.Dictionary")
    ReDim numbers(1 To data.Rows.Count): ReDim names(1 To data.Rows.Count)
    For i = 2 To data.Rows.Count
        If NormalText(data.Cells(i, 1).Value2) <> "YES" And NormalText(data.Cells(i, 1).Value2) <> "NO" Then Err.Raise vbObjectError + 995, , "Include must be Yes or No."
        If NormalText(data.Cells(i, 1).Value2) = "YES" Then
            If Not IsNumeric(data.Cells(i, 2).Value2) Then Err.Raise vbObjectError + 995, , "Use a positive whole-number section order."
            n = CDbl(data.Cells(i, 2).Value2)
            If n < 1 Or n > 999 Or n <> Fix(n) Then Err.Raise vbObjectError + 995, , "Use section order 1 to 999."
            If orders.Exists(CStr(n)) Then Err.Raise vbObjectError + 995, , "Selected sections have duplicate order numbers."
            orders.Add CStr(n), True: name = SafeText(data.Cells(i, 3).Value2)
            If selectedNames.Exists(name) Then Err.Raise vbObjectError + 995, , "Select each report sheet once."
            Set ws = Nothing: On Error Resume Next: Set ws = plan.Parent.Worksheets(name): On Error GoTo 0
            If ws Is Nothing Then Err.Raise vbObjectError + 995, , "Report sheet no longer exists: " & name
            If CAPSummaryRange(ws) Is Nothing Then Err.Raise vbObjectError + 995, , "Select CAP report sheets only."
            RequireFreshReport ws: mode = SafeText(data.Cells(i, 5).Value2)
            If mode <> "Tables" And mode <> "Charts" And mode <> "Both" Then Err.Raise vbObjectError + 995, , "Content must be Tables, Charts or Both."
            If mode = "Charts" And ws.ChartObjects.Count = 0 Then Err.Raise vbObjectError + 995, , name & " has no charts; select Tables."
            title = SafeText(data.Cells(i, 4).Value2): If Len(title) = 0 Then title = ws.Range("A1").Value2
            count = count + 1: numbers(count) = CLng(n): names(count) = name: selectedNames.Add name, True: modes.Add name, mode: titles.Add name, title
        End If
    Next
    If count = 0 Or count > 20 Then Err.Raise vbObjectError + 995, , "Select 1 to 20 report sections."
    For i = 1 To count - 1: For j = i + 1 To count
        If numbers(j) < numbers(i) Then temp = numbers(i): numbers(i) = numbers(j): numbers(j) = temp: s = names(i): names(i) = names(j): names(j) = s
    Next: Next
    For i = 1 To count: chosen.Add plan.Parent.Worksheets(names(i)): Next
    Set SelectedReportSections = chosen
End Function

Public Sub ExportSelectedReport(Optional ByVal plan As Worksheet, Optional ByVal format As String = "", Optional ByVal path As String = "", Optional ByVal showSummary As Boolean = True)
    Dim sections As Collection, modes As Object, titles As Object, choice As Variant
    On Error GoTo failed
    If plan Is Nothing Then Set plan = ActiveSheet
    Set sections = SelectedReportSections(plan, modes, titles)
    If Len(format) = 0 Then
        choice = Application.InputBox("Export selected sections: 1 = PDF; 2 = PowerPoint.", "CAP Index | Selected Report", "1", Type:=1)
        If VarType(choice) = vbBoolean Then Exit Sub
        If choice <> 1 And choice <> 2 Then Err.Raise vbObjectError + 995, , "Choose PDF or PowerPoint."
        format = IIf(choice = 1, "pdf", "pptx")
    End If
    If Len(path) = 0 Then path = ChooseExportPath(format, "Save selected CAP report")
    If Len(path) = 0 Then Exit Sub
    If format = "pptx" Then
        ExportBriefingPowerPoint sections, path, False, modes, titles
    ElseIf format = "pdf" Then
        ExportSelectedPDF sections, modes, titles, path
    Else
        Err.Raise vbObjectError + 995, , "Use pdf or pptx."
    End If
    If showSummary Then MsgBox "Saved " & path & " with " & sections.Count & " selected sections. Review each section's scope, period, measure and coverage before sharing.", vbInformation, "CAP Index"
    Exit Sub
failed:
    If Not showSummary Then Err.Raise Err.Number, "Selected Report", Err.Description
    MsgBox Err.Description, vbExclamation, "CAP Index | Selected Report"
End Sub

Private Sub ExportSelectedPDF(ByVal sections As Collection, ByVal modes As Object, ByVal titles As Object, ByVal path As String)
    Dim temporary As Workbook, ws As Worksheet, copy As Worksheet, area As Range, body As Range, picture As ChartObject, pasted As Shape
    Dim folder As String, imagePath As String, col As Long, fit As Double, count As Long, oldEvents As Boolean, oldScreen As Boolean, oldAlerts As Boolean, errNo As Long, description As String
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating: oldAlerts = Application.DisplayAlerts
    On Error GoTo failed
    RequireNewExportFile path: Application.EnableEvents = False
    folder = Environ$("TEMP") & "\CAPSelected_" & Format$(Now, "yyyymmddhhnnss") & "_" & Format$(Timer * 100, "0"): MkDir folder
    Set temporary = Workbooks.Add(xlWBATWorksheet)
    For Each ws In sections
        Application.StatusBar = "CAP Index: assembling " & ws.Name
        If modes(ws.Name) <> "Charts" Then
            Set area = CAPNamedRange(ws, "_CAPReportBody"): If area Is Nothing Then Set area = CAPSummaryRange(ws)
            ws.Copy After:=temporary.Worksheets(temporary.Worksheets.Count): Set copy = temporary.Worksheets(temporary.Worksheets.Count)
            count = count + 1: copy.Name = "Section_" & count
            For col = copy.ChartObjects.Count To 1 Step -1: copy.ChartObjects(col).Delete: Next
            copy.Range("A1").Value2 = titles(ws.Name)
            fit = Application.Min(1, 700 / copy.Range(area.Address).Width)
            For col = area.Column To area.Column + area.Columns.Count - 1: copy.Columns(col).ColumnWidth = copy.Columns(col).ColumnWidth * fit: Next
            copy.Range(area.Address).Font.Size = 9: copy.Range("A1").Font.Size = 18
            copy.Range("A2:H5").WrapText = True: copy.Rows("1:5").RowHeight = 36
            copy.Range(copy.Cells(6, area.Column), copy.Cells(area.Row + area.Rows.Count - 1, area.Column + area.Columns.Count - 1)).Rows.AutoFit
            With copy.PageSetup
                .PrintArea = area.Address: .PrintTitleRows = "$1:$6": .PrintTitleColumns = "": .Orientation = xlLandscape: .PaperSize = xlPaperLetter
                .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
                .LeftMargin = 36: .RightMargin = 36: .TopMargin = 36: .BottomMargin = 36
                .LeftHeader = "CAP Index": .LeftFooter = "&8Toolkit: " & CAP_COPYRIGHT: .RightFooter = "Page &P of &N"
            End With
        End If
        If modes(ws.Name) <> "Tables" Then
            For Each picture In ws.ChartObjects
                count = count + 1: imagePath = folder & "\chart" & count & ".png": ExportChartPNG picture.Chart, imagePath
                Set copy = temporary.Worksheets.Add(After:=temporary.Worksheets(temporary.Worksheets.Count)): copy.Name = "Section_" & count
                Set pasted = copy.Shapes.AddPicture(imagePath, 0, -1, 0, 0, -1, -1): pasted.LockAspectRatio = -1: If pasted.Width > 680 Then pasted.Width = 680
                If pasted.Height > 440 Then pasted.Height = 440
                copy.Columns("A").ColumnWidth = 100: copy.Columns("A").ColumnWidth = copy.Columns("A").ColumnWidth * 700 / copy.Columns("A").Width
                copy.Rows("1:30").RowHeight = 15
                With copy.PageSetup
                    .PrintArea = "$A$1:$A$30": .PrintTitleRows = "": .Orientation = xlLandscape: .PaperSize = xlPaperLetter: .Zoom = 100
                    .LeftMargin = 36: .RightMargin = 36: .TopMargin = 36: .BottomMargin = 36
                    .LeftHeader = titles(ws.Name): .LeftFooter = "&8Toolkit: " & CAP_COPYRIGHT: .RightFooter = "Page &P of &N"
                End With
            Next
        End If
    Next
    Application.DisplayAlerts = False: temporary.Worksheets(1).Delete: Application.DisplayAlerts = oldAlerts
    temporary.ExportAsFixedFormat xlTypePDF, path, xlQualityStandard, False, False
    temporary.Close False: Set temporary = Nothing
    On Error Resume Next
    Kill folder & "\*.png": RmDir folder
    On Error GoTo failed
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen: Application.StatusBar = False
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    On Error Resume Next: If Not temporary Is Nothing Then temporary.Close False
    If Len(folder) > 0 Then Kill folder & "\*.png": RmDir folder
    Application.DisplayAlerts = oldAlerts: Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen: Application.StatusBar = False
    On Error GoTo 0: Err.Raise errNo, "Selected Report PDF", description
End Sub
