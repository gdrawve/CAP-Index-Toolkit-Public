Attribute VB_Name = "modExports"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit

Public Function CAPSummaryRange(ByVal ws As Worksheet) As Range
    Dim named As Name, cache As Range
    For Each named In ws.Names
        If Right$(named.Name, Len("_CAPExportSummary")) = "_CAPExportSummary" Then Set CAPSummaryRange = named.RefersToRange: Exit Function
    Next
    Set cache = TimeReportCache(ws)
    If Not cache Is Nothing Then
        If cache.Parent.Range("N1").Value2 = "Heatmap" Then Set CAPSummaryRange = ws.Range("A1:Y27") Else Set CAPSummaryRange = ws.Range("A1:I" & Application.Min(38, ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1))
        Exit Function
    End If
    Set cache = ConcentrationCache(ws)
    If Not cache Is Nothing Then Set CAPSummaryRange = ws.Range("A1:H28")
End Function

Public Function ChooseExportPath(ByVal extension As String, ByVal title As String) As String
    Dim answer As Variant, initial As String
    initial = DefaultExportFilename(ActiveWorkbook.Path, extension)
    answer = Application.GetSaveAsFilename(initial, UCase$(extension) & " file (*." & extension & "), *." & extension, , title)
    If VarType(answer) = vbBoolean Then Exit Function
    ChooseExportPath = CStr(answer)
    If LCase$(Right$(ChooseExportPath, Len(extension) + 1)) <> "." & extension Then ChooseExportPath = ChooseExportPath & "." & extension
    RequireNewExportFile ChooseExportPath
End Function

Public Function DefaultExportFilename(ByVal workbookFolder As String, ByVal extension As String) As String
    Dim fs As Object
    DefaultExportFilename = "CAP_" & Format$(Now, "yyyymmdd-hhnnss") & "." & extension
    If Len(workbookFolder) = 0 Or InStr(workbookFolder, "://") > 0 Then Exit Function
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(workbookFolder) Then DefaultExportFilename = fs.BuildPath(workbookFolder, DefaultExportFilename)
End Function

Public Sub RequireNewExportFile(ByVal path As String)
    Dim fs As Object, folder As String, filename As String, ch As Variant
    If InStr(path, "://") > 0 Then Err.Raise vbObjectError + 840, , "Choose a local or synced folder in the Save dialog; web addresses cannot be export filenames."
    Set fs = CreateObject("Scripting.FileSystemObject")
    filename = fs.GetFileName(path): folder = fs.GetParentFolderName(path)
    If Len(filename) = 0 Then Err.Raise vbObjectError + 840, , "Enter an export filename."
    For Each ch In Array("<", ">", ":", Chr$(34), "/", "\", "|", "?", "*")
        If InStr(filename, ch) > 0 Then Err.Raise vbObjectError + 840, , "Remove invalid characters from the export filename."
    Next
    If Len(folder) > 0 Then
        If Not fs.FolderExists(folder) Then Err.Raise vbObjectError + 840, , "The export folder does not exist or is unavailable. Choose an accessible folder."
    End If
    If fs.FileExists(path) Then Err.Raise vbObjectError + 840, , "Export file already exists; choose a new filename."
End Sub

Public Sub ExportRangePNG(ByVal r As Range, ByVal path As String)
    Dim ws As Worksheet, oldScreen As Boolean, oldEvents As Boolean, errNo As Long, description As String
    Dim attempt As Long, rendered As Boolean, captureError As String, appearance As Long, imageFormat As Long
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    On Error GoTo failed
    If r.Areas.Count <> 1 Or r.Rows.Count > 100 Or r.Columns.Count > 80 Then Err.Raise vbObjectError + 841, , "Choose one report range up to 100 rows and 80 columns."
    RequireNewExportFile path
    Set ws = r.Parent: RequireFreshReport ws
    Application.EnableEvents = False: Application.ScreenUpdating = True
    For attempt = 1 To 5
        appearance = xlScreen: imageFormat = xlPicture
        If attempt = 2 Or attempt = 4 Then imageFormat = xlBitmap
        If attempt = 5 Then appearance = xlPrinter
        rendered = TryRangePNG(r, path, appearance, imageFormat, captureError)
        If rendered Then Exit For
        Application.CutCopyMode = False: DoEvents
    Next
    If Not rendered Then Err.Raise vbObjectError + 842, , "Excel could not capture this report after five attempts. Close any open Excel dialogs, activate the report workbook and retry. A default printer is needed for the final fallback. Last capture error: " & captureError
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    On Error GoTo 0: Err.Raise errNo, "PNG export", description
End Sub

Private Function TryRangePNG(ByVal r As Range, ByVal path As String, ByVal appearance As Long, ByVal imageFormat As Long, ByRef captureError As String) As Boolean
    Dim picture As ChartObject, ownsFile As Boolean
    On Error GoTo failed
    r.Parent.Parent.Activate: r.Parent.Activate: r.Cells(1, 1).Select
    ActiveWindow.ScrollRow = r.Row: ActiveWindow.ScrollColumn = r.Column
    Application.CutCopyMode = False: DoEvents
    r.CopyPicture Appearance:=appearance, Format:=imageFormat
    DoEvents
    Set picture = r.Parent.ChartObjects.Add(r.Left, r.Top + r.Height + 20, Application.Max(600, r.Width), Application.Max(120, r.Height))
    picture.Activate: picture.Chart.Paste: DoEvents: picture.Chart.Refresh
    ownsFile = True
    If Not picture.Chart.Export(path, "PNG") Then Err.Raise vbObjectError + 842, , "PNG export failed."
    If FileLen(path) < 2000 Then Err.Raise vbObjectError + 842, , "Excel produced a blank or very small image."
    picture.Delete: Set picture = Nothing
    TryRangePNG = True: Exit Function
failed:
    captureError = CStr(Err.Number) & ": " & Err.Description
    On Error Resume Next
    If Not picture Is Nothing Then picture.Delete
    If ownsFile Then Kill path
    On Error GoTo 0
End Function

Public Sub SaveReportImage()
    Dim r As Range, ws As Worksheet, path As String, chart As Chart, choice As Variant
    On Error GoTo failed
    RequireFreshReport ActiveSheet
    Set chart = ActiveChart
    If Not chart Is Nothing Then
        path = ChooseExportPath("png", "Save selected chart as PNG"): If Len(path) = 0 Then Exit Sub
        ExportChartPNG chart, path
    Else
        Set ws = ActiveSheet: Set r = CAPSummaryRange(ws)
        If r Is Nothing Then
            On Error Resume Next
            Set r = Application.InputBox("Select a formatted range including title, dates, measure and coverage notes. Limit 100 rows x 80 columns.", "CAP Index | Save Image", Type:=8)
            On Error GoTo failed
            If r Is Nothing Then Exit Sub
        Else
            choice = Application.InputBox("1 = Export this report summary; 2 = Select another range. To export a chart, select the chart before clicking Save Image.", "CAP Index | Save Image", "1", Type:=1)
            If VarType(choice) = vbBoolean Then Exit Sub
            If choice = 2 Then
                Set r = Nothing
                On Error Resume Next: Set r = Application.InputBox("Select the report range including its labels and notes.", "CAP Index | Save Image", Type:=8): On Error GoTo failed
                If r Is Nothing Then Exit Sub
            ElseIf choice <> 1 Then
                Err.Raise vbObjectError + 844, , "Choose 1 or 2."
            End If
        End If
        path = ChooseExportPath("png", "Save report image"): If Len(path) = 0 Then Exit Sub
        ExportRangePNG r, path
    End If
    MsgBox "Saved " & path, vbInformation, "CAP Index": Exit Sub
failed: MsgBox "Export failed (" & Err.Number & "): " & Err.Description, vbExclamation, "CAP Index | Save Image"
End Sub

Public Sub ExportReportPDF(Optional ByVal ws As Worksheet, Optional ByVal path As String = "", Optional ByVal showSummary As Boolean = True)
    Dim summary As Range, area As Range, picture As ChartObject, named As Name, temporary As Workbook
    Dim copySheet As Worksheet, chartSheet As Worksheet, pasted As Shape, folder As String, imagePath As String
    Dim fitRatio As Double, col As Long
    Dim oldEvents As Boolean, oldScreen As Boolean, created As Boolean, i As Long, errNo As Long, description As String, invoked As Boolean
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    On Error GoTo failed
    invoked = ws Is Nothing: If ws Is Nothing Then Set ws = ActiveSheet
    RequireFreshReport ws: Set summary = CAPSummaryRange(ws)
    If summary Is Nothing Then Err.Raise vbObjectError + 845, , "Select a CAP Index analysis/report sheet."
    If Len(path) = 0 Then path = ChooseExportPath("pdf", "Save CAP Index report as PDF")
    If Len(path) = 0 Then Exit Sub
    RequireNewExportFile path: Set area = summary
    For Each named In ws.Names
        If Right$(named.Name, Len("_CAPReportBody")) = "_CAPReportBody" Then Set area = named.RefersToRange
    Next
    folder = Environ$("TEMP") & "\CAPPDF_" & Format$(Now, "yyyymmddhhnnss") & "_" & Format$(Timer * 100, "0")
    MkDir folder: created = True: Application.EnableEvents = False
    ws.Copy: Set temporary = ActiveWorkbook: Set copySheet = temporary.Worksheets(1)
    For i = copySheet.ChartObjects.Count To 1 Step -1: copySheet.ChartObjects(i).Delete: Next
    fitRatio = Application.Min(1, 700 / copySheet.Range(area.Address).Width)
    For col = area.Column To area.Column + area.Columns.Count - 1
        copySheet.Columns(col).ColumnWidth = copySheet.Columns(col).ColumnWidth * fitRatio
    Next
    copySheet.Range(area.Address).Font.Size = 9
    copySheet.Range("A1").Font.Size = 18: copySheet.Rows(1).RowHeight = 32
    copySheet.Range("A2:H5").WrapText = True: copySheet.Rows("2:5").RowHeight = 34
    copySheet.Range(copySheet.Cells(6, area.Column), copySheet.Cells(area.Row + area.Rows.Count - 1, area.Column + area.Columns.Count - 1)).Rows.AutoFit
    With copySheet.PageSetup
        .PrintArea = area.Address: .PrintTitleRows = "$1:$6": .Orientation = xlLandscape: .PaperSize = xlPaperLetter
        .PrintTitleColumns = "": .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .LeftMargin = 36: .RightMargin = 36: .TopMargin = 36: .BottomMargin = 36
        .LeftHeader = "CAP Index": .LeftFooter = "&8Toolkit: " & CAP_COPYRIGHT: .RightFooter = "Page &P of &N"
    End With
    For Each picture In ws.ChartObjects
        i = i + 1: imagePath = folder & "\chart" & i & ".png"
        ExportChartPNG picture.Chart, imagePath
        Set chartSheet = temporary.Worksheets.Add(After:=temporary.Worksheets(temporary.Worksheets.Count))
        chartSheet.Name = "Chart_" & i
        Set pasted = chartSheet.Shapes.AddPicture(imagePath, 0, -1, 0, 0, -1, -1)
        pasted.LockAspectRatio = -1: If pasted.Width > 680 Then pasted.Width = 680
        If pasted.Height > 440 Then pasted.Height = 440
        chartSheet.Columns("A").ColumnWidth = 100
        chartSheet.Columns("A").ColumnWidth = chartSheet.Columns("A").ColumnWidth * 700 / chartSheet.Columns("A").Width
        chartSheet.Rows("1:30").RowHeight = 15
        With chartSheet.PageSetup
            .PrintArea = "$A$1:$A$30": .PrintTitleRows = "": .PrintTitleColumns = ""
            .Orientation = xlLandscape: .PaperSize = xlPaperLetter: .Zoom = 100
            .LeftMargin = 36: .RightMargin = 36: .TopMargin = 36: .BottomMargin = 36
            .LeftFooter = "&8Toolkit: " & CAP_COPYRIGHT: .RightFooter = "Page &P of &N"
        End With
    Next
    temporary.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path, Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    temporary.Close False: Set temporary = Nothing: CleanBriefingImages folder: created = False
    ws.Parent.Activate: ws.Activate: Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    If showSummary Then MsgBox "Saved " & path & ". Charts have separate pages for readability.", vbInformation, "CAP Index"
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    On Error Resume Next
    If Not temporary Is Nothing Then temporary.Close False
    If created Then CleanBriefingImages folder
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    If Not invoked Then Err.Raise errNo, "PDF export", description
    MsgBox "Export failed (" & errNo & "): " & description, vbExclamation, "CAP Index | PDF"
End Sub

Public Sub ExportBriefingPowerPoint(Optional ByVal selectedSheets As Collection, Optional ByVal path As String = "", Optional ByVal showSummary As Boolean = True, Optional ByVal contentModes As Object, Optional ByVal sectionTitles As Object)
    Dim wb As Workbook, ws As Worksheet, summary As Range, sheets As Collection, available As Collection, answer As Variant, tokens As Variant, token As Variant
    Dim ppt As Object, deck As Object, slide As Object, shape As Object, picture As ChartObject, tempFolder As String, tempImage As String
    Dim ownApp As Boolean, tempCreated As Boolean, sectionMode As String, sectionTitle As String
    Dim i As Long, index As Long, count As Long, list As String, seen As Object, invoked As Boolean, errNo As Long, description As String, oldScreen As Boolean
    On Error GoTo failed
    invoked = selectedSheets Is Nothing: oldScreen = Application.ScreenUpdating
    If invoked Then
        Set wb = ActiveWorkbook: Set available = New Collection: Set sheets = New Collection
        For Each ws In wb.Worksheets
            If Not CAPSummaryRange(ws) Is Nothing Then available.Add ws: list = list & available.Count & " = " & ws.Name & vbCrLf
        Next
        If available.Count = 0 Then Err.Raise vbObjectError + 846, , "No CAP Index reports in the active workbook."
        MsgBox Left$(list, 1800), vbInformation, "CAP Index | Available briefing reports"
        answer = Application.InputBox("Enter report numbers separated by commas, in slide order. Summary and chart slides will be created as pictures; no source workbook is embedded.", "CAP Index | PowerPoint", "1", Type:=2)
        If VarType(answer) = vbBoolean Then Exit Sub
        tokens = Split(CStr(answer), ","): Set seen = CreateObject("Scripting.Dictionary")
        For Each token In tokens
            If Not IsNumeric(Trim$(token)) Then Err.Raise vbObjectError + 847, , "Invalid report number."
            If CDbl(token) <> Fix(CDbl(token)) Then Err.Raise vbObjectError + 847, , "Report numbers must be integers."
            index = CLng(token)
            If index < 1 Or index > available.Count Then Err.Raise vbObjectError + 847, , "Report number is outside the list."
            If Not seen.Exists(index) Then sheets.Add available(index): seen.Add index, True
        Next
    Else
        Set sheets = selectedSheets
    End If
    If sheets.Count = 0 Or sheets.Count > 20 Then Err.Raise vbObjectError + 848, , "Choose 1 to 20 reports."
    For Each ws In sheets
        RequireFreshReport ws
        If CAPSummaryRange(ws) Is Nothing Then Err.Raise vbObjectError + 848, , "Select CAP Index reports only."
    Next
    If Len(path) = 0 Then path = ChooseExportPath("pptx", "Save CAP Index briefing")
    If Len(path) = 0 Then Exit Sub
    RequireNewExportFile path
    tempFolder = Environ$("TEMP") & "\CAPBriefing_" & Format$(Now, "yyyymmddhhnnss") & "_" & Format$(Timer * 100, "0")
    MkDir tempFolder: tempCreated = True
    On Error Resume Next: Set ppt = GetObject(, "PowerPoint.Application"): Err.Clear: On Error GoTo failed
    If ppt Is Nothing Then Set ppt = CreateObject("PowerPoint.Application"): ownApp = True
    Set deck = ppt.Presentations.Add(0)
    deck.PageSetup.SlideWidth = 960: deck.PageSetup.SlideHeight = 540
    For Each ws In sheets
        sectionMode = "Both": sectionTitle = ws.Range("A1").Value2
        If Not contentModes Is Nothing Then sectionMode = contentModes(ws.Name)
        If Not sectionTitles Is Nothing Then sectionTitle = sectionTitles(ws.Name)
        If sectionMode <> "Charts" Then
        Set summary = CAPSummaryRange(ws): count = count + 1: tempImage = tempFolder & "\summary" & count & ".png"
        ExportRangePNG summary, tempImage
        AddBriefingSlide deck, tempImage, sectionTitle, BriefingFooter(ws)
        End If
        If sectionMode <> "Tables" Then
        For Each picture In ws.ChartObjects
            If Not contentModes Is Nothing Or Left$(picture.Name, 4) = "CAP_" Or Left$(picture.Name, 8) = "CAPTime_" Or Left$(picture.Name, 10) = "CAPReport_" Then
                count = count + 1: tempImage = tempFolder & "\chart" & count & ".png"
                Application.ScreenUpdating = True: ws.Activate: picture.Activate: picture.Chart.Refresh: DoEvents
                If Not picture.Chart.Export(tempImage, "PNG") Then Err.Raise vbObjectError + 849, , "Briefing chart export failed."
                AddBriefingSlide deck, tempImage, sectionTitle & " | " & picture.Chart.ChartTitle.Text, BriefingFooter(ws)
            End If
        Next
        End If
    Next
    deck.SaveAs path, 24: deck.Close: Set deck = Nothing
    If ownApp Then ppt.Quit
    Set ppt = Nothing
    CleanBriefingImages tempFolder: Application.ScreenUpdating = oldScreen
    If showSummary Then MsgBox "Saved " & path & ". Static pictures only; no source workbook is embedded. Review labels/coverage before sharing.", vbInformation, "CAP Index"
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    On Error Resume Next
    If Not deck Is Nothing Then deck.Close
    If Not ppt Is Nothing And ownApp Then ppt.Quit
    If tempCreated Then CleanBriefingImages tempFolder
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    If Not invoked Then Err.Raise errNo, "PowerPoint briefing", description
    MsgBox "Export failed (" & errNo & "): " & description, vbExclamation, "CAP Index | PowerPoint"
End Sub

Private Function BriefingFooter(ByVal ws As Worksheet) As String
    Dim cache As Range
    Set cache = SingleSiteCache(ws)
    If Not cache Is Nothing Then
        BriefingFooter = "Crime: " & ws.Range("B3").Value2 & " | " & ws.Range("E3").Value2 & " | Site: " & ws.Range("B4").Value2 & ", " & ws.Range("D4").Value2 & " | " & SitePeriodLabel(ws) & " | Approximate straight-line miles; synthetic locations excluded."
    Else
    Set cache = ReportingCache(ws)
    If Not cache Is Nothing Then
        BriefingFooter = "Crime: " & ws.Range("B3").Value2 & " | " & ws.Range("E3").Value2 & " | " & ws.Range("A5").Value2
    Else
    Set cache = TimeReportCache(ws)
    If Not cache Is Nothing Then
        BriefingFooter = "Crime: " & ws.Range("B4").Value2 & " | " & ws.Range("B5").Value2 & " | " & TimeReportPeriod(cache) & " | " & ws.Range("B14").Value2
    Else
        Set cache = ConcentrationCache(ws)
        If Not cache Is Nothing Then BriefingFooter = "Crime: " & ws.Range("B4").Value2 & " | " & ws.Range("B5").Value2 & " | Period: selected source range (dates not mapped). Selected-crime populated blocks only." Else BriefingFooter = ws.Range("A3").Value2 & " | " & ws.Range("A5").Value2
    End If
    End If
    End If
    BriefingFooter = "Toolkit: " & CAP_COPYRIGHT & vbCrLf & BriefingFooter & " | " & SavedAnalysisScope(ws)
End Function

Private Sub AddBriefingSlide(ByVal deck As Object, ByVal imagePath As String, ByVal title As String, ByVal footer As String)
    Dim slide As Object, shape As Object, ratio As Double
    Set slide = deck.Slides.Add(deck.Slides.Count + 1, 12)
    Set shape = slide.Shapes.AddShape(1, 0, 0, 960, 58): shape.Fill.ForeColor.RGB = RGB(26, 55, 82): shape.Line.Visible = 0
    Set shape = slide.Shapes.AddTextbox(1, 20, 8, 920, 44): shape.TextFrame.TextRange.Text = title
    shape.TextFrame.TextRange.Font.Size = 23: shape.TextFrame.TextRange.Font.Color.RGB = vbWhite
    Set shape = slide.Shapes.AddPicture(imagePath, 0, -1, 0, 0, -1, -1): shape.LockAspectRatio = -1
    ratio = Application.Min(920 / shape.Width, 400 / shape.Height): shape.Width = shape.Width * ratio
    shape.Left = (960 - shape.Width) / 2: shape.Top = 66 + (400 - shape.Height) / 2
    Set shape = slide.Shapes.AddTextbox(1, 20, 474, 920, 58): shape.TextFrame.TextRange.Text = footer
    shape.TextFrame.TextRange.Font.Size = 11: shape.TextFrame.TextRange.Font.Color.RGB = RGB(60, 70, 80)
End Sub

Private Sub CleanBriefingImages(ByVal folder As String)
    Dim file As String
    file = Dir$(folder & "\*.png")
    Do While Len(file) > 0: Kill folder & "\" & file: file = Dir$(): Loop
    RmDir folder
End Sub

Public Sub RequireFreshReport(ByVal ws As Worksheet)
    Dim cache As Range, i As Long
    Set cache = SingleSiteCache(ws)
    If Not cache Is Nothing Then
        If SafeText(ws.Range("B3").Value2) <> SafeText(cache.Parent.Range("R1").Value2) Or ws.Range("B4").Value2 <> cache.Parent.Range("R2").Value2 Or ws.Range("D4").Value2 <> cache.Parent.Range("R3").Value2 Then Err.Raise vbObjectError + 850, , "Single Site controls changed; click Refresh Report."
        If cache.Parent.Range("M6").Value2 Then
            If SiteReportDate(ws.Range("F4").Value2, ws.Parent.Date1904) <> cache.Parent.Range("R4").Value2 Or SiteReportDate(ws.Range("H4").Value2, ws.Parent.Date1904) <> cache.Parent.Range("R5").Value2 Then Err.Raise vbObjectError + 850, , "Single Site dates changed; click Refresh Report."
        ElseIf ws.Range("F4").Value2 <> "Not mapped" Or ws.Range("H4").Value2 <> "Not mapped" Then
            Err.Raise vbObjectError + 850, , "Dates are not mapped. Rebuild Single Site with a date field."
        End If
        Exit Sub
    End If
    If Not CAPNamedRange(ws, "_CAPCommonCrime") Is Nothing Then
        If SafeText(ws.Range("B3").Value2) <> SafeText(ws.Range("B4").Value2) Then Err.Raise vbObjectError + 850, , "Crime filter changed; click Refresh Report."
    End If
    Set cache = ReportingCache(ws)
    If Not cache Is Nothing Then
        If SafeText(ws.Range("B3").Value2) <> SafeText(cache.Parent.Range("K13").Value2) Then Err.Raise vbObjectError + 850, , "Crime filter changed. Click Refresh Report before exporting."
        Exit Sub
    End If
    Set cache = TimeReportCache(ws)
    If Not cache Is Nothing Then
        For i = 4 To 9
            If SafeText(ws.Cells(i, 2).Value2) <> SafeText(cache.Parent.Cells(i, 16).Value2) Then Err.Raise vbObjectError + 850, , "Report controls changed. Click Refresh Report before exporting."
        Next
        If Left$(SafeText(ws.Range("B14").Value2), 14) = "Refresh failed" Then Err.Raise vbObjectError + 851, , "Resolve the failed report refresh before exporting."
    Else
        Set cache = ConcentrationCache(ws)
        If Not cache Is Nothing Then
            If SafeText(ws.Range("B4").Value2) <> SafeText(ws.Range("E4").Value2) Or SafeText(ws.Range("B5").Value2) <> SafeText(ws.Range("E5").Value2) Then Err.Raise vbObjectError + 850, , "Concentration controls changed. Refresh Report before exporting."
        End If
    End If
End Sub

Public Sub ExportChartPNG(ByVal sourceChart As Chart, ByVal path As String)
    Dim ws As Worksheet, canvas As ChartObject, picture As Shape, caption As Shape, tempImage As String, oldScreen As Boolean
    Dim w As Double, h As Double, errNo As Long, description As String, created As Boolean
    oldScreen = Application.ScreenUpdating
    On Error GoTo failed
    RequireNewExportFile path
    If TypeName(sourceChart.Parent) <> "ChartObject" Then Err.Raise vbObjectError + 852, , "Select a chart embedded in a worksheet."
    Set ws = sourceChart.Parent.Parent: RequireFreshReport ws
    tempImage = Environ$("TEMP") & "\CAPChart_" & Format$(Now, "yyyymmddhhnnss") & "_" & Format$(Timer * 100, "0") & ".png"
    RequireNewExportFile tempImage
    Application.ScreenUpdating = True: ws.Activate: sourceChart.Parent.Activate: sourceChart.Refresh: DoEvents
    If Not sourceChart.Export(tempImage, "PNG") Then Err.Raise vbObjectError + 853, , "Chart image export failed."
    created = True: w = sourceChart.ChartArea.Width: h = sourceChart.ChartArea.Height
    Set canvas = ws.ChartObjects.Add(0, ws.UsedRange.Top + ws.UsedRange.Height + 20, w, h + 65)
    canvas.Chart.ChartArea.Format.Fill.ForeColor.RGB = vbWhite
    canvas.Chart.ChartArea.Border.LineStyle = xlNone
    Set picture = canvas.Chart.Shapes.AddPicture(tempImage, 0, -1, 0, 0, w, h)
    Set caption = canvas.Chart.Shapes.AddTextbox(1, 8, h + 4, w - 16, 57)
    If Not CAPSummaryRange(ws) Is Nothing Then caption.TextFrame.Characters.Text = BriefingFooter(ws) Else caption.TextFrame.Characters.Text = "Source chart: " & ws.Name
    caption.TextFrame.Characters.Font.Size = 10: caption.TextFrame.Characters.Font.Name = "Calibri"
    canvas.Activate: canvas.Chart.Refresh: DoEvents
    If Not canvas.Chart.Export(path, "PNG") Then Err.Raise vbObjectError + 853, , "Captioned PNG export failed."
    If FileLen(path) = 0 Then Err.Raise vbObjectError + 853, , "Captioned PNG is empty."
    canvas.Delete: Kill tempImage: Application.ScreenUpdating = oldScreen: Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    On Error Resume Next
    If Not canvas Is Nothing Then canvas.Delete
    If created Then Kill tempImage
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0: Err.Raise errNo, "Chart PNG", description
End Sub

Private Function TimeReportPeriod(ByVal cache As Range) As String
    If Len(SafeText(cache.Parent.Range("N5").Value2)) > 0 Then
        TimeReportPeriod = "Day/hour inputs; calendar dates not mapped"
    ElseIf cache.Parent.Range("N1").Value2 = "RepeatLocations" Then
        TimeReportPeriod = "Address counts; dates not mapped"
    ElseIf cache.Parent.Range("N2").Value2 = 0 Or cache.Parent.Range("N3").Value2 = 0 Then
        TimeReportPeriod = "No valid calendar dates in source"
    Else
        TimeReportPeriod = "Source dates: " & Format$(CDate(cache.Parent.Range("N2").Value2), "mm/dd/yyyy") & " - " & Format$(CDate(cache.Parent.Range("N3").Value2), "mm/dd/yyyy")
    End If
End Function
