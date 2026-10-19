Attribute VB_Name = "modReportFilters"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
Option Explicit

Public Function ReportingCache(ByVal ws As Worksheet) As Range
    Dim named As Name, r As Range
    On Error Resume Next
    For Each named In ws.Names
        If Right$(named.Name, Len("_CAPReportingData")) = "_CAPReportingData" Then
            Set r = named.RefersToRange
            If Not r Is Nothing Then
                If r.Columns.Count = 6 And r.Parent.Parent Is ws.Parent Then Set ReportingCache = r
            End If
            Exit For
        End If
    Next
    On Error GoTo 0
End Function

Public Sub AttachReportingFilter(ByVal ws As Worksheet, ByVal ctx As Object, ByVal kind As String, ByVal unit As String, Optional ByVal aStart As Double = 0, Optional ByVal aEnd As Double = 0)
    Dim cache As Range, store As Worksheet, data As Variant, types As Object, i As Long, key As Variant, named As Name
    Dim oldEvents As Boolean, errNo As Long, description As String
    oldEvents = Application.EnableEvents
    On Error GoTo failed
    Application.EnableEvents = False
    Set cache = ReportingCache(ws)
    If cache Is Nothing Then
        Set store = ws.Parent.Worksheets.Add(After:=ws.Parent.Worksheets(ws.Parent.Worksheets.Count))
        store.Name = "_CAPR_" & ws.Name
        data = ctx("Data")
        store.Range("A1").Resize(UBound(data, 1), 6).NumberFormat = "@"
        store.Range("A1").Resize(UBound(data, 1), 6).Value2 = data
        Set named = ws.Names.Add(Name:="_CAPReportingData", RefersTo:="='" & store.Name & "'!" & store.Range("A1").Resize(UBound(data, 1), 6).Address): named.Visible = False
        i = 0
        For Each key In Array("Start", "End", "SourceStart", "SourceEnd", "IDMapped", "MissingDates", "LocationMapped", "Measure")
            i = i + 1: store.Cells(i, 10).Value2 = key: store.Cells(i, 11).Value2 = ctx(key)
        Next
        store.Range("K9").Value2 = kind: store.Range("K10").Value2 = unit
        store.Range("K11").Value2 = aStart: store.Range("K12").Value2 = aEnd
        Set types = CreateObject("Scripting.Dictionary"): types.Add "All crime types", True
        For i = 2 To UBound(data, 1)
            If Not types.Exists(data(i, 2)) Then types.Add data(i, 2), True
        Next
        i = 0: store.Columns("H").NumberFormat = "@"
        For Each key In types.Keys: i = i + 1: store.Cells(i, 8).Value2 = key: Next
        Set named = ws.Names.Add(Name:="_CAPReportingTypes", RefersTo:="='" & store.Name & "'!" & store.Range("H1:H" & i).Address): named.Visible = False
        ws.Activate: store.Visible = xlSheetVeryHidden
    Else
        Set store = cache.Parent
    End If
    store.Range("K13").Value2 = ctx("Crime")
    ws.Range("A3:H3").UnMerge: ws.Range("A3:H3").ClearContents
    ws.Range("A3").Value2 = "Crime type" & vbCrLf & "Click the blue field": ws.Range("B3:D3").Merge: ws.Range("B3").Value2 = ctx("Crime")
    ws.Range("B3:D3").Interior.Color = RGB(224, 239, 250)
    With ws.Range("B3").Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=_CAPReportingTypes"
        .InCellDropdown = True: .ShowError = True
    End With
    ws.Range("E3:H3").Merge
    ws.Range("E3").Value2 = ctx("Measure") & " | " & Format$(CDate(ctx("Start")), "mm/dd/yyyy") & " - " & Format$(CDate(ctx("End")), "mm/dd/yyyy")
    If kind = "Comparison" Then ws.Range("E3").Value2 = ctx("Measure") & " | A: " & Format$(CDate(aStart), "mm/dd/yyyy") & " - " & Format$(CDate(aEnd), "mm/dd/yyyy") & " | B: " & Format$(CDate(ctx("Start")), "mm/dd/yyyy") & " - " & Format$(CDate(ctx("End")), "mm/dd/yyyy")
    ws.Range("A3:H3").WrapText = True: ws.Rows(3).RowHeight = 48
    ws.Range("A5").Value2 = ws.Range("A5").Value2 & " Crime dropdown filters the saved snapshot; rebuild after source edits."
    InitializeCAPEvents
    StampReportStatus ws
    ws.Activate: ws.Range("B3").Select
    Application.EnableEvents = oldEvents
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    Application.EnableEvents = oldEvents
    Err.Raise errNo, "Report filter setup", description
End Sub

Public Sub RefreshReportingFilter(ByVal ws As Worksheet, Optional ByVal propagateErrors As Boolean = False)
    Dim cache As Range, ctx As Object, i As Long, data As Variant, oldEvents As Boolean, crime As String, found As Boolean
    Dim errNo As Long, description As String
    oldEvents = Application.EnableEvents
    On Error GoTo failed
    Set cache = ReportingCache(ws): If cache Is Nothing Then Exit Sub
    crime = SafeText(ws.Range("B3").Value2): data = cache.Value2
    found = crime = "All crime types"
    For i = 2 To UBound(data, 1): If data(i, 2) = crime Then found = True
    Next
    If Not found Then Err.Raise vbObjectError + 890, , "Choose a crime from the dropdown."
    Set ctx = CreateObject("Scripting.Dictionary")
    ctx.Add "Data", data: ctx.Add "Source", cache: ctx.Add "Crime", crime
    For i = 1 To 8: ctx.Add cache.Parent.Cells(i, 10).Value2, cache.Parent.Cells(i, 11).Value2: Next
    Application.EnableEvents = False
    If cache.Parent.Range("K9").Value2 = "Trend" Then
        BuildTrendReport source:=cache, unit:=cache.Parent.Range("K10").Value2, showSummary:=False, reportTarget:=ws, reportContext:=ctx
    ElseIf cache.Parent.Range("K9").Value2 = "YoY" Or cache.Parent.Range("K9").Value2 = "Seasonality" Then
        BuildCalendarReport source:=cache, kind:=cache.Parent.Range("K9").Value2, unit:=cache.Parent.Range("K10").Value2, showSummary:=False, reportTarget:=ws, reportContext:=ctx
    ElseIf cache.Parent.Range("K9").Value2 = "Profile" Then
        BuildCrimeProfile source:=cache, showSummary:=False, reportTarget:=ws, reportContext:=ctx
    Else
        BuildPeriodComparison source:=cache, aStart:=cache.Parent.Range("K11").Value2, aEnd:=cache.Parent.Range("K12").Value2, showSummary:=False, reportTarget:=ws, reportContext:=ctx
    End If
    Application.EnableEvents = oldEvents
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    Application.EnableEvents = oldEvents
    If propagateErrors Then Err.Raise errNo, "Report filter", description
    MsgBox description, vbExclamation, "CAP Index | Report filter"
End Sub

Public Sub ClearReportingResults(ByVal ws As Worksheet)
    Dim i As Long
    On Error Resume Next: ws.AutoFilterMode = False: On Error GoTo 0
    For i = ws.ChartObjects.Count To 1 Step -1
        If Left$(ws.ChartObjects(i).Name, 10) = "CAPReport_" Then ws.ChartObjects(i).Delete
    Next
    ws.Range("A3:H5").UnMerge
    ws.Range("A3:H" & Application.Max(7, ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1)).Clear
End Sub
