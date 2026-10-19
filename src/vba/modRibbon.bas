Attribute VB_Name = "modRibbon"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit
Public Sub CAP_OnLoad(ByVal ribbon As Office.IRibbonUI)
    Set CAPRibbon = ribbon
    InitializeCAPEvents
End Sub
Public Sub CAP_OnAction(ByVal control As Office.IRibbonControl)
    Dim wbBefore As Workbook, countBefore As Long
    On Error GoTo failed
    Set wbBefore = ActiveWorkbook
    If Not wbBefore Is Nothing Then countBefore = wbBefore.Worksheets.Count
    Application.StatusBar = "CAP Index: processing selected tool..."
    Select Case control.Id
        Case "capDateTime": BuildDateTime
        Case "capHour": BuildHourFields
        Case "capBlocks": BuildHundredBlocks
        Case "capBlockRandom": RandomizeHundredBlocks
        Case "capReview": ReviewCases
        Case "capOwnership": ShowCAPOwnership
        Case "capPareto": BuildConcentrationReport
        Case "capRefresh": UnifiedRefresh
        Case "capReset": ResetReportFilters
        Case "capSingleSite": BuildSingleSite
        Case "capHeatmap": BuildTemporalHeatmap
        Case "capRepeat": BuildRepeatLocations
        Case "capHotStreets": BuildHotStreets
        Case "capCensus": GeocodeCensus
        Case "capTrends": BuildTrendReport
        Case "capCompare": BuildPeriodComparison
        Case "capProfile": BuildCrimeProfile
        Case "capYoY": BuildYearOverYear
        Case "capSeasonality": BuildSeasonality
        Case "capCrosswalk": PrepareNIBRSCrosswalk
        Case "capNIBRS": ApplyNIBRSCrosswalk
        Case "capReportBuilder": CreateReportBuilder
        Case "capExportSelected": ExportSelectedReport
        Case "capSaveImage": SaveReportImage
        Case "capPowerPoint": ExportBriefingPowerPoint
        Case "capPDF": ExportReportPDF
        Case "capCheck": BuildDatasetCheck
        Case "capHelp": BuildToolGuide
        Case Else: Err.Raise vbObjectError + 930, , "Unrecognized command: " & control.Id
    End Select
    Application.StatusBar = False
    If control.Id = "capReview" Then CompleteCAPCommand wbBefore, countBefore
    Exit Sub
failed:
    Application.StatusBar = False
    MsgBox "Command failed (" & Err.Number & "): " & Err.Description, vbExclamation, "CAP Index " & CAP_VERSION
End Sub

Public Sub CAP100_OnLoad(ByVal ribbon As Office.IRibbonUI)
    Set CAPRibbon = ribbon
    InitializeCAPEvents
End Sub
Public Sub CAP100_OnAction(ByVal control As Office.IRibbonControl)
    CAP_OnAction control
End Sub
Public Sub CAP100_DatasetCheck(ByVal control As Office.IRibbonControl)
    On Error GoTo failed
    BuildDatasetCheck
    Exit Sub
failed:
    MsgBox "Dataset Check failed (" & Err.Number & "): " & Err.Description, vbExclamation, "CAP Index " & CAP_VERSION
End Sub
