VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsCAPEvents"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents ExcelApp As Excel.Application
Private Sub ExcelApp_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    Dim oldEvents As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo failed
    If Not SingleSiteCache(Sh) Is Nothing Then
        If Not Intersect(Target, Sh.Range("B3:D3,B4,D4,F4,H4")) Is Nothing Then RefreshSingleSite Sh
    ElseIf Not ReportingCache(Sh) Is Nothing Then
        If Not Intersect(Target, Sh.Range("B3:D3")) Is Nothing Then RefreshReportingFilter Sh
    ElseIf Not TimeReportCache(Sh) Is Nothing Or Not ConcentrationCache(Sh) Is Nothing Then
        If Not Intersect(Target, Sh.Range("B3:D3")) Is Nothing Then
            Application.EnableEvents = False: Sh.Range("B4").Value2 = Sh.Range("B3").Value2: Application.EnableEvents = oldEvents
            UnifiedRefresh Sh, False
        ElseIf Not Intersect(Target, Sh.Range("B4:B9")) Is Nothing Then
            If Not TimeReportCache(Sh) Is Nothing Then RefreshTimeReport Sh Else RefreshConcentration Sh
        End If
    End If
    Exit Sub
failed:
    Application.EnableEvents = oldEvents
    MsgBox Err.Description, vbExclamation, "CAP Index | Report controls"
End Sub

Private Sub ExcelApp_WorkbookActivate(ByVal Wb As Workbook)
    On Error Resume Next
    If Not CAPRibbon Is Nothing Then CAPRibbon.Invalidate
End Sub
