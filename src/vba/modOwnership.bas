Attribute VB_Name = "modOwnership"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit
Public Const CAP_OWNER As String = "CAP Index"
Public Const CAP_COPYRIGHT As String = "Copyright (c) 2026 CAP Index. Licensed under the MIT License."
Public Const CAP_ATTRIBUTION As String = "Toolkit: Copyright (c) 2026 CAP Index. Source data retains its original ownership."

Public Sub StampCAPAddinMetadata()
    ThisWorkbook.BuiltinDocumentProperties("Author").Value = CAP_OWNER
    ThisWorkbook.BuiltinDocumentProperties("Title").Value = "CAP Index Crime Analysis Excel Toolkit"
    ThisWorkbook.BuiltinDocumentProperties("Comments").Value = CAP_ATTRIBUTION
    ThisWorkbook.CustomDocumentProperties.Add "CAP Copyright", False, 4, CAP_COPYRIGHT
    ThisWorkbook.CustomDocumentProperties.Add "CAP Publisher", False, 4, CAP_OWNER
End Sub

Public Sub ShowCAPOwnership()
    frmCAPAbout.Show
End Sub

Public Function RunOwnershipTests() As String
    Dim wb As Workbook, ws As Worksheet, output As Worksheet
    Dim errNo As Long, description As String
    On Error GoTo failed
    If ThisWorkbook.BuiltinDocumentProperties("Author").Value <> CAP_OWNER Then Err.Raise vbObjectError + 880, , "Add-in author metadata missing"
    If ThisWorkbook.CustomDocumentProperties("CAP Copyright").Value <> CAP_COPYRIGHT Then Err.Raise vbObjectError + 881, , "Add-in copyright metadata missing"
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    ws.Range("A1").Value2 = "Address": ws.Range("A2").Value2 = "123 MAIN ST"
    Set output = NewOutput(wb, "Ownership_Test", ws.Range("A1:A2"), "Synthetic fixture")
    If InStr(output.Range("A2").Value2, CAP_ATTRIBUTION) = 0 Then Err.Raise vbObjectError + 882, , "Working output attribution missing"
    wb.Close False
    RunOwnershipTests = "Ownership tests passed: add-in author/copyright metadata and working output attribution"
    Exit Function
failed:
    errNo = Err.Number: description = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    Err.Raise errNo, "Ownership tests", description
End Function
