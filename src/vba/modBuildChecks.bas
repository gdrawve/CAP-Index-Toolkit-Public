Attribute VB_Name = "modBuildChecks"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit
Public Function RunBuildCheck(ByVal checkName As String, Optional ByVal previewDir As String = "") As String
    On Error GoTo failed
    If Len(previewDir) > 0 Then
        RunBuildCheck = Application.Run("'" & ThisWorkbook.Name & "'!" & checkName, previewDir)
    Else
        RunBuildCheck = Application.Run("'" & ThisWorkbook.Name & "'!" & checkName)
    End If
    Exit Function
failed:
    RunBuildCheck = "FAILED: " & checkName & " | " & Err.Description
End Function
