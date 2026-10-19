param($Excel,$Book,[Parameter(Mandatory=$true)][string]$CachePath)
$ErrorActionPreference = 'Stop'
[void](New-Item -ItemType Directory -Path $CachePath -Force)
$bootstrap = $Book.VBProject.VBComponents.Add(1)
$bootstrap.Name = 'CAPBuildEnvironmentOnly'
try {
    $bootstrap.CodeModule.AddFromString(@'
Option Explicit
#If VBA7 Then
Private Declare PtrSafe Function SetEnvironmentVariableW Lib "kernel32" (ByVal lpName As LongPtr, ByVal lpValue As LongPtr) As Long
#Else
Private Declare Function SetEnvironmentVariableW Lib "kernel32" (ByVal lpName As Long, ByVal lpValue As Long) As Long
#End If
Public Sub ConfigureCAPBuildTemp(ByVal folder As String)
    Dim variableName As String
    variableName = "TEMP"
    If SetEnvironmentVariableW(StrPtr(variableName), StrPtr(folder)) = 0 Then Err.Raise 5, , "Could not set build TEMP."
    variableName = "TMP"
    If SetEnvironmentVariableW(StrPtr(variableName), StrPtr(folder)) = 0 Then Err.Raise 5, , "Could not set build TMP."
End Sub
'@)
    $null = $Excel.Run("'" + $Book.Name + "'!ConfigureCAPBuildTemp", $CachePath)
} finally {
    $Book.VBProject.VBComponents.Remove($bootstrap)
}
# The helper is removed before importing release code. Changes affect this Excel
# process only. No machine/user environment or Office security policy is changed.
