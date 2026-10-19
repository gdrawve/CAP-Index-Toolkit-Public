VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsCensusFixture"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public CallCount As Long
Public FailAt As Long
Public MaxSeen As Long
Public Captured As String
Public Function PostBatch(ByVal csv As String) As String
    Dim records As Collection, row As Variant, i As Long, result As String, status As String, kind As String
    CallCount = CallCount + 1: Captured = Captured & csv
    If FailAt > 0 And CallCount = FailAt Then Err.Raise vbObjectError + 350, , "Synthetic batch failure"
    Set records = ReadCSV(csv)
    If records.Count > MaxSeen Then MaxSeen = records.Count
    ' Reverse response order deliberately: production must join by request ID.
    For i = records.Count To 1 Step -1
        row = records(i): status = "Match": kind = "Exact"
        If InStr(row(1), "MISSING") > 0 Then status = "No_Match"
        If InStr(row(1), "TIE") > 0 Then status = "Tie"
        If InStr(row(1), "NONEXACT") > 0 Then kind = "Non_Exact"
        result = result & CsvQuote(row(0)) & "," & CsvQuote(row(1)) & "," & status
        If status = "Match" Then result = result & "," & kind & "," & CsvQuote(UCase$(row(1)) & ", WASHINGTON, DC, 20233") & "," & CsvQuote("-76.927487,38.846016") & ",12345,L"
        result = result & vbCrLf
    Next
    PostBatch = result
End Function
