Attribute VB_Name = "modCensusTransport"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit
Public Const CENSUS_BENCHMARK As String = "Public_AR_Current"
Public Const CENSUS_MAX_ROWS As Long = 10000
Public Const CENSUS_MAX_BYTES As Long = 4900000

Public Function CsvQuote(ByVal s As String) As String
    CsvQuote = Chr$(34) & Replace(s, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Public Function ReadCSV(ByVal content As String) As Collection
    Dim records As New Collection, fields As Collection, field As String, quoted As Boolean
    Dim i As Long, ch As String, row As Variant, j As Long, pending As Boolean
    Set fields = New Collection
    For i = 1 To Len(content)
        ch = Mid$(content, i, 1): pending = True
        If ch = Chr$(34) Then
            If quoted And i < Len(content) Then
                If Mid$(content, i + 1, 1) = Chr$(34) Then
                    field = field & Chr$(34): i = i + 1
                Else
                    quoted = False
                End If
            Else
                quoted = Not quoted
            End If
        ElseIf ch = "," And Not quoted Then
            fields.Add field: field = ""
        ElseIf (ch = vbCr Or ch = vbLf) And Not quoted Then
            fields.Add field: field = ""
            ReDim row(0 To fields.Count - 1)
            For j = 1 To fields.Count: row(j - 1) = fields(j): Next
            If fields.Count > 1 Or Len(row(0)) > 0 Then records.Add row
            Set fields = New Collection: pending = False
            If ch = vbCr And i < Len(content) Then
                If Mid$(content, i + 1, 1) = vbLf Then i = i + 1
            End If
        Else
            field = field & ch
        End If
    Next
    If quoted Then Err.Raise vbObjectError + 300, , "Census returned an unterminated quoted CSV field."
    If pending Then
        fields.Add field: ReDim row(0 To fields.Count - 1)
        For j = 1 To fields.Count: row(j - 1) = fields(j): Next
        If fields.Count > 1 Or Len(row(0)) > 0 Then records.Add row
    End If
    Set ReadCSV = records
End Function

Public Function UTF8Bytes(ByVal text As String) As Variant
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText text: stream.Position = 0: stream.Type = 1: stream.Position = 3
    UTF8Bytes = stream.Read: stream.Close
End Function

Public Function UTF8Size(ByVal text As String) As Long
    Dim bytes As Variant: bytes = UTF8Bytes(text)
    UTF8Size = UBound(bytes) - LBound(bytes) + 1
End Function

Private Sub ResponsivePause(ByVal seconds As Double)
    Dim began As Date: began = Now
    Do While (Now - began) * 86400# < seconds: DoEvents: Loop
End Sub

Public Function CensusPost(ByVal csv As String) As String
    Dim http As Object, boundary As String, body As String, bytes As Variant, attempt As Long
    Dim began As Date, completed As Boolean, errNo As Long, description As String, code As Long
    boundary = "CAPIndexCensusBoundary0x3"
    body = "--" & boundary & vbCrLf & "Content-Disposition: form-data; name=""benchmark""" & vbCrLf & vbCrLf & CENSUS_BENCHMARK & vbCrLf
    body = body & "--" & boundary & vbCrLf & "Content-Disposition: form-data; name=""addressFile""; filename=""addresses.csv""" & vbCrLf & "Content-Type: text/csv; charset=utf-8" & vbCrLf & vbCrLf & csv & vbCrLf & "--" & boundary & "--" & vbCrLf
    bytes = UTF8Bytes(body)
    For attempt = 1 To 3
        errNo = 0: description = "": completed = False
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        http.SetTimeouts 30000, 30000, 60000, 180000
        http.Open "POST", "https://geocoding.geo.census.gov/geocoder/locations/addressbatch", True
        http.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
        On Error Resume Next
        http.Send bytes
        errNo = Err.Number: description = Err.Description: Err.Clear
        On Error GoTo 0
        began = Now
        If errNo = 0 Then
            Do
                On Error Resume Next
                completed = http.WaitForResponse(0.1)
                errNo = Err.Number: description = Err.Description: Err.Clear
                On Error GoTo 0
                DoEvents
                If completed Or errNo <> 0 Then Exit Do
                If (Now - began) * 86400# > 240 Then
                    http.Abort: errNo = vbObjectError + 301: description = "Census batch timed out.": Exit Do
                End If
            Loop
        End If
        If errNo = 18 Then Err.Raise 18, , "Geocoding cancelled."
        If errNo = 0 Then
            code = http.Status
            If code = 200 Then CensusPost = http.ResponseText: Exit Function
            description = "Census HTTP " & code
            If code <> 429 And code < 500 Then Err.Raise vbObjectError + 302, , description
        End If
        If attempt < 3 Then ResponsivePause 2 ^ attempt
    Next
    Err.Raise vbObjectError + 303, , description & "; three attempts failed."
End Function

Public Function CensusResponse(ByVal content As String, ByVal expected As Object) As Object
    Dim records As Collection, results As Object, item As Variant, id As String, status As String, kind As String
    Dim coordinates As Variant, longitude As Double, latitude As Double, value As Variant
    Set records = ReadCSV(content): Set results = CreateObject("Scripting.Dictionary")
    For Each item In records
        If UBound(item) < 2 Then Err.Raise vbObjectError + 304, , "Malformed Census CSV response."
        id = item(0)
        If Not expected.Exists(id) Or results.Exists(id) Then Err.Raise vbObjectError + 305, , "Unknown or repeated request ID in Census response."
        status = UCase$(Replace(Replace(Trim$(item(2)), " ", ""), "_", ""))
        value = Array("", Empty, Empty, "", "", "")
        Select Case status
            Case "MATCH"
                If UBound(item) < 5 Then Err.Raise vbObjectError + 306, , "Matched Census record lacks address or coordinates."
                coordinates = Split(item(5), ",")
                If UBound(coordinates) <> 1 Then Err.Raise vbObjectError + 307, , "Malformed Census coordinates."
                If Not InvariantCoordinate(coordinates(0), longitude) Or Not InvariantCoordinate(coordinates(1), latitude) Then Err.Raise vbObjectError + 308, , "Invalid Census coordinate values."
                If longitude < -180 Or longitude > 180 Or latitude < -90 Or latitude > 90 Then Err.Raise vbObjectError + 309, , "Census coordinates outside valid ranges."
                kind = UCase$(Replace(Replace(Trim$(item(3)), "-", "_"), " ", "_"))
                If kind <> "EXACT" And kind <> "NON_EXACT" Then Err.Raise vbObjectError + 310, , "Unknown Census match type."
                If Len(Trim$(item(4))) = 0 Then Err.Raise vbObjectError + 311, , "Matched address is missing."
                value(0) = item(4): value(1) = latitude: value(2) = longitude
                value(3) = "Match": value(4) = IIf(kind = "EXACT", "Exact", "Non-Exact")
                value(5) = IIf(kind = "EXACT", "", "Review non-exact match")
            Case "NOMATCH": value(3) = "No Match": value(5) = "Review unmatched address"
            Case "TIE": value(3) = "Tie": value(5) = "Review ambiguous address"
            Case Else: Err.Raise vbObjectError + 312, , "Unknown Census match status."
        End Select
        results.Add id, value
    Next
    If results.Count <> expected.Count Then Err.Raise vbObjectError + 313, , "Census response is missing request IDs; batch results were not applied."
    Set CensusResponse = results
End Function

Public Function InvariantCoordinate(ByVal text As String, ByRef number As Double) As Boolean
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^-?\d+(\.\d+)?$"
    text = Trim$(text)
    If Not re.Test(text) Then Exit Function
    number = Val(text): InvariantCoordinate = True
End Function
