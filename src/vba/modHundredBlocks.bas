Attribute VB_Name = "modHundredBlocks"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' Toolkit software only; source data and third-party material retain their ownership.
' See LICENSE and OWNERSHIP-NOTICE.txt. Retain copyright and license notices.
Option Explicit

Public Function BlockHeaders() As Variant
    BlockHeaders = Array("CAP_HundredBlockNumber", "CAP_HundredBlock", "CAP_BlockKey", "CAP_AddressPrecision", "CAP_SyntheticAddress", "CAP_SyntheticFlag", "CAP_RandomSeed", "CAP_BlockStatus")
End Function

Public Function ParseBlockAddress(ByVal raw As Variant, ByRef base As Long, ByRef street As String, ByRef precision As String, ByRef status As String, ByRef masked As Boolean) As Boolean
    Dim s As String, re As Object, match As Object, token As String, number As Long, labelled As Boolean
    base = 0: street = "": precision = "Unknown": masked = False: status = "Unsupported address; review"
    If IsError(raw) Then status = "Source cell error": Exit Function
    s = NormalText(raw)
    If Len(s) = 0 Then status = "Missing address": Exit Function
    If InStr(s, "&") > 0 Or InStr(s, " AND ") > 0 Or InStr(s, " AT ") > 0 Or InStr(s, "/") > 0 Then Exit Function
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^([0-9]{1,7}|[0-9]{0,5}XX) (.+)$"
    If Not re.Test(s) Then Exit Function
    Set match = re.Execute(s)(0): token = match.SubMatches(0): street = match.SubMatches(1)
    re.Pattern = "^(BLOCK|BLK)( OF)? "
    labelled = re.Test(street): street = re.Replace(street, "")
    If InStr(token, "XX") > 0 Then
        number = CLng(Replace(token, "XX", "00")): precision = "Hundred block (XX)": masked = True
    Else
        number = CLng(token)
        If labelled Then
            If number Mod 100 <> 0 Then status = "Block label is not a hundred boundary; review": street = "": Exit Function
            precision = "Hundred block (label)": masked = True
        ElseIf number Mod 100 = 0 Then
            precision = "Hundred block (00)": masked = True
        Else
            precision = "Street number"
        End If
    End If
    re.Pattern = "^([A-Z]|[0-9]+(ST|ND|RD|TH)\b|[0-9]+ (ST|STREET|AVE|AVENUE|RD|ROAD)\b)"
    If Not re.Test(street) Then status = "Unsupported street format; review": street = "": Exit Function
    If Left$(street, 6) = "PO BOX" Or Left$(street, 8) = "P.O. BOX" Then street = "": Exit Function
    re.Pattern = "\s+((APT|UNIT|SUITE|STE)(\s+|$)|#\s*).*$"
    street = Trim$(re.Replace(street, ""))
    If Len(street) = 0 Then Exit Function
    base = (number \ 100) * 100
    status = "Ready; street aliases not standardized"
    ParseBlockAddress = True
End Function

Public Function BlockRandomSuffix(ByVal seed As Long, ByVal stableKey As String) As Long
    Dim value As Double, i As Long, char As Long
    If seed < 1 Or seed > 2147483646 Then Err.Raise vbObjectError + 400, , "Seed must be an integer from 1 to 2147483646."
    value = seed
    For i = 1 To Len(stableKey)
        char = AscW(Mid$(stableKey, i, 1)): If char < 0 Then char = char + 65536
        value = value * 131# + char
        value = value - Fix(value / 2147483647#) * 2147483647#
    Next
    value = value * 48271#
    value = value - Fix(value / 2147483647#) * 2147483647#
    BlockRandomSuffix = CLng(value - Fix(value / 100#) * 100#)
End Function

Private Function ExistingBlockOutput(ByVal r As Range, ByVal col As Long) As Range
    Dim named As Name, marked As Range, headers As Variant, j As Long
    headers = BlockHeaders()
    For Each named In r.Parent.Names
        If InStr(1, named.Name, "_CAPBlocks_", vbTextCompare) > 0 Then
            Set marked = Nothing
            On Error Resume Next: Set marked = named.RefersToRange: On Error GoTo 0
            If Not marked Is Nothing Then
                If marked.Row = r.Row And marked.Column = col And marked.Columns.Count = 9 Then
                    For j = 0 To 7
                        If SafeText(marked.Cells(1, j + 2).Value2) <> headers(j) Then Err.Raise vbObjectError + 401, , "Generated block headers were changed; restore them before refresh."
                    Next
                    Set ExistingBlockOutput = marked.Cells(1, 2): Exit Function
                End If
            End If
        End If
    Next
End Function

Public Sub RandomizeHundredBlocks()
    BuildHundredBlocks generateSynthetic:=True
End Sub

Public Sub BuildHundredBlocksInPlace(Optional ByVal source As Range, Optional ByVal addressHeader As String = "Address", Optional ByVal cityHeader As String = "", Optional ByVal stateHeader As String = "", Optional ByVal incidentHeader As String = "", Optional ByVal generateSynthetic As Boolean = False, Optional ByVal seed As Long = 20261013, Optional ByVal showSummary As Boolean = True)
    Dim r As Range, ws As Worksheet, wb As Workbook, a As Variant, out() As Variant, n As Long, i As Long, j As Long
    Dim ac As Long, cc As Long, sc As Long, ic As Long, col As Long, base As Long, street As String, precision As String
    Dim status As String, masked As Boolean, city As String, state As String, key As String, identity As String
    Dim first As Range, lo As ListObject, pos As Long, inserted As Long, headers As Variant, marked As Name
    Dim valid As Long, synthetic As Long, invalid As Long, oldEvents As Boolean, oldScreen As Boolean, changing As Boolean
    Dim response As Variant, errNo As Long, description As String, unknownIDs As Long
    On Error GoTo failed
    Set r = source: If r Is Nothing Then Set r = PickSource()
    If r Is Nothing Then Exit Sub
    Set ws = r.Parent: Set wb = ws.Parent
    If r.Areas.Count <> 1 Or r.Rows.Count < 2 Or r.Rows.Count > 200001 Or r.Columns.Count > 200 Then Err.Raise vbObjectError + 402, , "Select one header/data rectangle, up to 200,000 rows and 200 columns."
    If wb Is ThisWorkbook Then Err.Raise vbObjectError + 403, , "Select a data workbook."
    If wb.ReadOnly Or ws.ProtectContents Then Err.Raise vbObjectError + 404, , "The worksheet must be unprotected and its workbook writable."
    If IsNull(r.MergeCells) Then Err.Raise vbObjectError + 405, , "Unmerge the source cells."
    If r.MergeCells Then Err.Raise vbObjectError + 405, , "Unmerge the source cells."
    a = r.Value2: n = UBound(a, 1) - 1
    If source Is Nothing Then
        ac = AskColumn(a, "Street address or agency hundred-block address.", "Address", False, "ADDRESS")
        cc = AskColumn(a, "Optional city for geographic block keys.", "", True, "CITY")
        sc = AskColumn(a, "Optional state for geographic block keys.", "", True, "STATE")
        If generateSynthetic Then
            ic = AskColumn(a, "Optional stable IncidentID: keeps offense rows of the same incident/block together. If absent, random values use source row number.", "", True, "INCIDENTID")
            response = Application.InputBox("Repeatable random seed (integer 1 to 2147483646). Keep the same seed to reproduce values; change it to reshuffle.", "CAP Index | Randomize Block Addresses", CStr(seed), Type:=1)
            If VarType(response) = vbBoolean Then Exit Sub
            If response < 1 Or response > 2147483646# Or response <> Fix(response) Then Err.Raise vbObjectError + 400, , "Invalid seed."
            seed = CLng(response)
        End If
    Else
        ac = HeaderColumn(a, addressHeader): cc = HeaderColumn(a, cityHeader)
        sc = HeaderColumn(a, stateHeader): ic = HeaderColumn(a, incidentHeader)
    End If
    If seed < 1 Or seed > 2147483646 Then Err.Raise vbObjectError + 400, , "Invalid seed."
    col = r.Column + ac - 1
    If col > ws.Columns.Count - 8 Then Err.Raise vbObjectError + 406, , "No room for eight block fields."
    ReDim out(1 To n, 1 To 8)
    For i = 2 To UBound(a, 1)
        city = "": state = ""
        If cc > 0 Then
            If Not IsError(a(i, cc)) Then city = NormalText(a(i, cc))
        End If
        If sc > 0 Then
            If Not IsError(a(i, sc)) Then state = NormalText(a(i, sc))
        End If
        If ParseBlockAddress(a(i, ac), base, street, precision, status, masked) Then
            valid = valid + 1
            out(i - 1, 1) = base: out(i - 1, 2) = CStr(base) & " BLOCK " & street
            key = KeyPart(state) & KeyPart(city) & KeyPart(out(i - 1, 2))
            If Len(city) > 0 And Len(state) > 0 Then
                out(i - 1, 3) = key
            Else
                status = status & "; missing city/state, no block key"
            End If
            out(i - 1, 4) = precision: out(i - 1, 6) = False
            If generateSynthetic And masked Then
                identity = ""
                If ic > 0 Then
                    If Not IsError(a(i, ic)) Then identity = SafeText(a(i, ic))
                End If
                If Len(identity) = 0 Then
                    identity = "ROW:" & r.Row + i - 1: unknownIDs = unknownIDs + 1
                Else
                    identity = "INCIDENT:" & identity
                End If
                out(i - 1, 5) = CStr(base + BlockRandomSuffix(seed, KeyPart(identity) & key)) & " " & street
                out(i - 1, 6) = True: out(i - 1, 7) = seed: synthetic = synthetic + 1
                status = "Synthetic within-block number; not an observed address. " & status
            End If
        Else
            invalid = invalid + 1: out(i - 1, 4) = precision: out(i - 1, 6) = False
        End If
        out(i - 1, 8) = status
    Next
    headers = BlockHeaders(): Set first = ExistingBlockOutput(r, col)
    On Error Resume Next: Set lo = ws.Cells(r.Row, col).ListObject: On Error GoTo failed
    If Not lo Is Nothing Then
        If r.Row <> lo.HeaderRowRange.Row Or n <> lo.ListRows.Count Or r.Column < lo.Range.Column Or r.Column + r.Columns.Count - 1 > lo.Range.Column + lo.Range.Columns.Count - 1 Then Err.Raise vbObjectError + 407, , "Select the table header and all data rows, excluding totals."
        pos = col - lo.Range.Column + 1
        If first Is Nothing Then
            For j = 0 To 7
                For i = 1 To lo.ListColumns.Count
                    If StrComp(lo.ListColumns(i).Name, headers(j), vbTextCompare) = 0 Then Err.Raise vbObjectError + 408, , "Table field already exists: " & headers(j)
                Next
            Next
        End If
    End If
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    changing = True: Application.EnableEvents = False: Application.ScreenUpdating = False
    If first Is Nothing Then
        If lo Is Nothing Then
            ws.Columns(col + 1).Resize(, 8).Insert Shift:=xlToRight: inserted = 8
        Else
            For j = 1 To 8: lo.ListColumns.Add Position:=pos + j: inserted = inserted + 1: Next
        End If
        Set first = ws.Cells(r.Row, col + 1)
        Set marked = ws.Names.Add(Name:="_CAPBlocks_" & col & "_" & r.Row, RefersTo:="='" & Replace(ws.Name, "'", "''") & "'!" & ws.Cells(r.Row, col).Resize(1, 9).Address)
        marked.Visible = False
    End If
    WriteHeaders ws, r.Row, first.Column, headers
    first.Offset(1).Resize(n, 8).NumberFormat = "@"
    first.Offset(1).Resize(n, 8).Value2 = out
    first.Offset(1).Resize(n, 1).NumberFormat = "0"
    first.Offset(1, 5).Resize(n, 1).NumberFormat = "General"
    first.Offset(1, 6).Resize(n, 1).NumberFormat = "0"
    With first.Resize(1, 8)
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(26, 55, 82): .WrapText = True
    End With
    first.Resize(n + 1, 8).Columns.AutoFit
    For j = 0 To 7
        If first.Offset(0, j).EntireColumn.ColumnWidth > 40 Then
            first.Offset(0, j).EntireColumn.ColumnWidth = 40: first.Offset(0, j).Resize(n + 1).WrapText = True
        End If
    Next
    first.Resize(n + 1, 8).Rows.AutoFit
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen: changing = False
    If showSummary Then MsgBox "Hundred-block fields updated." & vbCrLf & "Valid rows: " & valid & " | Review rows: " & invalid & vbCrLf & "Synthetic rows: " & synthetic & " | Seed: " & seed & vbCrLf & "Synthetic rows using row numbers rather than IncidentID: " & unknownIDs & vbCrLf & "00/XX endings are treated as agency hundred-block input. Synthetic numbers range from the block base through base + 99. They do not recover the original incident address.", vbInformation, "CAP Index"
    Exit Sub
failed:
    errNo = Err.Number: description = Err.Description
    On Error Resume Next
    If inserted > 0 Then
        If lo Is Nothing Then
            ws.Columns(col + 1).Resize(, inserted).Delete
        Else
            For j = 1 To inserted: lo.ListColumns(pos + 1).Delete: Next
        End If
        If Not marked Is Nothing Then marked.Delete
    End If
    If changing Then Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    If Not source Is Nothing Then Err.Raise errNo, "Hundred-Block Builder", description
    MsgBox "Hundred-Block Builder: " & description, vbExclamation, "CAP Index"
End Sub
