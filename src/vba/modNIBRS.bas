Attribute VB_Name = "modNIBRS"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
' FBI reference names/codes retain their government provenance; CAP claims no ownership of them.
Option Explicit
Public Const NIBRS_REFERENCE As String = "FBI NIBRS User Manual 2025.0 (06/23/2025), section 2.3"
Public Const NIBRS_URL As String = "https://le.fbi.gov/file-repository/nibrs-user-manual-2025-0-062625.pdf"

Public Function NIBRSCatalog() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    NCode d, "720", "Animal Cruelty", "Society"
    NCode d, "200", "Arson", "Property"
    NCode d, "13A", "Aggravated Assault", "Person"
    NCode d, "13B", "Simple Assault", "Person"
    NCode d, "13C", "Intimidation", "Person"
    NCode d, "510", "Bribery", "Property"
    NCode d, "220", "Burglary/Breaking & Entering", "Property"
    NCode d, "58A", "Import Violations", "Society", True
    NCode d, "58B", "Export Violations", "Society", True
    NCode d, "61A", "Federal Liquor Offenses", "Society", True
    NCode d, "61B", "Federal Tobacco Offenses", "Society", True
    NCode d, "620", "Wildlife Trafficking", "Society", True
    NCode d, "250", "Counterfeiting/Forgery", "Property"
    NCode d, "290", "Destruction/Damage/Vandalism of Property", "Property"
    NCode d, "35A", "Drug/Narcotic Violations", "Society"
    NCode d, "35B", "Drug Equipment Violations", "Society"
    NCode d, "270", "Embezzlement", "Property"
    NCode d, "103", "Espionage", "Society", True
    NCode d, "210", "Extortion/Blackmail", "Property"
    NCode d, "26A", "False Pretenses/Swindle/Confidence Game", "Property"
    NCode d, "26B", "Credit Card/Automated Teller Machine Fraud", "Property"
    NCode d, "26C", "Impersonation", "Property"
    NCode d, "26D", "Welfare Fraud", "Property"
    NCode d, "26E", "Wire Fraud", "Property"
    NCode d, "26F", "Identity Theft", "Property"
    NCode d, "26G", "Hacking/Computer Invasion", "Property"
    NCode d, "26H", "Money Laundering", "Society", True
    NCode d, "49A", "Harboring Escapee/Concealing from Arrest", "Society", True
    NCode d, "49B", "Flight to Avoid Prosecution", "Society", True
    NCode d, "49C", "Flight to Avoid Deportation", "Society", True
    NCode d, "39A", "Betting/Wagering", "Society"
    NCode d, "39B", "Operating/Promoting/Assisting Gambling", "Society"
    NCode d, "39C", "Gambling Equipment Violations", "Society"
    NCode d, "39D", "Sports Tampering", "Society"
    NCode d, "09A", "Murder and Nonnegligent Manslaughter", "Person"
    NCode d, "09B", "Negligent Manslaughter", "Person"
    NCode d, "09C", "Justifiable Homicide", "Not a Crime"
    NCode d, "64A", "Human Trafficking, Commercial Sex Acts", "Person"
    NCode d, "64B", "Human Trafficking, Involuntary Servitude", "Person"
    NCode d, "30A", "Illegal Entry into the United States", "Society", True
    NCode d, "30B", "False Citizenship", "Society", True
    NCode d, "30C", "Smuggling Aliens", "Society", True
    NCode d, "30D", "Re-entry after Deportation", "Society", True
    NCode d, "100", "Kidnapping/Abduction", "Person"
    NCode d, "23A", "Pocket-picking", "Property"
    NCode d, "23B", "Purse-snatching", "Property"
    NCode d, "23C", "Shoplifting", "Property"
    NCode d, "23D", "Theft From Building", "Property"
    NCode d, "23E", "Theft From Coin-Operated Machine or Device", "Property"
    NCode d, "23F", "Theft From Motor Vehicle", "Property"
    NCode d, "23G", "Theft of Motor Vehicle Parts or Accessories", "Property"
    NCode d, "23H", "All Other Larceny", "Property"
    NCode d, "240", "Motor Vehicle Theft", "Property"
    NCode d, "370", "Pornography/Obscene Material", "Society"
    NCode d, "40A", "Prostitution", "Society"
    NCode d, "40B", "Assisting or Promoting Prostitution", "Society"
    NCode d, "40C", "Purchasing Prostitution", "Society"
    NCode d, "120", "Robbery", "Property"
    NCode d, "11A", "Rape", "Person"
    NCode d, "11B", "Sodomy", "Person", False, "A", "Legacy code converts to 11A; review"
    NCode d, "11C", "Sexual Assault With An Object", "Person", False, "A", "Legacy code converts to 11A; review"
    NCode d, "11D", "Criminal Sexual Contact", "Person"
    NCode d, "36A", "Incest", "Person"
    NCode d, "36B", "Statutory Rape", "Person"
    NCode d, "360", "Failure to Register as a Sex Offender", "Society", True
    NCode d, "280", "Stolen Property Offenses", "Property"
    NCode d, "101", "Treason", "Society", True
    NCode d, "520", "Weapon Law Violations", "Society"
    NCode d, "521", "Violation of National Firearm Act of 1934", "Society", True
    NCode d, "522", "Weapons of Mass Destruction", "Society", True
    NCode d, "526", "Explosives", "Society", True
    NCode d, "90K", "Failure to Appear", "", True, "B"
    NCode d, "90B", "Curfew/Loitering/Vagrancy Violations", "", False, "B"
    NCode d, "90C", "Disorderly Conduct", "", False, "B"
    NCode d, "90D", "Driving Under the Influence", "", False, "B"
    NCode d, "90F", "Family Offenses, Nonviolent", "", False, "B"
    NCode d, "90L", "Federal Resource Violations", "", True, "B"
    NCode d, "90G", "Liquor Law Violations", "", False, "B"
    NCode d, "90M", "Perjury", "", True, "B"
    NCode d, "90J", "Trespass of Real Property", "", False, "B"
    NCode d, "90Z", "All Other Offenses", "", False, "B"
    Set NIBRSCatalog = d
End Function

Private Sub NCode(ByVal catalog As Object, ByVal code As String, ByVal label As String, ByVal against As String, Optional ByVal restricted As Boolean = False, Optional ByVal group As String = "A", Optional ByVal note As String = "")
    catalog.Add code, Array(label, group, against, restricted, note)
End Sub

Public Sub PrepareNIBRSCrosswalk(Optional ByVal source As Range, Optional ByVal crimeHeader As String = "CrimeType", Optional ByVal agencyHeader As String = "", Optional ByVal allowFederal As Boolean = False, Optional ByVal showSummary As Boolean = True)
    Dim r As Range, a As Variant, ws As Worksheet, reference As Worksheet, catalog As Object, pairs As Object, i As Long, cc As Long, ac As Long, rowNo As Long
    Dim agency As String, label As String, key As Variant, code As String, k As Variant, values As Variant, choices As Variant
    On Error GoTo failed
    If source Is Nothing Then OpenToolSetup "capCrosswalk": Exit Sub
    Set r = source
    If r Is Nothing Then Set r = PickSource("NIBRS crosswalk: select headers/data; map Crime Type or supplied NIBRS code and optional Agency. Proposals require analyst approval. Original labels are retained.")
    If r Is Nothing Then Exit Sub
    ValidateOutputSource r: a = ScopedValues(r)
    If source Is Nothing Then
        cc = AskColumn(a, "Crime type or supplied NIBRS code.", "CrimeType", False, "CRIME")
        ac = AskColumn(a, "Optional Agency to keep agency-specific crosswalks separate.", "", True, "AGENCY")
        allowFederal = MsgBox("Is this crosswalk for federal/tribal reporting? Select No for local/state agencies.", vbYesNo + vbQuestion, "CAP Index | NIBRS applicability") = vbYes
    Else
        cc = HeaderColumn(a, crimeHeader): ac = HeaderColumn(a, agencyHeader)
    End If
    Set catalog = NIBRSCatalog(): Set pairs = CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(a, 1)
        agency = "": If ac > 0 Then agency = NormalText(a(i, ac))
        label = NormalText(a(i, cc)): key = KeyPart(agency) & KeyPart(label)
        If Not pairs.Exists(key) Then pairs.Add key, Array(agency, label)
    Next
    Set reference = NewOutput(r.Parent.Parent, "NIBRS_Reference", r, NIBRS_REFERENCE & " | " & NIBRS_URL)
    WriteHeaders reference, 6, 1, Array("Code", "Offense", "Group", "Crime against", "Federal/tribal only", "Note")
    rowNo = 7: reference.Columns("A:F").NumberFormat = "@"
    For Each k In catalog.Keys
        values = catalog(k): reference.Cells(rowNo, 1).Value2 = k
        reference.Cells(rowNo, 2).Value2 = values(0): reference.Cells(rowNo, 3).Value2 = values(1): reference.Cells(rowNo, 4).Value2 = values(2)
        reference.Cells(rowNo, 5).Value2 = IIf(values(3), "Yes", "No"): reference.Cells(rowNo, 6).Value2 = values(4)
        If values(1) = "B" Then reference.Cells(rowNo, 6).Value2 = "Group B collects arrest data only; do not mix with Group A offense counts"
        rowNo = rowNo + 1
    Next
    FinishOutput reference, rowNo - 1, 6
    Set ws = NewOutput(r.Parent.Parent, "NIBRS_Crosswalk", r, "Fill/verify Code and set Approved=Yes only after reviewing agency definitions against " & NIBRS_REFERENCE & ". This is an analytical crosswalk, not NIBRS submission validation.")
    AnnotateAnalysisScope ws, r
    ws.Range("A3:F3").Merge: ws.Range("A3").WrapText = True: ws.Rows(3).RowHeight = 45
    WriteHeaders ws, 6, 1, Array("Agency", "Source value", "Code", "Proposed offense", "Approved", "Review note")
    rowNo = 7: ws.Range("A7:F" & pairs.Count + 6).NumberFormat = "@"
    For Each k In pairs.Keys
        choices = pairs(k): agency = choices(0): label = choices(1): code = ""
        If catalog.Exists(label) Then
            code = label
        Else
            For Each key In catalog.Keys
                values = catalog(key): If NormalText(values(0)) = label Then code = key: Exit For
            Next
        End If
        If code = "11B" Or code = "11C" Then code = "11A"
        ws.Cells(rowNo, 1).Value2 = agency: ws.Cells(rowNo, 2).Value2 = label: ws.Cells(rowNo, 3).Value2 = code: ws.Cells(rowNo, 5).Value2 = "No"
        ws.Cells(rowNo, 6).Value2 = "Needs analyst review; ambiguous descriptions are not inferred"
        If Len(code) > 0 Then values = catalog(code): ws.Cells(rowNo, 4).Value2 = values(0)
        ws.Cells(rowNo, 4).NumberFormat = "General"
        ws.Cells(rowNo, 4).Formula = "=IFERROR(VLOOKUP(C" & rowNo & ",'" & reference.Name & "'!$A$7:$F$" & catalog.Count + 6 & ",2,FALSE),"""")"
        rowNo = rowNo + 1
    Next
    ws.Names.Add Name:="_CAPNIBRSSource", RefersTo:="='" & Replace(r.Parent.Name, "'", "''") & "'!" & r.Address
    ws.Range("H1").Value2 = a(1, cc): ws.Range("H2").Value2 = "": If ac > 0 Then ws.Range("H2").Value2 = a(1, ac)
    ws.Range("H3").Value2 = allowFederal
    ws.Names.Add Name:="_CAPNIBRSCodes", RefersTo:="='" & reference.Name & "'!$A$7:$A$" & catalog.Count + 6
    ws.Range("C7:C" & rowNo - 1).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=_CAPNIBRSCodes"
    ws.Range("E7:E" & rowNo - 1).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes" & Application.International(xlListSeparator) & "No"
    ws.Range("C7:C" & rowNo - 1 & ",E7:E" & rowNo - 1).Interior.Color = RGB(224, 239, 250)
    FinishOutput ws, rowNo - 1, 6
    If showSummary Then MsgBox "Output: " & ws.Name & " plus " & reference.Name & ". Review each code, set Approved=Yes, then click Apply NIBRS on the crosswalk sheet. Broad labels such as Theft remain unmapped.", vbInformation, "CAP Index"
    Exit Sub
failed:
    If Not source Is Nothing Then Err.Raise Err.Number, "NIBRS crosswalk", Err.Description
    MsgBox Err.Description, vbExclamation, "CAP Index | NIBRS"
End Sub

Public Sub ApplyNIBRSCrosswalk(Optional ByVal mapping As Worksheet, Optional ByVal showSummary As Boolean = True)
    Dim r As Range, a As Variant, ws As Worksheet, catalog As Object, pairs As Object, conflicts As Object, rowNo As Long, i As Long, j As Long
    Dim key As String, agency As String, label As String, code As String, status As String, cc As Long, ac As Long, values As Variant, proposal As Variant, output() As Variant
    On Error GoTo failed
    If mapping Is Nothing Then Set mapping = ActiveSheet
    Set r = CAPNamedRange(mapping, "_CAPNIBRSSource")
    If r Is Nothing Then Err.Raise vbObjectError + 990, , "Select the NIBRS_Crosswalk sheet and review its mappings first."
    a = ScopedValues(r): cc = HeaderColumn(a, mapping.Range("H1").Value2): ac = HeaderColumn(a, SafeText(mapping.Range("H2").Value2))
    Set catalog = NIBRSCatalog(): Set pairs = CreateObject("Scripting.Dictionary"): Set conflicts = CreateObject("Scripting.Dictionary")
    For rowNo = 7 To mapping.Cells(mapping.Rows.Count, 2).End(xlUp).Row
        key = KeyPart(NormalText(mapping.Cells(rowNo, 1).Value2)) & KeyPart(NormalText(mapping.Cells(rowNo, 2).Value2))
        proposal = Array(NormalText(mapping.Cells(rowNo, 3).Value2), NormalText(mapping.Cells(rowNo, 5).Value2))
        If pairs.Exists(key) Then conflicts(key) = True Else pairs.Add key, proposal
    Next
    ReDim output(1 To UBound(a, 1), 1 To UBound(a, 2) + 5)
    For i = 1 To UBound(a, 1): For j = 1 To UBound(a, 2): output(i, j) = a(i, j): Next: Next
    output(1, UBound(a, 2) + 1) = "CAP_NIBRSCode": output(1, UBound(a, 2) + 2) = "CAP_NIBRSOffense": output(1, UBound(a, 2) + 3) = "CAP_NIBRSGroup"
    output(1, UBound(a, 2) + 4) = "CAP_CrimeAgainst": output(1, UBound(a, 2) + 5) = "CAP_ClassificationStatus"
    For i = 2 To UBound(a, 1)
        agency = "": If ac > 0 Then agency = NormalText(a(i, ac))
        label = NormalText(a(i, cc)): key = KeyPart(agency) & KeyPart(label): code = "": status = "Unmapped; analyst review required"
        If conflicts.Exists(key) Then
            status = "Duplicate crosswalk key; resolve before classification"
        ElseIf pairs.Exists(key) Then
            proposal = pairs(key)
            If proposal(1) <> "YES" Then
                status = "Not approved; analyst review required"
            ElseIf Not catalog.Exists(proposal(0)) Then
                status = "Invalid/obsolete code; review reference"
            Else
                values = catalog(proposal(0))
                If Len(values(4)) > 0 Then
                    status = "Legacy code; review conversion to 11A"
                ElseIf values(3) And Not mapping.Range("H3").Value2 Then
                    status = "Federal/tribal-only code; applicability review required"
                Else
                    code = proposal(0): status = "Reviewed analytical crosswalk; not submission validation"
                    If values(1) = "B" Then status = "Reviewed Group B arrest-only classification; keep separate from Group A offense counts"
                End If
            End If
        End If
        If Len(code) > 0 Then
            output(i, UBound(a, 2) + 1) = code: output(i, UBound(a, 2) + 2) = values(0): output(i, UBound(a, 2) + 3) = values(1): output(i, UBound(a, 2) + 4) = values(2)
        End If
        output(i, UBound(a, 2) + 5) = status
    Next
    Set ws = NewOutput(r.Parent.Parent, "NIBRS_Classified", r, NIBRS_REFERENCE & ". Original fields/offense rows retained. Unapproved, invalid and ambiguous mappings remain unclassified. Group B represents arrest data; not Group A offense counts.")
    AnnotateAnalysisScope ws, r
    ws.Range("A6").Resize(UBound(output, 1), UBound(output, 2)).NumberFormat = "@": ws.Range("A6").Resize(UBound(output, 1), UBound(output, 2)).Value2 = output
    CopySourceFormats r, ws.Range("A6").Resize(UBound(output, 1), UBound(a, 2))
    FinishOutput ws, UBound(output, 1) + 5, UBound(output, 2)
    ws.ListObjects.Add(xlSrcRange, ws.Range("A6").Resize(UBound(output, 1), UBound(output, 2)), , xlYes).TableStyle = "TableStyleMedium2"
    If showSummary Then MsgBox "Output: " & ws.Name & ". Review CAP_ClassificationStatus and unmapped rows. Source data is unchanged.", vbInformation, "CAP Index"
    Exit Sub
failed:
    If Not showSummary Then Err.Raise Err.Number, "Apply NIBRS", Err.Description
    MsgBox Err.Description, vbExclamation, "CAP Index | Apply NIBRS"
End Sub
