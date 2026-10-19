Attribute VB_Name = "modAbout"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
Option Explicit
Public Function CAPAboutOverview() As String
    CAPAboutOverview = "CAP Index" & vbCrLf & _
        "CAP Index provides crime risk data, assessments and analytical tools for evaluating location-based risk." & vbCrLf & vbCrLf & _
        "Crime Analysis Excel Toolkit" & vbCrLf & _
        "Prepare crime data, classify reviewed NIBRS categories, examine crime concentration and temporal patterns, evaluate crime near a site, and create reports in Windows desktop Excel." & vbCrLf & vbCrLf & _
        "Ownership" & vbCrLf & _
        "Original toolkit code and documentation are licensed under the MIT License. Retain copyright and license notices. CAP Index names and logos are covered separately by BRANDING.md; the software license does not grant trademark rights or imply endorsement." & vbCrLf & vbCrLf & _
        "Agency data, third-party services and materials retain their respective ownership. No claim is made to ownership of crime facts or analytical methods."
End Function
Public Function CAPAboutChanges() As String
    CAPAboutChanges = "Version " & CAP_VERSION & " | Released October 19, 2026" & vbCrLf & vbCrLf & _
        "- Sixteen data and analysis tools now use one branded setup window with header dropdowns, selected-range counts and validation before Run." & vbCrLf & _
        "- Tool-specific help stays inside the window. Cancel creates no output." & vbCrLf & _
        "- Correctable run errors keep the setup choices available for retry." & vbCrLf & _
        "- Remembered mappings use explicit field roles; contaminated older Address/City/Block mappings are ignored." & vbCrLf & _
        "- Public-release packaging checks personal build paths and the approved sensitivity classification. Installation guidance includes Unblock and SHA256 verification." & vbCrLf & vbCrLf & _
        "Contact: Grant Drawve, gdrawve@capindex.com." & vbCrLf & vbCrLf & _
        "Unsigned release. Internal deployment follows CAP Index IT policy, including signing where required. Windows desktop Excel is required."
End Function
Public Function CAPAboutMethods() As String
    CAPAboutMethods = "Data handling" & vbCrLf & _
        "Census Geocoder sends address components and generated request IDs to the US Census Bureau after the user chooses that tool. Other analyses calculate locally in the workbook. Reports retain hidden local snapshots that can contain incident IDs and location keys; share reviewed exports when the complete workbook is sensitive." & vbCrLf & vbCrLf & _
        "Counts and scope" & vbCrLf & _
        "Unique incidents and source/offense records are distinct measures. Qualify IDs across agencies. New analyses use the chosen Analysis Scope; preparation tools process all selected rows. Refresh uses saved snapshots. Rebuild to reload source edits or change scope." & vbCrLf & vbCrLf & _
        "Accuracy and coverage" & vbCrLf & _
        "Geocoding match information does not establish positional accuracy. Randomized block addresses are synthetic estimates, not observed locations. Single Site uses approximate straight-line distances. Observed date spans and zero-record periods do not establish complete reporting." & vbCrLf & vbCrLf & _
        "Classification" & vbCrLf & _
        "NIBRS mappings require analyst approval. Group B is arrest reporting. The crosswalk supports analysis and does not validate a complete NIBRS submission. See the Tool Guide for inputs, outputs and interpretation limits."
End Function
Public Sub OpenCAPAboutLink(ByVal address As String)
    On Error GoTo failed
    ThisWorkbook.FollowHyperlink Address:=address
    Exit Sub
failed:
    MsgBox "Could not open the link. Website: https://capindex.com/" & vbCrLf & "Contact Grant Drawve: gdrawve@capindex.com", vbInformation, "CAP Index"
End Sub
Public Sub DesignCAPAbout(ByVal path As String)
    Dim component As Object, form As Object, logo As Object, pages As Object, page As Object, body As Object, button As Object
    Dim i As Long, captions As Variant, names As Variant, buttons As Variant, item As Variant
    Set component = ThisWorkbook.VBProject.VBComponents("frmCAPAbout"): Set form = component.Designer
    component.Properties("Width") = 600: component.Properties("Height") = 460: component.Properties("StartUpPosition") = 1
    form.BackColor = RGB(255, 255, 255)
    Set logo = form.Controls.Add("Forms.Image.1", "logo", True)
    logo.Left = 20: logo.Top = 12: logo.Width = 118: logo.Height = 46: logo.BorderStyle = 0: logo.BackStyle = 0: logo.PictureSizeMode = 3
    Set logo.Picture = LoadPicture(path)
    AboutLabel form, "lblTitle", "CAP Index", 155, 12, 410, 28, 20, True
    AboutLabel form, "lblVersion", "Crime Analysis Excel Toolkit", 155, 43, 410, 20, 11
    Set pages = form.Controls.Add("Forms.MultiPage.1", "pages", True)
    pages.Left = 20: pages.Top = 78: pages.Width = 558: pages.Height = 248
    pages.Font.Name = "Calibri": pages.Font.Size = 11
    Do While pages.Pages.Count < 3: pages.Pages.Add: Loop
    captions = Array("Overview", "What's New", "Data & Methods"): names = Array("txtOverview", "txtChanges", "txtMethods")
    For i = 0 To 2
        Set page = pages.Pages(i): page.Caption = captions(i)
        Set body = page.Controls.Add("Forms.TextBox.1", names(i), True)
        body.Left = 10: body.Top = 10: body.Width = 528: body.Height = 200
        body.MultiLine = True: body.WordWrap = True: body.Locked = True: body.ScrollBars = 2: body.BorderStyle = 0
        body.BackColor = RGB(255, 255, 255): body.ForeColor = RGB(51, 51, 51): body.Font.Name = "Calibri": body.Font.Size = 11
    Next
    AboutLabel form, "lblContact", "Contact: Grant Drawve | gdrawve@capindex.com", 20, 336, 558, 20, 11, True
    AboutLabel form, "lblCopyright", CAP_COPYRIGHT, 20, 358, 558, 18, 10
    buttons = Array(Array("cmdWebsite", "CAP Website", 20, 110), Array("cmdContact", "Email Grant Drawve", 142, 150), Array("cmdGuide", "Tool Guide", 304, 130), Array("cmdClose", "Close", 446, 132))
    For Each item In buttons
        Set button = form.Controls.Add("Forms.CommandButton.1", item(0), True)
        button.Caption = item(1): button.Left = item(2): button.Top = 388: button.Width = item(3): button.Height = 26
        button.Font.Name = "Calibri": button.Font.Size = 11
    Next
    form.Controls("cmdClose").Cancel = True
End Sub
Private Sub AboutLabel(ByVal form As Object, ByVal name As String, ByVal text As String, ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double, ByVal size As Double, Optional ByVal bold As Boolean = False)
    Dim label As Object
    Set label = form.Controls.Add("Forms.Label.1", name, True)
    label.Caption = text: label.Left = x: label.Top = y: label.Width = width: label.Height = height
    label.Font.Name = "Calibri": label.Font.Size = size: label.Font.Bold = bold
    label.BackStyle = 0: label.ForeColor = RGB(33, 28, 101)
End Sub
Public Function CheckCAPAbout() As String
    Load frmCAPAbout
    If frmCAPAbout.Controls("pages").Pages.Count <> 3 Then Err.Raise vbObjectError + 1001, , "About pages missing"
    If frmCAPAbout.Controls("logo").Picture Is Nothing Then Err.Raise vbObjectError + 1001, , "About logo missing"
    If InStr(frmCAPAbout.Controls("lblContact").Caption, "gdrawve@capindex.com") = 0 Then Err.Raise vbObjectError + 1001, , "Contact missing"
    If InStr(frmCAPAbout.Controls("lblVersion").Caption, CAP_VERSION) = 0 Then Err.Raise vbObjectError + 1001, , "About version missing"
    Unload frmCAPAbout
    CheckCAPAbout = "PASS: About form loads with embedded logo, three readable tabs, current version and Grant Drawve contact"
End Function
