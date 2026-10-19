Attribute VB_Name = "modSetupDesign"
' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
Option Explicit
Public Sub DesignCAPSetup(ByVal logoPath As String)
    Dim component As Object, form As Object, control As Object, i As Long, y As Double
    Set component = ThisWorkbook.VBProject.VBComponents("frmCAPSetup"): Set form = component.Designer
    component.Properties("Width") = 752: component.Properties("Height") = 500: component.Properties("StartUpPosition") = 1
    form.BackColor = RGB(255, 255, 255)
    Set control = form.Controls.Add("Forms.Image.1", "logo", True)
    control.Left = 16: control.Top = 10: control.Width = 76: control.Height = 32
    control.BorderStyle = 0: control.BackStyle = 0: control.PictureSizeMode = 3: Set control.Picture = LoadPicture(logoPath)
    SetupControl form, "Forms.Label.1", "lblTitle", "Tool setup", 106, 10, 610, 26, 18
    SetupControl form, "Forms.Label.1", "lblDescription", "", 106, 39, 610, 28, 10
    SetupControl form, "Forms.Label.1", "lblRange", "Select headers and data.", 18, 78, 590, 28, 10
    SetupControl form, "Forms.CommandButton.1", "cmdSelect", "Select range", 618, 78, 110, 25, 11
    SetupControl form, "Forms.Label.1", "lblCounts", "", 18, 110, 710, 17, 10
    SetupControl form, "Forms.Label.1", "lblFields", "FIELDS", 18, 135, 348, 16, 10
    SetupControl form, "Forms.Label.1", "lblOptions", "OPTIONS", 398, 135, 330, 16, 10
    For i = 1 To 10
        y = 155 + (i - 1) * 30
        SetupControl form, "Forms.Label.1", "lblField" & i, "", 18, y + 2, 106, 18, 10
        SetupControl form, "Forms.ComboBox.1", "cboField" & i, "", 128, y, 238, 19, 11
        Set control = form.Controls("cboField" & i): control.Style = 2: control.ListRows = 12: control.MatchEntry = 1
        SetupControl form, "Forms.Label.1", "lblHelp" & i, "", 128, y + 19, 238, 11, 8
    Next
    For i = 1 To 6
        y = 155 + (i - 1) * 38
        SetupControl form, "Forms.Label.1", "lblOption" & i, "", 398, y, 330, 15, 10
        SetupControl form, "Forms.ComboBox.1", "cboOption" & i, "", 398, y + 15, 330, 20, 11
        form.Controls("cboOption" & i).Style = 2: form.Controls("cboOption" & i).ListRows = 14
        SetupControl form, "Forms.TextBox.1", "txtOption" & i, "", 398, y + 15, 330, 20, 11
    Next
    SetupControl form, "Forms.TextBox.1", "txtHelp", "", 18, 155, 710, 220, 11
    With form.Controls("txtHelp")
        .MultiLine = True: .WordWrap = True: .Locked = True: .ScrollBars = 2: .Visible = False
    End With
    SetupControl form, "Forms.Label.1", "lblObserved", "", 398, 386, 330, 28, 10
    SetupControl form, "Forms.Label.1", "lblNotice", "", 18, 425, 710, 16, 10
    SetupControl form, "Forms.Label.1", "lblStatus", "", 18, 445, 710, 30, 10
    form.Controls("lblStatus").ForeColor = RGB(160, 36, 36)
    SetupControl form, "Forms.CommandButton.1", "cmdGuide", "Tool Guide", 18, 480, 112, 26, 11
    SetupControl form, "Forms.CommandButton.1", "cmdCancel", "Cancel", 492, 480, 112, 26, 11
    SetupControl form, "Forms.CommandButton.1", "cmdRun", "Run", 616, 480, 112, 26, 11
    form.Controls("cmdRun").Default = True: form.Controls("cmdCancel").Cancel = True
End Sub
Private Sub SetupControl(ByVal form As Object, ByVal kind As String, ByVal name As String, ByVal caption As String, ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double, ByVal fontSize As Double)
    Dim c As Object
    Set c = form.Controls.Add(kind, name, True)
    c.Left = x: c.Top = y: c.Width = width: c.Height = height: c.Font.Name = "Calibri": c.Font.Size = fontSize
    If kind = "Forms.Label.1" Or kind = "Forms.CommandButton.1" Then c.Caption = caption
    If kind = "Forms.Label.1" Then c.BackStyle = 0: c.ForeColor = RGB(33, 28, 101): c.WordWrap = True
End Sub
Public Function CheckCAPSetup() As String
    Dim i As Long
    Load frmCAPSetup
    For i = 1 To 10
        If TypeName(frmCAPSetup.Controls("cboField" & i)) <> "ComboBox" Then Err.Raise vbObjectError + 1202, , "Setup field missing"
    Next
    For i = 1 To 6
        If TypeName(frmCAPSetup.Controls("cboOption" & i)) <> "ComboBox" Or TypeName(frmCAPSetup.Controls("txtOption" & i)) <> "TextBox" Then Err.Raise vbObjectError + 1202, , "Setup option missing"
    Next
    If frmCAPSetup.Controls("logo").Picture Is Nothing Then Err.Raise vbObjectError + 1202, , "Setup logo missing"
    If Not frmCAPSetup.Controls("cmdRun").Default Or Not frmCAPSetup.Controls("cmdCancel").Cancel Then Err.Raise vbObjectError + 1202, , "Setup buttons missing"
    Unload frmCAPSetup
    CheckCAPSetup = "PASS: Setup form loads with logo, ten field slots, six option slots, Run/Cancel/Tool Guide"
End Function
