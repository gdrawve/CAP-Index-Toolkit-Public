' Copyright (c) 2026 CAP Index. Licensed under the MIT License.
Option Explicit
Private Sub UserForm_Initialize()
    Dim wb As Workbook
    Me.Caption = "CAP Index | About | " & CAP_VERSION
    Me.Controls("lblVersion").Caption = "Crime Analysis Excel Toolkit  |  Version " & CAP_VERSION
    Me.Controls("lblContact").Caption = "Contact: Grant Drawve  |  gdrawve@capindex.com"
    Me.Controls("lblCopyright").Caption = CAP_COPYRIGHT
    Me.Controls("pages").Pages(0).Controls("txtOverview").Text = CAPAboutOverview()
    Me.Controls("pages").Pages(1).Controls("txtChanges").Text = CAPAboutChanges()
    Me.Controls("pages").Pages(2).Controls("txtMethods").Text = CAPAboutMethods()
    Me.Controls("pages").Value = 0
    Me.Controls("cmdGuide").Enabled = False
    Set wb = ActiveWorkbook
    If Not wb Is Nothing Then
        If Not wb Is ThisWorkbook Then Me.Controls("cmdGuide").Enabled = Not wb.ReadOnly And Not wb.ProtectStructure
    End If
End Sub
Private Sub cmdWebsite_Click()
    OpenCAPAboutLink "https://capindex.com/"
End Sub
Private Sub cmdContact_Click()
    OpenCAPAboutLink "mailto:gdrawve@capindex.com?subject=CAP%20Index%20Toolkit%20" & CAP_VERSION
End Sub
Private Sub cmdGuide_Click()
    Me.Hide
    BuildToolGuide
    Unload Me
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
