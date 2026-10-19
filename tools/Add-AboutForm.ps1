param($Excel,$Book,[string]$ProjectRoot)
$component = $Book.VBProject.VBComponents.Add(3)
$component.Name = 'frmCAPAbout'
$null = $Excel.Run("'" + $Book.Name + "'!DesignCAPAbout", (Join-Path $ProjectRoot 'src\ribbon\brand\CAP-About-Logo.bmp'))
$component.CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $ProjectRoot 'src\forms\frmCAPAbout.vba')))
