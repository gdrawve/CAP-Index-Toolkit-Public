param($Excel,$Book,[string]$ProjectRoot)
$component = $Book.VBProject.VBComponents.Add(3)
$component.Name = 'frmCAPSetup'
$null = $Excel.Run("'" + $Book.Name + "'!DesignCAPSetup", (Join-Path $ProjectRoot 'src\ribbon\brand\CAP-About-Logo.bmp'))
$component.CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $ProjectRoot 'src\forms\frmCAPSetup.vba')))
