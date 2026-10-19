param([string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot), [switch]$LiveCensusSmoke)
$ErrorActionPreference = 'Stop'
$ProjectRoot = [IO.Path]::GetFullPath($ProjectRoot)
$coreText = [IO.File]::ReadAllText((Join-Path $ProjectRoot 'src\vba\modCore.bas'))
$versionMatch = [regex]::Match($coreText, 'CAP_VERSION As String = "([^"]+)"')
if (-not $versionMatch.Success) { throw 'CAP_VERSION was not found.' }
$version = $versionMatch.Groups[1].Value
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$buildDir = Join-Path $ProjectRoot ('build\' + $stamp)
[void](New-Item -ItemType Directory -Path $buildDir)
$releasePath = Join-Path $buildDir ('CAP-Index-Toolkit-' + $version + '.xlam')
$addinPath = Join-Path $buildDir ('CAP-Index-Toolkit-' + $version + '-UNVALIDATED.xlam')
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $book = $excel.Workbooks.Add()
    try { $null = $book.VBProject.VBComponents.Count } catch { throw 'Excel blocks VBA project access. Use an IT-approved build workstation or manual VBA import. This script does not change Trust Center settings.' }
    $importDir = Join-Path $buildDir 'imports'
    [void](New-Item -ItemType Directory -Path $importDir)
    foreach ($file in Get-ChildItem -LiteralPath (Join-Path $ProjectRoot 'src\vba') -File | Where-Object { $_.Extension -in @('.bas','.cls') }) {
        $sourceLines = [IO.File]::ReadAllLines($file.FullName)
        foreach ($sourceLine in $sourceLines) { if ($sourceLine.Length -gt 1023) { throw ('VBA source line exceeds 1023 characters: ' + $file.Name) } }
        $importPath = Join-Path $importDir $file.Name
        $importText = [IO.File]::ReadAllText($file.FullName) -replace "\r?\n", "`r`n"
        [IO.File]::WriteAllText($importPath, $importText, [Text.Encoding]::GetEncoding(1252))
        $component = $book.VBProject.VBComponents.Import($importPath)
        if ($file.Extension -eq '.cls' -and $component.Type -ne 2) { throw ('Class module import failed: ' + $file.Name) }
    }
    & (Join-Path $PSScriptRoot 'Add-AboutForm.ps1') -Excel $excel -Book $book -ProjectRoot $ProjectRoot
    & (Join-Path $PSScriptRoot 'Add-SetupForm.ps1') -Excel $excel -Book $book -ProjectRoot $ProjectRoot
    $setupFormResult = $excel.Run("'" + $book.Name + "'!CheckCAPSetup")
    if ($setupFormResult -notlike 'PASS:*') { throw 'Setup form validation failed.' }
    $aboutResult = $excel.Run("'" + $book.Name + "'!CheckCAPAbout")
    if ($aboutResult -notlike 'PASS:*') { throw 'About form validation failed.' }
    $book.VBProject.VBComponents.Item('ThisWorkbook').CodeModule.AddFromString("Private Sub Workbook_Open()`r`nInitializeCAPEvents`r`nEnd Sub`r`nPrivate Sub Workbook_BeforeClose(Cancel As Boolean)`r`nSet CAPEvents = Nothing`r`nEnd Sub")
    $excel.Run("'" + $book.Name + "'!StampCAPAddinMetadata")
    $book.IsAddin = $true
    $book.SaveAs($addinPath, 55)
    $book.Close($false)
    $book = $excel.Workbooks.Open($addinPath)
    $ownershipResult = $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunOwnershipTests")
    $testResult = $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunSelfTests")
    $smokeResult = $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunOutputSmokeTest")
    $commandResult = $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunCommandTests")
    $timestampResult = $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunTimestampTests")
    $censusResult = $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunCensusTests")
    $batchResult = $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunCensusBatchLimitTest")
    $blockResult = $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunBlockTests")
    $concentrationResult = $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunConcentrationTests", $buildDir)
    $timeResult = $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunTimeAnalysisTests", $buildDir)
    $workflowResult = $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunWorkflowTests", $buildDir)
    $datasetResult = $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunDatasetTests", $buildDir)
    $reportingResult = $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunReportingTests", $buildDir)
    $qualityResult = $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunDatasetCheckTests", $buildDir)
    $uxResult = $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunUsabilityTests", $buildDir)
    $release015Result = $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunRelease015Tests", $buildDir)
    foreach ($result in @($release015Result,$uxResult,$qualityResult,$ownershipResult,$testResult,$smokeResult,$commandResult,$timestampResult,$censusResult,$batchResult,$blockResult,$concentrationResult,$timeResult,$workflowResult,$datasetResult,$reportingResult)) { if ([string]::IsNullOrWhiteSpace([string]$result)) { throw 'A native check returned no result.' }; if ($result -like 'FAILED:*') { throw $result } }
    $setupResult = $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunSetupTests")
    if ($setupResult -notlike 'PASS:*') { throw ('Setup tests failed: ' + $setupResult) }
    $liveResult = 'Live Census test not requested'
    if ($LiveCensusSmoke) { $liveResult = $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunCensusLiveSmoke") }
    $maps = @()
    foreach ($id in @('ChartTypeMapsInsertGallery','PowerMapLaunch','Filter','SortCustomExcel','ViewFreezePanesGallery','PivotTableInsert','TableInsertExcel','SortClear','ConvertTextToTable','FindDialogExcel','ConditionalFormattingMenu')) {
        try { $maps += $id + ': ' + $excel.CommandBars.GetLabelMso($id) } catch { throw "Built-in Maps control could not be verified: $id" }
    }
    [xml]$ribbonDoc = [IO.File]::ReadAllText((Join-Path $ProjectRoot 'src\ribbon\customUI14.xml'))
    if ($ribbonDoc.SelectNodes('//*[@id="capScores"]').Count -ne 0) { throw 'CAP Score button must be absent.' }
    if ($ribbonDoc.SelectNodes('//*[@id="capOwnership"]').Count -ne 1) { throw 'Ownership button missing.' }
    if ($ribbonDoc.SelectSingleNode('//*[@id="capCheck"]').ParentNode.GetAttribute('id') -ne 'capPrepGroup') { throw 'Dataset Check must be in Crime Data Prep.' }
    foreach ($mapId in @('ChartTypeMapsInsertGallery','PowerMapLaunch')) {
        if ($ribbonDoc.SelectSingleNode(('//*[@idMso="' + $mapId + '"]')).ParentNode.GetAttribute('id') -ne 'capGeoGroup') { throw 'Maps must be in Geography.' }
    }
    if ($ribbonDoc.SelectNodes('//*[@id="capRefresh" or @id="capConcentrationRefresh" or @id="capTimeRefresh" ]').Count -ne 1) { throw 'Expected one unified Refresh button.' }
    if ($ribbonDoc.SelectSingleNode('//*[@id="capSingleSite" ]').ParentNode.GetAttribute('id') -ne 'capAnalysisGroup') { throw 'Single Site must be in Crime Analysis.' }
    foreach ($essentialId in @('Filter','SortCustomExcel','ViewFreezePanesGallery','PivotTableInsert','TableInsertExcel','SortClear','ConvertTextToTable','FindDialogExcel','ConditionalFormattingMenu')) { if ($ribbonDoc.SelectSingleNode(('//*[@idMso="' + $essentialId + '" ]')).ParentNode.GetAttribute('id') -ne 'capExcelGroup') { throw 'Excel Essentials group is incorrect.' } }
    $ribbonCode = $book.VBProject.VBComponents.Item('modRibbon').CodeModule
    foreach ($callback in @('CAP100_OnLoad','CAP100_OnAction','CAP100_DatasetCheck')) {
        if ($ribbonCode.ProcCountLines($callback, 0) -lt 1) { throw ('Missing ribbon callback: ' + $callback) }
    }
    if ($ribbonDoc.SelectSingleNode('//*[@id="capCheck"]').GetAttribute('onAction') -ne 'CAP100_DatasetCheck') { throw 'Dataset Check callback is incorrect.' }
    $groupOrder = @($ribbonDoc.SelectSingleNode('//*[@id="capIndexTab"]').ChildNodes | Where-Object { $_.LocalName -eq 'group' } | ForEach-Object { $_.GetAttribute('id') })
    if (($groupOrder[0..2] -join ',') -ne 'capPrepGroup,capScopeGroup,capClassificationGroup') { throw 'Expected Crime Data Prep, Analysis Scope, Classification first.' }
    $helpMenu = $ribbonDoc.SelectSingleNode('//*[@id="capHelpMenu"]')
    if ($null -eq $helpMenu -or $helpMenu.LocalName -ne 'menu' -or $helpMenu.GetAttribute('image') -ne 'cap-logo-32') { throw 'Branded Help dropdown is missing.' }
    foreach ($helpId in @('capHelp','capOwnership')) { if ($ribbonDoc.SelectSingleNode(('//*[@id="' + $helpId + '"]')).ParentNode.GetAttribute('id') -ne 'capHelpMenu') { throw 'Help commands must be inside the CAP menu.' } }
    Add-Type -AssemblyName System.Drawing
    $icons = @()
    foreach ($button in $ribbonDoc.SelectNodes('//*[local-name()="button" or local-name()="toggleButton" or local-name()="menu"]')) {
        $customId = $button.GetAttribute('image')
        if (-not [string]::IsNullOrEmpty($customId)) {
            $iconPath = Join-Path $ProjectRoot ('src\ribbon\icons\' + $customId + '.png')
            if (-not (Test-Path -LiteralPath $iconPath)) { throw ('Missing custom icon: ' + $customId) }
            $bitmap = [Drawing.Bitmap]::new($iconPath)
            try {
                $size = if ($button.GetAttribute('size') -eq 'large') { 32 } else { 16 }
                if ($bitmap.Width -ne $size -or $bitmap.Height -ne $size) { throw ('Incorrect icon dimensions: ' + $customId) }
                $opaque = 0; $transparent = 0
                for ($y = 0; $y -lt $size; $y++) { for ($x = 0; $x -lt $size; $x++) {
                    if ($bitmap.GetPixel($x,$y).A -gt 0) { $opaque++ } else { $transparent++ }
                } }
                if ($opaque -lt 10 -or $transparent -lt 10) { throw ('Empty icon or missing transparency: ' + $customId) }
            } finally { $bitmap.Dispose() }
            $icons += 'Embedded icon verified: ' + $customId
            continue
        }
        $id = $button.GetAttribute('imageMso')
        if ([string]::IsNullOrEmpty($id)) { throw ('Missing icon: ' + $button.GetAttribute('id')) }
        try {
            $iconValid = $excel.Run("'" + $book.Name + "'!VerifyRibbonIcon", $id)
            if (-not $iconValid) { throw 'Image was empty' }
            $icons += 'Icon verified: ' + $id
        } catch { throw ('Ribbon icon could not be verified: ' + $id) }
    }
    $report = @(('CAP Index Toolkit ' + $version),('Build: ' + $stamp),('Excel version: ' + $excel.Version),$aboutResult,$setupFormResult,$setupResult,$ownershipResult,'Ribbon: About CAP Index present',$testResult,$smokeResult,$commandResult,$timestampResult,$censusResult,$batchResult,$blockResult,$concentrationResult,$timeResult,$workflowResult,$datasetResult,$reportingResult,$qualityResult,$uxResult,$release015Result,$liveResult) + $maps + $icons
    $book.Close($false)
    $book = $null
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    $excel = $null
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::Open($addinPath, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $xml = [IO.File]::ReadAllText((Join-Path $ProjectRoot 'src\ribbon\customUI14.xml'))
        $null = [xml]$xml
        $entry = $zip.CreateEntry('customUI/customUI14.xml')
        $writer = New-Object IO.StreamWriter($entry.Open(), (New-Object Text.UTF8Encoding($false)))
        $writer.Write($xml); $writer.Dispose()
        # PNG image relationship IDs are used directly by RibbonX image attributes.
        $iconAssets = @(Get-ChildItem -LiteralPath (Join-Path $ProjectRoot 'src\ribbon\icons') -Filter '*.png' -File)
        [xml]$imageRels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />'
        foreach ($asset in $iconAssets) {
            $iconEntry = $zip.CreateEntry('customUI/images/' + $asset.Name)
            $inputStream = [IO.File]::OpenRead($asset.FullName); $outputStream = $iconEntry.Open()
            try { $inputStream.CopyTo($outputStream) } finally { $inputStream.Dispose(); $outputStream.Dispose() }
            $imageRel = $imageRels.CreateElement('Relationship', 'http://schemas.openxmlformats.org/package/2006/relationships')
            $imageRel.SetAttribute('Id', $asset.BaseName)
            $imageRel.SetAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
            $imageRel.SetAttribute('Target', 'images/' + $asset.Name)
            $null = $imageRels.DocumentElement.AppendChild($imageRel)
        }
        $imageRelEntry = $zip.CreateEntry('customUI/_rels/customUI14.xml.rels')
        $writer = New-Object IO.StreamWriter($imageRelEntry.Open(), (New-Object Text.UTF8Encoding($false)))
        $writer.Write($imageRels.OuterXml); $writer.Dispose()
        $entry = $zip.GetEntry('_rels/.rels')
        $reader = New-Object IO.StreamReader($entry.Open())
        [xml]$rels = $reader.ReadToEnd(); $reader.Dispose(); $entry.Delete()
        $rel = $rels.CreateElement('Relationship', 'http://schemas.openxmlformats.org/package/2006/relationships')
        $rel.SetAttribute('Id','rIdCAPRibbon')
        $rel.SetAttribute('Type','http://schemas.microsoft.com/office/2007/relationships/ui/extensibility')
        $rel.SetAttribute('Target','customUI/customUI14.xml')
        $null = $rels.DocumentElement.AppendChild($rel)
        $entry = $zip.CreateEntry('_rels/.rels')
        $writer = New-Object IO.StreamWriter($entry.Open(), (New-Object Text.UTF8Encoding($false)))
        $writer.Write($rels.OuterXml); $writer.Dispose()
        # Explicit content type keeps the package valid even if Excel omits a default xml entry.
        $entry = $zip.GetEntry('[Content_Types].xml')
        $reader = New-Object IO.StreamReader($entry.Open())
        [xml]$types = $reader.ReadToEnd(); $reader.Dispose(); $entry.Delete()
        $override = $types.CreateElement('Override', 'http://schemas.openxmlformats.org/package/2006/content-types')
        $override.SetAttribute('PartName','/customUI/customUI14.xml'); $override.SetAttribute('ContentType','application/xml')
        $null = $types.DocumentElement.AppendChild($override)
        if (-not $types.DocumentElement.ChildNodes.Where({ $_.LocalName -eq 'Default' -and $_.GetAttribute('Extension') -eq 'png' })) {
            $pngType = $types.CreateElement('Default','http://schemas.openxmlformats.org/package/2006/content-types')
            $pngType.SetAttribute('Extension','png'); $pngType.SetAttribute('ContentType','image/png')
            $null = $types.DocumentElement.AppendChild($pngType)
        }
        $entry = $zip.CreateEntry('[Content_Types].xml')
        $writer = New-Object IO.StreamWriter($entry.Open(), (New-Object Text.UTF8Encoding($false)))
        $writer.Write($types.OuterXml); $writer.Dispose()
    } finally { $zip.Dispose() }
    $zip = [IO.Compression.ZipFile]::OpenRead($addinPath)
    try {
        foreach ($button in $ribbonDoc.SelectNodes('//*[@image]')) {
            $iconId = $button.GetAttribute('image')
            if ($null -eq $zip.GetEntry('customUI/images/' + $iconId + '.png')) { throw ('Packaged image is missing: ' + $iconId) }
            if ($null -eq $imageRels.SelectSingleNode(('//*[local-name()="Relationship" and @Id="' + $iconId + '"]'))) { throw ('Packaged image relationship is missing: ' + $iconId) }
        }
        $report += 'Ribbon: first groups reordered; CAP logo Help menu; all custom images embedded with matching relationships and transparent 16/32 pixel assets'
    } finally { $zip.Dispose() }
    # Reopen the finished package and rerun tests to detect packaging damage.
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false
    $book = $excel.Workbooks.Open($addinPath)
    $report += 'Packaged About: ' + $excel.Run("'" + $book.Name + "'!CheckCAPAbout")
    $report += 'Packaged reopen: ' + $excel.Run("'" + $book.Name + "'!RunBuildCheck", "RunSelfTests")
    $report += 'Ribbon display and Maps launch: manual acceptance pending'
    $report += 'Unsigned release package; company deployment pending applicable signing and acceptance'
    Add-Type -AssemblyName System.Drawing
    foreach ($png in Get-ChildItem -LiteralPath $buildDir -Filter '*.png' -File) {
        $bitmap = [Drawing.Bitmap]::new($png.FullName)
        try {
            $ink = 0
            $stepX = [Math]::Max(1, [int]($bitmap.Width / 100))
            $stepY = [Math]::Max(1, [int]($bitmap.Height / 100))
            for ($y = 3; $y -lt $bitmap.Height - 3; $y += $stepY) {
                for ($x = 3; $x -lt $bitmap.Width - 3; $x += $stepX) {
                    $pixel = $bitmap.GetPixel($x, $y)
                    if ($pixel.R -lt 160 -or $pixel.G -lt 160 -or $pixel.B -lt 160) { $ink++ }
                }
            }
            if ($ink -lt 5) { throw ('Blank image detected: ' + $png.Name) }
        } finally { $bitmap.Dispose() }
    }
    $report += 'Generated PNGs: nonblank pixel validation passed'
    $report | Set-Content -LiteralPath (Join-Path $buildDir 'validation.txt') -Encoding UTF8
    $book.Close($false); $book = $null
    $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel); $excel = $null
    Move-Item -LiteralPath $addinPath -Destination $releasePath
    $addinPath = $releasePath
    $hash = Get-FileHash -LiteralPath $addinPath -Algorithm SHA256
    ($hash.Hash + '  ' + [IO.Path]::GetFileName($addinPath)) | Set-Content -LiteralPath (Join-Path $buildDir 'SHA256.txt')
    foreach ($notice in @('LICENSE','BRANDING.md','OWNERSHIP-NOTICE.txt')) { Copy-Item -LiteralPath (Join-Path $ProjectRoot $notice) -Destination $buildDir }
    Write-Output ('BUILD=' + $addinPath)
    $report | Write-Output
} catch {
    $_.Exception.Message | Set-Content -LiteralPath (Join-Path $buildDir 'FAILED.txt')
    throw
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
