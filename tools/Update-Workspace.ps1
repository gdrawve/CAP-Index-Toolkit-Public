param([Parameter(Mandatory=$true)][string]$Destination)
$ErrorActionPreference = 'Stop'
$projectRoot = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$targetRoot = [IO.Path]::GetFullPath($Destination)
if (-not (Test-Path -LiteralPath (Join-Path $targetRoot 'README.md'))) { throw 'The target is not an existing toolkit workspace.' }
$latest = Get-ChildItem -LiteralPath (Join-Path $projectRoot 'build') -Directory | Where-Object { (Test-Path -LiteralPath (Join-Path $_.FullName 'validation.txt')) -and -not (Test-Path -LiteralPath (Join-Path $_.FullName 'FAILED.txt')) } | Sort-Object Name -Descending | Select-Object -First 1
if ($null -eq $latest) { throw 'No validated build is available.' }
$buildTarget = Join-Path (Join-Path $targetRoot 'build') $latest.Name
if (Test-Path -LiteralPath $buildTarget) { throw 'This build already exists in the target workspace.' }
$backup = Join-Path $targetRoot ('history\before-update-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
[void](New-Item -ItemType Directory -Path $backup)
foreach ($name in @('src','docs','tools','README.md')) {
    $existing = Join-Path $targetRoot $name
    if (Test-Path -LiteralPath $existing) { Copy-Item -LiteralPath $existing -Destination $backup -Recurse }
}
foreach ($relative in @('src\vba\modAbout.bas','src\vba\modAnalysisScope.bas','src\vba\modCalendarAnalysis.bas','src\vba\modNIBRS.bas','src\vba\modReportBuilder.bas','src\vba\modRelease015Tests.bas','docs\ANALYSIS-WORKFLOWS.md','src\vba\modSingleSite.bas','src\vba\modUsability.bas','src\vba\modUsabilityTests.bas','src\vba\modDatasetCheck.bas','src\vba\modToolGuide.bas','src\vba\modCheckTests.bas','src\vba\modReportFilters.bas','LICENSE','BRANDING.md','OWNERSHIP-NOTICE.txt','src\vba\modOwnership.bas','README.md','src\vba\modCore.bas','src\vba\modDateTime.bas','src\vba\modRibbon.bas','src\vba\modTests.bas','src\vba\modCensusGeocode.bas','src\vba\modCensusTransport.bas','src\vba\modCensusTests.bas','src\vba\clsCensusFixture.cls','src\ribbon\customUI14.xml','docs\DATA-CONTRACT.md','docs\BUILD-PLAN.md','docs\TIMESTAMP.md','docs\CENSUS-GEOCODER.md','tools\Build-Addin.ps1','tools\Publish-Workspace.ps1','tools\Update-Workspace.ps1')) {
    Copy-Item -LiteralPath (Join-Path $projectRoot $relative) -Destination (Join-Path $targetRoot $relative) -Force
}
foreach ($relative in @('src\vba\modGeography.bas','src\vba\modHundredBlocks.bas','src\vba\modBlockTests.bas','docs\HUNDRED-BLOCKS.md')) {
    Copy-Item -LiteralPath (Join-Path $projectRoot $relative) -Destination (Join-Path $targetRoot $relative) -Force
}
foreach ($relative in @('src\vba\modConcentration.bas','src\vba\modConcentrationTests.bas','src\vba\clsCAPEvents.cls','docs\CRIME-CONCENTRATION.md')) {
    Copy-Item -LiteralPath (Join-Path $projectRoot $relative) -Destination (Join-Path $targetRoot $relative) -Force
}
foreach ($relative in @('src\vba\modTimeReports.bas','src\vba\modTimeAnalysis.bas','src\vba\modTimeTests.bas','src\vba\modSimpleOutputs.bas','src\vba\modHour.bas','src\vba\modAddressRepeats.bas','src\vba\modWorkflowTests.bas','src\vba\modBuildChecks.bas','src\vba\modDatasetTests.bas','src\vba\modReportInput.bas','src\vba\modReporting.bas','src\vba\modExports.bas','src\vba\modReportingTests.bas','docs\REPORTING-EXPORT.md','src\vba\modReview.bas','docs\DATA-PREP.md','docs\TIME-LOCATION-ANALYSIS.md')) {
    Copy-Item -LiteralPath (Join-Path $projectRoot $relative) -Destination (Join-Path $targetRoot $relative) -Force
}
foreach ($relative in @('src\ribbon\icons','src\ribbon\brand','src\forms')) {
    $assetTarget = Join-Path $targetRoot $relative
    [void](New-Item -ItemType Directory -Path $assetTarget -Force)
    Copy-Item -Path (Join-Path (Join-Path $projectRoot $relative) '*') -Destination $assetTarget -Force
}
foreach ($relative in @('tools\Add-AboutForm.ps1','tools\Create-RibbonIcons.cjs','docs\RIBBON-DESIGN.md','docs\ribbon-icon-preview.png','docs\ribbon-icon-preview-dark.png')) { Copy-Item -LiteralPath (Join-Path $projectRoot $relative) -Destination (Join-Path $targetRoot $relative) -Force }
Copy-Item -LiteralPath (Join-Path $projectRoot 'docs\DEPLOYMENT.md') -Destination (Join-Path $targetRoot 'docs\DEPLOYMENT.md') -Force
Copy-Item -LiteralPath $latest.FullName -Destination (Join-Path $targetRoot 'build') -Recurse
$addin = Get-ChildItem -LiteralPath $latest.FullName -Filter '*.xlam' -File | Select-Object -First 1
$copied = Join-Path $buildTarget $addin.Name
if ((Get-FileHash -LiteralPath $addin.FullName).Hash -ne (Get-FileHash -LiteralPath $copied).Hash) { throw 'Copied add-in hash mismatch.' }
Write-Output ('ADDIN=' + $copied)
Write-Output ('SOURCE_BACKUP=' + $backup)
