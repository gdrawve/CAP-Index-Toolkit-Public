param([Parameter(Mandatory=$true)][string]$DestinationParent)
$ErrorActionPreference = 'Stop'
$projectRoot = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$parentPath = [IO.Path]::GetFullPath($DestinationParent)
if (-not (Test-Path -LiteralPath $parentPath -PathType Container)) { throw 'The selected working folder does not exist.' }
$destination = Join-Path $parentPath 'CAP-Index-Toolkit'
if (Test-Path -LiteralPath $destination) { throw 'CAP-Index-Toolkit already exists in the selected folder. No existing files have been replaced.' }
$latest = Get-ChildItem -LiteralPath (Join-Path $projectRoot 'build') -Directory | Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'validation.txt') } | Sort-Object Name -Descending | Select-Object -First 1
if ($null -eq $latest) { throw 'No successfully validated build is available.' }
[void](New-Item -ItemType Directory -Path $destination)
foreach ($name in @('src','docs','tests','tools','README.md','LICENSE','BRANDING.md','OWNERSHIP-NOTICE.txt')) {
    Copy-Item -LiteralPath (Join-Path $projectRoot $name) -Destination $destination -Recurse
}
$buildParent = Join-Path $destination 'build'
[void](New-Item -ItemType Directory -Path $buildParent)
Copy-Item -LiteralPath $latest.FullName -Destination $buildParent -Recurse
$original = (Get-ChildItem -LiteralPath $latest.FullName -Filter '*.xlam' -File | Select-Object -First 1).FullName
$copied = Join-Path (Join-Path $buildParent $latest.Name) ([IO.Path]::GetFileName($original))
if ((Get-FileHash -LiteralPath $original).Hash -ne (Get-FileHash -LiteralPath $copied).Hash) { throw 'Copied add-in hash mismatch.' }
Write-Output ('WORKSPACE=' + $destination)
Write-Output ('ADDIN=' + $copied)
