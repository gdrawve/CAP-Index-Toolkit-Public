param([Parameter(Mandatory=$true)][string]$Path, [switch]$VerifyOnly, [string]$ExpectedSensitivityLabelId = '')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$mode = if ($VerifyOnly) { [IO.Compression.ZipArchiveMode]::Read } else { [IO.Compression.ZipArchiveMode]::Update }
$zip = [IO.Compression.ZipFile]::Open($Path, $mode)
function Read-PackageText($entry) {
    $reader = [IO.StreamReader]::new($entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}
function Write-PackageText($archive, [string]$name, [string]$value) {
    $old = $archive.GetEntry($name)
    if ($null -ne $old) { $old.Delete() }
    $entry = $archive.CreateEntry($name)
    $writer = [IO.StreamWriter]::new($entry.Open(), [Text.UTF8Encoding]::new($false))
    try { $writer.Write($value) } finally { $writer.Dispose() }
}
try {
    if (-not $VerifyOnly) {
        $entry = $zip.GetEntry('xl/workbook.xml')
        [xml]$xml = Read-PackageText $entry
        foreach ($node in @($xml.SelectNodes('//*[local-name()="absPath"]'))) {
            $parent = $node.ParentNode
            if ($parent.LocalName -eq 'Choice' -and $parent.ChildNodes.Count -eq 1 -and $parent.ParentNode.LocalName -eq 'AlternateContent' -and $parent.ParentNode.ChildNodes.Count -eq 1) {
                $wrapper = $parent.ParentNode
                [void]$wrapper.ParentNode.RemoveChild($wrapper)
            } else { [void]$parent.RemoveChild($node) }
        }
        Write-PackageText $zip 'xl/workbook.xml' $xml.OuterXml
        $entry = $zip.GetEntry('docProps/core.xml')
        if ($null -ne $entry) {
            [xml]$xml = Read-PackageText $entry
            $modifiedBy = $xml.SelectSingleNode('//*[local-name()="lastModifiedBy"]')
            if ($null -ne $modifiedBy) { $modifiedBy.InnerText = 'CAP Index' }
            Write-PackageText $zip 'docProps/core.xml' $xml.OuterXml
        }
    }
    foreach ($entry in $zip.Entries) {
        if ($entry.FullName -match '\.(xml|rels)$') {
            $text = Read-PackageText $entry
            if ($text -match '(?i)[a-z]:[\\/]Users[\\/]' -or $text -match '<(?:\w+:)?absPath\b') { throw ('Private build path remains in ' + $entry.FullName) }
        }
    }
    $entry = $zip.GetEntry('xl/vbaProject.bin')
    $buffer = [IO.MemoryStream]::new()
    $stream = $entry.Open()
    try { $stream.CopyTo($buffer) } finally { $stream.Dispose() }
    $bytes = $buffer.ToArray(); $buffer.Dispose()
    foreach ($encoding in @([Text.Encoding]::GetEncoding(28591), [Text.Encoding]::Unicode)) {
        if ($encoding.GetString($bytes) -match '(?i)[a-z]:\\Users\\') { throw 'Personal build profile remains in VBA reference/cache metadata.' }
    }
    $labelEntries = @($zip.Entries | Where-Object { $_.FullName -match '(?i)LabelInfo\.xml$' })
    if ($ExpectedSensitivityLabelId) {
        $ids = @()
        foreach ($entry in $labelEntries) {
            [xml]$labels = Read-PackageText $entry
            foreach ($label in $labels.SelectNodes('//*[local-name()="label"]')) {
                if ($label.GetAttribute('enabled') -eq '1' -and $label.GetAttribute('removed') -ne '1') { $ids += $label.GetAttribute('id').Trim('{}') }
            }
        }
        if ($ids.Count -ne 1 -or $ids[0] -ne $ExpectedSensitivityLabelId.Trim('{}')) { throw 'Saved sensitivity label does not match the approved public classification.' }
        Write-Output 'PASS: saved sensitivity label matches the approved public classification; no label metadata was stripped.'
    } else {
        Write-Output 'Sensitivity classification preserved; public classification approval not supplied.'
    }
    Write-Output 'PASS: workbook save-folder hint removed; no personal user-folder paths in XML or readable VBA reference metadata.'
} finally { $zip.Dispose() }
