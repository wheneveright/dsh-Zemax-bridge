# normalize-encoding.ps1 -- ASCII only on purpose.
# PowerShell 5.1 decodes a BOM-less file using the system ANSI code page (GBK on
# Chinese Windows), so any UTF-8 Chinese text turns into garbage and breaks
# parsing -- this bit both the *.ps1 scripts and config.json (its Chinese
# "_comment" made ConvertFrom-Json fail outright).
#
# Run this after editing any *.ps1 in this folder or config.json:
#     powershell.exe -NoProfile -ExecutionPolicy Bypass -File normalize-encoding.ps1
#
# It never touches run outputs (data.json etc.): the Python report builder reads
# those as strict UTF-8 without BOM.
param([string]$Dir)

if (-not $Dir) { $Dir = $PSScriptRoot }

$utf8Bom = New-Object System.Text.UTF8Encoding($true)

Get-ChildItem -LiteralPath $Dir -Filter *.ps1 -File | ForEach-Object {
    $text = [System.IO.File]::ReadAllText($_.FullName, [System.Text.Encoding]::UTF8)
    [System.IO.File]::WriteAllText($_.FullName, $text, $utf8Bom)
    Write-Host ("normalized: {0}" -f $_.Name)
}

# config.json lives one level up from the scripts folder.
$configPath = Join-Path (Split-Path -Parent $Dir) 'config.json'
if (Test-Path -LiteralPath $configPath) {
    $text = [System.IO.File]::ReadAllText($configPath, [System.Text.Encoding]::UTF8)
    [System.IO.File]::WriteAllText($configPath, $text, $utf8Bom)
    Write-Host ("normalized: {0}" -f (Split-Path -Leaf $configPath))
}
