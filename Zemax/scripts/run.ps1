# run.ps1 -- ASCII only on purpose (so it can never be broken by an encoding problem).
#
# PowerShell 5.1 decodes a BOM-less .ps1 using the system ANSI code page (GBK on
# Chinese Windows). The bridge scripts contain Chinese text, so any editor that
# rewrites them as UTF-8 without BOM breaks parsing. This launcher repairs the
# encoding first, then forwards every argument to zemax.ps1.
#
# Always call this instead of zemax.ps1 directly:
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File run.ps1 report -File "x.zmx"
#
# NOTE: no [Parameter()] attribute on purpose -- an advanced parameter would add the
# common parameters (-OutVariable/-OutBuffer) and then a bridge argument like "-Out"
# becomes ambiguous. Plain $args forwards everything untouched.
param()

$dir = $PSScriptRoot
$utf8Bom = New-Object System.Text.UTF8Encoding($true)

function Repair-Encoding([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { return }
    $text = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    [System.IO.File]::WriteAllText($Path, $text, $utf8Bom)
    Write-Host ("[run.ps1] repaired UTF-8 BOM: {0}" -f (Split-Path -Leaf $Path))
}

Repair-Encoding (Join-Path $dir 'zemax-lib.ps1')
Repair-Encoding (Join-Path $dir 'zemax.ps1')
Repair-Encoding (Join-Path (Split-Path -Parent $dir) 'config.json')

& (Join-Path $dir 'zemax.ps1') @args
exit $LASTEXITCODE
