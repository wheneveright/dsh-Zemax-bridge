# install.ps1 -- ASCII only on purpose (never breaks on an encoding problem).
#
# Installs the DSH <-> Ansys Zemax OpticStudio bridge into a target DSH workspace:
#   1. copies the bridge scripts + examples + reference lists
#   2. copies the DSH skill (workspace .dsh\skills\zemax, and/or the user-level ~\.dsh\skills)
#   3. auto-detects OpticStudio / ZOSAPI_NetHelper.dll / a Python for Office reports
#      and writes config.json (backing up any existing one)
#   4. runs a self test (info) and reports license / API availability
#
# Usage (from the package root):
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -Workspace "D:\MyWorkspace"
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -Workspace "D:\MyWorkspace" -Force
#
# Notes:
#   * No admin rights needed: everything is written under the target workspace and, for
#     -SkillScope user, under %USERPROFILE%\.dsh\skills.
#   * Safe to re-run: scripts are overwritten, but an existing config.json is kept unless
#     -Force is given (so your machine-specific paths are not clobbered).
param(
    [string]$Workspace = (Get-Location).Path,
    [ValidateSet('workspace', 'user', 'both', 'none')]
    [string]$SkillScope = 'workspace',
    [string]$ZemaxInstall,
    [string]$PythonExe,
    [switch]$Force,
    [switch]$SkipSelfTest
)

$ErrorActionPreference = 'Stop'
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Say([string]$m) { Write-Host $m }
function Step([string]$m) { Write-Host ""; Write-Host ("== " + $m) }

# ---------------------------------------------------------------- 0. locate payload
# Two layouts are supported, so the same script works from the release zip and from a git clone:
#   release zip : <root>\install.ps1  +  <root>\payload\Zemax\scripts\...   (payload = install content)
#   git repo    : <root>\install.ps1  +  <root>\Zemax\scripts\...           (root IS the install content)
$pkgRoot = $PSScriptRoot
$payload = Join-Path $pkgRoot 'payload'
if (Test-Path -LiteralPath (Join-Path $payload 'Zemax\scripts')) {
    $layout = 'release (payload)'
} elseif (Test-Path -LiteralPath (Join-Path $pkgRoot 'Zemax\scripts')) {
    $payload = $pkgRoot
    $layout = 'git repo'
} else {
    throw "cannot find the bridge files next to install.ps1: expected 'payload\Zemax\scripts' or 'Zemax\scripts'"
}
if (-not (Test-Path -LiteralPath $Workspace)) { New-Item -ItemType Directory -Path $Workspace -Force | Out-Null }
$Workspace = (Resolve-Path -LiteralPath $Workspace).Path

Say "DSH Zemax bridge installer"
Say ("  package : " + $pkgRoot)
Say ("  layout  : " + $layout)
Say ("  target  : " + $Workspace)

# ---------------------------------------------------------------- 1. copy files
Step "1/5 copy bridge files"
$toolDir = Join-Path $Workspace 'Zemax'
$scriptSrc = Join-Path $payload 'Zemax\scripts'
$scriptDst = Join-Path $toolDir 'scripts'

foreach ($d in @($toolDir, $scriptDst,
        (Join-Path $toolDir 'examples'), (Join-Path $toolDir 'reference'),
        (Join-Path $toolDir 'runs'), (Join-Path $toolDir 'scratch'))) {
    if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

Copy-Item (Join-Path $scriptSrc '*') -Destination $scriptDst -Force
foreach ($f in @('README.md', 'config.example.json')) {
    $src = Join-Path $payload ("Zemax\" + $f)
    if (Test-Path -LiteralPath $src) { Copy-Item $src -Destination (Join-Path $toolDir $f) -Force }
}
# examples / reference live under Zemax\ in the release zip and at the repo root in a git clone
$examplesSrc = Join-Path $payload 'Zemax\examples'
if (-not (Test-Path -LiteralPath $examplesSrc)) { $examplesSrc = Join-Path $payload 'examples' }
$referenceSrc = Join-Path $payload 'Zemax\reference'
if (-not (Test-Path -LiteralPath $referenceSrc)) { $referenceSrc = Join-Path $payload 'reference' }
if (Test-Path -LiteralPath $examplesSrc) { Copy-Item (Join-Path $examplesSrc '*') -Destination (Join-Path $toolDir 'examples') -Recurse -Force }
if (Test-Path -LiteralPath $referenceSrc) { Copy-Item (Join-Path $referenceSrc '*') -Destination (Join-Path $toolDir 'reference') -Force }
Say ("  scripts  -> " + $scriptDst)
Say ("  examples -> " + (Join-Path $toolDir 'examples'))
Say ("  reference-> " + (Join-Path $toolDir 'reference'))

# Repair UTF-8 BOM on the copied PowerShell files (PowerShell 5.1 needs it for Chinese
# text; pure-ASCII files are left untouched so their bytes match the source exactly).
Get-ChildItem -LiteralPath $scriptDst -Filter *.ps1 -File | ForEach-Object {
    $bytes = [IO.File]::ReadAllBytes($_.FullName)
    $hasNonAscii = $false
    foreach ($b in $bytes) { if ($b -gt 127) { $hasNonAscii = $true; break } }
    if ($hasNonAscii -and -not ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)) {
        $t = [IO.File]::ReadAllText($_.FullName, [Text.Encoding]::UTF8)
        [IO.File]::WriteAllText($_.FullName, $t, $utf8Bom)
        Say ("  BOM added: " + $_.Name)
    }
}

# ---------------------------------------------------------------- 2. copy the skill
Step "2/5 install DSH skill"
$skillSrc = Join-Path $payload '.dsh\skills\zemax\SKILL.md'
$targets = New-Object Collections.Generic.List[string]
if ($SkillScope -eq 'workspace' -or $SkillScope -eq 'both') {
    $targets.Add((Join-Path $Workspace '.dsh\skills\zemax'))
}
if ($SkillScope -eq 'user' -or $SkillScope -eq 'both') {
    $targets.Add((Join-Path $env:USERPROFILE '.dsh\skills\zemax'))
}
if ($targets.Count -eq 0) { Say "  skipped (-SkillScope none)" }
foreach ($t in $targets) {
    if (-not (Test-Path -LiteralPath $t)) { New-Item -ItemType Directory -Path $t -Force | Out-Null }
    Copy-Item $skillSrc -Destination (Join-Path $t 'SKILL.md') -Force
    Say ("  SKILL.md -> " + (Join-Path $t 'SKILL.md'))
}
Say "  (DSH scans <workspace>\.dsh\skills and %USERPROFILE%\.dsh\skills; new skills appear"
Say "   in the next model step, no restart needed)"

# ---------------------------------------------------------------- 3. auto detect
Step "3/5 detect OpticStudio / helper dll / python"

function Find-ZemaxInstall {
    # registry: HKLM\SOFTWARE\ANSYS, Inc.\ZemaxOpticStudio v<ver> -> installFolder
    foreach ($hive in @('HKLM:\SOFTWARE\ANSYS, Inc.', 'HKLM:\SOFTWARE\WOW6432Node\ANSYS, Inc.')) {
        if (-not (Test-Path $hive)) { continue }
        $keys = Get-ChildItem $hive -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -like 'ZemaxOpticStudio*' }
        foreach ($k in $keys) {
            $folder = (Get-ItemProperty $k.PSPath -ErrorAction SilentlyContinue).installFolder
            if ($folder -and (Test-Path -LiteralPath (Join-Path $folder 'ZOSAPI.dll'))) { return $folder }
            if ($folder -and (Test-Path -LiteralPath $folder)) { return $folder }
        }
    }
    # fallback: common locations containing ZOSAPI.dll
    $candidates = @(
        'C:\Program Files\Zemax OpticStudio', 'C:\Program Files\Ansys Zemax OpticStudio',
        'C:\Program Files\ANSYS Inc\Zemax', 'C:\Zemax', 'D:\Zemax', 'D:\Ansys Zemax', 'E:\Ansys Zemax'
    )
    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath (Join-Path $c 'ZOSAPI.dll')) { return $c }
    }
    foreach ($drive in (Get-PSDrive -PSProvider FileSystem).Name) {
        foreach ($sub in @('Ansys Zemax', 'Zemax OpticStudio', 'Program Files\Zemax OpticStudio')) {
            $p = "${drive}:\$sub"
            if (Test-Path -LiteralPath (Join-Path $p 'ZOSAPI.dll')) { return $p }
        }
    }
    return $null
}

function Find-Python {
    $found = New-Object Collections.Generic.List[string]
    $runtimeRoot = Join-Path $env:USERPROFILE '.dsh\dsh-runtimes'
    if (Test-Path -LiteralPath $runtimeRoot) {
        foreach ($p in (Get-ChildItem $runtimeRoot -Recurse -Filter python.exe -ErrorAction SilentlyContinue |
                Where-Object { $_.FullName -like '*dependencies\python\python.exe' })) {
            $found.Add($p.FullName)
        }
    }
    $cmd = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($cmd) { $found.Add($cmd.Source) }
    foreach ($f in $found) {
        # prefer an interpreter that can actually build the reports
        $probe = & $f -c "import openpyxl, docx; print('ok')" 2>$null
        if ($LASTEXITCODE -eq 0) { return $f }
    }
    if ($found.Count -gt 0) { return $found[0] }
    return $null
}

if (-not $ZemaxInstall) { $ZemaxInstall = Find-ZemaxInstall }
$helper = Join-Path $env:USERPROFILE 'Documents\Zemax\ZOS-API\Libraries\ZOSAPI_NetHelper.dll'
if (-not (Test-Path -LiteralPath $helper) -and $ZemaxInstall) {
    $alt = Join-Path $ZemaxInstall 'ZOSAPI_NetHelper.dll'
    if (Test-Path -LiteralPath $alt) { $helper = $alt }
}
if (-not $PythonExe) { $PythonExe = Find-Python }
$samples = Join-Path $env:USERPROFILE 'Documents\Zemax\SAMPLES'

Say ("  OpticStudio      : " + $(if ($ZemaxInstall) { $ZemaxInstall } else { "NOT FOUND - set zemaxInstall in config.json manually" }))
Say ("  ZOSAPI_NetHelper : " + $(if (Test-Path -LiteralPath $helper) { $helper } else { "NOT FOUND" }))
Say ("  python (reports) : " + $(if ($PythonExe) { $PythonExe } else { "NOT FOUND - Office reports will be skipped" }))
Say ("  samples          : " + $(if (Test-Path -LiteralPath $samples) { $samples } else { "not found (optional)" }))

# ---------------------------------------------------------------- 4. write config.json
Step "4/5 write config.json"
$configPath = Join-Path $toolDir 'config.json'
$writeConfig = $true
if ((Test-Path -LiteralPath $configPath) -and -not $Force) {
    $writeConfig = $false
    Say "  existing config.json kept (use -Force to overwrite); backup not needed"
}
if ($writeConfig) {
    if (Test-Path -LiteralPath $configPath) {
        Copy-Item $configPath ($configPath + '.bak') -Force
        Say "  existing config.json backed up to config.json.bak"
    }
    $cfg = [ordered]@{
        '_comment'          = 'Path config for the DSH <-> Ansys Zemax OpticStudio bridge. Generated by install.ps1; edit only this file when moving machines or folders. Override with ZEMAX_CONFIG / ZEMAX_OUTPUT_ROOT / ZEMAX_SCRATCH_ROOT.'
        outputRoot          = (Join-Path $toolDir 'runs')
        scratchRoot         = (Join-Path $toolDir 'scratch')
        zemaxInstall        = $(if ($ZemaxInstall) { $ZemaxInstall } else { '' })
        zosApiHelper        = $helper
        samplesDir          = $samples
        pythonExe           = $(if ($PythonExe) { $PythonExe } else { '' })
        windowsPowerShell   = (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe')
        analysis            = [ordered]@{
            mtfMaxFrequency = 100.0
            mtfFreqs        = @(30.0, 50.0, 100.0)
            mtfSampleSize   = 0
            spotFieldNumber = 0
            copyInputFile   = $true
        }
        limits              = [ordered]@{
            analysisTimeoutSeconds = 600
            maxSeriesPoints        = 20000
        }
    }
    $json = $cfg | ConvertTo-Json -Depth 6
    [IO.File]::WriteAllText($configPath, $json, $utf8Bom)
    Say ("  written -> " + $configPath)
}

# ---------------------------------------------------------------- 5. self test
Step "5/5 self test"
if ($SkipSelfTest) {
    Say "  skipped (-SkipSelfTest)"
} else {
    $pwsh = (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe')
    $runner = Join-Path $scriptDst 'run.ps1'
    Say "  running: run.ps1 info"
    $out = & $pwsh -NoProfile -ExecutionPolicy Bypass -File $runner info 2>&1
    $code = $LASTEXITCODE
    foreach ($line in @($out)) { Say ("    " + $line) }
    if ($code -ne 0) {
        Say ""
        Say "  SELF TEST FAILED. Check: zemaxInstall / zosApiHelper in config.json,"
        Say "  and that the ANSYS license service is running (OpticStudio must be licensed for ZOS-API)."
    } else {
        Say "  self test OK"
    }
}

# ---------------------------------------------------------------- done
$runnerPath = Join-Path $scriptDst 'run.ps1'
$exampleSpec = Join-Path $toolDir 'examples\achromat_design.json'
Say ""
Say "Done. Next steps:"
Say "  1) open the target workspace in DSH and run a report:"
Say ("       powershell.exe -NoProfile -ExecutionPolicy Bypass -File " + $runnerPath + " report -File <lens.zmx>")
Say "  2) try a design run with the bundled example spec:"
Say ("       powershell.exe -NoProfile -ExecutionPolicy Bypass -File " + $runnerPath + " design -Spec " + $exampleSpec + " -Freqs 30,50,100")
Say "  3) read the tool readme and the skill:"
Say ("       " + (Join-Path $toolDir 'README.md'))
Say ("       " + (Join-Path $Workspace '.dsh\skills\zemax\SKILL.md'))
Say "  4) if OpticStudio or python was NOT FOUND above, edit config.json and re-run the self test"
