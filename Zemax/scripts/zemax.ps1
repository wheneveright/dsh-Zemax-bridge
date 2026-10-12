<#
.SYNOPSIS
  DSH <-> Ansys Zemax OpticStudio 2024 桥接命令行（Phase 0）。

.DESCRIPTION
  每次调用会拉起一个无界面 OpticStudio 实例、执行一个命令、把结果落到一次性的 run 目录，
  然后关闭实例。冷启动约 5-15 秒，属于正常开销。

  必须用 Windows PowerShell 5.1 运行（ZOS-API 是 .NET Framework 程序集）：
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File zemax.ps1 <命令> [参数]

  命令：
    info            环境与许可自检（不读写镜头文件）
    list-analyses   导出全部 AnalysisIDM / MeritOperandType 名称，便于查可用分析
    prescription    导出系统参数 + 逐面处方表
    mtf             导出 FFT MTF 全曲线 + 指定空间频率处的取值
    spot            导出点列相关分析（RMS 波前/点列半径 / 包围能量与 r50-r90）
    report          处方 + MTF + 点列汇总，并生成 Excel / Word 报告
    design          按规格 JSON 建模 → 设变量 → 建评价函数 → 优化 → 对焦 → 存 .zos，
                    再对存盘文件做全套提取与验收比对（-Spec 指定规格文件）

.EXAMPLE
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File zemax.ps1 info
.EXAMPLE
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File zemax.ps1 report -File "C:\...\Doublet.zmx"
.EXAMPLE
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File zemax.ps1 mtf -File x.zmx -MaxFrequency 120 -Freqs 30,60,90
.EXAMPLE
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File zemax.ps1 design -Spec ".\design.json" -Out ".\run"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet('info', 'list-analyses', 'prescription', 'mtf', 'spot', 'report', 'design')]
    [string]$Command,

    [Parameter(Position = 1)]
    [string]$File,

    # design 命令的设计规格 JSON
    [string]$Spec,

    [string]$Out,
    [double]$MaxFrequency = 0,
    # 注意：类型故意用 object。用 -File 调用时参数以字符串传入，
    # 若声明为 [double[]]，"10,20,30" 会被按区域设置当成千位分隔符解析成 10203040。
    [object]$Freqs,
    [string]$Config,
    [switch]$NoOffice,
    [switch]$KeepOpen
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'zemax-lib.ps1')

$script:ToolName = 'dsh-zemax-bridge'
$script:ToolVersion = $script:ZemaxLibVersion

function ConvertTo-FreqList {
    <#  把 "10,20,30" / @(10,20) / 混合嵌套统一成 double[]。 #>
    param($Value)
    $list = New-Object Collections.Generic.List[double]
    if ($null -eq $Value) { return $list.ToArray() }
    foreach ($item in @($Value)) {
        if ($null -eq $item) { continue }
        if ($item -is [string]) {
            foreach ($piece in ($item -split '[,;\s]+')) {
                if ($piece -ne '') { $list.Add([double]$piece) }
            }
        } elseif ($item -is [System.Array]) {
            foreach ($piece in $item) { $list.Add([double]$piece) }
        } else {
            $list.Add([double]$item)
        }
    }
    return $list.ToArray()
}

$cfg = Get-ZemaxConfig -Path $Config
if ($MaxFrequency -le 0) { $MaxFrequency = [double]$cfg.analysis.mtfMaxFrequency }
$Freqs = ConvertTo-FreqList $Freqs
if ($Freqs.Count -eq 0) { $Freqs = ConvertTo-FreqList $cfg.analysis.mtfFreqs }

$needsFile = @('prescription', 'mtf', 'spot', 'report') -contains $Command
if ($needsFile) {
    if (-not $File) { throw "-File 必填：$Command 命令需要指定镜头文件（.zmx/.zos）" }
    if (-not (Test-Path -LiteralPath $File)) { throw "镜头文件不存在：$File" }
    $File = (Resolve-Path -LiteralPath $File).Path
}

# design 命令：先读规格（run 目录名要取自规格里的设计名）
$designSpec = $null
if ($Command -eq 'design') {
    if (-not $Spec) { throw "-Spec 必填：design 命令需要设计规格 JSON" }
    if (-not (Test-Path -LiteralPath $Spec)) { throw "找不到设计规格：$Spec" }
    $Spec = (Resolve-Path -LiteralPath $Spec).Path
    $designSpec = Get-Content -LiteralPath $Spec -Raw -Encoding UTF8 | ConvertFrom-Json
}

# ---- run 目录 ----
if ($Out) {
    New-Item -ItemType Directory -Path $Out -Force | Out-Null
    $runDir = (Resolve-Path -LiteralPath $Out).Path
} else {
    $label = if ($designSpec -and $designSpec.name) { [string]$designSpec.name }
    elseif ($File) { [IO.Path]::GetFileNameWithoutExtension($File) }
    else { $Command }
    $runDir = New-ZemaxRunDirectory -Config $cfg -LensPath $File -Label $label
}
Start-ZemaxLog -Path (Join-Path $runDir 'run.log')
Write-ZemaxLog "$script:ToolName v$script:ToolVersion  命令=$Command"
Write-ZemaxLog "run 目录：$runDir"

# ---- 连接 ----
$app = $null
$payload = $null
try {
    $installDir = Initialize-ZemaxApi -Config $cfg
    $app = Connect-ZemaxApplication -Config $cfg

    $payload = [ordered]@{
        tool        = $script:ToolName
        version     = $script:ToolVersion
        command     = $Command
        generatedAt = (Get-Date).ToString('s')
        host        = [ordered]@{
            machine    = $env:COMPUTERNAME
            user       = $env:USERNAME
            powershell = $PSVersionTable.PSVersion.ToString()
            clr        = [System.Runtime.InteropServices.RuntimeInformation]::FrameworkDescription
        }
        zemax       = [ordered]@{
            install    = $installDir
            license    = "$($app.LicenseStatus)"
            apiValid   = [bool]$app.IsValidLicenseForAPI
            samplesDir = "$($app.SamplesDir)"
        }
        source      = $null
        system      = $null
        prescription = @()
        mtf         = $null
        spot        = @()
        spotDiagram = $null
        design      = $null
        metrics     = $null
        acceptance  = @()
        errors      = @()
    }

    # ---- design：按规格建模 → 优化 → 存盘，然后以落盘文件为源继续走提取流程 ----
    if ($Command -eq 'design') {
        $system = $app.PrimarySystem
        $designLog = New-Object Collections.Generic.List[string]
        Write-ZemaxLog "设计规格：$($designSpec.name)"

        foreach ($line in (New-ZemaxSequentialSystem -System $system -Spec $designSpec)) { $designLog.Add($line); Write-ZemaxLog "  $line" }
        foreach ($line in (Set-ZemaxSurfaceData -System $system -Spec $designSpec)) { $designLog.Add($line); Write-ZemaxLog "  $line" }
        foreach ($line in (Set-ZemaxVariables -System $system -Spec $designSpec)) { $designLog.Add($line); Write-ZemaxLog "  $line" }
        foreach ($line in (Set-ZemaxMeritFunction -System $system -Spec $designSpec)) { $designLog.Add($line); Write-ZemaxLog "  $line" }

        $opt = Invoke-ZemaxOptimization -System $system -Spec $designSpec
        foreach ($line in $opt.log) { $designLog.Add($line); Write-ZemaxLog "  $line" }

        $zosName = if ($designSpec.save -and $designSpec.save.fileName) { [string]$designSpec.save.fileName } else { 'design.zos' }
        $zosPath = Save-ZemaxSystem -System $system -Path (Join-Path $runDir $zosName)
        $designLog.Add("已保存设计文件：$zosPath")
        Write-ZemaxLog "已保存设计文件：$zosPath"

        $payload.design = [ordered]@{
            name     = "$($designSpec.name)"
            notes    = "$($designSpec.notes)"
            specFile = $Spec
            spec     = $designSpec
            log      = @($designLog)
            optimize = $opt.result
            saved    = $zosPath
        }

        # 以落盘文件为唯一真源：重新加载它，后面的处方/指标/MTF/点列全部基于该文件
        $File = $zosPath
        Write-ZemaxLog "重新加载已存盘的设计文件，后续提取以它为准"
    }

    if ($File) {
        $fp = Get-FileFingerprint -Path $File
        $payload.source = $fp
        Write-ZemaxLog "源文件：$($fp.path)  ($($fp.bytes) B, sha256=$($fp.sha256.Substring(0,12))…)"

        $system = $app.PrimarySystem
        $loaded = $system.LoadFile($File, $false)
        if ($loaded -eq $false) { throw "OpticStudio 无法加载文件：$File" }

        $payload.system = Get-ZemaxSystemSummary -System $system
        Write-ZemaxLog ("系统载入：mode={0} 面数={1} 视场={2} 波长={3}" -f `
                $payload.system.mode, (Get-Prop $system.LDE 'NumberOfSurfaces' 0), `
                $payload.system.fields.Count, $payload.system.wavelengths.Count)

        if ($cfg.analysis.copyInputFile -and $Command -ne 'design') {
            $copyName = [IO.Path]::GetFileName($File)
            $copyPath = Join-Path $runDir "input_$copyName"
            Copy-Item -LiteralPath $File -Destination $copyPath -Force
            $payload.source['copiedTo'] = $copyPath
            Write-ZemaxLog "已复制输入到 run 目录：$copyPath"
        }

        if (@('prescription', 'report', 'design') -contains $Command) {
            $payload.prescription = Get-ZemaxPrescription -System $system
            $csv = Write-ZemaxCsv -Rows $payload.prescription -Path (Join-Path $runDir 'prescription.csv')
            Write-ZemaxLog "处方表：$($payload.prescription.Count) 面 -> $csv"
        }

        if (@('mtf', 'report', 'design') -contains $Command) {
            Write-ZemaxLog "运行 FFT MTF（最高频率 $MaxFrequency lp/mm）…"
            $mtf = Get-ZemaxMtf -System $system -MaxFrequency $MaxFrequency -Freqs $Freqs `
                -MaxPoints ([int]$cfg.limits.maxSeriesPoints)
            $payload.mtf = [ordered]@{
                maxFrequency = ConvertTo-SafeNumber $MaxFrequency 6
                requested    = @($Freqs)
                analysis     = $mtf.analysis
                atFreqs      = $mtf.atFreqs
            }
            $rows = @()
            foreach ($s in $mtf.analysis.series) {
                foreach ($p in $s.points) {
                    $rows += [ordered]@{ group = $s.group; series = $s.label; frequency = $p[0]; mtf = $p[1] }
                }
            }
            $csv = Write-ZemaxCsv -Rows $rows -Path (Join-Path $runDir 'mtf.csv')
            Write-ZemaxLog "MTF：状态=$($mtf.analysis.status) 曲线组=$($mtf.analysis.dataSeriesCount) 曲线=$($mtf.analysis.series.Count) 行=$($rows.Count) -> $csv"
            foreach ($row in $mtf.atFreqs) {
                foreach ($ps in $row.perSeries) {
                    Write-ZemaxLog ("  MTF@{0} lp/mm [{1} | {2}] = {3}" -f $row.frequency, $ps.group, $ps.label, $ps.value)
                }
            }
        }

        if (@('spot', 'report', 'design') -contains $Command) {
            Write-ZemaxLog '运行点列相关分析…'
            $spotData = Get-ZemaxSpot -System $system -MaxPoints ([int]$cfg.limits.maxSeriesPoints)
            $payload.spot = @($spotData.series)
            $payload.spotDiagram = $spotData.diagram

            $rows = @()
            $summary = @()
            foreach ($a in $payload.spot) {
                foreach ($s in $a.series) {
                    foreach ($p in $s.points) {
                        $rows += [ordered]@{
                            analysis = $a.id; quantity = $a.quantity; group = $s.group
                            series = $s.label; x = $p[0]; y = $p[1]
                        }
                    }
                    # 每条曲线的极值/末值，便于人快速读和交叉检查
                    $ys = @($s.points | ForEach-Object { $_[1] })
                    $xs = @($s.points | ForEach-Object { $_[0] })
                    $r50 = $null; $r80 = $null; $r90 = $null
                    if ($s.Contains('energyRadii')) {
                        foreach ($er in $s.energyRadii) {
                            switch ($er.fraction) {
                                0.5 { $r50 = $er.radiusUm }
                                0.8 { $r80 = $er.radiusUm }
                                0.9 { $r90 = $er.radiusUm }
                            }
                        }
                    }
                    $summary += [ordered]@{
                        analysis = $a.id
                        quantity = $a.quantity
                        group    = $s.group
                        series   = $s.label
                        xLabel   = $s.xLabel
                        xMin     = ($xs | Measure-Object -Minimum).Minimum
                        xMax     = ($xs | Measure-Object -Maximum).Maximum
                        yMin     = ($ys | Measure-Object -Minimum).Minimum
                        yMax     = ($ys | Measure-Object -Maximum).Maximum
                        yLast    = $ys[-1]
                        r50_um   = $r50
                        r80_um   = $r80
                        r90_um   = $r90
                    }
                }
                $q = if ($a.settings.Data) { "Data=$($a.settings.Data)" } elseif ($a.settings.Type) { "Type=$($a.settings.Type)" } else { '' }
                Write-ZemaxLog ("  {0} [{1}] {2}：状态={3} 曲线组={4} 曲线={5} $($a.note)" -f `
                        $a.id, $q, $a.quantity, $a.status, $a.dataSeriesCount, $a.series.Count)
                foreach ($row in ($summary | Where-Object { $_.analysis -eq $a.id -and $_.quantity -eq $a.quantity })) {
                    $radii = if ($null -ne $row.r50_um) { "  r50/r80/r90 = {0}/{1}/{2} µm" -f $row.r50_um, $row.r80_um, $row.r90_um } else { '' }
                    Write-ZemaxLog ("     {0}/{1}: y 范围 {2} … {3}（x: {4} → {5}）{6}" -f `
                            $row.group, $row.series, $row.yMin, $row.yMax, $row.xMin, $row.xMax, $radii)
                }
            }
            $csv = Write-ZemaxCsv -Rows $rows -Path (Join-Path $runDir 'spot.csv')
            Write-ZemaxLog "点列曲线数据：$($rows.Count) 行 -> $csv"
            $csv2 = Write-ZemaxCsv -Rows $summary -Path (Join-Path $runDir 'spot-summary.csv')
            Write-ZemaxLog "点列曲线摘要：$($summary.Count) 行 -> $csv2"

            $d = $payload.spotDiagram
            if ($d) {
                Write-ZemaxLog ("  {0}（点列图原始聚合量，口径存疑、不入报告）：最大半径读数={1}  平均半径读数={2}  半宽={3}×{4} mm  视场×波长={5}×{6}" -f `
                        $d.id, $d.metrics.maxRadiusRaw, $d.metrics.meanRadiusRaw,
                        $d.metrics.halfWidthX, $d.metrics.halfWidthY, $d.metrics.numberOfFields, $d.metrics.numberOfWaves)
                if ($d.note) { Write-ZemaxLog "     $($d.note)" }
            }
        }

        # ---- design：把验收指标（评价函数操作数）与规格判据逐条比对 ----
        if ($Command -eq 'design') {
            Write-ZemaxLog '读取设计验收指标（评价函数操作数 + 由分析数据派生的指标）…'
            $payload.metrics = Get-ZemaxDesignMetrics -System $system -Spec $designSpec

            # 把 MTF / RMSField / 包围能量里派生的指标并进同一张表，供判据引用
            $derived = Get-ZemaxDerivedMetrics -Mtf $payload.mtf -Spot $payload.spot -Freqs $Freqs
            foreach ($k in $derived.Keys) { $payload.metrics[$k] = $derived[$k] }

            # 纵向像差曲线（每波长一条，横轴相对光瞳）：x=0 处的 y 就是该波长的近轴焦点偏移，
            # 这正是经典意义上的"轴向色差"数据表，可逐点复核。
            # 注：CardinalPoints 分析在 ZOS-API 里不返回任何数值（曲线/网格/消息全为 0），
            # 所以 EFL / F# 只能走 EFFL / WFNO 操作数，无法从该分析读。
            Write-ZemaxLog '读取纵向像差曲线（各波长近轴焦点偏移）…'
            $long = Invoke-ZemaxAnalysis -System $system -AnalysisName 'LongitudinalAberration' -MaxPoints 2000
            $payload.longitudinalAberration = $long
            $paraxial = @()
            foreach ($s in $long.series) {
                $zeroPoint = $s.points | Sort-Object { $_[0] } | Select-Object -First 1
                if ($null -ne $zeroPoint) {
                    $paraxial += [ordered]@{ series = $s.label; relPupil = $zeroPoint[0]; focusOffsetMm = $zeroPoint[1] }
                }
            }
            foreach ($p in $paraxial) {
                Write-ZemaxLog ("  波长 {0}：近轴焦点偏移 = {1} mm（相对光瞳 {2}）" -f $p.series, $p.focusOffsetMm, $p.relPupil)
            }
            if ($paraxial.Count -ge 3) {
                $fcPara = [double]$paraxial[0].focusOffsetMm - [double]$paraxial[-1].focusOffsetMm
                $vals = @($paraxial | ForEach-Object { [double]$_.focusOffsetMm })
                $spanPara = ($vals | Measure-Object -Maximum).Maximum - ($vals | Measure-Object -Minimum).Minimum
                Write-ZemaxLog ("  近轴 F–C = {0} mm，近轴跨度 = {1} mm" -f $fcPara, $spanPara)
                $payload.metrics['axial_color_paraxial_fc_mm'] = [ordered]@{
                    label = '近轴轴向色差 F–C（纵向像差曲线 x=0）'; unit = 'mm'; value = ConvertTo-SafeNumber $fcPara 6
                    operand = 'derived:LongitudinalAberration'
                }
                $payload.metrics['axial_color_paraxial_span_mm'] = [ordered]@{
                    label = '近轴焦点跨度（纵向像差曲线 x=0）'; unit = 'mm'; value = ConvertTo-SafeNumber $spanPara 6
                    operand = 'derived:LongitudinalAberration'
                }
            }

            # 轴向色差的独立测量放在最后：它会临时改波长权重并对焦，不写盘
            Write-ZemaxLog '逐波长单独对焦，测轴向色差（独立于 AXCL 的近轴定义）…'
            $chroma = Get-ZemaxChromaticFocus -System $system
            $payload.chromaticFocus = $chroma
            foreach ($p in $chroma.positions) {
                Write-ZemaxLog ("  波长 {0} ({1} µm) 最佳像面 = {2} mm" -f $p.wave, $p.lambdaUm, $p.focusMm)
            }
            Write-ZemaxLog ("  F–C 焦点差 = {0} mm，全场跨度 = {1} mm" -f $chroma.shiftFCmm, $chroma.spanMm)
            $payload.metrics['focus_shift_fc_mm'] = [ordered]@{
                label = '轴向色差 F–C（逐波长最佳焦点差）'; unit = 'mm'; value = $chroma.shiftFCmm
                operand = 'derived:QuickFocus-per-wavelength'
            }
            $payload.metrics['focus_span_mm'] = [ordered]@{
                label = '各波长最佳焦点跨度（二次光谱量级）'; unit = 'mm'; value = $chroma.spanMm
                operand = 'derived:QuickFocus-per-wavelength'
            }

            foreach ($k in $payload.metrics.Keys) {
                $m = $payload.metrics[$k]
                Write-ZemaxLog ("  {0} = {1} {2}   [{3} Wave={4}]" -f $m.label, $m.value, $m.unit, $m.operand, (Get-Prop $m 'wave' 0))
            }
            $payload.acceptance = @(Test-ZemaxAcceptance -Metrics $payload.metrics -Spec $designSpec)
            $passCount = @($payload.acceptance | Where-Object { $_.pass -eq $true }).Count
            Write-ZemaxLog "验收：$passCount / $(@($payload.acceptance).Count) 条通过"
            foreach ($row in $payload.acceptance) {
                $mark = if ($row.pass -eq $true) { '通过' } elseif ($row.pass -eq $false) { '未通过' } else { '无数据' }
                Write-ZemaxLog ("  [{0}] {1}：判据 {2} {3}，实测 {4} {3}" -f $mark, $row.label, $row.rule, $row.unit, $row.actual)
            }
            $csvA = Write-ZemaxCsv -Rows $payload.acceptance -Path (Join-Path $runDir 'acceptance.csv')
            Write-ZemaxLog "验收表：$csvA"
            $csvM = Write-ZemaxCsv -Rows @($payload.metrics.Keys | ForEach-Object {
                    $m = $payload.metrics[$_]
                    [ordered]@{ metric = $_; label = $m.label; operand = $m.operand; wave = $m.wave; unit = $m.unit; value = $m.value }
                }) -Path (Join-Path $runDir 'metrics.csv')
            Write-ZemaxLog "指标表：$csvM"
        }
    }
    else {
        # info / list-analyses
        $system = $app.PrimarySystem
        $payload.system = [ordered]@{
            mode         = "$(Get-Prop $system 'Mode' 'Unknown')"
            surfaces     = [int](Get-Prop $system.LDE 'NumberOfSurfaces' 0)
            samplesDir   = "$($app.SamplesDir)"
            zemaxRoot    = "$env:USERPROFILE\Documents\Zemax"
        }
        Write-ZemaxLog ("info：license={0} apiValid={1} install={2}" -f $app.LicenseStatus, $app.IsValidLicenseForAPI, $installDir)
    }

    if ($Command -eq 'list-analyses') {
        $ids = [enum]::GetNames([ZOSAPI.Analysis.AnalysisIDM])
        $ops = [enum]::GetNames([ZOSAPI.Editors.MFE.MeritOperandType])
        $solves = [enum]::GetNames([ZOSAPI.Editors.SolveType])
        [IO.File]::WriteAllLines((Join-Path $runDir 'analysis-ids.txt'), $ids, (New-Object Text.UTF8Encoding($false)))
        [IO.File]::WriteAllLines((Join-Path $runDir 'merit-operand-types.txt'), $ops, (New-Object Text.UTF8Encoding($false)))
        [IO.File]::WriteAllLines((Join-Path $runDir 'solve-types.txt'), $solves, (New-Object Text.UTF8Encoding($false)))
        Write-ZemaxLog ("AnalysisIDM={0}  MeritOperandType={1}  SolveType={2}" -f $ids.Count, $ops.Count, $solves.Count)
        Write-ZemaxLog "已写出 analysis-ids.txt / merit-operand-types.txt / solve-types.txt"
    }

    $dataPath = Write-ZemaxJson -Object $payload -Path (Join-Path $runDir 'data.json')
    Write-ZemaxLog "数据：$dataPath"

    # ---- Excel / Word 报告 ----
    if (@('report', 'design') -contains $Command -and -not $NoOffice) {
        $py = $cfg.pythonExe
        $builder = Join-Path $PSScriptRoot 'build_report.py'
        if (-not (Test-Path -LiteralPath $py)) {
            Write-ZemaxLog "警告：找不到 pythonExe（$py），跳过 Excel/Word 生成。可用 -NoOffice 显式跳过。"
        }
        elseif (-not (Test-Path -LiteralPath $builder)) {
            Write-ZemaxLog "警告：找不到 build_report.py（$builder），跳过 Excel/Word 生成。"
        }
        else {
            Write-ZemaxLog '正在生成 Excel / Word 报告…'
            $old = $ErrorActionPreference
            $ErrorActionPreference = 'Continue'
            $out = & $py $builder '--data' $dataPath '--outdir' $runDir 2>&1
            $code = $LASTEXITCODE
            $ErrorActionPreference = $old
            foreach ($line in @($out)) { Write-ZemaxLog "  [report] $line" }
            if ($code -ne 0) { throw "报告生成失败（build_report.py 退出码 $code）" }
        }
    }

    Write-ZemaxLog '完成'
}
finally {
    if ($null -ne $app -and -not $KeepOpen) {
        try { $app.CloseApplication(); Write-ZemaxLog 'OpticStudio 实例已关闭' } catch { }
    }
}

# 供 agent 解析的机器可读尾行
Write-Host ""
Write-Host "RUN_DIR=$runDir"
Write-Host "DATA_JSON=$(Join-Path $runDir 'data.json')"
if ($payload -and $payload.mtf) {
    foreach ($row in $payload.mtf.atFreqs) {
        foreach ($ps in $row.perSeries) {
            Write-Host ("MTF@{0}lp_per_mm[{1}|{2}]={3}" -f $row.frequency, $ps.group, $ps.label, $ps.value)
        }
    }
}
if ($payload -and $payload.design) {
    Write-Host "DESIGN_SAVED=$($payload.design.saved)"
    foreach ($k in $payload.metrics.Keys) {
        Write-Host ("METRIC[{0}]={1}{2}" -f $k, $payload.metrics[$k].value, $payload.metrics[$k].unit)
    }
    $passCount = @($payload.acceptance | Where-Object { $_.pass -eq $true }).Count
    Write-Host ("ACCEPTANCE={0}/{1}" -f $passCount, @($payload.acceptance).Count)
    foreach ($row in $payload.acceptance) {
        Write-Host ("CRITERION[{0}]={1} actual={2} rule={3}" -f $row.metric, `
            $(if ($row.pass -eq $true) { 'PASS' } elseif ($row.pass -eq $false) { 'FAIL' } else { 'NODATA' }), $row.actual, $row.rule)
    }
}
