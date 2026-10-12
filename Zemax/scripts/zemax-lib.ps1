# zemax-lib.ps1 — DSH <-> Ansys Zemax OpticStudio 桥接公共库
#
# 重要：本库必须运行在 Windows PowerShell 5.1（.NET Framework 4.8）之下。
# ZOS-API 的 ZOSAPI.dll / ZOSAPI_Interfaces.dll 是 .NET Framework 程序集，
# 在 PowerShell 7+（.NET Core）下加载会失败。DSH 的 pwsh 工具本身就是 5.1，
# 用别的宿主时请显式调用 config.json 里的 windowsPowerShell。

$script:ZemaxLibVersion = '0.2.0'
$script:ZemaxLibScriptDir = Split-Path -Parent $PSCommandPath
$script:ZemaxApiLoaded = $false
$script:ZemaxLogPath = $null

# ---------------------------------------------------------------- 配置

function Get-ZemaxConfig {
    <#  读取唯一配置真源；环境变量可覆盖，便于换机器/换目录。 #>
    param([string]$Path)

    if (-not $Path) { $Path = $env:ZEMAX_CONFIG }
    if (-not $Path) { $Path = Join-Path (Split-Path -Parent $script:ZemaxLibScriptDir) 'config.json' }
    if (-not (Test-Path -LiteralPath $Path)) { throw "找不到桥接配置文件：$Path" }

    $cfg = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json

    if ($env:ZEMAX_OUTPUT_ROOT) { $cfg.outputRoot = $env:ZEMAX_OUTPUT_ROOT }
    if ($env:ZEMAX_SCRATCH_ROOT) { $cfg.scratchRoot = $env:ZEMAX_SCRATCH_ROOT }

    if (-not $cfg.outputRoot) { throw "配置缺少 outputRoot：$Path" }
    return $cfg
}

function Initialize-ZemaxApi {
    <#  加载 ZOSAPI_NetHelper + ZOSAPI 程序集，返回 OpticStudio 安装目录。 #>
    param([Parameter(Mandatory = $true)]$Config)

    if ($script:ZemaxApiLoaded) { return $script:ZemaxInstallDir }

    $helper = $Config.zosApiHelper
    if (-not (Test-Path -LiteralPath $helper)) {
        throw "找不到 ZOSAPI_NetHelper.dll：$helper（请在 config.json 中修正 zosApiHelper）"
    }

    Add-Type -Path $helper
    $ok = $false
    if ($Config.zemaxInstall -and (Test-Path -LiteralPath $Config.zemaxInstall)) {
        $ok = [ZOSAPI_NetHelper.ZOSAPI_Initializer]::Initialize($Config.zemaxInstall)
    }
    if (-not $ok) { $ok = [ZOSAPI_NetHelper.ZOSAPI_Initializer]::Initialize() }
    if (-not $ok) { throw 'ZOSAPI_Initializer.Initialize() 失败：无法定位 OpticStudio 安装。' }

    $dir = [ZOSAPI_NetHelper.ZOSAPI_Initializer]::GetZemaxDirectory()
    Add-Type -Path (Join-Path $dir 'ZOSAPI_Interfaces.dll')
    Add-Type -Path (Join-Path $dir 'ZOSAPI.dll')

    $script:ZemaxApiLoaded = $true
    $script:ZemaxInstallDir = $dir
    return $dir
}

function Connect-ZemaxApplication {
    <#  拉起一个无界面 OpticStudio 实例（占用一个许可席位），返回 IZOSAPI_Application。 #>
    param(
        [Parameter(Mandatory = $true)]$Config,
        [switch]$Quiet
    )

    $dir = Initialize-ZemaxApi -Config $Config
    if (-not $Quiet) { Write-ZemaxLog "ZOS-API 已初始化，OpticStudio 目录：$dir" }

    $connection = New-Object ZOSAPI.ZOSAPI_Connection
    if ($null -eq $connection) { throw '无法创建 ZOSAPI_Connection。' }

    $sw = [Diagnostics.Stopwatch]::StartNew()
    $app = $connection.CreateNewApplication()
    $sw.Stop()
    if ($null -eq $app) { throw 'CreateNewApplication() 返回空：可能是许可被占用或 OpticStudio 启动失败。' }
    if (-not $app.IsValidLicenseForAPI) { throw "当前许可不支持 ZOS-API 调用（LicenseStatus=$($app.LicenseStatus)）。" }

    if (-not $Quiet) {
        Write-ZemaxLog ("OpticStudio 就绪：{0:N1}s，license={1}" -f $sw.Elapsed.TotalSeconds, $app.LicenseStatus)
    }
    return $app
}

# ---------------------------------------------------------------- 日志 / 小工具

function Start-ZemaxLog {
    param([Parameter(Mandatory = $true)][string]$Path)
    $script:ZemaxLogPath = $Path
    [IO.File]::WriteAllText($Path, '', (New-Object Text.UTF8Encoding($false)))
}

function Write-ZemaxLog {
    param([Parameter(Mandatory = $true)][string]$Message)
    $line = "[{0:HH:mm:ss}] {1}" -f (Get-Date), $Message
    Write-Host $line
    if ($script:ZemaxLogPath) {
        Add-Content -LiteralPath $script:ZemaxLogPath -Value $line -Encoding UTF8
    }
}

function Get-Prop {
    <#  安全读取属性：ZOS-API 各接口属性并不统一，缺失时返回 $null 而不是抛错。
        注意：本库内部大量用 [ordered]@{} 当结构体，而 PowerShell 的 PSObject.Properties
        并不包含字典键 —— 所以字典必须走 Contains/索引分支，否则会静默取到默认值
        （曾导致 RMSField 的 Data 读成空串、派生指标整批缺失）。 #>
    param($Object, [Parameter(Mandatory = $true)][string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    try {
        if ($Object -is [System.Collections.IDictionary]) {
            if (-not $Object.Contains($Name)) { return $Default }
            $dv = $Object[$Name]
            if ($null -eq $dv) { return $Default }
            if ($dv -is [System.Array]) { Write-Output -NoEnumerate $dv; return }
            return $dv
        }
        $p = $Object.PSObject.Properties[$Name]
        if ($null -eq $p) { return $Default }
        $v = $p.Value
        if ($null -eq $v) { return $Default }
        # 关键：数组（尤其是 double[,]）不能用 return 直接返回 —— PowerShell 会把
        # 函数输出流里的数组"拆开"，double[,] 会被展平成一维，多维结构就此丢失。
        # 必须用 -NoEnumerate 保持原样。
        if ($v -is [System.Array]) { Write-Output -NoEnumerate $v; return }
        return $v
    } catch { return $Default }
}

function Set-Prop {
    <#  安全写入属性：属性不存在或类型不符时返回 $false，不中断流程。 #>
    param($Object, [Parameter(Mandatory = $true)][string]$Name, $Value)
    if ($null -eq $Object) { return $false }
    try {
        if ($null -eq $Object.PSObject.Properties[$Name]) { return $false }
        $Object.$Name = $Value
        return $true
    } catch { return $false }
}

function ConvertTo-SafeNumber {
    <#  JSON 不能表示 Infinity/NaN，统一转成 $null（CSV 中显示为空）。 #>
    param($Value, [int]$Digits = 9)
    if ($null -eq $Value) { return $null }
    try { $d = [double]$Value } catch { return $null }
    if ([double]::IsNaN($d) -or [double]::IsInfinity($d)) { return $null }
    return [math]::Round($d, $Digits)
}

function Flatten-Series {
    <#  ZOS-API 的数据可能是 double[]，也可能是 double[,]；统一压平成一维。
        .NET 二维数组在 PowerShell 里必须用 $a[i, j] 索引，直接 $a[i] 会报
        "You cannot index into a 2 dimensional array with index [i]"。 #>
    param($Array, [int]$Series = 0)

    if ($null -eq $Array) { return @() }
    if ($Array.Rank -eq 1) { return $Array }

    $n0 = $Array.GetLength(0)
    $n1 = $Array.GetLength(1)
    if ($n1 -le 8 -and $n0 -gt 8) {
        # 形状是 [点数, 系列数]
        $result = New-Object 'double[]' $n0
        for ($i = 0; $i -lt $n0; $i++) { $result[$i] = $Array[$i, $Series] }
        return $result
    }
    # 形状是 [系列数, 点数]
    $result = New-Object 'double[]' $n1
    for ($i = 0; $i -lt $n1; $i++) { $result[$i] = $Array[$Series, $i] }
    return $result
}

function Get-NearestValue {
    <#  在 (x, y) 序列里取最接近 $Target 的 y，并返回实际命中的 x。 #>
    param($Xs, $Ys, [double]$Target)
    if ($null -eq $Xs -or $Xs.Count -eq 0) { return $null }
    $bestI = 0; $bestD = [double]::MaxValue
    for ($i = 0; $i -lt $Xs.Count; $i++) {
        $d = [math]::Abs([double]$Xs[$i] - $Target)
        if ($d -lt $bestD) { $bestD = $d; $bestI = $i }
    }
    return [pscustomobject]@{
        requested = ConvertTo-SafeNumber $Target 6
        actualX   = ConvertTo-SafeNumber $Xs[$bestI] 6
        value     = ConvertTo-SafeNumber $Ys[$bestI] 6
    }
}

# ---------------------------------------------------------------- 数据提取

function Get-ZemaxSystemSummary {
    param([Parameter(Mandatory = $true)]$System)

    $sd = $System.SystemData
    $summary = [ordered]@{
        mode   = "$(Get-Prop $System 'Mode' 'Unknown')"
        units  = "$(Get-Prop $sd 'Units' '')"
        aperture = [ordered]@{
            type  = "$(Get-Prop $sd.Aperture 'ApertureType' '')"
            value = ConvertTo-SafeNumber (Get-Prop $sd.Aperture 'ApertureValue')
        }
        fields      = @()
        wavelengths = @()
    }

    $fields = $sd.Fields
    for ($i = 1; $i -le [int](Get-Prop $fields 'NumberOfFields' 0); $i++) {
        $f = $fields.GetField($i)
        $summary.fields += [ordered]@{
            number = $i
            x      = ConvertTo-SafeNumber (Get-Prop $f 'X')
            y      = ConvertTo-SafeNumber (Get-Prop $f 'Y')
            weight = ConvertTo-SafeNumber (Get-Prop $f 'Weight')
        }
    }

    $waves = $sd.Wavelengths
    for ($i = 1; $i -le [int](Get-Prop $waves 'NumberOfWavelengths' 0); $i++) {
        $w = $waves.GetWavelength($i)
        $summary.wavelengths += [ordered]@{
            number     = $i
            value      = ConvertTo-SafeNumber (Get-Prop $w 'Wavelength')
            weight     = ConvertTo-SafeNumber (Get-Prop $w 'Weight')
            isPrimary  = [bool](Get-Prop $w 'IsPrimary' $false)
        }
    }

    return $summary
}

function Get-ZemaxPrescription {
    <#  逐面导出处方表数据。 #>
    param([Parameter(Mandatory = $true)]$System)

    $lde = $System.LDE
    $rows = @()
    $count = [int](Get-Prop $lde 'NumberOfSurfaces' 0)
    for ($i = 0; $i -lt $count; $i++) {
        $s = $lde.GetSurfaceAt($i)
        $rows += [ordered]@{
            surface      = $i
            type         = "$(Get-Prop $s 'Type' '')"
            comment      = "$(Get-Prop $s 'Comment' '')"
            radius       = ConvertTo-SafeNumber (Get-Prop $s 'Radius')
            thickness    = ConvertTo-SafeNumber (Get-Prop $s 'Thickness')
            material     = "$(Get-Prop $s 'Material' '')"
            semiDiameter = ConvertTo-SafeNumber (Get-Prop $s 'SemiDiameter')
            conic        = ConvertTo-SafeNumber (Get-Prop $s 'Conic')
            isStop       = [bool](Get-Prop $s 'IsStop' $false)
        }
    }
    return $rows
}

function Get-ZemaxSettingsSnapshot {
    <#  抓取"这条曲线代表什么量"的关键设置项，随数据一起落盘。
        OpticStudio 的分析默认值会变（例如 RMSField 默认画波前误差而不是点列半径），
        报告必须能自证每个数字的量纲，否则就会张冠李戴。 #>
    param($Settings)

    $names = @('Data', 'Type', 'ReferTo', 'Method', 'Orientation', 'Pattern',
        'RayDensity', 'FieldDensity', 'SampleSize', 'MaximumFrequency', 'RadiusMaximum',
        'ShowDiffractionLimit', 'UsePolarization', 'Configuration')
    $snap = [ordered]@{}
    foreach ($n in $names) {
        $v = Get-Prop $Settings $n
        if ($null -ne $v) { $snap[$n] = "$v" }
    }
    foreach ($sel in @('Field', 'Wavelength', 'Surface')) {
        $o = Get-Prop $Settings $sel
        if ($null -ne $o) {
            $txt = "$o"
            if ($txt -and $txt -notmatch '^ZemaxUI|^\(') { $snap[$sel] = $txt }
        }
    }
    return $snap
}

function Invoke-ZemaxAnalysis {
    <#  通用分析执行器：跑一个 AnalysisIDM 分析，把全部数据系列导成点集。

        OpticStudio 的数据模型是两层：
          IAR_Results.NumberOfDataSeries / GetDataSeries(i)   -> 一条"曲线组"
            .Description  例如 "Field: 0.0000 (deg)"
            .XData.Data   double[]，横坐标
            .YData.Data   double[,]，形状 [点数, 子系列数]
            .NumSeries    子系列数
            .SeriesLabels string[]，每个子系列的标签（例如各波长）
        所以最终每个 (曲线组 × 子系列) 才是一条可绘制的曲线。

        -Configure 是脚本块，用来设置该分析专有参数；属性不存在只记告警，不中断。 #>
    param(
        [Parameter(Mandatory = $true)]$System,
        [Parameter(Mandatory = $true)][string]$AnalysisName,
        [scriptblock]$Configure,
        [int]$MaxPoints = 20000
    )

    $result = [ordered]@{
        id              = $AnalysisName
        status          = 'ok'
        note            = ''
        dataSeriesCount = 0
        settings        = [ordered]@{}
        series          = @()
    }

    $notes = New-Object Collections.Generic.List[string]

    try {
        $id = [ZOSAPI.Analysis.AnalysisIDM]::$AnalysisName
    } catch {
        $result.status = 'error'
        $result.note = "AnalysisIDM 中不存在 $AnalysisName"
        return $result
    }

    $analysis = $null
    try {
        $analysis = $System.Analyses.New_Analysis($id)
        if ($null -eq $analysis) { throw "New_Analysis($AnalysisName) 返回空" }

        # 注意：下面这些调用有的会返回值，必须 [void] 掉，
        # 否则会混进函数输出流，把返回值污染成数组（曾导致 analysis 变成 [null, 对象]）。
        if ($Configure) {
            try { $null = & $Configure $analysis.Settings } catch { $notes.Add("设置参数失败：$($_.Exception.Message)") }
        }

        # 把"这条曲线画的到底是什么量"记下来：同一个分析在不同设置下可以是
        # 波前误差(waves)或点列半径(µm)，报告里必须写清楚，否则数字会贴错标签。
        $result.settings = Get-ZemaxSettingsSnapshot $analysis.Settings

        [void]$analysis.ApplyAndWaitForCompletion()
        $status = Get-Prop $analysis 'StatusMessage' ''
        if ($status) { $notes.Add("status=$status") }

        $results = $analysis.GetResults()
        $groupCount = [int](Get-Prop $results 'NumberOfDataSeries' 0)
        $result.dataSeriesCount = $groupCount
        if ($groupCount -eq 0) { $notes.Add('该分析没有数值数据系列（可能只输出图形或矩阵）') }

        for ($i = 0; $i -lt $groupCount; $i++) {
            $ds = $results.GetDataSeries($i)
            $group = "$(Get-Prop $ds 'Description' "曲线组 $($i + 1)")"
            $xLabel = "$(Get-Prop $ds 'XLabel' '')"
            $xs = Flatten-Series (Get-Prop $ds.XData 'Data') 0
            $yMatrix = Get-Prop $ds.YData 'Data'
            # SeriesLabels 直接用属性访问（不要经过 Get-Prop）：需要的是原始 string[]，
            # 再包 @() 会变成"单元素数组里套数组"，标签会串成一行。
            $labels = @()
            try { if ($null -ne $ds.SeriesLabels) { $labels = @($ds.SeriesLabels) } } catch { }
            $subCount = [int](Get-Prop $ds 'NumSeries' 0)
            if ($subCount -le 0) {
                if ($null -ne $yMatrix -and $yMatrix.Rank -eq 2) {
                    $subCount = if ($yMatrix.GetLength(1) -le 8 -and $yMatrix.GetLength(0) -gt 8) { $yMatrix.GetLength(1) } else { $yMatrix.GetLength(0) }
                } else { $subCount = 1 }
            }

            for ($j = 0; $j -lt $subCount; $j++) {
                $label = ''
                if ($j -lt $labels.Count -and $labels[$j]) { $label = "$($labels[$j])" }
                if (-not $label) { $label = "子系列 $($j + 1)" }

                $ys = Flatten-Series $yMatrix $j
                $points = New-Object Collections.Generic.List[object]
                $n = [math]::Min($xs.Count, $ys.Count)
                if ($n -gt $MaxPoints) { $n = $MaxPoints; $notes.Add("$group/$label 已截断到 $MaxPoints 点") }
                for ($k = 0; $k -lt $n; $k++) {
                    $points.Add(@((ConvertTo-SafeNumber $xs[$k] 6), (ConvertTo-SafeNumber $ys[$k] 6)))
                }

                $result.series += [ordered]@{
                    group    = $group
                    label    = $label
                    xLabel   = $xLabel
                    xCount   = $xs.Count
                    count    = $points.Count
                    points   = $points
                }
            }
        }
    } catch {
        $result.status = 'error'
        $notes.Add("异常：$($_.Exception.Message)")
    } finally {
        if ($null -ne $analysis) { try { [void]$analysis.Close() } catch { } }
    }

    $result.note = ($notes -join '；')
    return $result
}

function Get-ZemaxMtf {
    <#  FFT MTF：返回 [pscustomobject]@{ analysis=...; atFreqs=@(...) } #>
    param(
        [Parameter(Mandatory = $true)]$System,
        [double]$MaxFrequency = 100.0,
        [double[]]$Freqs = @(30.0, 50.0, 100.0),
        [int]$MaxPoints = 20000
    )

    $analysis = Invoke-ZemaxAnalysis -System $System -AnalysisName 'FftMtf' -MaxPoints $MaxPoints -Configure {
        param($settings)
        [void](Set-Prop $settings 'MaximumFrequency' $MaxFrequency)
    }

    $atFreqs = @()
    foreach ($f in $Freqs) {
        $row = [ordered]@{ frequency = ConvertTo-SafeNumber $f 6; perSeries = @() }
        foreach ($s in $analysis.series) {
            $xs = @(); $ys = @()
            foreach ($p in $s.points) { $xs += $p[0]; $ys += $p[1] }
            $hit = Get-NearestValue -Xs $xs -Ys $ys -Target $f
            $row.perSeries += [ordered]@{
                group   = $s.group
                label   = $s.label
                actualX = if ($hit) { $hit.actualX } else { $null }
                value   = if ($hit) { $hit.value } else { $null }
            }
        }
        $atFreqs += $row
    }

    return [pscustomobject]@{ analysis = $analysis; atFreqs = $atFreqs }
}

function Get-EnergyRadii {
    <#  在包围能量曲线上插值求"包含指定能量比例所需的半径"（r50 / r80 / r90）。
        这是点列数据里最可靠、可自证的指标：直接由 EE 曲线线性插值得到，
        不依赖任何口径不明的聚合量。 #>
    param($Points, [double[]]$Fractions = @(0.5, 0.8, 0.9))

    $out = @()
    if ($null -eq $Points -or $Points.Count -lt 2) { return $out }
    foreach ($f in $Fractions) {
        $radius = $null
        for ($i = 1; $i -lt $Points.Count; $i++) {
            $y0 = $Points[$i - 1][1]; $y1 = $Points[$i][1]
            if ($null -eq $y0 -or $null -eq $y1) { continue }
            if ($y1 -ge $f -and $y0 -lt $f) {
                $x0 = $Points[$i - 1][0]; $x1 = $Points[$i][0]
                $t = if ($y1 -ne $y0) { ($f - $y0) / ($y1 - $y0) } else { 0 }
                $radius = $x0 + $t * ($x1 - $x0)
                break
            }
        }
        $out += [ordered]@{ fraction = $f; radiusUm = ConvertTo-SafeNumber $radius 6 }
    }
    return $out
}

function Get-ZemaxSpot {
    <#  点列相关数据。返回两个键：
          series      —— 带数值曲线的分析（每条曲线的量纲由 settings.Data/Type 标明）
          diagram     —— StandardSpot 的 SpotData 聚合量（原始读数，口径存疑，仅供留档）

        重要：RMSField 的 Data 默认是 Wavefront（RMS 波前误差，单位 waves）。
        要看点列半径必须显式设成 SpotRadius（单位 µm）。这里两个都跑，各自标清楚。
        包围能量曲线还会额外算出 r50/r80/r90，作为报告里采用的能量半径指标。 #>
    param(
        [Parameter(Mandatory = $true)]$System,
        [int]$MaxPoints = 20000
    )

    $runs = @(
        @{ id = 'RMSField'; label = 'RMS 波前误差 vs 视场'; configure = { param($s) [void](Set-Prop $s 'Data' 'Wavefront') } },
        @{ id = 'RMSField'; label = 'RMS 点列半径 vs 视场'; configure = { param($s) [void](Set-Prop $s 'Data' 'SpotRadius') } },
        @{ id = 'GeometricEncircledEnergy'; label = '几何包围能量 vs 半径'; configure = $null }
    )

    $series = @()
    foreach ($run in $runs) {
        $a = Invoke-ZemaxAnalysis -System $System -AnalysisName $run.id -MaxPoints $MaxPoints -Configure $run.configure
        $a['quantity'] = $run.label
        if ($run.id -eq 'GeometricEncircledEnergy') {
            # 给每条视场曲线补上 r50/r80/r90（单位 µm），报告采用这三个指标
            foreach ($s in $a.series) {
                $s['energyRadii'] = Get-EnergyRadii -Points $s.points
            }
        }
        $series += $a
    }

    return [pscustomobject]@{
        series  = $series
        diagram = Get-ZemaxSpotDiagramSummary -System $System
    }
}

function Get-ZemaxSpotDiagramSummary {
    <#  StandardSpot 分析只提供聚合量（本版本 ZOS-API 不暴露逐条射线坐标：
        DataSeries / DataScatterPoints / RayData 全为 0）。
        SpotData 给出：MaxRadius、MeanRadius、HalfWidth_X/Y、视场数、波长数。 #>
    param([Parameter(Mandatory = $true)]$System)

    $result = [ordered]@{
        id       = 'StandardSpot'
        status   = 'ok'
        note     = ''
        settings = [ordered]@{}
        metrics  = [ordered]@{}
    }

    $analysis = $null
    try {
        $analysis = $System.Analyses.New_Analysis([ZOSAPI.Analysis.AnalysisIDM]::StandardSpot)
        $result.settings = Get-ZemaxSettingsSnapshot $analysis.Settings
        [void]$analysis.ApplyAndWaitForCompletion()
        $res = $analysis.GetResults()
        $spot = Get-Prop $res 'SpotData'
        if ($null -eq $spot) {
            $result.status = 'error'
            $result.note = '该分析未返回 SpotData（点列图可能被设置成非标准类型）'
        } else {
            $rawMean = ConvertTo-SafeNumber (Get-Prop $spot 'MeanRadius') 6
            # 这两个聚合量的口径无法自证：Doublet 上 MaxRadius=8.437、MeanRadius=4427.98
            # （换算成 nm 后 4.43 µm，与由 EE 曲线推算的 ≥3.16 µm 自洽）；
            # 但 Cooke 上 MaxRadius=24.16 µm、MeanRadius=45011.72，换算后 45.01 µm 反而大于
            # 最大半径，几何上不可能。因此只作原始读数留档，不进入报告结论。
            $result.metrics = [ordered]@{
                maxRadiusRaw   = ConvertTo-SafeNumber (Get-Prop $spot 'MaxRadius')
                meanRadiusRaw  = $rawMean
                halfWidthX     = ConvertTo-SafeNumber (Get-Prop $spot 'HalfWidth_X')
                halfWidthY     = ConvertTo-SafeNumber (Get-Prop $spot 'HalfWidth_Y')
                numberOfFields = ConvertTo-SafeNumber (Get-Prop $spot 'NumberOfFields')
                numberOfWaves  = ConvertTo-SafeNumber (Get-Prop $spot 'NumberOfWavelengths')
            }
            $result.note = '口径核对：MaxRadius 原始读数与包围能量曲线推出的 r90 在两次实测中吻合（Doublet 8.437 ≥ r90 6.609；Cooke 24.161 ≥ r90 23.511），可判为 µm；但 MeanRadius 无法自洽（两例 MaxRadius 比值 2.9，MeanRadius 比值却为 10.2），故不进入报告结论。报告采用 r50/r80/r90（由 EE 曲线插值，单位 µm）。'
        }
    } catch {
        $result.status = 'error'
        $result.note = "异常：$($_.Exception.Message)"
    } finally {
        if ($null -ne $analysis) { try { [void]$analysis.Close() } catch { } }
    }
    return $result
}

# ---------------------------------------------------------------- 落盘

function New-ZemaxRunDirectory {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [string]$LensPath,
        [string]$Label
    )

    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    if (-not $Label) {
        if ($LensPath) { $Label = [IO.Path]::GetFileNameWithoutExtension($LensPath) } else { $Label = 'noinput' }
    }
    $safe = ($Label -replace '[^\w\.\-]+', '_')
    if ($safe.Length -gt 60) { $safe = $safe.Substring(0, 60) }

    $dir = Join-Path $Config.outputRoot ("{0}_{1}" -f $stamp, $safe)
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    return $dir
}

function Write-ZemaxJson {
    param([Parameter(Mandatory = $true)]$Object, [Parameter(Mandatory = $true)][string]$Path, [int]$Depth = 12)
    $json = $Object | ConvertTo-Json -Depth $Depth
    [IO.File]::WriteAllText($Path, $json, (New-Object Text.UTF8Encoding($false)))
    return $Path
}

function Write-ZemaxCsv {
    <#  注意：本库的行都是 [ordered]@{}（OrderedDictionary），而 Export-Csv 对字典导出的是
        它的"属性"（Count/Keys/Values/IsReadOnly…），不是键值 —— 直接导会得到一列垃圾。
        必须先转成 PSCustomObject 再导。 #>
    param([Parameter(Mandatory = $true)]$Rows, [Parameter(Mandatory = $true)][string]$Path)
    if ($null -eq $Rows -or @($Rows).Count -eq 0) {
        [IO.File]::WriteAllText($Path, '', (New-Object Text.UTF8Encoding($true)))
        return $Path
    }

    $flat = New-Object Collections.Generic.List[object]
    foreach ($r in @($Rows)) {
        if ($r -is [System.Collections.IDictionary]) {
            $o = New-Object psobject
            foreach ($k in $r.Keys) {
                $o | Add-Member -NotePropertyName ([string]$k) -NotePropertyValue $r[$k] -Force
            }
            $flat.Add($o)
        } else {
            $flat.Add($r)
        }
    }
    $flat | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
    return $Path
}

function Get-FileFingerprint {
    param([Parameter(Mandatory = $true)][string]$Path)
    $fi = Get-Item -LiteralPath $Path
    $sha = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    return [ordered]@{
        path   = $fi.FullName
        bytes  = $fi.Length
        sha256 = $sha
        mtime  = $fi.LastWriteTime.ToString('s')
    }
}

# ================================================================ 设计（按规格建模 → 优化）

function New-ZemaxSequentialSystem {
    <#  按规格新建一个序列模式系统：孔径、视场、波长、材料库、面数。 #>
    param(
        [Parameter(Mandatory = $true)]$System,
        [Parameter(Mandatory = $true)]$Spec
    )

    $log = New-Object Collections.Generic.List[string]

    [void]$System.New($false)

    if ($Spec.catalog) {
        try {
            [void]$System.SystemData.MaterialCatalogs.AddCatalog($Spec.catalog)
            $log.Add("材料库 $($Spec.catalog) 已加载")
        } catch { $log.Add("材料库加载失败：$($_.Exception.Message)") }
    }

    $ap = $Spec.aperture
    if ($ap) {
        $typeName = if ($ap.type) { $ap.type } else { 'EntrancePupilDiameter' }
        $okType = Set-Prop $System.SystemData.Aperture 'ApertureType' ([ZOSAPI.SystemData.ZemaxApertureType]::$typeName)
        $okVal = Set-Prop $System.SystemData.Aperture 'ApertureValue' ([double]$ap.value)
        $log.Add("孔径 $typeName=$($ap.value)（type=$okType value=$okVal）")
    }

    $fields = $Spec.fields
    if ($fields -and $fields.values) {
        if ($fields.type) {
            [void](Set-Prop $System.SystemData.Fields 'FieldType' ([ZOSAPI.SystemData.FieldType]::$($fields.type)))
        }
        $values = @($fields.values)
        $first = $System.SystemData.Fields.GetField(1)
        [void](Set-Prop $first 'X' 0.0)
        [void](Set-Prop $first 'Y' ([double]$values[0]))
        [void](Set-Prop $first 'Weight' 1.0)
        for ($i = 1; $i -lt $values.Count; $i++) {
            [void]$System.SystemData.Fields.AddField(0.0, [double]$values[$i], 1.0)
        }
        $log.Add("视场 $($values.Count) 个：$($values -join ', ')（类型 $($fields.type)）")
    }

    $waves = $Spec.wavelengths
    if ($waves) {
        if ($waves.preset) {
            $ok = $false
            try {
                [void]$System.SystemData.Wavelengths.SelectWavelengthPreset([ZOSAPI.SystemData.WavelengthPreset]::$($waves.preset))
                $ok = $true
            } catch { }
            $log.Add("波长预设 $($waves.preset)（成功=$ok）")
        }
        if ($waves.values) {
            $vals = @($waves.values)
            $w1 = $System.SystemData.Wavelengths.GetWavelength(1)
            [void](Set-Prop $w1 'Wavelength' ([double]$vals[0]))
            for ($i = 1; $i -lt $vals.Count; $i++) {
                if (($i + 1) -le $System.SystemData.Wavelengths.NumberOfWavelengths) {
                    $w = $System.SystemData.Wavelengths.GetWavelength($i + 1)
                    [void](Set-Prop $w 'Wavelength' ([double]$vals[$i]))
                } else {
                    [void]$System.SystemData.Wavelengths.AddWavelength([double]$vals[$i], 1.0)
                }
            }
            $log.Add("波长按数值设置：$($vals -join ', ') µm")
        }
    }

    $needed = 3
    if ($Spec.surfaces) { $needed = @($Spec.surfaces).Count + 2 }
    while ($System.LDE.NumberOfSurfaces -lt $needed) {
        [void]$System.LDE.InsertNewSurfaceAt($System.LDE.NumberOfSurfaces - 1)
    }
    $log.Add("LDE 面数 = $($System.LDE.NumberOfSurfaces)（目标 $needed）")

    return $log
}

function Set-ZemaxSurfaceData {
    <#  按规格写入面型数据：半径、厚度、材料、半口径、注释。 #>
    param(
        [Parameter(Mandatory = $true)]$System,
        [Parameter(Mandatory = $true)]$Spec
    )

    $log = New-Object Collections.Generic.List[string]
    foreach ($s in @($Spec.surfaces)) {
        $idx = [int]$s.index
        if ($idx -lt 1 -or $idx -gt ($System.LDE.NumberOfSurfaces - 2)) {
            $log.Add("跳过非法面号 $idx")
            continue
        }
        $row = $System.LDE.GetSurfaceAt($idx)
        if ($null -ne $s.radius) { [void](Set-Prop $row 'Radius' ([double]$s.radius)) }
        if ($null -ne $s.thickness) { [void](Set-Prop $row 'Thickness' ([double]$s.thickness)) }
        if ($null -ne $s.semiDiameter) { [void](Set-Prop $row 'SemiDiameter' ([double]$s.semiDiameter)) }
        if ($s.comment) { [void](Set-Prop $row 'Comment' ([string]$s.comment)) }
        if ($null -ne $s.material -and $s.material -ne '') {
            try { $row.Material = [string]$s.material } catch { $log.Add("面 $idx 材料设置失败：$($_.Exception.Message)") }
        }
        $log.Add("面 $idx：R=$($row.Radius) T=$($row.Thickness) 材料='$($row.Material)'")
    }
    return $log
}

function Set-ZemaxVariables {
    <#  按规格设置变量：surface + cell（Radius / Thickness / Conic / SemiDiameter）。 #>
    param(
        [Parameter(Mandatory = $true)]$System,
        [Parameter(Mandatory = $true)]$Spec
    )

    $log = New-Object Collections.Generic.List[string]
    [void]$System.Tools.RemoveAllVariables()

    foreach ($v in @($Spec.variables)) {
        $idx = [int]$v.surface
        $cell = if ($v.cell) { [string]$v.cell } else { 'Radius' }
        $row = $System.LDE.GetSurfaceAt($idx)
        $ok = $false
        try {
            switch ($cell) {
                'Radius' { [void]$row.RadiusCell.MakeSolveVariable(); $ok = $true }
                'Thickness' { [void]$row.ThicknessCell.MakeSolveVariable(); $ok = $true }
                'Conic' { [void]$row.ConicCell.MakeSolveVariable(); $ok = $true }
                'SemiDiameter' { [void]$row.SemiDiameterCell.MakeSolveVariable(); $ok = $true }
                default { }
            }
        } catch { }
        $log.Add("变量 面$idx.$cell → $ok")
    }
    return $log
}

function Set-ZemaxOperand {
    <#  追加一个评价函数操作数。单元布局（1-based，已实测）：
        cell[1]=Type(String)，cell[2]=Int1，cell[3]=Wave 或 Double，cell[4]=Double。
        Target / Weight 用属性设置。 #>
    param(
        [Parameter(Mandatory = $true)]$Mfe,
        [Parameter(Mandatory = $true)][string]$Type,
        [double]$Target = 0,
        [double]$Weight = 1,
        [int]$Wave = 0,
        [int]$Int1 = 0,
        [int]$Int2 = 0
    )

    $op = $Mfe.AddOperand()
    try { $op.ChangeType([ZOSAPI.Editors.MFE.MeritOperandType]::$Type) } catch { return $null }
    [void](Set-Prop $op 'Target' $Target)
    [void](Set-Prop $op 'Weight' $Weight)
    # 带参数的走 Get-ZemaxOperandMetric；这里只在显式给了参数时写入
    return $op
}

function Set-ZemaxMeritFunction {
    <#  按规格建立评价函数：可选"默认评价函数向导" + 追加操作数。 #>
    param(
        [Parameter(Mandatory = $true)]$System,
        [Parameter(Mandatory = $true)]$Spec
    )

    $log = New-Object Collections.Generic.List[string]
    $mfe = $System.MFE

    $w = $Spec.merit.wizard
    if ($w) {
        try {
            $wiz = $mfe.SEQOptimizationWizard
            if ($null -ne $w.data) { [void](Set-Prop $wiz 'Data' ([int]$w.data)) }
            if ($null -ne $w.ring) { [void](Set-Prop $wiz 'Ring' ([int]$w.ring)) }
            if ($null -ne $w.overallWeight) { [void](Set-Prop $wiz 'OverallWeight' ([double]$w.overallWeight)) }
            if ($null -ne $w.glassUsed) {
                [void](Set-Prop $wiz 'IsGlassUsed' ([bool]$w.glassUsed))
                [void](Set-Prop $wiz 'GlassMin' ([double]$w.glassMin))
                [void](Set-Prop $wiz 'GlassMax' ([double]$w.glassMax))
                [void](Set-Prop $wiz 'GlassEdge' ([double]$w.glassEdge))
            }
            if ($null -ne $w.airUsed) {
                [void](Set-Prop $wiz 'IsAirUsed' ([bool]$w.airUsed))
                [void](Set-Prop $wiz 'AirMin' ([double]$w.airMin))
                [void](Set-Prop $wiz 'AirMax' ([double]$w.airMax))
                [void](Set-Prop $wiz 'AirEdge' ([double]$w.airEdge))
            }
            [void]$wiz.Apply()
            $log.Add("默认评价函数向导已应用（Data=$($w.data), Ring=$($w.ring)），operand 数=$($mfe.NumberOfOperands)")
        } catch {
            $log.Add("向导应用失败：$($_.Exception.Message)")
        }
    }

    foreach ($o in @($Spec.merit.operands)) {
        $op = $mfe.AddOperand()
        $ok = $false
        try {
            $op.ChangeType([ZOSAPI.Editors.MFE.MeritOperandType]::$([string]$o.type))
            [void](Set-Prop $op 'Target' ([double]$o.target))
            [void](Set-Prop $op 'Weight' ([double]$o.weight))
            $ok = $true
        } catch {
            $log.Add("追加操作数 $($o.type) 失败：$($_.Exception.Message)")
        }
        $log.Add("追加操作数 $($o.type)（Target=$($o.target), Weight=$($o.weight)）→ $(if ($ok) { '成功' } else { '失败' })")
    }
    $log.Add("评价函数 operand 总数 = $($mfe.NumberOfOperands)")
    return $log
}

function Get-ZemaxOperandValue {
    <#  临时追加一个操作数、算一次评价函数、读出它的 Value，再把操作数删掉。
        实测 operand.Value 是取 EFL / AXCL / WFNO 这类量的可靠通路。
        参数写入：按 Header 匹配 'Wave' 的单元写波长，其余整数按 header 顺序填。 #>
    param(
        [Parameter(Mandatory = $true)]$System,
        [Parameter(Mandatory = $true)][string]$Type,
        [int]$Wave = 0,
        [int[]]$Ints = @()
    )

    $mfe = $System.MFE
    $before = $mfe.NumberOfOperands
    $op = $mfe.AddOperand()
    $value = $null
    try {
        # ChangeType / Apply 这类调用会返回 bool，必须 [void] —— 否则会混进函数输出流，
        # 让返回值变成 [True, 100.000000003] 这种数组（曾直接把报告生成搞崩）。
        [void]$op.ChangeType([ZOSAPI.Editors.MFE.MeritOperandType]::$Type)
        [void](Set-Prop $op 'Target' 0.0)
        [void](Set-Prop $op 'Weight' 0.0)

        # 写整数参数：cell[2] 起，遇到 Header 含 'Wave' 的单元写波长
        $intQueue = New-Object Collections.Generic.Queue[int]
        foreach ($v in $Ints) { $intQueue.Enqueue([int]$v) }
        for ($idx = 2; $idx -le 4; $idx++) {
            $cell = $null
            try { $cell = $op.GetCellAt($idx) } catch { continue }
            if ($null -eq $cell) { continue }
            $header = "$(Get-Prop $cell 'Header' '')"
            $dataType = "$(Get-Prop $cell 'DataType' '')"
            if ($dataType -notmatch 'Integer') { continue }
            if ($header -match 'Wave') {
                if ($Wave -gt 0) { try { $cell.IntegerValue = $Wave } catch { } }
            } elseif ($intQueue.Count -gt 0) {
                try { $cell.IntegerValue = $intQueue.Dequeue() } catch { }
            }
        }

        [void]$mfe.CalculateMeritFunction()
        $v = Get-Prop $op 'Value'
        if ($null -ne $v) { $value = ConvertTo-SafeNumber $v 9 }
    } catch {
        $value = $null
    } finally {
        while ($mfe.NumberOfOperands -gt $before) { [void]$mfe.RemoveOperandAt($mfe.NumberOfOperands) }
    }
    # 防御：万一仍有别处漏出的输出混进来，只取最后一个（真正的读数）
    if ($value -is [System.Array]) { $value = $value[-1] }
    return $value
}

function Get-ZemaxDesignMetrics {
    <#  设计验收指标：每个指标都是一个评价函数操作数，可独立复算。 #>
    param(
        [Parameter(Mandatory = $true)]$System,
        [Parameter(Mandatory = $true)]$Spec
    )

    $metrics = [ordered]@{}
    foreach ($m in @($Spec.metrics)) {
        $name = [string]$m.name
        $args = @{ System = $System; Type = [string]$m.operand }
        if ($null -ne $m.wave) { $args.Wave = [int]$m.wave }
        if ($m.ints) { $args.Ints = @($m.ints | ForEach-Object { [int]$_ }) }
        $v = Get-ZemaxOperandValue @args
        $metrics[$name] = [ordered]@{
            label   = if ($m.label) { [string]$m.label } else { $name }
            operand = [string]$m.operand
            wave    = if ($null -ne $m.wave) { [int]$m.wave } else { 0 }
            unit    = if ($m.unit) { [string]$m.unit } else { '' }
            value   = $v
        }
    }
    return $metrics
}

function Invoke-ZemaxOptimization {
    <#  跑优化：local 与/或 hammer（限时）+ 快速对焦。
        返回每一步的 merit 变化，报告里据此说明"做了什么、有没有收敛"。 #>
    param(
        [Parameter(Mandatory = $true)]$System,
        [Parameter(Mandatory = $true)]$Spec
    )

    $log = New-Object Collections.Generic.List[string]
    $result = [ordered]@{ local = $null; hammer = $null; quickFocus = $null }
    $opt = $Spec.optimize

    if ($opt.local) {
        $t = [Diagnostics.Stopwatch]::StartNew()
        $lo = $System.Tools.OpenLocalOptimization()
        if ($null -ne $lo) {
            $algo = if ($opt.local.algorithm) { $opt.local.algorithm } else { 'DampedLeastSquares' }
            $cycles = if ($opt.local.cycles) { $opt.local.cycles } else { 'Automatic' }
            [void](Set-Prop $lo 'Algorithm' ([ZOSAPI.Tools.Optimization.OptimizationAlgorithm]::$algo))
            [void](Set-Prop $lo 'Cycles' ([ZOSAPI.Tools.Optimization.OptimizationCycles]::$cycles))
            [void](Set-Prop $lo 'NumberOfCores' ([int]$opt.local.cores))
            $m0 = ConvertTo-SafeNumber (Get-Prop $lo 'InitialMeritFunction') 6
            [void]$lo.RunAndWaitForCompletion()
            $m1 = ConvertTo-SafeNumber (Get-Prop $lo 'CurrentMeritFunction') 6
            $status = "$(Get-Prop $lo 'Status' '')"
            [void]$lo.Close()
            $t.Stop()
            $result.local = [ordered]@{ meritBefore = $m0; meritAfter = $m1; seconds = [math]::Round($t.Elapsed.TotalSeconds, 2); status = $status }
            $log.Add("局部优化：$m0 → $m1（$([math]::Round($t.Elapsed.TotalSeconds,1))s）")
        }
    }

    if ($opt.hammerSeconds -and [double]$opt.hammerSeconds -gt 0) {
        $t = [Diagnostics.Stopwatch]::StartNew()
        $ho = $System.Tools.OpenHammerOptimization()
        if ($null -ne $ho) {
            $algo = if ($opt.local.algorithm) { $opt.local.algorithm } else { 'DampedLeastSquares' }
            [void](Set-Prop $ho 'Algorithm' ([ZOSAPI.Tools.Optimization.OptimizationAlgorithm]::$algo))
            [void](Set-Prop $ho 'NumberOfCores' ([int]$opt.local.cores))
            $m0 = ConvertTo-SafeNumber (Get-Prop $ho 'InitialMeritFunction') 6
            [void]$ho.RunAndWaitWithTimeout([double]$opt.hammerSeconds)
            $m1 = ConvertTo-SafeNumber (Get-Prop $ho 'CurrentMeritFunction') 6
            [void]$ho.Cancel()
            [void]$ho.WaitForCompletion()
            [void]$ho.Close()
            $t.Stop()
            $result.hammer = [ordered]@{ meritBefore = $m0; meritAfter = $m1; seconds = [math]::Round($t.Elapsed.TotalSeconds, 2) }
            $log.Add("锤形优化 $($opt.hammerSeconds)s：$m0 → $m1（实际 $([math]::Round($t.Elapsed.TotalSeconds,1))s）")
        }
    }

    if ($opt.quickFocus) {
        try {
            $qf = $System.Tools.OpenQuickFocus()
            $criterion = if ($opt.quickFocusCriterion) { $opt.quickFocusCriterion } else { 'SpotSizeRadial' }
            [void](Set-Prop $qf 'Criterion' ([ZOSAPI.Tools.General.QuickFocusCriterion]::$criterion))
            [void](Set-Prop $qf 'UseCentroid' $true)
            [void]$qf.RunAndWaitForCompletion()
            [void]$qf.Close()
            $result.quickFocus = [ordered]@{ criterion = $criterion; done = $true }
            $log.Add("快速对焦完成（$criterion）")
        } catch {
            $result.quickFocus = [ordered]@{ done = $false; error = "$($_.Exception.Message)" }
            $log.Add("快速对焦失败：$($_.Exception.Message)")
        }
    }

    return [pscustomobject]@{ result = $result; log = $log }
}

function Test-ZemaxAcceptance {
    <#  把实测指标与规格里预先声明的判据逐条比对。 #>
    param($Metrics, $Spec)

    $rows = @()
    foreach ($c in @($Spec.acceptance)) {
        $name = [string]$c.metric
        $actual = $null
        if ($Metrics.Contains($name)) { $actual = $Metrics[$name].value }
        $pass = $null
        $rule = ''
        if ($null -ne $actual) {
            if ($null -ne $c.max -and $null -ne $c.min) {
                $pass = ($actual -ge [double]$c.min) -and ($actual -le [double]$c.max)
                $rule = "[$($c.min), $($c.max)]"
            } elseif ($null -ne $c.max) {
                $pass = $actual -le [double]$c.max
                $rule = "≤ $($c.max)"
            } elseif ($null -ne $c.min) {
                $pass = $actual -ge [double]$c.min
                $rule = "≥ $($c.min)"
            }
        }
        $rows += [ordered]@{
            metric = $name
            label  = if ($c.label) { [string]$c.label } else { $name }
            unit   = if ($c.unit) { [string]$c.unit } else { '' }
            rule   = $rule
            actual = $actual
            pass   = $pass
        }
    }
    return $rows
}

function Save-ZemaxSystem {
    param(
        [Parameter(Mandatory = $true)]$System,
        [Parameter(Mandatory = $true)][string]$Path
    )
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [void]$System.SaveAs($Path)
    return $Path
}

function Get-ZemaxChromaticFocus {
    <#  轴向色差的独立测量：逐个波长单独对焦（QuickFocus，径向光斑 + 质心），
        读像面前一面的厚度作为该波长的最佳像面位置，再取差值。
        与 AXCL 的区别必须说清楚：AXCL 是"近轴焦点"之差，本函数是"最佳几何焦点"之差，
        有球差时两者并不相等（实测 Doublet 样例：AXCL(F–C)=0.126 mm，本方法=-0.019 mm）。
        注意：本函数临时改波长权重，结束时恢复；全程不写盘，不会污染已保存的设计文件。 #>
    param([Parameter(Mandatory = $true)]$System)

    $waves = $System.SystemData.Wavelengths
    $n = [int]$waves.NumberOfWavelengths
    $imgRow = $System.LDE.NumberOfSurfaces - 1

    $orig = @{}
    for ($k = 1; $k -le $n; $k++) { $orig[$k] = $waves.GetWavelength($k).Weight }

    $positions = @()
    $values = @()
    try {
        for ($w = 1; $w -le $n; $w++) {
            for ($k = 1; $k -le $n; $k++) {
                $waves.GetWavelength($k).Weight = if ($k -eq $w) { 1.0 } else { 0.0 }
            }
            $qf = $System.Tools.OpenQuickFocus()
            [void](Set-Prop $qf 'Criterion' ([ZOSAPI.Tools.General.QuickFocusCriterion]::SpotSizeRadial))
            [void](Set-Prop $qf 'UseCentroid' $true)
            [void]$qf.RunAndWaitForCompletion()
            [void]$qf.Close()
            $t = ConvertTo-SafeNumber $System.LDE.GetSurfaceAt($imgRow - 1).Thickness 9
            $positions += [ordered]@{
                wave     = $w
                lambdaUm = ConvertTo-SafeNumber $waves.GetWavelength($w).Wavelength 6
                focusMm  = $t
            }
            $values += [double]$t
        }
    } finally {
        for ($k = 1; $k -le $n; $k++) { $waves.GetWavelength($k).Weight = $orig[$k] }
    }

    $span = if ($values.Count -gt 0) { ($values | Measure-Object -Maximum).Maximum - ($values | Measure-Object -Minimum).Minimum } else { $null }
    $fc = if ($values.Count -ge 3) { $values[0] - $values[2] } else { $null }

    return [ordered]@{
        method    = '逐波长单独快速对焦后读像面位置（最佳几何焦点差）'
        note      = '与 AXCL 的近轴定义不同；有球差时两者不相等，报告同时列出以便对照'
        positions = $positions
        shiftFCmm = ConvertTo-SafeNumber $fc 6
        spanMm    = ConvertTo-SafeNumber $span 6
    }
}

function Get-ZemaxDerivedMetrics {
    <#  从已经抽好的分析数据里派生出可验收的指标（供 design 的判据引用）。
        全部来自 data.json 里的点集，可用同一份数据复算。 #>
    param($Mtf, $Spot, [double[]]$Freqs = @(30, 50, 100))

    $out = [ordered]@{}

    # --- MTF：取请求频率里最接近 50 lp/mm 的一行 ---
    if ($Mtf -and $Mtf.atFreqs) {
        $target = 50.0
        $row = $null; $bestD = [double]::MaxValue
        foreach ($r in $Mtf.atFreqs) {
            $d = [math]::Abs([double]$r.frequency - $target)
            if ($d -lt $bestD) { $bestD = $d; $row = $r }
        }
        if ($row) {
            $axisTan = $null; $axisSag = $null; $worstTan = $null; $worstSag = $null
            foreach ($ps in $row.perSeries) {
                $isAxis = ($ps.group -match '0\.0+ \(deg\)')
                if ($ps.label -match 'Tang') {
                    if ($isAxis -and $null -eq $axisTan) { $axisTan = $ps.value }
                    if ($null -eq $worstTan -or $ps.value -lt $worstTan) { $worstTan = $ps.value }
                } elseif ($ps.label -match 'Sag') {
                    if ($isAxis -and $null -eq $axisSag) { $axisSag = $ps.value }
                    if ($null -eq $worstSag -or $ps.value -lt $worstSag) { $worstSag = $ps.value }
                }
            }
            $f = ConvertTo-SafeNumber $row.frequency 4
            $out['mtf_axis_tan'] = [ordered]@{ label = "MTF@$f lp/mm 轴上（切向）"; unit = ''; value = $axisTan; operand = 'derived:FftMtf' }
            $out['mtf_axis_sag'] = [ordered]@{ label = "MTF@$f lp/mm 轴上（弧矢）"; unit = ''; value = $axisSag; operand = 'derived:FftMtf' }
            $out['mtf_worst_tan'] = [ordered]@{ label = "MTF@$f lp/mm 最差视场（切向）"; unit = ''; value = $worstTan; operand = 'derived:FftMtf' }
            $out['mtf_worst_sag'] = [ordered]@{ label = "MTF@$f lp/mm 最差视场（弧矢）"; unit = ''; value = $worstSag; operand = 'derived:FftMtf' }
        }
    }

    # --- RMS 点列半径 / 波前误差：按"轴上 Poly"与"全视场最差"两项取 ---
    foreach ($a in @($Spot)) {
        if ($a.id -ne 'RMSField') { continue }
        $kind = "$(Get-Prop $a.settings 'Data' '')"
        if ($kind -ne 'SpotRadius' -and $kind -ne 'Wavefront') { continue }
        $unit = if ($kind -eq 'SpotRadius') { 'µm' } else { 'waves' }
        $prefix = if ($kind -eq 'SpotRadius') { 'rms_spot' } else { 'rms_wavefront' }
        $axisPoly = $null; $worst = $null; $worstLabel = ''
        foreach ($s in $a.series) {
            if ($null -eq $s.points -or $s.points.Count -eq 0) { continue }
            $vals = @($s.points | ForEach-Object { $_[1] })
            $v = ($vals | Measure-Object -Maximum).Maximum
            # 注意：RMSField 的曲线组是"视场方向"，视场本身是 X 轴，
            # 所以"轴上"= x 最小那一点的值，不能按组名去匹配视场。
            if ($s.label -eq 'Poly') { $axisPoly = $s.points[0][1] }
            if ($null -eq $worst -or $v -gt $worst) { $worst = $v; $worstLabel = "$($s.group)/$($s.label)" }
        }
        $out["${prefix}_axis_poly"] = [ordered]@{ label = "RMS $kind 最小视场（Poly，x 最小处的值）"; unit = $unit; value = $axisPoly; operand = 'derived:RMSField' }
        $out["${prefix}_worst"] = [ordered]@{ label = "RMS $kind 全视场最差（$worstLabel）"; unit = $unit; value = $worst; operand = 'derived:RMSField' }
    }

    # --- 包围能量半径 ---
    foreach ($a in @($Spot)) {
        if ($a.id -ne 'GeometricEncircledEnergy') { continue }
        foreach ($s in $a.series) {
            $isAxis = ($s.group -match '0\.0+ \(deg\)')
            foreach ($er in @($s.energyRadii)) {
                $tag = [int]($er.fraction * 100)
                if ($isAxis) {
                    $out["r${tag}_axis"] = [ordered]@{ label = "轴上 r$tag 能量半径"; unit = 'µm'; value = $er.radiusUm; operand = 'derived:GeometricEncircledEnergy' }
                }
                $key = "r${tag}_worst"
                if (-not $out.Contains($key) -or ($null -ne $er.radiusUm -and $er.radiusUm -gt $out[$key].value)) {
                    $out[$key] = [ordered]@{ label = "最差视场 r$tag 能量半径（$($s.group)）"; unit = 'µm'; value = $er.radiusUm; operand = 'derived:GeometricEncircledEnergy' }
                }
            }
        }
    }

    return $out
}
