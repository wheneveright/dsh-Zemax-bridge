# Zemax 桥接（DSH ↔ Ansys Zemax OpticStudio 2024 R1）

这个目录把 DSH 和本机的 Ansys Zemax OpticStudio 2024 R1 连起来：用 ZOS-API 拉起无界面实例，
读取镜头文件、算出光学指标，再落成 JSON / CSV / Excel / Word，全程可复现、可追溯。

## 目录用途

| 路径 | 说明 |
|---|---|
| `config.json` | **路径配置真源**。换机器或换目录只改这里；也可用环境变量 `ZEMAX_CONFIG` / `ZEMAX_OUTPUT_ROOT` / `ZEMAX_SCRATCH_ROOT` 覆盖。 |
| `scripts\zemax.ps1` | 命令行入口，参数见 `-?` 或脚本头部注释。 |
| `scripts\zemax-lib.ps1` | ZOS-API 公共库（连接、系统/处方提取、通用分析执行器、落盘）。 |
| `scripts\build_report.py` | 把 `data.json` 渲染成 Excel / Word（openpyxl + python-docx）。 |
| `scripts\normalize-encoding.ps1` | 改完 `.ps1` 后跑一次：转成 UTF-8 带 BOM。**不做会因 GBK 解码而语法报错。** |
| `runs\<时间戳>_<镜头名>\` | 每次运行的产物目录，只增不改；输入文件会复制进来保证可追溯。 |
| `_early-probe\` | 可行性验证阶段的早期产物，仅作留档。 |

## 快速开始

```powershell
$dir = '<workspace>\Zemax\scripts'
$ps  = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"

# 环境/许可自检
& $ps -NoProfile -ExecutionPolicy Bypass -File "$dir\run.ps1" info

# 出一个完整报告（Excel + Word）
& $ps -NoProfile -ExecutionPolicy Bypass -File "$dir\run.ps1" report `
    -File "$env:USERPROFILE\Documents\Zemax\SAMPLES\Sequential\Objectives\Doublet.zmx"

# 指定空间频率
& $ps -NoProfile -ExecutionPolicy Bypass -File "$dir\run.ps1" mtf `
    -File "<某个 .zmx / .zos>" -MaxFrequency 100 -Freqs 30,50,100

# 按规格做设计：建模 → 设变量 → 建评价函数 → 优化 → 对焦 → 存 .zos → 出验收报告
& $ps -NoProfile -ExecutionPolicy Bypass -File "$dir\run.ps1" design `
    -Spec "<...>\Design01_Achromatic_Objective_100mm_f5\design.json" `
    -Out  "<...>\Design01_Achromatic_Objective_100mm_f5\run" `
    -Freqs 30,50,100
```

⚠️ 一律用 `run.ps1` 调用（纯 ASCII 启动器，会在运行前自动把 `zemax.ps1` / `zemax-lib.ps1` /
`config.json` 修成 UTF-8 带 BOM）。直接调 `zemax.ps1` 时，任何一次"无 BOM 的编辑"都会让
PowerShell 5.1 按 GBK 解码中文并当场语法报错。

一次 `report` 约 6–10 秒（冷启动 OpticStudio 约 5 秒；`design` 还要加优化时间）。产物：`data.json`、
若干 CSV（`prescription` / `mtf` / `spot` / `spot-summary`，设计任务另有 `metrics` / `acceptance`）、
`<镜头名>_report.xlsx`（7 个工作表，含 MTF 折线图）、`<镜头名>_report.docx`、输入副本、`run.log`；
设计任务另写出设计文件 `.zos` 与其 `.ZDA`。

点列数据说明：ZOS-API 不暴露逐条射线坐标，能拿到的只有聚合指标。报告采用**包围能量半径
r50/r80/r90（µm）**——由几何包围能量曲线线性插值得到，可用 `spot.csv` 独立复现；
`StandardSpot` 的 `SpotData` 原始读数（MaxRadius / MeanRadius）口径不自洽，只作留档，
不进入报告结论。

## 为什么是 PowerShell 而不是 Python

ZOS-API 是 .NET Framework 程序集，本机 DSH 的 shell 工具本身就是 Windows PowerShell 5.1（.NET Framework 4.8），
可以直接 `Add-Type` 加载 `ZOSAPI.dll`，**零安装、零依赖**。
Python 路线需要 pythonnet：系统 Python 是 3.14（缺少可用轮子）、自带 Python 3.12 未装 pythonnet、
pypi 直连超时——所以 Python 只用来做报告渲染（numpy/pandas/openpyxl/python-docx 都已就绪）。

## 已知限制

- 每次调用都会冷启动一个无界面 OpticStudio 实例并占用一个许可席位（本机 Premium 许可，`IsValidRequestforAPI=True`）。
- 已覆盖"数据提取与报告"和"按规格做设计（建模/评价函数/局部+锤形优化/对焦/验收）"两类任务；
  **公差分析、多重组态、热/压力环境、非序列模式尚未接入**，但库里的 `Invoke-ZemaxAnalysis`、
  `Get-ZemaxOperandValue` 等接口可以直接扩展。
- 锤形优化是**限时随机过程**，重跑可能得到略有差异的处方；要完全复现请直接使用已保存的 `.zos`。
- 日志里的 `FRU__delta_init(): Attempt to start when running!` 是本机许可环境的噪声，退出码为 0，可忽略。
- 大体积中间产物（`.zrd`/`.zbf`，可达 GB 级）目前没有本地 scratch 隔离，仍会落在工作区；接入这类分析前
  建议把 `config.json` 的 `scratchRoot` 指到非 OneDrive 的本地盘，避免同步风暴。

## 相关文档

- 面向模型的技能说明：`..\.dsh\skills\zemax\SKILL.md`（含硬性约定与踩坑清单，改脚本前先读）
- ZOS-API 帮助：`E:\Ansys Zemax\ZOS-API\Help\ZOS-API.chm`
- 官方样例：`C:\Users\<you>\Documents\Zemax\ZOS-API Sample Code`
