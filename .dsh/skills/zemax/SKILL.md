---
name: zemax
description: 用本机 Ansys Zemax OpticStudio 2024 R1 的 ZOS-API 读取镜头文件、抽取处方与成像指标（FFT MTF、RMS 点列、包围能量），并生成 Excel / Word 报告。当任务涉及 Zemax / OpticStudio / .zmx / .zos 镜头文件、光学处方、MTF、点列图、光斑半径时使用。
whenToUse: 用户要求分析、汇总或报告某个 Zemax 镜头文件；需要从 OpticStudio 取数值而不是凭经验估计；需要把光学指标整理成表格/报告交付。
---

# Zemax 桥接（OpticStudio 2024 R1 via ZOS-API）

## 这个技能做什么

本机需装有 Ansys Zemax OpticStudio（2024 R1 已实测）且许可支持 `IsValidLicenseForAPI`。
技能通过 ZOS-API 以**无界面实例**方式驱动它：读文件 → 计算结果 → 落盘成 JSON/CSV/Excel/Word，全程可复现。

所有工具与产物都在**当前工作区**的 `Zemax\` 目录内（路径不写死，全部从 `config.json` 读）：

```
<工作区>\Zemax\config.json                 路径配置真源（安装器生成；换机器只需改这里）
<工作区>\Zemax\scripts\run.ps1             命令行入口（纯 ASCII 启动器，优先用它）
<工作区>\Zemax\scripts\zemax.ps1           命令分发（不要直接调，见下方警告）
<工作区>\Zemax\scripts\zemax-lib.ps1       ZOS-API 公共库
<工作区>\Zemax\scripts\build_report.py     data.json -> Excel/Word
<工作区>\Zemax\scripts\normalize-encoding.ps1  改完脚本后修 UTF-8 BOM
<工作区>\Zemax\examples\                  示例设计规格 JSON
<工作区>\Zemax\reference\                 156 个 AnalysisIDM / 437 个 operand / 39 个解法类型清单
<工作区>\Zemax\runs\<时间戳>_<镜头名>\      每次运行的产物（永不覆盖源文件）
```

## 怎么调用

**必须用 Windows PowerShell 5.1**（ZOS-API 是 .NET Framework 程序集，PowerShell 7 加载会失败）。
`powershell.exe` 的完整路径在 `config.json` 的 `windowsPowerShell` 里。

```powershell
# <工作区> 就是本会话的工作区根目录；脚本目录固定为 <工作区>\Zemax\scripts
$dir = '<工作区>\Zemax\scripts'
& "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile `
    -ExecutionPolicy Bypass -File "$dir\run.ps1" report -File "<镜头文件的绝对路径>"
```

⚠️ **一律走 `run.ps1`，不要直接调 `zemax.ps1`**：`run.ps1` 是纯 ASCII 启动器，会在调用前把
`zemax.ps1` / `zemax-lib.ps1` / `config.json` 自动修成 UTF-8 带 BOM。直接调 `zemax.ps1` 时，
任何一次"无 BOM 的编辑"都会让 PS 5.1 按 GBK 解码中文并当场语法报错。

| 命令 | 作用 | 主要参数 |
|---|---|---|
| `info` | 环境/许可自检 | — |
| `list-analyses` | 导出全部 `AnalysisIDM` / `MeritOperandType` / `SolveType` 名称 | — |
| `prescription` | 系统参数 + 逐面处方表 | `-File` |
| `mtf` | FFT MTF 全曲线 + 指定频率取值 | `-File` `-MaxFrequency 100` `-Freqs 30,50,100` |
| `spot` | `RMSField`（波前/点列半径）+ `GeometricEncircledEnergy`（含 r50/r80/r90） | `-File` |
| `report` | 处方 + MTF + 点列，并出 Excel/Word | `-File`，可加 `-NoOffice` |
| `design` | **按规格 JSON 建模 → 设变量 → 建评价函数 → 优化 → 对焦 → 存 `.zos`**，再对落盘文件做全套提取与验收比对 | `-Spec <design.json>`，可加 `-Out` `-Freqs` |

输出目录可以显式指定 `-Out <目录>`；默认每次新建 `Zemax\runs\<yyyyMMdd-HHmmss>_<镜头名>\`。

**命令的输出解析**：日志是给人看的，末尾有机器可读行，直接取这些而不是猜路径：

```
RUN_DIR=<本次运行目录>
DATA_JSON=<run>\data.json
MTF@30lp_per_mm[Field: 0.00 (deg)|Tangential]=0.694586
```

一次 `report` 典型耗时 6–10 秒（冷启动 OpticStudio 约 5 秒），产物：

| 文件 | 内容 |
|---|---|
| `data.json` | 全部原始数据（系统参数、处方、MTF 点集、点列点集与能量半径、各分析设置快照、指纹、许可） |
| `prescription.csv` / `mtf.csv` / `spot.csv` | 长表 CSV，可直接给 pandas |
| `spot-summary.csv` | 点列每条曲线的摘要：X/Y 范围 + r50/r80/r90（µm） |
| `<镜头名>_report.xlsx` | 概览 / 处方表 / MTF 关键频率 / MTF 曲线（含折线图）/ 点列摘要 / 点列曲线 / 元数据 |
| `<镜头名>_report.docx` | 概览 + 处方表 + MTF 关键值 + 点列数据与能量半径 + 方法与口径 |
| `input_<文件名>` | 输入文件副本（保证可追溯） |
| `run.log` | 完整运行日志 |

## 报告交付前的自检（必做）

先按 Office 技能的规矩，用 `check_office.py` 查结构与关键文本（DSH 自带 Office 技能会给出该脚本路径）：

```powershell
$cfg = Get-Content '<工作区>\Zemax\config.json' -Raw -Encoding UTF8 | ConvertFrom-Json
$py  = $cfg.pythonExe          # 安装器会自动填好；没有就跳过 Excel/Word（仍可用 -NoOffice）
$chk = '<DSH 安装目录>\resources\runtime\office-skills\scripts\check_office.py'
& $py $chk "<run>\<镜头名>_report.xlsx" --out "$env:TEMP\chk-x.json" --contains 处方表 --count 7
& $py $chk "<run>\<镜头名>_report.docx" --out "$env:TEMP\chk-d.json" --contains "FFT MTF"
```

还要核对数值本身：打开 xlsx 的「处方表」「MTF 关键频率」与 `data.json` 逐格比对，别只看结构检查通过。
设计任务另有「设计验收」工作表与 `acceptance.csv`，两者应与 `data.json` 的 `acceptance` 完全对应。

## 硬性约定

1. **绝不原地改写用户的镜头文件**。桥接只读源文件；一切写操作都落在 run 目录。需要修改镜头的后续阶段必须 `SaveAs` 到新文件。
2. **路径只能从 `config.json` 取**。`outputRoot`（默认 `Zemax\runs`）、`scratchRoot`、`zemaxInstall`、`zosApiHelper`、`pythonExe` 都在那里；也可用环境变量 `ZEMAX_CONFIG` / `ZEMAX_OUTPUT_ROOT` / `ZEMAX_SCRATCH_ROOT` 覆盖。不要把路径硬编码进新脚本。
3. **每个结论都要能追到数据**：报告里出现的每个数字都应能在同目录 `data.json` 里找到来源（源文件 SHA256 一并记录）。
4. 一次只跑一个实例：脚本会自己拉起并关闭无界面 OpticStudio；不要同时让 GUI 打开同一份文件做实验。

## 已踩过的坑（改脚本前务必读）

1. **`.ps1` 必须存为 UTF-8 带 BOM。** PowerShell 5.1 对无 BOM 文件按系统 ANSI（本机 GBK）解码，中文会变乱码并导致语法错误。用 `edit`/`write` 改过任何 `.ps1` 后，跑一次 `normalize-encoding.ps1`。
2. **多维数组不能用 `return` 返回。** PowerShell 会把函数输出流里的数组拆开，`double[,]` 会被展平成一维，两条曲线就读到同一段数据（曾导致切向/弧矢 MTF 完全相同）。取数组属性必须 `Write-Output -NoEnumerate`，库里的 `Get-Prop` 已经这么处理。
3. **`-File` 传入的 `"10,20,30"` 若参数声明为 `[double[]]`，会按区域设置当千位分隔符解析成 `10203040`。** 所以 `-Freqs` 参数类型故意是 `object`，由 `ConvertTo-FreqList` 显式切分。
4. **分析结果的层次是两层**：`IAR_Results.GetDataSeries(i)` 是"曲线组"（带 `Description`，如 `Field: 14.00 (deg)`），组内 `NumSeries`/`SeriesLabels` 才是子曲线（`Tangential`/`Sagittal` 或各波长）。`SeriesLabels` 要直接读属性，包一层 `@()` 会变成套娃并串成一个标签。
5. **分析名称必须真实存在**：`RmsVsField` 不存在，正确的是 `RMSField`。名字不确定时先跑 `list-analyses`，它会导出 156 个 `AnalysisIDM`、437 个评价函数 operand、39 个解法类型。
6. **每次运行都会打印 `FRU__delta_init(): Attempt to start when running!`**（来自本机许可环境的噪声），退出码仍为 0，可以忽略。
7. **指标要选对**：评价函数值下降不等于成像质量提升。实测中"最小 RMS 光斑"的默认评价函数优化后，评价函数 0.5716→0.4403，但 30 lp/mm 切向 MTF 从 0.6946 掉到 0.0108。做优化类任务必须用独立指标（MTF、包围能量、畸变等）验收，不要只报评价函数。
8. **提取结果这样自证**：`YData.Data` 是 `[点数, 子系列数]` 矩阵，取样点索引 `k` 处的切向/弧矢值就是 `Y[k,0]`、`Y[k,1]`，`XData.Data[k]` 是空间频率。报告里的每个 MTF 值都应能在原始矩阵里按索引复现（已按此交叉验证过）。
9. **评价函数操作数的单元是 1-based，且各操作数布局不同**：`cell[1]=Type(String)`；`EFFL` 的 `cell[3]` 表头是 `Wave`；`AXCL` 的 `cell[2]='Wave1'`、`cell[3]='Wave2'`、`cell[4]='Zone'(Double)`。`Target`/`Weight` 用属性写。`GetCellAt(0)` 会抛 `Value does not fall within the expected range`。**读值用 `$op.Value`**（已实测可用），取完记得把临时操作数删掉。
10. **同一个分析画什么量取决于设置**：`RMSField` 的 `Data` 默认是 `Wavefront`（RMS 波前误差，单位 waves）；要看点列半径必须显式设成 `SpotRadius`（单位 µm）。库里两个都跑、分别标注，每条曲线画的量都记在 `data.json` 的 `spot[].settings` 里。
11. **ZOS-API 不暴露逐条射线坐标**：`StandardSpot` / `FullFieldSpot` 的 `DataSeries`、`DataScatterPoints`、`DataGrids`、`RayData` 全部为 0，只有 `SpotData` 聚合量。所以"点列图数据"在数值层面只能拿到聚合指标，拿不到每个光线的落点。
12. **`SpotData` 的聚合量要挑着用**：`MaxRadius` 与包围能量曲线推出的 r90 在两个样例上吻合（可判为 µm），但 `MeanRadius` 无法自洽（两例 MaxRadius 比值 2.9，MeanRadius 比值却 10.2）。报告因此采用**由包围能量曲线插值得到的 r50/r80/r90**（单位 µm，可独立重算复核），`SpotData` 只作原始读数留档。
13. **能量半径可自证**：`E[R]` 一类问题都能由 EE 曲线直接算，报告里每个 r50/r80/r90 都能用 `spot.csv` 的 EE 采样点线性插值复现（校验脚本已按此实现）。
14. **`config.json` 也必须 UTF-8 带 BOM**：它的 `_comment` 是中文，PS 5.1 用默认编码读会解析失败（`ConvertFrom-Json` 直接报“传入的对象无效”）。`normalize-encoding.ps1` 现在会连 `config.json` 一起规范化；但注意 **run 产物（`data.json`）要保持无 BOM**，因为 Python 端按严格 UTF-8 读取。
15. **路径覆盖已实测可用**：`-Config <另一份配置>` 与环境变量 `ZEMAX_OUTPUT_ROOT` / `ZEMAX_SCRATCH_ROOT` 都能改变产物落点，换机器或换目录不需要改脚本。
16. **返回值别被"漏出的输出"污染**：`ChangeType()` / `MakeSolveVariable()` / `Apply()` / `RemoveOperandAt()` / `RunAndWaitForCompletion()` / `RunAndWaitWithTimeout()` / `Cancel()` / `WaitForCompletion()` 都会返回 bool，不 `[void]` 掉就会混进函数输出流，把对象返回值变成 `[True, {…}]` 这种数组。**这种污染很隐蔽**：PowerShell 的成员枚举会让 `.字段` 照常读到值、日志看起来完全正常，只有落到 JSON 里才暴露（`chromaticFocus` 变成 list、`design.optimize` 变成 `[true, {...}]`）。写完新函数后请检查 `data.json` 里对应字段的 `type` 是不是 `dict`。
17. **`Export-Csv` 对 `[ordered]@{}` 导出的是字典属性而不是键值**：会得到 `Count/Keys/Values/IsReadOnly…` 一列垃圾。`Write-ZemaxCsv` 已统一转成 `PSCustomObject`；自己写新导出时别直接用 `Export-Csv`。
18. **`Get-Prop` 必须显式支持字典**：PowerShell 的 `PSObject.Properties` 不含 `[ordered]@{}` 的键，用属性方式读字典键会静默拿到默认值（曾导致 `RMSField` 的 `Data` 读成空串、整批派生指标缺失）。库里的 `Get-Prop` 已加 `IDictionary` 分支。
19. **`AXCL`（近轴焦点差）与"逐波长最佳几何焦点差"是两个量，别把其中一个当错误**：实测本次设计 AXCL=+0.0758 mm，而逐波长 `QuickFocus` 得到 −0.0106 mm；两者都对，因为球差随波长变化。**第三个独立通路**是 `LongitudinalAberration` 分析（每波长一条曲线，取相对光瞳 x=0 处的 y 就是近轴焦点偏移），它与 AXCL 一致到 1e-5 mm，可用来互相验证。
20. **`RMSField` 的视场在 X 轴，不在曲线组里**：它的曲线组名是"Field is oriented along the +y direction."，子系列才是 `Poly`/各波长。所以"轴上"值 = x 最小那一端的 y，不能按组名匹配视场（这一点曾经写错，导致轴上指标整批取不到）。
21. **`CardinalPoints` 分析在 ZOS-API 里取不到数值**：曲线/网格/散点/RayData/消息全为 0。EFL 与 F 数只能走评价函数操作数 `EFFL` / `WFNO`（用 `Get-ZemaxOperandValue` 读 `$op.Value`）。
22. **近轴消色差 ≠ 成像消色差**：把 `AXCL(target=0)` 加进评价函数能把近轴色差压到 ~1e-5 mm，但本次实测最佳焦点色差反而从 0.011 mm 恶化到 0.062 mm、MTF@50 从 0.771 掉到 0.725。**验收判据要按仪器的用途来定**（成像看 MTF/点列，近轴量当诊断），否则很容易写出互相冲突的判据。

## 按规格做设计（`design` 命令）

`design` 读一份规格 JSON（示例见工作区根目录 `Design01_Achromatic_Objective_100mm_f5\design.json`），
流程是：**新建系统 → 写面型/材料 → 设变量 → 建评价函数 → 局部优化 → 锤形优化（限时）→ 快速对焦 → SaveAs `.zos`
→ 重新加载落盘文件 → 处方/MTF/点列/指标/验收全套提取 → 出报告**。规格字段：

| 字段 | 含义 |
|---|---|
| `aperture` / `fields` / `wavelengths` / `catalog` | 孔径类型与值、视场类型与列表、波长预设或数值、材料库 |
| `surfaces[]` | 起始处方：`index`（LDE 行号）/ `radius` / `thickness` / `material` / `semiDiameter` / `comment` |
| `variables[]` | 变量表：`surface` + `cell`（`Radius`/`Thickness`/`Conic`/`SemiDiameter`） |
| `merit.wizard` | 默认评价函数向导参数（`data`=1 为 RMS 光斑、`ring`、玻璃/空气边界） |
| `merit.operands[]` | 追加操作数（如 `EFFL` target=100 weight=10） |
| `optimize` | `local`（算法/cycles/cores）、`hammerSeconds`、`quickFocus` |
| `save.fileName` | 设计文件写进 run 目录的文件名 |
| `metrics[]` | 要读的操作数指标：`name`/`operand`/`wave`/`unit`/`label` |
| `acceptance[]` | **预先声明的判据**：`metric` + `min`/`max` + `label` |

判据可以引用两类指标：操作数指标（`efl`/`wfno`/`axcl`…）与派生指标（`mtf_axis_tan`、`mtf_worst_tan`、
`rms_wavefront_axis_poly`、`rms_spot_axis_poly`、`rms_spot_worst`、`r50_axis`/`r80_axis`/`r90_axis`、
`focus_shift_fc_mm`、`focus_span_mm`、`axial_color_paraxial_fc_mm`、`axial_color_paraxial_span_mm`）。
**先把判据写进规格再跑**，否则就是在往结果上凑；但也要按仪器用途来写（成像看 MTF/点列，近轴色差当诊断），
别同时塞进互相冲突的判据（见坑 22）。

## 参考资源（只读，具体路径见 `config.json`）

- ZOS-API 帮助：`<zemaxInstall>\ZOS-API\Help\ZOS-API.chm`
- 官方样例代码：`<%USERPROFILE%>\Documents\Zemax\ZOS-API Sample Code\{C#,Python,MATLAB,...}`
- 样例镜头：`config.json` 的 `samplesDir`（`Sequential\Objectives` 下有 Doublet、Cooke 40 degree field 等）
- 名称速查（离线可查）：`<工作区>\Zemax\reference\` 下的 156 个 `AnalysisIDM`、437 个 `MeritOperandType`、39 个 `SolveType`
- 授权：需要本机 ANSYS 许可服务支持 ZOS-API（`info` 命令会打印 `IsValidLicenseForAPI`）

## 扩展方向

- **优化/设计已经可用**：走 `design` 命令 + 规格 JSON（见上）。要加新的评价函数操作数，直接在规格的 `merit.operands` 里加一条即可；要加新的验收指标，加在 `metrics` + `acceptance` 里。
- **还没接的**：公差分析（灵敏度/Monte Carlo）、多重组态、热/压力环境、非序列模式。这些需要在 `zemax-lib.ps1` 里加新的 `Get-Zemax*` / `Invoke-Zemax*` 函数，并接到 `zemax.ps1` 的命令分发；对应工具入口是 `$sys.Tools.OpenTolerance()`、`$sys.MCE`、`$sys.SystemData.Environment`。
- 要减少每次 5 秒的启动开销、或让模型直接以原生工具形式调用时，再按 Phase 1 计划做常驻桥接 + MCP server（DSH 已内置 `@deepseek-ai/dsh-mcp-client`，工具会以 `mcp__zemax__*` 出现）。
