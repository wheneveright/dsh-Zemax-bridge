# 变更记录

## v0.2.0

**新增**
- `design` 命令：按规格 JSON 完成"新建系统 → 写面型/材料 → 设变量 → 建评价函数（可选默认函数向导 + 追加操作数）
  → 局部优化 → 限时锤形优化 → 快速对焦 → SaveAs `.zos` → 重新加载落盘文件 → 全套提取 → 逐条验收比对"。
- 设计验收体系：`metrics.csv` / `acceptance.csv`、Excel「设计验收」工作表、Word 验收表；
  支持操作数指标（`efl`/`wfno`/`axcl`…）与派生指标（MTF、RMS 波前/点列、r50/r80/r90、最佳焦点色差、近轴色差）。
- 轴向色差的独立测量：逐波长单独 `QuickFocus` 读像面位置；并从 `LongitudinalAberration` 曲线取
  近轴焦点偏移（与 `AXCL` 互为独立验证）。
- `run.ps1`：**纯 ASCII 启动器**，调用前自动把 `zemax.ps1` / `zemax-lib.ps1` / `config.json`
  修成 UTF-8 带 BOM，从机制上消除 GBK 解码导致的语法崩溃。
- 安装包：`install.ps1`（自动检测 OpticStudio / `ZOSAPI_NetHelper.dll` / DSH 自带 Python 并生成 `config.json`，
  自带 `info` 自检）、`README.md`、`INSTALL.md`、`manifest.json`、示例设计规格、`compare_runs.py`、
  离线名称速查（156 个 `AnalysisIDM` / 437 个 `MeritOperandType` / 39 个 `SolveType`）。

**修复**
- `Write-ZemaxCsv`：`Export-Csv` 对 `[ordered]@{}` 导出的是字典属性（`Count/Keys/Values…`）而不是键值，
  导致全部 CSV 内容为垃圾。现统一转 `PSCustomObject` 再导出。
- `Get-Prop`：`PSObject.Properties` 不包含字典键，读 `[ordered]@{}` 的键会静默取默认值
  （曾使 `RMSField` 的 `Data` 读成空串、整批派生指标缺失）。现支持 `IDictionary`。
- 未 `[void]` 的布尔返回调用（`ChangeType` / `MakeSolveVariable` / `Apply` / `RemoveOperandAt` /
  `RunAndWaitForCompletion` / `RunAndWaitWithTimeout` / `Cancel` / `WaitForCompletion`）会污染函数返回值，
  把对象变成 `[True, {…}]`；此类污染会被 PowerShell 的成员枚举掩盖，只在 JSON 结构里暴露。
- 分析结果解析：`RMSField` 的视场在 X 轴（曲线组名是视场方向），轴上值应取 x 最小处而非按组名匹配。
- 提取层修正：`double[,]` 经函数返回会被 PowerShell 展平，导致切向/弧矢 MTF 读到同一段数据（v0.1 修复，此处记录）。

**已知限制**
- 评价函数操作数单元 1-based（`cell[1]=Type`，`AXCL` 的 `cell[2]/[3]` 是 `Wave1`/`Wave2`，`cell[4]=Zone`）。
- `CardinalPoints` 分析在 ZOS-API 里不返回任何数值，EFL/F# 只能走 `EFFL`/`WFNO` 操作数。
- ZOS-API 不暴露逐条射线坐标（`StandardSpot` 的 `DataSeries`/`DataScatterPoints`/`RayData` 全为 0）；
  点列图数据只能用聚合量，报告采用可由包围能量曲线独立复算的 r50/r80/r90。
- 尚未接入：公差分析、多重组态、热/压力环境、非序列模式。

## v0.1.0

- 首个可用版本：`info` / `list-analyses` / `prescription` / `mtf` / `spot` / `report` 六个命令。
- ZOS-API 无界面实例连接、系统与处方提取、通用分析执行器（曲线组 × 子系列两层解析）、
  FFT MTF 与点列/包围能量提取、`data.json` + CSV + Excel/Word 报告。
- `config.json` 单一路径真源 + 环境变量覆盖；每次运行独立 `runs\<时间戳>_<镜头名>\` 目录，绝不改写源文件。
