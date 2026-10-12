# dsh-zemax-bridge

把 **DeepSeek Harness (DSH)** 接到本机 **Ansys Zemax OpticStudio** 上：读镜头文件、算光学指标、
出 Excel/Word 报告，并支持**按规格做设计**（建模 → 评价函数 → 优化 → 对焦 → 存 `.zos` → 逐条验收）。

全部通过 ZOS-API 驱动**无界面实例**，不需要安装任何第三方库：DSH 自带的 shell 就是
Windows PowerShell 5.1（.NET Framework 4.8），可以直接加载 ZOS-API 的 .NET 程序集。

```
$ dsh-zemax-bridge
  install.ps1                 一键安装（自动检测 OpticStudio / NetHelper / Python，写配置、跑自检）
  Zemax/scripts/              桥接脚本（PowerShell 5.1 + 一个 Python 报告渲染器）
  .dsh/skills/zemax/SKILL.md  DSH 技能说明（含 22 条踩坑记录）
  examples/                   示例设计规格、复现性比对工具、完整设计案例
  reference/                  156 个 AnalysisIDM / 437 个 merit operand / 39 个 solve type 离线清单
  INSTALL.md  CHANGELOG.md  LICENSE
```

## 它能做什么

| 命令 | 作用 |
|---|---|
| `info` | 环境/许可自检（`IsValidLicenseForAPI`、安装目录、Python） |
| `list-analyses` | 导出全部 `AnalysisIDM` / `MeritOperandType` / `SolveType` 名称 |
| `prescription` | 系统参数 + 逐面处方表（半径/厚度/材料/半口径/光阑） |
| `mtf` | FFT MTF 全曲线 + 指定空间频率处的取值（按视场、切向/弧矢） |
| `spot` | RMS 波前误差 / RMS 点列半径 vs 视场 + 几何包围能量与 r50/r80/r90 |
| `report` | 处方 + MTF + 点列 → `data.json` / CSV / **Excel（8 表含折线图）** / **Word** |
| `design` | 按规格 JSON 建模 → 设变量 → 建评价函数 → 局部+锤形优化 → 快速对焦 → 存 `.zos` → 逐条验收比对 |

## 快速开始

```powershell
# 1) 安装到目标 DSH 工作区（OpticStudio 需已安装并授权支持 ZOS-API）
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -Workspace "D:\MyDshWorkspace"

# 2) 出第一份报告
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "D:\MyDshWorkspace\Zemax\scripts\run.ps1" `
    report -File "<某个 .zmx / .zos>"

# 3) 跑一遍示例设计（消色差双胶合物镜 f=100mm F/5，13 条验收判据逐条打印）
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "D:\MyDshWorkspace\Zemax\scripts\run.ps1" `
    design -Spec "D:\MyDshWorkspace\Zemax\examples\achromat_design.json" -Freqs 30,50,100
```

安装细节、换机器改什么、故障对照表见 [INSTALL.md](INSTALL.md)。

## 设计案例（`examples/design01_achromat/`）

一次完整的设计迭代被原样留档，包括**失败的那一轮**：

| 指标 | 第 1 轮（定稿） | 第 2 轮（把 AXCL 加进评价函数） |
|---|---|---|
| 近轴 F–C（AXCL） | 0.0758 mm | **2.5e-5 mm**（近轴完美消色差） |
| 逐波长最佳焦点 F–C | **−0.0106 mm** | 0.0621 mm |
| 轴上 MTF@50 lp/mm | **0.7712** | 0.7247 |
| 轴上 RMS 点列半径 | **1.800 µm** | 2.852 µm |

结论：**近轴消色差 ≠ 成像消色差**。把近轴焦点压重合，反而把边缘光线的焦点按波长拉开。
最终按"成像优先"定稿第 1 轮，近轴色差降级为有参考标定的诊断带。完整数据、处方、验收表与
查验方法见 `examples/design01_achromat/DESIGN_NOTES.md`。

## 设计上的几个刻意取舍

* **路径只有一处真源**：`config.json`；另有 `ZEMAX_CONFIG` / `ZEMAX_OUTPUT_ROOT` / `ZEMAX_SCRATCH_ROOT`
  可覆盖，换机器不用改脚本。
* **绝不改写源文件**：桥接只读镜头文件；产物一律落进 `Zemax\runs\<时间戳>_<镜头名>\`，输入会复制一份留痕。
* **每个数字可追**：`data.json` 是唯一数值真源（含源文件 SHA256、各分析的设置快照），报告里每个数都能在
  它里面找到来源，也能用 `examples/compare_runs.py` 两次运行逐值比对。
* **宁可少报也不乱报**：口径无法自证的量（如 `SpotData.MeanRadius`）只留档、不进结论；点列图数据改用可由
  包围能量曲线独立复算的 r50/r80/r90。
* **纯 ASCII 启动器**：`run.ps1` 不含中文，每次调用前会把其余脚本与配置修成 UTF-8 带 BOM，
  规避 PowerShell 5.1 按 GBK 解码中文导致的语法崩溃。

## 前置条件

| 项 | 要求 |
|---|---|
| OpticStudio | 已安装且许可支持 ZOS-API（`info` 中 `IsValidLicenseForAPI=True`）。已在 2024 R1 实测 |
| Shell | Windows PowerShell 5.1（`powershell.exe`）。**不要用 PowerShell 7**，它加载 ZOS-API 会失败 |
| Office 报告 | 可选。DSH 自带 Python 含 `openpyxl`/`python-docx`；缺省则只出 JSON/CSV |
| 许可席位 | 每次调用会拉起一个无界面实例并占用一个席位，跑完即关闭 |

## 许可

[MIT](LICENSE)。本项目仅调用你本机已安装的 OpticStudio，**不包含也不分发**任何 Zemax/Ansys 二进制或
许可文件；相关商标归其各自所有者。
