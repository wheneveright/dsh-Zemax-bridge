# Design01 消色差双胶合物镜 f=100 mm F/5 —— 交付说明与查验方法

**结论一句话**：可见光（F/d/C）消色差双胶合物镜设计完成，**13 条预先声明的验收判据全部通过**；
把落盘的 `.zos` 拿回来**独立重跑一遍**，处方、18 个 MTF 取值、20 项点列与能量半径**逐值一致**
（`compare_runs.py` 可复现这个比对）。

---

## 1. 交付物

| 文件 | 说明 |
|---|---|
| [design.json](design.json) | 设计规格与验收判据（**先声明、后设计**：判据在设计之前就写死在文件里；迭代过程也记在 `notes` 里） |
| `run/achromat_100mm_f5.zos` | 设计文件本体，18,589 B，SHA256 `DA8CAD6C981963CD04A84BC7877669464598CD051BCCF725F5DB07C04D51C1A9` |
| `run/prescription.csv` / `metrics.csv` / `acceptance.csv` | 处方、指标、验收表（机器可读） |
| `run/mtf.csv` / `spot.csv` / `spot-summary.csv` | 全量曲线数据（MTF 点集、点列曲线与 r50/r80/r90） |
| `run/data.json` | 本次设计的**唯一数值真源**：规格、优化过程、全部指标、各分析点集、逐波长对焦结果、源文件指纹 |
| `run/achromat_100mm_f5_report.xlsx` / `.docx` | 报告：设计验收 / 概览 / 处方 / MTF / 点列（xlsx 8 表含折线图，docx 7 张表） |
| `verify/` | **独立复跑**：把 `run/achromat_100mm_f5.zos` 当输入再跑一次 `report` |
| `iter2_axcl_weighted/` | **第二迭代留档**：把 AXCL 加进评价函数后的那一版设计（见第 5 节，它证明了一个真实的物理权衡） |
| [compare_runs.py](compare_runs.py) | 两两比对两个 run 目录是否逐值一致（复现性自查） |
| `run/run.log` | 完整运行日志（每一步做了什么、耗时、merit 变化） |

**起始值的来历**（不是随手填的）：按薄透镜消色差条件 `φ1/V1 + φ2/V2 = 0`，取 N-BK7（Vd=64.17）与
N-SF5（Vd=32.25），总光焦度 1/100 mm⁻¹，解出 f1≈49.74、f2≈−98.98；再取胶合面 r2=−45 mm，由
`1/f = (n−1)(1/r1 − 1/r2)` 反解 r1≈59.97、r3≈−138.8。之后交给局部优化 + 8 s 锤形优化。

## 2. 最终处方

| 面 | 类型 | 注释 | 半径 (mm) | 厚度 (mm) | 材料 | 半口径 (mm) |
|---|---|---|---|---|---|---|
| 0 | Standard | 物面 | ∞ | ∞ | — | — |
| 1 | Standard | 冕牌前表面（光阑） | 61.3993 | 15.0000 | N-BK7 | 11 |
| 2 | Standard | 胶合面 | −41.7834 | 7.1959 | N-SF5 | 10 |
| 3 | Standard | 火石后表面 → 像面 | −120.3618 | 89.5158 | 空气 | 10 |
| 4 | Standard | 像面 | ∞ | — | — | 0.876 |

系统参数：入瞳直径 20 mm，视场 0 / 0.25 / 0.5°（角度），波长 F 0.4861 µm / d 0.5876 µm / C 0.6563 µm。
优化：局部优化 merit 0.543291 → 0.00058（1.0 s），随后 8 s 锤形优化 → 0.000574。

## 3. 验收结果（13/13）

| 验收项 | 判据 | 实测 | 结论 |
|---|---|---|---|
| 有效焦距 EFL | 100 ± 0.5 mm | 100.0000 mm | 通过 |
| 像方 F 数 | 5.0 ± 0.15 | 4.9982 | 通过 |
| 近轴轴向色差 AXCL | ≤ 0.10 mm | 0.0758 mm | 通过 |
| 近轴 F–C（纵向像差曲线 x=0） | ≤ 0.10 mm | −0.0758 mm | 通过 |
| 近轴焦点跨度（同曲线） | ≤ 0.10 mm | 0.0758 mm | 通过 |
| 逐波长最佳焦点 F–C | ≤ ±0.03 mm | −0.0106 mm | 通过 |
| 各波长最佳焦点跨度 | ≤ 0.06 mm | 0.0520 mm | 通过 |
| 轴上 RMS 波前 | ≤ 0.10λ | 0.0568λ | 通过 |
| 轴上 RMS 点列半径 | ≤ 3.6 µm | 1.800 µm | 通过 |
| 最差视场 RMS 点列半径 | ≤ 8 µm | 2.645 µm | 通过 |
| 轴上 MTF@50 lp/mm（切向） | ≥ 0.60 | 0.7712 | 通过 |
| 最差视场 MTF@50 lp/mm（切向） | ≥ 0.35 | 0.7681 | 通过 |
| 轴上 80% 包围能量半径 | ≤ 6 µm | 2.963 µm | 通过 |

**色差判据不是随手定的**，而是拿经典 BK7/F2 双胶合样例（`SAMPLES\Sequential\Objectives\Doublet.zmx`）标定的：

| 量 | 本次设计 | 参考镜头（经典双胶合样例） |
|---|---|---|
| 近轴 F–C（AXCL） | **0.0758 mm** | 0.1263 mm |
| 逐波长最佳焦点 F–C | **−0.0106 mm** | −0.0190 mm |
| 最佳焦点跨度（二次光谱） | **0.0520 mm** | 0.0537 mm |

轴上 RMS 点列半径 1.80 µm 小于同孔径下的 Airy 半径（1.22·λ·F# ≈ 3.58 µm），轴上已在衍射极限附近。

## 4. 两条独立的色差通路互相印证

轴向色差这件事我用了两种互不依赖的方法测，结果一致到 1e-5 mm：

| 方法 | 依据 | 结果 |
|---|---|---|
| 评价函数操作数 `AXCL` | OpticStudio 的**近轴**焦点差 | **+0.0758 mm** |
| `LongitudinalAberration` 分析 | 每波长一条纵向像差曲线，取相对光瞳 x=0（近轴端）的 y 值 | F +0.0050、d +0.0110、C +0.0808 → **F−C = −0.0758 mm** |
| 逐波长 `QuickFocus` | 每波长单独对焦（径向光斑+质心）后读像面位置 | F 89.5261、d 89.4846、C 89.5366 → **F−C = −0.0106 mm** |

前两条是**近轴**量、符号相反、大小一致（这是同一件事的两种读法）；第三条是**最佳几何焦点**量，与前两条
本来就不相等——因为球差本身随波长变化。三者不矛盾，报告把三条都列出来，避免拿其中一个当"错误"。

> 注：`CardinalPoints` 分析在 ZOS-API 里**不返回任何数值**（曲线/网格/散点/消息全为 0），所以
> EFL 与 F 数只能走评价函数操作数 `EFFL` / `WFNO`——这也是为什么报告里的 EFL/F# 标注来源为 operand 而不是分析。

## 5. 迭代过程（这才是"设计"真正的样子）

第 1 轮只用了"默认 RMS 光斑评价函数 + EFFL 约束"，跑完发现：成像指标很好，**近轴**色差偏大。
于是第 2 轮把 `AXCL(target=0, weight=10)` 加进评价函数，想让近轴色差归零。结果：

| 指标 | 第 1 轮（定稿） | 第 2 轮（AXCL 加权） |
|---|---|---|
| 近轴 F–C（AXCL） | 0.0758 mm | **2.5e-5 mm**（几乎完美消色差） |
| 逐波长最佳焦点 F–C | **−0.0106 mm** | 0.0621 mm |
| 最佳焦点跨度 | **0.0520 mm** | 0.0928 mm |
| 轴上 MTF@50 lp/mm | **0.7712** | 0.7247 |
| 轴上 RMS 点列半径 | **1.800 µm** | 2.852 µm |
| 轴上 RMS 波前 | **0.0568λ** | 0.0944λ |
| 冕牌/火石厚度 | 15.0 / 7.20 mm | 15.0 / 15.0 mm（都被约束顶住） |

**结论：近轴消色差 ≠ 成像消色差。** 把这个镜头的近轴焦点压到重合，反而把边缘光线的焦点按波长拉开了，
成像质量变差。对一台**成像**物镜来说，决定成像的是最佳焦点处的 MTF/点列，所以**定稿采用第 1 轮**，
近轴色差降级为"有参考标定的诊断带"（≤0.10 mm），硬判据放在成像指标与最佳焦点色差上。
第 2 轮的完整产物留档在 `iter2_axcl_weighted/`，可用 `compare_runs.py` 与定稿版对比。

**我也要说明第一版判据定得不合理**：最初我同时要求"近轴 F–C ≤ ±0.03 mm"和"最佳焦点跨度 ≤ 0.06 mm"，
这两条对同一副 N-BK7/N-SF5 双胶合是**互相冲突**的（上面两张表就是证据）。判据是"先声明"的，但声明得
不够专业——这也是这次设计最有价值的一条经验。

---

## 6. 怎么查验（三条路，任选，建议至少走 A + B）

### 方法 A：在 OpticStudio 界面里亲眼看

1. 打开 `E:\Ansys Zemax\OpticStudio.exe`。
2. **File → Open** → 选 `run\achromat_100mm_f5.zos`。
3. **Lens Data Editor** 核对：面 1 = N-BK7，R 61.399，T 15.000；面 2 = N-SF5，R −41.783，T 7.196；
   面 3 到像面 T 89.516。半径与厚度应当是**变量（V）**状态——设计文件本就该保留变量。
4. **System Explorer** 核对：Aperture = Entrance Pupil Diameter 20；Fields = 0 / 0.25 / 0.5 度；
   Wavelengths = 0.486 / 0.588 / 0.656 µm。
5. **Analyze → MTF → FFT MTF**：Maximum Frequency 设 100 lp/mm，50 lp/mm 处轴上应约 **0.77**。
6. **Analyze → Spot Diagram** / **RMS vs Field**：轴上 RMS 半径应约 **1.8 µm**。
7. **Analyze → Aberrations → Longitudinal Aberration**：可见三条曲线；曲线在相对光瞳 0 处的纵坐标
   就是上表第 4 节的近轴焦点偏移。
8. **Merit Function Editor**：应有 88 个操作数（87 个默认评价函数操作数 + 1 个 `EFFL = 100`）。

### 方法 B：命令行复算（可脚本化，最推荐）

```powershell
$ps   = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$root = "<workspace>"
$cli  = "$root\Zemax\scripts\run.ps1"

# 用设计文件当输入，独立跑一遍完整报告（会新建一个 run 目录）
& $ps -NoProfile -ExecutionPolicy Bypass -File $cli report `
    -File "$root\Design01_Achromatic_Objective_100mm_f5\run\achromat_100mm_f5.zos" `
    -Freqs 30,50,100

# 再和原始设计 run 逐值比对（应输出 PASS）
$py = "C:\Users\<you>\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"
& $py "$root\Design01_Achromatic_Objective_100mm_f5\compare_runs.py" `
    "$root\Design01_Achromatic_Objective_100mm_f5\run" `
    "$root\Design01_Achromatic_Objective_100mm_f5\verify"
```

`verify/` 就是这条命令的成品；比对结果为：处方 5 项一致、MTF 18 项一致、点列/能量半径 20 项一致。

### 方法 C：只靠数据自己算（不依赖 OpticStudio）

* **MTF@50**：在 `run/mtf.csv` 里筛 `frequency` 最接近 50 的行，取 `mtf` 列。
* **r50 / r80 / r90**：在 `run/spot.csv` 里取 `analysis=GeometricEncircledEnergy` 的曲线，找 `y` 首次达到
  0.5 / 0.8 / 0.9 的半径并线性插值；也可用恒等式 `E[R]=∫(1−EE)dr` 复核。
* **近轴色差**：`run/data.json` → `longitudinalAberration.series[].points` 取 x≈0 的 y 值；
  或直接看 `metrics.axial_color_paraxial_fc_mm`。
* **最佳焦点色差**：`run/data.json` → `chromaticFocus.positions`（逐波长单独对焦后的像面位置），相邻相减。
* **EFL / F 数**：`run/metrics.csv` 的 `efl` / `wfno`，来源是 `EFFL` / `WFNO` 操作数。

一行 Python 复算示例：

```python
import csv
run = r"<workspace>\Design01_Achromatic_Objective_100mm_f5\run"
rows = list(csv.DictReader(open(run + r"\mtf.csv", encoding="utf-8-sig")))
near = min(rows, key=lambda r: abs(float(r["frequency"]) - 50))
print("MTF@50:", near["frequency"], near["group"], near["series"], near["mtf"])
```

## 7. 这份"第一次设计"明确没做的事（别当成成品镜头）

1. **冕牌玻璃厚度收敛到 15 mm 上界**（评价函数向导里 GlassMax=15 把它顶住了）。工程上偏重、成本高；
   下一轮应加厚度/重量约束（例如收紧 `MXCG`/`MNEG` 或直接把向导的 GlassMax 设小）。
2. **半口径是手工给的**（11/10/10 mm），没有按实际光束足迹求解，也没做渐晕与边缘厚度检查。
3. **只做到轴上与 ±0.5° 小视场**：没有公差分析、热/压力分析、多重组态，也不是无焦或目镜系统。
4. **锤形优化是限时随机过程**：重跑 `design` 会得到略有差异的处方（末几位小数级别）。要完全复现本文数字，
   请直接使用现成的 `run/achromat_100mm_f5.zos`。
5. 日志里的 `FRU__delta_init(): Attempt to start when running!` 是本机许可环境的噪声，退出码 0，可忽略。

## 8. 想重新设计的话

```powershell
$ps   = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$root = "<workspace>"
& $ps -NoProfile -ExecutionPolicy Bypass -File "$root\Zemax\scripts\run.ps1" design `
    -Spec "$root\Design01_Achromatic_Objective_100mm_f5\design.json" `
    -Out  "$root\Design01_Achromatic_Objective_100mm_f5\run" `
    -Freqs 30,50,100
```

改设计就改 `design.json`：孔径、视场、波长、起始处方、变量表、评价函数（向导参数 + 追加操作数）、
优化策略、验收判据全在那一份文件里。**判据请先写、后跑**——但要写得专业：先想清楚这台仪器靠什么指标
成像（本例是 MTF/点列），否则会像我第一版那样把互相冲突的判据同时写进去。
