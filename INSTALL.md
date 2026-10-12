# 安装与排错

## 0. 它到底做了什么

安装器把两部分东西放进**目标工作区**：

| 放哪 | 东西 | 作用 |
|---|---|---|
| `<工作区>\.dsh\skills\zemax\SKILL.md` | DSH 技能 | 让模型知道何时、如何调用这套桥接（DSH 自动扫描该目录） |
| `<工作区>\Zemax\...` | 工具本体 | 脚本、配置、示例、参考清单、运行产物 |

不写注册表、不需要管理员、不碰系统目录。想卸载直接删这两个目录即可。

## 1. 前置检查（30 秒）

在一台装好 OpticStudio 的机器上：

```powershell
# OpticStudio 是否装了（有 ZOSAPI.dll 即可）
Get-ChildItem 'HKLM:\SOFTWARE\ANSYS, Inc.' | Where-Object PSChildName -like 'ZemaxOpticStudio*' |
    ForEach-Object { (Get-ItemProperty $_.PSPath).installFolder }

# 许可服务是否在跑（ZOS-API 需要）
Get-Service | Where-Object { $_.Name -match 'ANSYS|FlexNet|lmgrd' } | Select Status,Name
```

要求：
1. **OpticStudio 的许可必须支持 ZOS-API**。安装器最后会跑 `info` 自检，输出里
   `license=...` 与 `apiValid=True` 才算可用；如果 `apiValid=False`，说明当前许可不含 API 权限。
2. **必须用 Windows PowerShell 5.1**（`powershell.exe`）。ZOS-API 是 .NET Framework 程序集，
   PowerShell 7（`pwsh`）加载会失败——这是硬约束，不是配置问题。
3. Office 报告需要 Python + `openpyxl`/`python-docx`。DSH 自带 Python 里已经有；
   安装器会在 `%USERPROFILE%\.dsh\dsh-runtimes\*\dependencies\python\python.exe` 里找。
   找不到也能装，只是报告退化成 JSON/CSV（用 `-NoOffice` 显式跳过）。

## 2. 安装

```powershell
# 最简：装进指定工作区
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -Workspace "D:\MyDshWorkspace"

# 技能同时装到用户级（对所有工作区可见，含没有 .git 的临时目录）
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -Workspace "D:\MyDshWorkspace" -SkillScope both

# 自动检测失败时手工指定
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -Workspace "D:\MyDshWorkspace" `
    -ZemaxInstall "E:\Ansys Zemax" -PythonExe "C:\Users\me\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"

# 只想更新脚本、保留现有 config.json（默认行为就是保留；要强制覆盖加 -Force）
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -Workspace "D:\MyDshWorkspace"
```

安装器是**幂等**的：重复运行只覆盖脚本与技能；已存在的 `config.json` 默认**不动**（要覆盖用 `-Force`，
覆盖前会自动备份为 `config.json.bak`）。

预期输出（关键几行）：

```
== 3/5 detect OpticStudio / helper dll / python
  OpticStudio      : E:\Ansys Zemax\
  ZOSAPI_NetHelper : C:\Users\<you>\Documents\Zemax\ZOS-API\Libraries\ZOSAPI_NetHelper.dll
  python (reports) : C:\Users\<you>\.dsh\dsh-runtimes\...\python.exe
== 5/5 self test
  running: run.ps1 info
    [..] OpticStudio 就绪：4.9s，license=PremiumEdition
    [..] info：license=PremiumEdition apiValid=True install=E:\Ansys Zemax\
  self test OK
```

## 3. 装完怎么用

```powershell
$cli = "D:\MyDshWorkspace\Zemax\scripts\run.ps1"

# 环境/许可自检
powershell.exe -NoProfile -ExecutionPolicy Bypass -File $cli info

# 完整报告（处方 + FFT MTF + 点列 + Excel/Word）
powershell.exe -NoProfile -ExecutionPolicy Bypass -File $cli report -File "C:\path\to\lens.zmx"

# 只看 MTF，指定空间频率
powershell.exe -NoProfile -ExecutionPolicy Bypass -File $cli mtf -File "lens.zmx" -MaxFrequency 100 -Freqs 30,50,100

# 按规格做设计（示例：消色差双胶合物镜 f=100mm F/5，13 条验收判据会逐条打印）
powershell.exe -NoProfile -ExecutionPolicy Bypass -File $cli design `
    -Spec "D:\MyDshWorkspace\Zemax\examples\achromat_design.json" -Freqs 30,50,100
```

每次运行都会输出 `RUN_DIR=<目录>`，产物都在那里面：

| 文件 | 内容 |
|---|---|
| `data.json` | 唯一数值真源（系统/处方/MTF 点集/点列/各分析设置快照/源文件 SHA256/许可） |
| `prescription.csv` `mtf.csv` `spot.csv` `spot-summary.csv` | 长表 CSV，直接给 pandas |
| `<镜头名>_report.xlsx` | 7–8 个工作表，含 MTF 折线图与「设计验收」 |
| `<镜头名>_report.docx` | 概览 + 处方 + MTF 关键值 + 点列 + 验收 |
| `metrics.csv` `acceptance.csv` | 仅设计任务：指标表与验收表 |
| `input_<文件名>` / `run.log` | 输入副本 / 完整日志 |

## 4. 排错

| 现象 | 原因 | 处理 |
|---|---|---|
| `InstallationException: Unable to locate Zemax OpticStudio` | `zemaxInstall` 为空或错 | 编辑 `config.json` 的 `zemaxInstall`，或在运行时设环境变量 |
| `License is not valid for ZOSAPI use` / `apiValid=False` | 许可不含 API 权限或许可服务没起 | 确认 ANSYS 许可服务在跑；检查许可类型 |
| `CreateNewApplication() 返回空` | 许可席位被占满 / OpticStudio 启动失败 | 关掉 GUI 里开着的实例，或换一个席位 |
| 脚本一跑就报一堆 `Unexpected token` 且中文变乱码 | `.ps1` 被存成 **UTF-8 无 BOM**，PS 5.1 按 GBK 解码 | 用 `run.ps1`（它会自动修），或手工跑 `scripts\normalize-encoding.ps1` |
| 报告没有 xlsx/docx，日志说"找不到 pythonExe" | `config.json` 的 `pythonExe` 为空或路径失效 | 在 config.json 里填上 DSH 自带 Python 的绝对路径，或加 `-NoOffice` 明确跳过 |
| 每次都弹出/起一个 OpticStudio | 正常：API 走的是无界面实例 | 跑完会自动关闭；不要在 GUI 里同时打开同一文件做实验 |
| 日志里有 `FRU__delta_init(): Attempt to start when running!` | 本机许可环境的噪声 | 退出码 0，忽略 |
| `-Freqs 10,20,30` 变成了 `10203040` | 旧版本的参数类型坑（v0.1 已修） | 用 v0.2.0 |

### 换机器要改什么

只改 `<工作区>\Zemax\config.json`：

```json
{
  "outputRoot":   "<新工作区>\\Zemax\\runs",
  "scratchRoot":  "<新工作区>\\Zemax\\scratch",
  "zemaxInstall": "E:\\Ansys Zemax",
  "zosApiHelper": "C:\\Users\\<你>\\Documents\\Zemax\\ZOS-API\\Libraries\\ZOSAPI_NetHelper.dll",
  "samplesDir":   "C:\\Users\\<你>\\Documents\\Zemax\\SAMPLES",
  "pythonExe":    "C:\\Users\\<你>\\.dsh\\dsh-runtimes\\...\\python.exe"
}
```

或者干脆重跑一次 `install.ps1 -Force`，让它重新检测并生成。

## 5. 在另一个 DSH 里让技能生效

* 技能目录：`<工作区>\.dsh\skills\`（DSH 扫描的"项目级"技能根）。
  若目标工作区**在某个 git 仓库内部**，项目根会是仓库根目录，此时请把技能装到
  `<仓库根>\.dsh\skills\`，或用 `-SkillScope user` 装到 `%USERPROFILE%\.dsh\skills`（对所有工作区可见）。
* 装好后无需重启 DSH：技能会在**下一个模型步骤**出现在目录里。
* 验证方式：直接问模型"现在有哪些 zemax 相关技能/工具"，或让它跑一次
  `run.ps1 info` 看输出。

## 6. 包内自检（可选）

```powershell
# 逐文件核对包是否被改动过（与 manifest.json 里的 SHA256 比对）
$manifest = Get-Content .\manifest.json -Raw -Encoding UTF8 | ConvertFrom-Json
$bad = 0
foreach ($f in $manifest.files) {
    $p = Join-Path $PSScriptRoot $f.path
    $h = (Get-FileHash $p -Algorithm SHA256).Hash
    if ($h -ne $f.sha256) { $bad++; "改动或缺失: $($f.path)" }
}
if ($bad -eq 0) { "manifest 校验通过（$($manifest.files.Count) 个文件）" }
```
