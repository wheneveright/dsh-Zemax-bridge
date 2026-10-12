#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""build_report.py — 把 zemax.ps1 产出的 data.json 变成 Excel / Word 报告。

用法：
    <python> build_report.py --data <run>/data.json --outdir <run>

输出（写入 --outdir）：
    <lens>_report.xlsx   概览 / 处方 / MTF 曲线 + 关键频率 / 点列 / 元数据（含 MTF 折线图）
    <lens>_report.docx   概览 + 处方表 + MTF 关键值 + 方法说明

依赖：openpyxl、python-docx（均由 DSH 自带 Python 提供，不需要安装任何东西）。
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from datetime import datetime

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

INF = "∞"      # 半径/厚度为无穷（平面或虚拟面）
NA = "—"       # 无数据


# ------------------------------------------------------------------ 小工具

def num(value, digits=6):
    """把 JSON 里的数值格式化；None 视为无穷/缺失。"""
    if value is None:
        return None
    if isinstance(value, str):
        try:
            value = float(value)
        except ValueError:
            return value
    if isinstance(value, (int, float)):
        if value == 0:
            return 0
        text = f"{value:.{digits}f}".rstrip("0").rstrip(".")
        return float(text) if text else 0
    return value


def cell_value(value):
    """openpyxl 只接受标量；列表/字典转成可读文本。"""
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, dict):
        return "; ".join(f"{k}={v}" for k, v in value.items())
    if isinstance(value, (list, tuple, set)):
        return ", ".join(str(cell_value(v)) for v in value)
    return str(value)


def spot_quantity(a):
    """返回 (量的名称, 单位)。必须依据分析设置判断：RMSField 的 Data 可以是
    SpotRadius（µm）或 Wavefront（waves），默认值是 Wavefront，张冠李戴会让报告失真。"""
    settings = a.get("settings") or {}
    kind = settings.get("Data") or settings.get("Type") or ""
    if a.get("id") == "RMSField":
        if kind == "SpotRadius":
            return "RMS 点列半径", "µm"
        if kind == "Wavefront":
            return "RMS 波前误差", "waves"
    if a.get("id") == "GeometricEncircledEnergy":
        return "几何包围能量", "Fraction"
    return a.get("quantity") or a.get("id") or "未知量", ""


def series_name(item):
    """曲线名：OpticStudio 用「曲线组（视场）+ 子系列（波长）」两层描述。"""
    group = (item.get("group") or "").strip()
    label = (item.get("label") or "").strip()
    if group and label:
        return f"{group} / {label}"
    return group or label or "未命名曲线"


def lens_label(data):
    src = (data.get("source") or {}).get("path") or ""
    label = os.path.splitext(os.path.basename(src))[0] if src else "zemax"
    return re.sub(r"[^\w.\-]+", "_", label) or "zemax"


def style_header(ws, row=1, ncols=1):
    fill = PatternFill("solid", fgColor="1F4E79")
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = ws.cell(row=row + 1, column=1)


def autosize(ws, min_width=9, max_width=46):
    for col in range(1, ws.max_column + 1):
        width = min_width
        for row in range(1, ws.max_row + 1):
            value = ws.cell(row=row, column=col).value
            if value is None:
                continue
            # 中文字符按 2 个宽度估算
            length = sum(2 if ord(ch) > 127 else 1 for ch in str(value))
            width = max(width, min(length + 2, max_width))
        ws.column_dimensions[get_column_letter(col)].width = width


def set_run_font(run, size=None, bold=None):
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold


# ------------------------------------------------------------------ Excel

def build_xlsx(data, outdir, label):
    path = os.path.join(outdir, f"{label}_report.xlsx")
    wb = Workbook()

    # ---- 概览 ----
    ws = wb.active
    ws.title = "概览"
    system = data.get("system") or {}
    zemax = data.get("zemax") or {}
    src = data.get("source") or {}
    ws.append(["项", "值"])
    rows = [
        ("源文件", src.get("path", NA)),
        ("文件大小 (B)", src.get("bytes", NA)),
        ("SHA256", src.get("sha256", NA)),
        ("OpticStudio 版本", "2024 R1 (ZOS-API)"),
        ("许可", zemax.get("license", NA)),
        ("安装目录", zemax.get("install", NA)),
        ("系统模式", system.get("mode", NA)),
        ("孔径类型", (system.get("aperture") or {}).get("type", NA)),
        ("孔径值", (system.get("aperture") or {}).get("value", NA)),
        ("视场数", len(system.get("fields") or [])),
        ("波长数", len(system.get("wavelengths") or [])),
        ("面数", len(data.get("prescription") or [])),
        ("生成时间", data.get("generatedAt", NA)),
        ("桥接工具", f"{data.get('tool', '')} v{data.get('version', '')}"),
    ]
    for item in rows:
        ws.append(list(item))

    ws.append([])
    ws.append(["视场", "X", "Y", "权重"])
    field_header = ws.max_row
    for f in system.get("fields") or []:
        ws.append([f.get("number"), num(f.get("x")), num(f.get("y")), num(f.get("weight"), 4)])

    ws.append([])
    ws.append(["波长序号", "波长 (µm)", "权重", "主波长"])
    wave_header = ws.max_row
    for w in system.get("wavelengths") or []:
        ws.append([w.get("number"), num(w.get("value")), num(w.get("weight"), 4),
                   "是" if w.get("isPrimary") else ""])

    for header_row in (1, field_header, wave_header):
        style_header(ws, row=header_row, ncols=4)
    autosize(ws)
    ws.column_dimensions["B"].width = 56

    # ---- 设计验收（仅设计任务）----
    design = data.get("design")
    if design:
        wsd = wb.create_sheet("设计验收", 1)
        wsd.append(["设计", design.get("name", ""), "", ""])
        wsd.append(["说明", design.get("notes", ""), "", ""])
        saved = design.get("saved", "")
        wsd.append(["设计文件", saved, "", ""])
        opt = design.get("optimize") or {}
        local = opt.get("local") or {}
        hammer = opt.get("hammer") or {}
        wsd.append(["局部优化 merit", f"{local.get('meritBefore')} → {local.get('meritAfter')}",
                    "耗时 (s)", local.get("seconds")])
        if hammer:
            wsd.append(["锤形优化 merit", f"{hammer.get('meritBefore')} → {hammer.get('meritAfter')}",
                        "耗时 (s)", hammer.get("seconds")])
        wsd.append([])
        head = wsd.max_row + 1
        wsd.append(["验收项", "判据", "实测", "单位", "结论", "指标来源"])
        metrics = data.get("metrics") or {}
        for row in data.get("acceptance") or []:
            source = (metrics.get(row.get("metric")) or {}).get("operand", "")
            verdict = "通过" if row.get("pass") is True else ("未通过" if row.get("pass") is False else "无数据")
            wsd.append([cell_value(row.get("label", row.get("metric"))), cell_value(row.get("rule", "")),
                        cell_value(row.get("actual")), cell_value(row.get("unit", "")), verdict,
                        cell_value(source)])
        style_header(wsd, row=head, ncols=6)
        autosize(wsd)
        wsd.column_dimensions["A"].width = 46
        wsd.column_dimensions["B"].width = 16

    # ---- 处方表 ----
    ws2 = wb.create_sheet("处方表")
    ws2.append(["面号", "类型", "注释", "半径 (mm)", "厚度 (mm)", "材料", "半口径 (mm)", "圆锥系数", "光阑"])
    for s in data.get("prescription") or []:
        radius = num(s.get("radius"))
        thick = num(s.get("thickness"))
        conic = num(s.get("conic"))
        semi = num(s.get("semiDiameter"))
        ws2.append([
            s.get("surface"),
            s.get("type", ""),
            s.get("comment", ""),
            INF if radius is None else radius,
            INF if thick is None else thick,
            s.get("material", ""),
            INF if semi is None else semi,
            INF if conic is None else conic,
            "是" if s.get("isStop") else "",
        ])
    style_header(ws2)
    autosize(ws2)

    # ---- MTF 关键频率 ----
    mtf = data.get("mtf") or {}
    ws3 = wb.create_sheet("MTF 关键频率")
    ws3.append(["空间频率 (lp/mm)", "曲线组", "曲线", "实际频率 (lp/mm)", "MTF"])
    for row in mtf.get("atFreqs") or []:
        for per in row.get("perSeries") or []:
            value = num(per.get("value"))
            ws3.append([num(row.get("frequency")), per.get("group", ""), per.get("label", ""),
                        num(per.get("actualX")), NA if value is None else value])
    style_header(ws3)
    autosize(ws3)

    # ---- MTF 曲线（长表 + 折线图）----
    analysis = mtf.get("analysis")
    if not isinstance(analysis, dict):
        analysis = {}
    series = analysis.get("series") or []
    ws4 = wb.create_sheet("MTF 曲线")
    ws4.append(["频率 (lp/mm)"] + [series_name(s) for s in series])
    by_series = []
    for s in series:
        pts = s.get("points") or []
        by_series.append([(p[0], p[1]) for p in pts])
    if by_series:
        base = by_series[0]
        for k, (freq, value) in enumerate(base):
            row = [freq]
            for series_pts in by_series:
                row.append(series_pts[k][1] if k < len(series_pts) else None)
            ws4.append(row)
    style_header(ws4)
    autosize(ws4)
    if by_series and ws4.max_row > 1:
        chart = LineChart()
        chart.title = "FFT MTF"
        chart.x_axis.title = "空间频率 (lp/mm)"
        chart.y_axis.title = "调制传递函数"
        chart.height, chart.width = 9, 18
        data_ref = Reference(ws4, min_col=2, max_col=1 + len(series), min_row=1, max_row=ws4.max_row)
        cats_ref = Reference(ws4, min_col=1, min_row=2, max_row=ws4.max_row)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        ws4.add_chart(chart, f"{get_column_letter(len(series) + 3)}2")

    # ---- 点列摘要（每条曲线的量、单位、取值范围、能量半径）----
    ws7 = wb.create_sheet("点列摘要")
    ws7.append(["分析", "量", "单位", "曲线组", "曲线", "X 范围", "最小值", "最大值", "末值",
                "r50 (µm)", "r80 (µm)", "r90 (µm)"])
    for a in data.get("spot") or []:
        name, unit = spot_quantity(a)
        for s in a.get("series") or []:
            pts = s.get("points") or []
            if not pts:
                continue
            xs = [p[0] for p in pts]
            ys = [p[1] for p in pts]
            radii = {r.get("fraction"): r.get("radiusUm") for r in (s.get("energyRadii") or [])}
            ws7.append([a.get("id"), name, unit, s.get("group", ""), s.get("label", ""),
                        f"{min(xs)} … {max(xs)}", min(ys), max(ys), ys[-1],
                        radii.get(0.5, ""), radii.get(0.8, ""), radii.get(0.9, "")])
    style_header(ws7)
    autosize(ws7)

    diagram = data.get("spotDiagram") or {}
    if diagram:
        metrics = diagram.get("metrics") or {}
        ws7.append([])
        ws7.append(["点列图原始读数（口径未确认，报告未采用）", "值", "单位"])
        block = ws7.max_row
        for label, key, unit in (("最大半径（原始读数）", "maxRadiusRaw", "读数单位未确认"),
                                 ("平均半径（原始读数）", "meanRadiusRaw", "读数单位未确认"),
                                 ("半宽 X", "halfWidthX", "mm"), ("半宽 Y", "halfWidthY", "mm"),
                                 ("视场数", "numberOfFields", ""), ("波长数", "numberOfWaves", "")):
            ws7.append([label, metrics.get(key), unit])
        ws7.append(["原因", diagram.get("note", ""), ""])
        style_header(ws7, row=block, ncols=3)

    # ---- 点列 ----
    ws5 = wb.create_sheet("点列曲线")
    ws5.append(["分析", "量", "单位", "曲线组", "曲线", "X", "Y"])
    for a in data.get("spot") or []:
        name, unit = spot_quantity(a)
        for s in a.get("series") or []:
            for p in s.get("points") or []:
                ws5.append([a.get("id"), name, unit, s.get("group", ""), s.get("label", ""), p[0], p[1]])
    style_header(ws5)
    autosize(ws5)

    # ---- 元数据 ----
    ws6 = wb.create_sheet("元数据")
    ws6.append(["键", "值"])
    host = data.get("host") or {}
    acceptance_rows = data.get("acceptance") or []
    acceptance_summary = (f"通过 {sum(1 for r in acceptance_rows if r.get('pass') is True)} / {len(acceptance_rows)}"
                          if acceptance_rows else NA)
    meta = [
        ("工具", data.get("tool")),
        ("版本", data.get("version")),
        ("命令", data.get("command")),
        ("生成时间", data.get("generatedAt")),
        ("主机", host.get("machine")),
        ("用户", host.get("user")),
        ("PowerShell", host.get("powershell")),
        ("CLR", host.get("clr")),
        ("API 许可有效", zemax.get("apiValid")),
        ("MTF 最高频率", mtf.get("maxFrequency")),
        ("请求频率", mtf.get("requested")),
        ("分析状态", analysis.get("status")),
        ("分析备注", analysis.get("note")),
        ("点列分析", ", ".join(f"{a.get('id')}({spot_quantity(a)[0]}):{a.get('status')}"
                               for a in data.get("spot") or [])),
        ("点列图原始读数", (f"最大半径 {((data.get('spotDiagram') or {}).get('metrics') or {}).get('maxRadiusRaw')}"
                      f"、平均半径 {((data.get('spotDiagram') or {}).get('metrics') or {}).get('meanRadiusRaw')}"
                      f"（口径未确认，报告未采用）" if data.get("spotDiagram") else NA)),
        ("设计文件", (design or {}).get("saved", NA) if design else NA),
        ("验收结论", acceptance_summary),
        ("错误", "; ".join(data.get("errors") or []) or "无"),
    ]
    for item in meta:
        ws6.append([cell_value(item[0]), cell_value(item[1])])
    style_header(ws6)
    autosize(ws6)

    wb.save(path)
    return path


# ------------------------------------------------------------------ Word

def build_docx(data, outdir, label):
    path = os.path.join(outdir, f"{label}_report.docx")
    doc = Document()

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(10.5)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    doc.add_heading("OpticStudio 光学系统数据报告", level=0)

    src = data.get("source") or {}
    zemax = data.get("zemax") or {}
    system = data.get("system") or {}
    mtf = data.get("mtf") or {}
    analysis = mtf.get("analysis")
    if not isinstance(analysis, dict):
        analysis = {}

    p = doc.add_paragraph()
    set_run_font(p.add_run(f"源文件：{src.get('path', NA)}"))
    p = doc.add_paragraph()
    set_run_font(p.add_run(
        f"由 {data.get('tool')} v{data.get('version')} 经 ZOS-API 于 {data.get('generatedAt')} 生成；"
        f"许可 {zemax.get('license', NA)}。"))

    design = data.get("design")
    if design:
        doc.add_heading("设计任务与验收", level=1)
        para = doc.add_paragraph()
        set_run_font(para.add_run(str(design.get("name", ""))), bold=True)
        note = design.get("notes")
        if note:
            pn = doc.add_paragraph()
            set_run_font(pn.add_run(str(note)))
        opt = design.get("optimize") or {}
        local = opt.get("local") or {}
        hammer = opt.get("hammer") or {}
        steps = [f"设计文件：{design.get('saved', NA)}",
                 f"局部优化 merit：{local.get('meritBefore')} → {local.get('meritAfter')}（{local.get('seconds')} s）"]
        if hammer:
            steps.append(f"锤形优化 merit：{hammer.get('meritBefore')} → {hammer.get('meritAfter')}（{hammer.get('seconds')} s）")
        if (opt.get("quickFocus") or {}).get("done"):
            steps.append(f"快速对焦：{(opt.get('quickFocus') or {}).get('criterion')}")
        for text in steps:
            pp = doc.add_paragraph(text, style="List Bullet")
            for run in pp.runs:
                set_run_font(run)

        acceptance_rows = data.get("acceptance") or []
        if acceptance_rows:
            passed = sum(1 for r in acceptance_rows if r.get("pass") is True)
            ph = doc.add_paragraph()
            set_run_font(ph.add_run(f"验收判据逐条比对：通过 {passed} / {len(acceptance_rows)}"), bold=True)
            table = doc.add_table(rows=1, cols=6)
            table.style = "Light Grid Accent 1"
            metrics = data.get("metrics") or {}
            for i, head in enumerate(["验收项", "判据", "实测", "单位", "结论", "来源"]):
                set_run_font(table.rows[0].cells[i].paragraphs[0].add_run(head), bold=True)
            for row in acceptance_rows:
                cells = table.add_row().cells
                verdict = "通过" if row.get("pass") is True else ("未通过" if row.get("pass") is False else "无数据")
                source = (metrics.get(row.get("metric")) or {}).get("operand", "")
                values = [row.get("label", row.get("metric")), row.get("rule", ""),
                          row.get("actual"), row.get("unit", ""), verdict, source]
                for i, value in enumerate(values):
                    text = cell_value(value)
                    set_run_font(cells[i].paragraphs[0].add_run(NA if text is None else str(text)))

    doc.add_heading("1. 系统概览", level=1)
    ap = system.get("aperture") or {}
    table = doc.add_table(rows=0, cols=2)
    table.style = "Light Grid Accent 1"
    overview = [
        ("系统模式", system.get("mode", NA)),
        ("孔径", f"{ap.get('type', NA)} = {ap.get('value', NA)}"),
        ("视场数", len(system.get("fields") or [])),
        ("波长数", len(system.get("wavelengths") or [])),
        ("面数", len(data.get("prescription") or [])),
        ("波长 (µm)", ", ".join(str(num(w.get("value"))) for w in system.get("wavelengths") or [])),
    ]
    for key, value in overview:
        cells = table.add_row().cells
        set_run_font(cells[0].paragraphs[0].add_run(str(key)), bold=True)
        set_run_font(cells[1].paragraphs[0].add_run(str(value)))

    doc.add_heading("2. 镜头处方", level=1)
    prescription = data.get("prescription") or []
    columns = ["面号", "类型", "注释", "半径 (mm)", "厚度 (mm)", "材料", "半口径 (mm)", "光阑"]
    table = doc.add_table(rows=1, cols=len(columns))
    table.style = "Light Grid Accent 1"
    for i, name in enumerate(columns):
        set_run_font(table.rows[0].cells[i].paragraphs[0].add_run(name), bold=True)
    for s in prescription:
        cells = table.add_row().cells
        radius = num(s.get("radius"))
        thick = num(s.get("thickness"))
        semi = num(s.get("semiDiameter"))
        values = [
            s.get("surface"), s.get("type", ""), s.get("comment", ""),
            INF if radius is None else radius,
            INF if thick is None else thick,
            s.get("material", ""),
            INF if semi is None else semi,
            "是" if s.get("isStop") else "",
        ]
        for i, value in enumerate(values):
            set_run_font(cells[i].paragraphs[0].add_run(NA if value is None else str(value)))

    doc.add_heading("3. FFT MTF 关键取值", level=1)
    note = doc.add_paragraph()
    set_run_font(note.add_run(
        f"最高空间频率 {mtf.get('maxFrequency')} lp/mm；分析状态 {analysis.get('status')}"
        + (f"（{analysis.get('note')}）" if analysis.get("note") else "")))
    freqs = mtf.get("atFreqs") or []
    series_labels = []
    for row in freqs:
        for per in row.get("perSeries") or []:
            name = series_name(per)
            if name not in series_labels:
                series_labels.append(name)
    if freqs and series_labels:
        table = doc.add_table(rows=1, cols=1 + len(series_labels))
        table.style = "Light Grid Accent 1"
        set_run_font(table.rows[0].cells[0].paragraphs[0].add_run("空间频率 (lp/mm)"), bold=True)
        for i, label in enumerate(series_labels):
            set_run_font(table.rows[0].cells[i + 1].paragraphs[0].add_run(str(label)), bold=True)
        for row in freqs:
            cells = table.add_row().cells
            set_run_font(cells[0].paragraphs[0].add_run(str(num(row.get("frequency")))))
            for i, label in enumerate(series_labels):
                value = None
                for per in row.get("perSeries") or []:
                    if series_name(per) == label:
                        value = num(per.get("value"))
                set_run_font(cells[i + 1].paragraphs[0].add_run(NA if value is None else str(value)))
    else:
        set_run_font(doc.add_paragraph().add_run("本次运行未取得 MTF 数值数据。"))

    spot = data.get("spot") or []
    diagram = data.get("spotDiagram") or {}
    if spot:
        doc.add_heading("4. 点列数据", level=1)
        for a in spot:
            name, unit = spot_quantity(a)
            title = doc.add_paragraph()
            set_run_font(title.add_run(
                f"{name}（{unit}）" if unit else name), bold=True)
            settings = a.get("settings") or {}
            detail = "，".join(f"{k}={v}" for k, v in settings.items() if k in
                               ("Data", "Type", "ReferTo", "RayDensity", "FieldDensity", "SampleSize"))
            para = doc.add_paragraph()
            set_run_font(para.add_run(
                f"分析 {a.get('id')}，状态 {a.get('status')}" + (f"，{detail}" if detail else "")
                + (f"；{a.get('note')}" if a.get("note") else "")))

            series = a.get("series") or []
            if series:
                table = doc.add_table(rows=1, cols=4)
                table.style = "Light Grid Accent 1"
                for i, head in enumerate(["曲线组（视场）", "曲线", f"最小值 ({unit})", f"最大值 ({unit})"]):
                    set_run_font(table.rows[0].cells[i].paragraphs[0].add_run(head), bold=True)
                for s in series:
                    pts = s.get("points") or []
                    if not pts:
                        continue
                    ys = [p[1] for p in pts]
                    cells = table.add_row().cells
                    for i, value in enumerate([s.get("group", ""), s.get("label", ""),
                                               num(min(ys)), num(max(ys))]):
                        set_run_font(cells[i].paragraphs[0].add_run(NA if value is None else str(value)))

    if diagram:
        metrics = diagram.get("metrics") or {}
        doc.add_heading("5. 点列能量半径", level=1)
        ee = [a for a in spot if a.get("id") == "GeometricEncircledEnergy"]
        if ee:
            para = doc.add_paragraph()
            set_run_font(para.add_run(
                "包围能量半径由几何包围能量曲线线性插值得到：r50/r80/r90 表示包含 50%/80%/90% 能量所需的半径（µm）。"))
            series = ee[0].get("series") or []
            rows = [s for s in series if s.get("energyRadii")]
            if rows:
                table = doc.add_table(rows=1, cols=4)
                table.style = "Light Grid Accent 1"
                for i, head in enumerate(["视场", "r50 (µm)", "r80 (µm)", "r90 (µm)"]):
                    set_run_font(table.rows[0].cells[i].paragraphs[0].add_run(head), bold=True)
                for s in rows:
                    radii = {r.get("fraction"): r.get("radiusUm") for r in s.get("energyRadii")}
                    cells = table.add_row().cells
                    for i, value in enumerate([s.get("group", ""), num(radii.get(0.5)),
                                               num(radii.get(0.8)), num(radii.get(0.9))]):
                        set_run_font(cells[i].paragraphs[0].add_run(NA if value is None else str(value)))
        note = doc.add_paragraph()
        set_run_font(note.add_run(
            f"说明：StandardSpot 的聚合量中，最大半径原始读数（{metrics.get('maxRadiusRaw')}）与上面的 r90 吻合、可判为 µm；"
            f"但平均半径原始读数（{metrics.get('meanRadiusRaw')}）无法用同一套单位自洽解释，故本报告不采用，"
            "仅保留在 data.json 与 Excel 的「点列摘要」中备查。"))

    doc.add_heading("附：方法与口径", level=1)
    for text in [
        "数据全部由 OpticStudio 2024 R1 的 ZOS-API 在独立（无界面）实例中计算并导出，未经过任何人工录入或近似换算。",
        "MTF 取 FFT MTF 分析结果，按数据系列逐点记录；关键取值取与请求频率最接近的采样点，实际采样频率同表列出。",
        "处方表中半径或厚度为 ∞ 表示该量为无穷（平面或虚拟面），并非缺失；“—”表示该分析无此项数据。",
        f"源文件指纹：SHA256 {src.get('sha256', NA)}，字节数 {src.get('bytes', NA)}。",
        "本报告只描述当前文件当前状态；对文件的任何修改都会使上述指纹失效。",
    ]:
        paragraph = doc.add_paragraph(text, style="List Bullet")
        for run in paragraph.runs:
            set_run_font(run)

    doc.save(path)
    return path


# ------------------------------------------------------------------ main

def main():
    parser = argparse.ArgumentParser(description="把 Zemax 桥接的 data.json 转成 Excel/Word 报告")
    parser.add_argument("--data", required=True, help="zemax.ps1 产出的 data.json")
    parser.add_argument("--outdir", required=True, help="输出目录（通常就是该 run 目录）")
    args = parser.parse_args()

    with open(args.data, "r", encoding="utf-8") as handle:
        data = json.load(handle)

    if not os.path.isdir(args.outdir):
        os.makedirs(args.outdir, exist_ok=True)

    label = lens_label(data)
    xlsx = build_xlsx(data, args.outdir, label)
    docx = build_docx(data, args.outdir, label)

    print(json.dumps({
        "xlsx": xlsx,
        "docx": docx,
        "label": label,
        "generated": datetime.now().isoformat(timespec="seconds"),
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
