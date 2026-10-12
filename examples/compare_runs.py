# -*- coding: utf-8 -*-
"""比较两个 run 目录的 data.json 是否逐值一致（复现性自查）。

用法：
    <python> compare_runs.py <runA> <runB>

比较三类东西：处方（面号/半径/厚度/材料）、MTF 关键频率取值、点列曲线极值与能量半径。
"""
import json
import os
import sys


def load(run_dir):
    with open(os.path.join(run_dir, "data.json"), encoding="utf-8") as handle:
        data = json.load(handle)

    mtf = {}
    for row in (data.get("mtf") or {}).get("atFreqs") or []:
        for per in row.get("perSeries") or []:
            mtf[(row["frequency"], per["group"], per["label"])] = per["value"]

    spot = {}
    for a in data.get("spot") or []:
        settings = a.get("settings") or {}
        kind = settings.get("Data") or settings.get("Type") or ""
        for s in a.get("series") or []:
            vals = [p[1] for p in s["points"]]
            key = (a["id"], kind, s["group"], s["label"])
            spot[key] = (min(vals), max(vals))
            for er in s.get("energyRadii") or []:
                spot[key + (er["fraction"],)] = er["radiusUm"]

    prescription = [(s["surface"], s["radius"], s["thickness"], s["material"])
                    for s in data.get("prescription") or []]
    return mtf, spot, prescription


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return 2

    a, b = sys.argv[1], sys.argv[2]
    mtf_a, spot_a, pres_a = load(a)
    mtf_b, spot_b, pres_b = load(b)

    ok = True
    for name, x, y in (("处方", pres_a, pres_b), ("MTF 取值", mtf_a, mtf_b), ("点列/能量半径", spot_a, spot_b)):
        same = x == y
        ok = ok and same
        print(f"{name}: {'一致' if same else '不一致'}（{len(x)} 项）")
        if not same:
            keys = set(x) | set(y)
            for key in sorted(keys, key=str):
                if x.get(key) != y.get(key):
                    print(f"   差异 {key}: A={x.get(key)!r}  B={y.get(key)!r}")
    print()
    print("PASS: 两次运行逐值一致" if ok else "FAIL: 存在差异")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
