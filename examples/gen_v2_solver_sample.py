"""產生 solver-integration 版 v2 sample:worker 需求 + Cluster_Menu(control-plane 菜單)
+ New_Build 清單,全部假資料。並驗證 worker 部分可被現行 import_v2 匯入(新表被略過)。

用法:.venv\\Scripts\\python.exe examples\\gen_v2_solver_sample.py

註:各角色 vcore 為佔位假值,待實際規格取代(見 spec §2)。
"""
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

BOLD = Font(bold=True)
YELLOW = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
MONTHS = ["2026-07", "2026-08", "2026-09"]

SKUS = [("std-64", 64, 0.8), ("big-128", 128, 0.8), ("gpu-80", 80, 0.8)]

# Worker 需求(粗粒度 vcore):(fab, product, group, {month: vcore})
WORKER = [
    ("A", "web-svc", "network1", {"2026-07": 200, "2026-08": 120}),
    ("B", "batch", "network1", {"2026-07": 300, "2026-08": 300, "2026-09": 300}),
    ("C", "api", "network2", {"2026-07": 150, "2026-09": 150}),
]
# HW_Current 帶 AG(最後一欄):(fab, bm_group, sku, count, ag)
# A/network1、C/network2 需 ≥5 個 AG(菜單 master×5 要能分散);在庫分佈在部分 AG。
CURRENT = [
    ("A", "network1", "std-64", 2, "ag1"), ("A", "network1", "std-64", 2, "ag2"),
    ("A", "network1", "std-64", 2, "ag3"),
    ("A", "network1", "big-128", 1, "ag1"), ("A", "network1", "big-128", 1, "ag2"),
    ("A", "network1", "big-128", 1, "ag3"),
    ("B", "network1", "std-64", 4, "ag1"), ("B", "network1", "std-64", 3, "ag2"),
    ("B", "network1", "std-64", 3, "ag3"),
    ("C", "network2", "std-64", 2, "ag1"), ("C", "network2", "std-64", 2, "ag2"),
    ("C", "network2", "std-64", 1, "ag3"),
]
MOVEIN = [("B", "network1", "big-128", "2026-08", 2)]
# HW_Caps:每 AG 槽位上限 (fab, network, ag, max_bm)
CAPS = (
    [("A", "network1", f"ag{i}", 20) for i in range(1, 6)]    # 5 個 AG
    + [("B", "network1", f"ag{i}", 20) for i in range(1, 4)]  # 3 個 AG
    + [("C", "network2", f"ag{i}", 20) for i in range(1, 6)]  # 5 個 AG
)

# 角色 vcore 佔位假值(待實際規格取代)
ROLE_VCORE = {"master": 16, "infra": 8, "l4lb": 8, "learner": 32,
              "F5": 32, "Bastion": 8, "HA": 16, "VT": 16}

# 共住底座(A–E 皆含):(role, count)
BASE = [("master", 5), ("infra", 5), ("l4lb", 3), ("learner", 5)]
# 各菜單額外的獨佔元件:(role, count)
EXTRA = {
    "A": [],
    "B": [("F5", 3)],
    "C": [("Bastion", 1)],
    "D": [("F5", 3), ("Bastion", 1)],
    "E": [("F5", 3), ("Bastion", 1), ("HA", 3), ("VT", 3)],
}

# New build 清單:(fab, network, cluster, month, menu)
NEW_BUILD = [
    ("A", "network1", "c2", "2026-08", "B"),
    ("C", "network2", "c5", "2026-09", "E"),
]


def build(path):
    wb = Workbook()
    wb.remove(wb.active)

    fabs = sorted({r[0] for r in WORKER} | {c[0] for c in CURRENT})
    for fab in fabs:
        wf = wb.create_sheet(fab)
        wf["A4"], wf["B4"] = "Product", "BM Group"
        wf["C4"] = "VM vcore"
        wf["D4"] = "爆炸半徑 (1:x), EX: 2"
        wf["E4"] = "Tenant (自由=留空/指定群組/獨佔)"
        wf.cell(row=4, column=20, value="Product")
        wf.cell(row=4, column=21, value="BM Group")
        wf.cell(row=4, column=22, value="機型")
        for i, m in enumerate(MONTHS):
            wf.cell(row=4, column=6 + i, value=m)
            wf.cell(row=4, column=23 + i, value=m)
        r = 5
        for (f, prod, grp, mvals) in WORKER:
            if f != fab:
                continue
            wf.cell(row=r, column=1, value=prod)
            wf.cell(row=r, column=2, value=grp)
            for i, m in enumerate(MONTHS):
                if m in mvals:
                    wf.cell(row=r, column=6 + i, value=mvals[m])
            r += 1
        for c in range(1, 6):
            wf.cell(row=4, column=c).font = BOLD
        for c in (20, 21, 22):
            wf.cell(row=4, column=c).font = BOLD

    hs = wb.create_sheet("HW_SKU")
    hs.append(["name", "vcore_per_node", "usable_ratio"])
    for s in SKUS:
        hs.append(list(s))
    hc = wb.create_sheet("HW_Current")
    hc.append(["fab", "bm_group", "sku", "count", "ag"])   # ag 最後一欄(向前相容)
    for row in CURRENT:
        hc.append(list(row))
    hm = wb.create_sheet("HW_MoveIn")
    hm.append(["fab", "bm_group", "sku", "month", "count"])
    for row in MOVEIN:
        hm.append(list(row))
    hcap = wb.create_sheet("HW_Caps")
    hcap.append(["fab", "network", "ag", "max_bm"])
    for row in CAPS:
        hcap.append(list(row))

    # Cluster_Menu:每菜單列出全部元件(底座 shared + 額外 exclusive)
    cm = wb.create_sheet("Cluster_Menu")
    cm.append(["menu", "role", "count", "co_residency", "vm_vcore"])
    for menu in ("A", "B", "C", "D", "E"):
        for role, cnt in BASE:
            cm.append([menu, role, cnt, "shared", ROLE_VCORE[role]])
        for role, cnt in EXTRA[menu]:
            cm.append([menu, role, cnt, "exclusive", ROLE_VCORE[role]])

    nb = wb.create_sheet("New_Build")
    nb.append(["fab", "network", "cluster", "month", "menu"])
    for row in NEW_BUILD:
        nb.append(list(row))

    for sheet in (hs, hc, hm, hcap, cm, nb):
        for cell in sheet[1]:
            cell.font = BOLD
    wb.save(path)


def verify(path):
    from captool.importer import import_v2
    pi = import_v2(path)
    print("worker 匯入 OK — 機型:", list(pi.skus),
          "pools:", [p.label for p in pi.pools], "月份:", pi.months)
    print("(Cluster_Menu / New_Build 目前被現行 importer 略過,待整合階段解析)")
    # 展示菜單展開(概念:依 New_Build 的 menu 展開角色需求)
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    menus: dict[str, list] = {}
    for row in wb["Cluster_Menu"].iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        menus.setdefault(row[0], []).append(row[1:])
    print("\n菜單展開預覽:")
    for (fab, net, cluster, month, menu) in NEW_BUILD:
        comps = menus.get(menu, [])
        summary = ", ".join(f"{role}×{cnt}[{cor}]" for role, cnt, cor, _ in comps)
        print(f"  {fab}/{net} {month} new build {cluster} 菜單{menu} → {summary}")


if __name__ == "__main__":
    out = ROOT / "examples" / "v2_sample_solver.xlsx"
    build(out)
    print("wrote", out, "\n")
    verify(out)
