"""產生統一版面 v2 sample:worker 與 new build 都寫在 A/B/C fab 頁(用 Menu 欄區分),
塞各種情境的假資料;HW 表帶 AG。並驗證 import_v2 可解析(worker + new_builds + menus)。

fab 頁需求區欄位:Product | BM Group | VM vcore | 爆炸半徑(1:x) | Tenant | Menu | 月份(G 起)
  Menu = Worker → 一般 worker 需求列(月份=vcore)
  Menu = 菜單名(A~E)→ new build 列,Product=cluster,月份=當月建幾個 cluster
Return 區:Product(T)| BM Group(U)| 機型(V)| 月份(W 起)

註:角色 vcore 為佔位假值(見 spec §2)。
"""
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

BOLD = Font(bold=True)
YELLOW = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
BLUE = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
MONTHS = ["2026-07", "2026-08", "2026-09"]

SKUS = [("std-64", 64, 0.8), ("big-128", 128, 0.8), ("gpu-80", 80, 0.8)]

# 需求列:(fab, product, group, vm_vcore|None, blast|None, tenant, menu, {month: 值})
#   menu="Worker" → worker(值=vcore);menu=菜單名 → new build(值=cluster 數)
DEMAND_ROWS = [
    # --- worker 各種情境 ---
    ("A", "web", "network1", None, None, "", "Worker", {"2026-07": 200, "2026-08": 120}),   # 粗粒度
    ("A", "cache", "network1", 60, 2, "", "Worker", {"2026-07": 180, "2026-08": 120}),      # detail + 爆炸半徑2
    ("A", "iso", "network1", 51, 1, "獨佔", "Worker", {"2026-07": 51}),                     # 獨佔 tenant
    ("A", "db", "network2", 48, None, "teamA", "Worker", {"2026-07": 144, "2026-08": 96}),  # 群組 tenant
    ("B", "batch", "network1", None, None, "", "Worker", {"2026-07": 300, "2026-08": 300, "2026-09": 300}),
    ("C", "api", "network2", 32, 3, "", "Worker", {"2026-07": 150, "2026-09": 150}),        # detail + 爆炸半徑3
    # --- new build(Menu=菜單名,值=當月建幾個 cluster)---
    ("A", "c-happy", "network1", None, None, "", "B", {"2026-08": 1}),   # 8月建1個 menu B
    ("A", "c-lite", "network2", None, None, "", "A", {"2026-07": 1}),    # 7月建1個 menu A
    ("B", "c-mega", "network1", None, None, "", "E", {"2026-09": 1}),    # 9月建1個 menu E(重)
    ("C", "c-twin", "network2", None, None, "", "D", {"2026-07": 2}),    # 7月建2個 menu D
]
# 退回列:(fab, product, group, sku, {month: count})
RETURN_ROWS = [
    ("A", "old-a", "network1", "std-64", {"2026-07": 3}),
    ("B", "old-b", "network1", "std-64", {"2026-09": 2}),
]
# HW_Current 帶 AG:(fab, bm_group, sku, count, ag);A/C 佈 5 個 AG(菜單 master×5 要分得開)
CURRENT = [
    ("A", "network1", "std-64", 2, "ag1"), ("A", "network1", "std-64", 2, "ag2"),
    ("A", "network1", "std-64", 2, "ag3"),
    ("A", "network1", "big-128", 1, "ag1"), ("A", "network1", "big-128", 1, "ag2"),
    ("A", "network2", "std-64", 3, "ag1"), ("A", "network2", "std-64", 2, "ag2"),
    ("B", "network1", "std-64", 4, "ag1"), ("B", "network1", "std-64", 3, "ag2"),
    ("B", "network1", "std-64", 3, "ag3"),
    ("C", "network2", "std-64", 2, "ag1"), ("C", "network2", "std-64", 2, "ag2"),
    ("C", "network2", "std-64", 1, "ag3"),
]
# HW_MoveIn 帶選配 ag:(fab, bm_group, sku, month, count, ag)  ag 留空 = solver 自選
MOVEIN = [("B", "network1", "big-128", "2026-08", 2, "ag1")]
# HW_Caps:(fab, network, ag, max_bm)
CAPS = (
    [("A", "network1", f"ag{i}", 20) for i in range(1, 6)]
    + [("A", "network2", f"ag{i}", 20) for i in range(1, 6)]
    + [("B", "network1", f"ag{i}", 20) for i in range(1, 4)]
    + [("C", "network2", f"ag{i}", 20) for i in range(1, 6)]
)

ROLE_VCORE = {"master": 16, "infra": 8, "l4lb": 8, "learner": 32,
              "F5": 32, "Bastion": 8, "HA": 16, "VT": 16}
BASE = [("master", 5), ("infra", 5), ("l4lb", 3), ("learner", 5)]
EXTRA = {"A": [], "B": [("F5", 3)], "C": [("Bastion", 1)],
         "D": [("F5", 3), ("Bastion", 1)],
         "E": [("F5", 3), ("Bastion", 1), ("HA", 3), ("VT", 3)]}


def build(path):
    wb = Workbook()
    wb.remove(wb.active)
    fabs = sorted({r[0] for r in DEMAND_ROWS} | {c[0] for c in CURRENT})
    for fab in fabs:
        wf = wb.create_sheet(fab)
        wf["A4"], wf["B4"], wf["C4"] = "Product", "BM Group", "VM vcore"
        wf["D4"] = "爆炸半徑 (1:x), EX: 2"
        wf["E4"] = "Tenant (自由=留空/指定群組/獨佔)"
        wf["F4"] = "Menu"
        wf.cell(row=4, column=20, value="Product")
        wf.cell(row=4, column=21, value="BM Group")
        wf.cell(row=4, column=22, value="機型")
        for i, m in enumerate(MONTHS):
            wf.cell(row=4, column=7 + i, value=m)     # 需求月份 G 起
            wf.cell(row=4, column=23 + i, value=m)    # 退回月份 W 起
        r = 5
        for (f, prod, grp, vc, blast, tenant, menu, mvals) in DEMAND_ROWS:
            if f != fab:
                continue
            wf.cell(row=r, column=1, value=prod)
            wf.cell(row=r, column=2, value=grp)
            if vc is not None:
                wf.cell(row=r, column=3, value=vc)
            if blast is not None:
                wf.cell(row=r, column=4, value=blast)
            if tenant:
                wf.cell(row=r, column=5, value=tenant)
            wf.cell(row=r, column=6, value=menu)
            for i, m in enumerate(MONTHS):
                if m in mvals:
                    wf.cell(row=r, column=7 + i, value=mvals[m])
            if menu != "Worker":
                for c in range(1, 7):
                    wf.cell(row=r, column=c).fill = BLUE   # new build 列標藍底
            r += 1
        r = 5
        for (f, prod, grp, sku, mvals) in RETURN_ROWS:
            if f != fab:
                continue
            wf.cell(row=r, column=20, value=prod)
            wf.cell(row=r, column=21, value=grp)
            wf.cell(row=r, column=22, value=sku)
            for i, m in enumerate(MONTHS):
                if m in mvals:
                    wf.cell(row=r, column=23 + i, value=mvals[m])
            r += 1
        for c in list(range(1, 7)) + [20, 21, 22]:
            wf.cell(row=4, column=c).font = BOLD

    hs = wb.create_sheet("HW_SKU")
    hs.append(["name", "vcore_per_node", "usable_ratio"])
    for s in SKUS:
        hs.append(list(s))
    hc = wb.create_sheet("HW_Current")
    hc.append(["fab", "bm_group", "sku", "count", "ag"])
    for row in CURRENT:
        hc.append(list(row))
    hm = wb.create_sheet("HW_MoveIn")
    hm.append(["fab", "bm_group", "sku", "month", "count", "ag"])
    for row in MOVEIN:
        hm.append(list(row))
    hcap = wb.create_sheet("HW_Caps")
    hcap.append(["fab", "network", "ag", "max_bm"])
    for row in CAPS:
        hcap.append(list(row))
    cm = wb.create_sheet("Cluster_Menu")
    cm.append(["menu", "role", "count", "co_residency", "vm_vcore"])
    for menu in ("A", "B", "C", "D", "E"):
        for role, cnt in BASE:
            cm.append([menu, role, cnt, "shared", ROLE_VCORE[role]])
        for role, cnt in EXTRA[menu]:
            cm.append([menu, role, cnt, "exclusive", ROLE_VCORE[role]])

    for sheet in (hs, hc, hm, hcap, cm):
        for cell in sheet[1]:
            cell.font = BOLD
    wb.save(path)


def verify(path):
    from captool.importer import import_v2
    pi = import_v2(path)
    print("機型:", list(pi.skus), "pools:", [p.label for p in pi.pools],
          "月份:", pi.months)
    print("worker 需求列:", len(pi.demands), "detail:", len(pi.vm_demands),
          "policies:", len(pi.policies), "currents(AG):",
          sorted({c.ag for c in pi.currents}), "caps:", len(pi.caps))
    issues = [i for i in pi.issues if i.severity == "error"]
    print("錯誤:", [i.message for i in issues] or "無")
    print("\nnew build 列(擷取):")
    for nb in pi.new_builds:
        print(f"  {nb.pool.label} {nb.month} 用菜單 {nb.menu} 建 {nb.count} 個 "
              f"cluster '{nb.cluster}'")
    print("\n菜單目錄(parsed):", {k: len(v) for k, v in pi.menus.items()})


if __name__ == "__main__":
    out = ROOT / "examples" / "v2_sample_solver.xlsx"
    build(out)
    print("wrote", out, "\n")
    verify(out)
