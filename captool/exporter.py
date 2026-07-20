"""匯出:summary 報表(工具計算結果)與 v2 範本(多機型格式)。"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from captool.models import PlanInput
from captool.planner import PlanResult

_ARIAL = Font(name="Arial")
_BOLD = Font(name="Arial", bold=True)
_RED = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
_YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

_SUMMARY_HEADERS = ["Fab", "BM Group", "月份", "需求 vcore", "VM 顆數", "進機",
                    "退回", "新啟用機台", "建議採購", "月末總機台", "月末空機",
                    "剩餘可售 vcore", "狀態", "缺口 vcore"]


def _fmt(d: dict[str, int]) -> str:
    return "; ".join(f"{k}×{v}" for k, v in sorted(d.items())) if d else ""


def _style_row(ws, row_idx: int, fill=None) -> None:
    for cell in ws[row_idx]:
        cell.font = _ARIAL
        if fill is not None:
            cell.fill = fill


def export_summary(plan_input: PlanInput, plan_result: PlanResult, path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "summary"
    ws.append(_SUMMARY_HEADERS)
    _style_row(ws, 1)
    for cell in ws[1]:
        cell.font = _BOLD
    for o in plan_result.outcomes:
        ws.append([
            o.pool.fab, o.pool.bm_group, o.month, o.demand_vcore,
            sum(c for _, c in o.vm_demand),
            _fmt(o.movein), _fmt(o.returns), _fmt(o.machines_opened),
            _fmt(o.suggested_purchases), _fmt(o.stock_total), _fmt(o.stock_empty),
            round(o.free_vcore_total, 2),
            "OK" if o.feasible else "缺口",
            round(o.shortfall_vcore, 2) if not o.feasible else 0,
        ])
        _style_row(ws, ws.max_row, fill=None if o.feasible else _RED)
    ws.append([])
    ws.append([f"本表由容量規劃工具產出(模式:{plan_result.mode}),數值為 solver 計算結果,非公式。"])
    _style_row(ws, ws.max_row)

    wp = wb.create_sheet("placements")
    wp.append(["Fab", "BM Group", "月份", "機型", "新啟用台數"])
    _style_row(wp, 1)
    for cell in wp[1]:
        cell.font = _BOLD
    for o in plan_result.outcomes:
        for sku_name, cnt in sorted(o.machines_opened.items()):
            wp.append([o.pool.fab, o.pool.bm_group, o.month, sku_name, cnt])
            _style_row(wp, wp.max_row)
    wb.save(path)


def generate_v2_template(plan_input: PlanInput, path) -> None:
    default_sku_name = next(iter(plan_input.skus))
    wb = Workbook()

    # README(第一個 sheet)
    ws = wb.active
    ws.title = "README"
    lines = [
        "容量規劃 v2 範本填表說明",
        "",
        "各廠區 tab:PM 維護需求區(C 欄起,每月新增 vcore);Return 區(K 欄起)含機型欄。",
        "HW_SKU:機型目錄(平台/硬體 team 維護)。",
        "HW_Current:各 pool 期初完全空置機台數(per 機型)。",
        "HW_MoveIn:進機計畫(硬體 team 維護),月份格式 YYYY-MM。",
        "VM_Spec:需要 VM 規格明細的 product 才填(size 為單顆 vcore 數)。",
        "注意:機型名稱必須存在於 HW_SKU;月份一律 YYYY-MM。",
    ]
    for line in lines:
        ws.append([line])
    for row in ws.iter_rows():
        for cell in row:
            cell.font = _ARIAL
    ws["A1"].font = _BOLD

    # 廠區 tabs
    fabs = sorted({p.fab for p in plan_input.pools})
    for fab in fabs:
        wf = wb.create_sheet(fab)
        wf["C3"] = "User Demand (vcore)"
        wf["M3"] = "Server Return"
        wf["A4"], wf["B4"] = "Product", "BM Group"
        wf["K4"], wf["L4"], wf["M4"] = "Product", "BM Group", "機型"
        for i, month in enumerate(plan_input.months):
            wf.cell(row=4, column=3 + i, value=month)       # 需求月份 C4 起
            wf.cell(row=4, column=14 + i, value=month)      # Return 月份 N4 起
        # 帶入既有需求
        fab_demands = [d for d in plan_input.demands if d.pool.fab == fab]
        rows: dict[tuple[str, str], int] = {}
        r = 5
        for d in fab_demands:
            key = (d.product, d.pool.bm_group)
            if key not in rows:
                rows[key] = r
                wf.cell(row=r, column=1, value=d.product)
                wf.cell(row=r, column=2, value=d.pool.bm_group)
                r += 1
            col = 3 + plan_input.months.index(d.month)
            wf.cell(row=rows[key], column=col, value=d.vcore)
        # 帶入既有 Return(機型填 default)
        fab_returns = [x for x in plan_input.returns if x.pool.fab == fab]
        rrows: dict[tuple[str, str, str], int] = {}
        r = 5
        for x in fab_returns:
            key = (x.product, x.pool.bm_group, x.sku_name)
            if key not in rrows:
                rrows[key] = r
                wf.cell(row=r, column=11, value=x.product)
                wf.cell(row=r, column=12, value=x.pool.bm_group)
                wf.cell(row=r, column=13, value=x.sku_name)
                r += 1
            col = 14 + plan_input.months.index(x.month)
            wf.cell(row=rrows[key], column=col, value=x.count)
        for row in wf.iter_rows():
            for cell in row:
                cell.font = _ARIAL
        wf["C3"].font = _BOLD
        wf["M3"].font = _BOLD

    # HW_SKU
    ws1 = wb.create_sheet("HW_SKU")
    ws1.append(["name", "vcore_per_node", "usable_ratio"])
    for sku in plan_input.skus.values():
        ws1.append([sku.name, sku.vcore_per_node, sku.usable_ratio])
    # HW_Current
    ws2 = wb.create_sheet("HW_Current")
    ws2.append(["fab", "bm_group", "sku", "count"])
    for c in plan_input.currents:
        ws2.append([c.pool.fab, c.pool.bm_group, c.sku_name, c.count])
    # HW_MoveIn
    ws3 = wb.create_sheet("HW_MoveIn")
    ws3.append(["fab", "bm_group", "sku", "month", "count"])
    for m in plan_input.moveins:
        ws3.append([m.pool.fab, m.pool.bm_group, m.sku_name, m.month, m.count])
    # VM_Spec(含示範列)
    ws4 = wb.create_sheet("VM_Spec")
    ws4.append(["fab", "bm_group", "product", "month", "vm_size_vcore", "count"])
    example_fab = fabs[0] if fabs else "A"
    example_month = plan_input.months[0] if plan_input.months else "2026-07"
    ws4.append([example_fab, "network1", "(示範列,請刪除後填入實際資料)",
                example_month, 32, 2])
    for cell in ws4[2]:
        cell.fill = _YELLOW
    for v in plan_input.vm_demands:
        ws4.append([v.pool.fab, v.pool.bm_group, v.product, v.month,
                    v.vm_size_vcore, v.count])
    for sheet in (ws1, ws2, ws3, ws4):
        for row in sheet.iter_rows():
            for cell in row:
                cell.font = _ARIAL
        for cell in sheet[1]:
            cell.font = _BOLD
    wb.save(path)
