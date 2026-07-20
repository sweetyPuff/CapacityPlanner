"""Excel 匯入:legacy(現行單機型格式)與 v2(多機型格式,Task 10)。"""
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from captool.models import (CurrentStock, DemandDelta, ImportIssue, MoveIn,
                            NodeReturn, PlanInput, Pool, Sku)
from captool.months import parse_month

DEFAULT_SKU = Sku(name="default-64", vcore_per_node=64, usable_ratio=0.8)

V2_SHEETS = {"HW_SKU", "HW_Current", "HW_MoveIn", "VM_Spec"}
NON_FAB_SHEETS = V2_SHEETS | {"summary", "README"}


class CapacityImportError(Exception):
    def __init__(self, issues: list[ImportIssue]):
        self.issues = issues
        super().__init__("; ".join(i.message for i in issues))


def detect_format(path) -> str:
    wb = load_workbook(path, read_only=True)
    try:
        return "v2" if "HW_SKU" in wb.sheetnames else "legacy"
    finally:
        wb.close()


def import_any(path) -> PlanInput:
    if detect_format(path) == "v2":
        return import_v2(path)  # Task 10
    return import_legacy(path)


def import_v2(path) -> PlanInput:
    raise NotImplementedError("v2 匯入於 Task 10 實作")


class _MonthParser:
    """包裝 parse_month:相同原始字串的警告只回報一次。"""

    def __init__(self, issues: list[ImportIssue]):
        self.issues = issues
        self._seen: set[str] = set()

    def parse(self, raw, sheet: str, cell: str):
        value, warning = parse_month(raw)
        if warning and str(raw) not in self._seen:
            self._seen.add(str(raw))
            self.issues.append(ImportIssue("warning", sheet, cell, warning))
        return value


def _numeric(value, issues, sheet, cell) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    issues.append(ImportIssue(
        "error", sheet, cell, f"儲存格 {cell} 應為數值,實際為 {value!r},已以 0 計"))
    return 0.0


def _month_columns(ws, header_row: int, start_col: int, parser, sheet: str):
    """自 start_col 向右讀月份標頭至空白,回傳 [(col_index, 'YYYY-MM')]。"""
    cols = []
    col = start_col
    while True:
        raw = ws.cell(row=header_row, column=col).value
        if raw is None:
            break
        coord = f"{get_column_letter(col)}{header_row}"
        month = parser.parse(raw, sheet, coord)
        if month:
            cols.append((col, month))
        col += 1
    return cols


def import_legacy(path, default_sku: Sku = DEFAULT_SKU) -> PlanInput:
    issues: list[ImportIssue] = []
    parser = _MonthParser(issues)
    wb_v = load_workbook(path, data_only=True)
    wb_f = load_workbook(path)

    if "summary" not in wb_v.sheetnames:
        raise CapacityImportError(
            [ImportIssue("error", "summary", "", "找不到 summary sheet")])

    demands: list[DemandDelta] = []
    returns: list[NodeReturn] = []
    months: set[str] = set()

    fab_sheets = [n for n in wb_v.sheetnames if n not in NON_FAB_SHEETS]
    for name in fab_sheets:
        ws = wb_v[name]
        # 需求區:月份 row 4 自 C(3) 向右;資料自 row 5,A=product、B=group
        demand_cols = _month_columns(ws, 4, 3, parser, name)
        row = 5
        while ws.cell(row=row, column=1).value is not None:
            product = str(ws.cell(row=row, column=1).value)
            group = str(ws.cell(row=row, column=2).value)
            pool = Pool(fab=name, bm_group=group)
            for col, month in demand_cols:
                coord = f"{get_column_letter(col)}{row}"
                vcore = _numeric(ws.cell(row=row, column=col).value, issues, name, coord)
                if vcore:
                    demands.append(DemandDelta(pool=pool, product=product,
                                               month=month, vcore=vcore))
                months.add(month)
            row += 1
        # Return 區:月份 row 4 自 M(13) 向右;資料自 row 5,K=product、L=group
        return_cols = _month_columns(ws, 4, 13, parser, name)
        row = 5
        while ws.cell(row=row, column=11).value is not None:
            product = str(ws.cell(row=row, column=11).value)
            group = str(ws.cell(row=row, column=12).value)
            pool = Pool(fab=name, bm_group=group)
            for col, month in return_cols:
                coord = f"{get_column_letter(col)}{row}"
                cnt = _numeric(ws.cell(row=row, column=col).value, issues, name, coord)
                if cnt:
                    returns.append(NodeReturn(pool=pool, product=product,
                                              sku_name=default_sku.name,
                                              month=month, count=int(cnt)))
                months.add(month)
            row += 1

    # summary:pools、Current、MoveIn
    ws = wb_v["summary"]
    ws_f = wb_f["summary"]
    current_col = None
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=2, column=col).value == "Current":
            current_col = col
            break
    if current_col is None:
        raise CapacityImportError(
            [ImportIssue("error", "summary", "", "summary 找不到 Current 欄")])
    movein_cols = []
    for rng in ws_f.merged_cells.ranges:
        anchor = ws_f.cell(row=rng.min_row, column=rng.min_col).value
        if rng.min_row == 1 and isinstance(anchor, str) and anchor.startswith("Server MoveIn"):
            movein_cols = [(c, parser.parse(ws.cell(row=2, column=c).value,
                                            "summary", f"{get_column_letter(c)}2"))
                           for c in range(rng.min_col, rng.max_col + 1)]
            movein_cols = [(c, m) for c, m in movein_cols if m]
            break
    if not movein_cols:
        issues.append(ImportIssue("warning", "summary", "",
                                  "summary 找不到 Server MoveIn 區塊,進機計畫視為 0"))

    pools: list[Pool] = []
    currents: list[CurrentStock] = []
    moveins: list[MoveIn] = []
    row = 3
    while ws.cell(row=row, column=1).value is not None:
        pool = Pool(fab=str(ws.cell(row=row, column=1).value),
                    bm_group=str(ws.cell(row=row, column=2).value))
        pools.append(pool)
        cur_coord = f"{get_column_letter(current_col)}{row}"
        cur = _numeric(ws.cell(row=row, column=current_col).value, issues,
                       "summary", cur_coord)
        currents.append(CurrentStock(pool=pool, sku_name=default_sku.name,
                                     count=int(cur)))
        _warn_if_formula(ws_f, row, current_col, issues)
        for col, month in movein_cols:
            coord = f"{get_column_letter(col)}{row}"
            cnt = _numeric(ws.cell(row=row, column=col).value, issues, "summary", coord)
            if cnt:
                moveins.append(MoveIn(pool=pool, sku_name=default_sku.name,
                                      month=month, count=int(cnt)))
            months.add(month)
            _warn_if_formula(ws_f, row, col, issues)
        row += 1

    return PlanInput(
        skus={default_sku.name: default_sku},
        months=sorted(months),
        pools=pools,
        demands=demands, vm_demands=[],
        moveins=moveins, returns=returns, currents=currents,
        issues=issues)


def _warn_if_formula(ws_f, row: int, col: int, issues: list[ImportIssue]) -> None:
    value = ws_f.cell(row=row, column=col).value
    if isinstance(value, str) and value.startswith("="):
        coord = f"{get_column_letter(col)}{row}"
        issues.append(ImportIssue(
            "warning", ws_f.title, coord,
            f"{coord} 為手填區卻含公式 {value},請確認是否錯位(已採計算後數值)"))
