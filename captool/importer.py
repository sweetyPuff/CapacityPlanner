"""Excel 匯入:legacy(現行單機型格式)與 v2(多機型格式,Task 10)。"""
import math

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from captool.models import (CurrentStock, DemandDelta, ImportIssue, MoveIn,
                            NodeReturn, PlanInput, Pool, ProductPolicy, Sku,
                            VmSpecDemand)
from captool.months import parse_month

DEFAULT_SKU = Sku(name="default-64", vcore_per_node=64, usable_ratio=0.8)

V2_SHEETS = {"HW_SKU", "HW_Current", "HW_MoveIn"}
V2_OPTIONAL_SHEETS = {"VM_Spec"}
NON_FAB_SHEETS = V2_SHEETS | V2_OPTIONAL_SHEETS | {"summary", "README"}


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


def _read_table(ws, headers: list[str], issues, sheet: str):
    """讀長表:row 1 為標頭,row 2 起至全空列。回傳 list[dict]。"""
    actual = [c.value for c in ws[1][:len(headers)]]
    if actual != headers:
        issues.append(ImportIssue(
            "error", sheet, "A1", f"{sheet} 標頭應為 {headers},實際為 {actual}"))
        return []
    rows = []
    for r in range(2, ws.max_row + 1):
        values = [ws.cell(row=r, column=c).value for c in range(1, len(headers) + 1)]
        if all(v is None for v in values):
            continue
        rows.append({"_row": r, **dict(zip(headers, values))})
    return rows


def import_v2(path) -> PlanInput:
    issues: list[ImportIssue] = []
    parser = _MonthParser(issues)
    wb_v = load_workbook(path, data_only=True)

    for required in V2_SHEETS:
        if required not in wb_v.sheetnames:
            raise CapacityImportError(
                [ImportIssue("error", required, "", f"找不到 {required} sheet")])

    # HW_SKU
    skus: dict[str, Sku] = {}
    for rec in _read_table(wb_v["HW_SKU"], ["name", "vcore_per_node", "usable_ratio"],
                           issues, "HW_SKU"):
        name = str(rec["name"])
        if name in skus:
            issues.append(ImportIssue("error", "HW_SKU", f"A{rec['_row']}",
                                      f"機型 {name} 重複定義,僅取第一筆"))
            continue
        try:
            skus[name] = Sku(name=name, vcore_per_node=int(rec["vcore_per_node"]),
                             usable_ratio=float(rec["usable_ratio"]))
        except (TypeError, ValueError):
            issues.append(ImportIssue("error", "HW_SKU", f"A{rec['_row']}",
                                      f"機型 {name} 的參數非數值,已跳過"))
    if not skus:
        raise CapacityImportError(
            [ImportIssue("error", "HW_SKU", "", "HW_SKU 無有效機型")])
    first_sku_name = next(iter(skus))

    def valid_sku(name, sheet, row) -> str | None:
        if name is None or str(name) not in skus:
            issues.append(ImportIssue(
                "error", sheet, f"C{row}",
                f"機型 '{name}' 不存在於 HW_SKU,該列已跳過"))
            return None
        return str(name)

    months: set[str] = set()

    # HW_Current / HW_MoveIn
    currents: list[CurrentStock] = []
    for rec in _read_table(wb_v["HW_Current"], ["fab", "bm_group", "sku", "count"],
                           issues, "HW_Current"):
        sku_name = valid_sku(rec["sku"], "HW_Current", rec["_row"])
        if sku_name is None:
            continue
        pool = Pool(fab=str(rec["fab"]), bm_group=str(rec["bm_group"]))
        count = int(_numeric(rec["count"], issues, "HW_Current", f"D{rec['_row']}"))
        currents.append(CurrentStock(pool=pool, sku_name=sku_name, count=count))
    moveins: list[MoveIn] = []
    for rec in _read_table(wb_v["HW_MoveIn"],
                           ["fab", "bm_group", "sku", "month", "count"],
                           issues, "HW_MoveIn"):
        sku_name = valid_sku(rec["sku"], "HW_MoveIn", rec["_row"])
        if sku_name is None:
            continue
        month = parser.parse(rec["month"], "HW_MoveIn", f"D{rec['_row']}")
        if month is None:
            continue
        months.add(month)
        pool = Pool(fab=str(rec["fab"]), bm_group=str(rec["bm_group"]))
        count = int(_numeric(rec["count"], issues, "HW_MoveIn", f"E{rec['_row']}"))
        moveins.append(MoveIn(pool=pool, sku_name=sku_name, month=month, count=count))

    # VM_Spec (optional in new v2 format)
    vm_demands: list[VmSpecDemand] = []
    if "VM_Spec" in wb_v.sheetnames:
        for rec in _read_table(wb_v["VM_Spec"],
                               ["fab", "bm_group", "product", "month",
                                "vm_size_vcore", "count"], issues, "VM_Spec"):
            product = str(rec["product"])
            if "示範" in product:
                continue
            month = parser.parse(rec["month"], "VM_Spec", f"D{rec['_row']}")
            if month is None:
                continue
            months.add(month)
            pool = Pool(fab=str(rec["fab"]), bm_group=str(rec["bm_group"]))
            vm_size_vcore = int(_numeric(rec["vm_size_vcore"], issues, "VM_Spec", f"E{rec['_row']}"))
            count = int(_numeric(rec["count"], issues, "VM_Spec", f"F{rec['_row']}"))
            vm_demands.append(VmSpecDemand(
                pool=pool, product=product, month=month,
                vm_size_vcore=vm_size_vcore, count=count))

    # 廠區 tabs:需求區新 v2 格式 (A/B/C/D/E + F+) 或舊格式 (A/B + C+)
    demands: list[DemandDelta] = []
    returns: list[NodeReturn] = []
    policies: list[ProductPolicy] = []
    fab_sheets = [n for n in wb_v.sheetnames if n not in NON_FAB_SHEETS]
    for name in fab_sheets:
        ws = wb_v[name]
        # 判斷格式:檢查 C4 是否為 "VM vcore"(新格式) 或月份(舊格式)
        c4_value = ws.cell(row=4, column=3).value
        is_new_format = c4_value == "VM vcore"

        if is_new_format:
            # 新 v2 格式:需求月份自 F(6),Return 月份自 W(23)
            demand_cols = _month_columns(ws, 4, 6, parser, name)
            row = 5
            while ws.cell(row=row, column=1).value is not None:
                product = str(ws.cell(row=row, column=1).value)
                if "示範" in product:
                    row += 1
                    continue
                group = str(ws.cell(row=row, column=2).value)
                pool = Pool(fab=name, bm_group=group)
                vm_size = ws.cell(row=row, column=3).value
                max_per = ws.cell(row=row, column=4).value
                coresid_raw = ws.cell(row=row, column=5).value
                if vm_size in (None, ""):
                    # 粗粒度需求
                    for col, month in demand_cols:
                        coord = f"{get_column_letter(col)}{row}"
                        vcore = _numeric(ws.cell(row=row, column=col).value, issues, name, coord)
                        if vcore:
                            demands.append(DemandDelta(pool=pool, product=product,
                                                       month=month, vcore=vcore))
                            months.add(month)
                else:
                    # detail:每顆 VM 的 vcore 尺寸,顆數由 vcore ÷ size 反推
                    # Guard C (VM vcore):必須是正整數
                    try:
                        size = int(vm_size)
                        if size <= 0:
                            raise ValueError("非正整數")
                    except (TypeError, ValueError):
                        issues.append(ImportIssue(
                            "error", name, f"C{row}",
                            f"C{row} 產品 {product} 的 VM vcore '{vm_size}' 非正整數,該列已跳過"))
                        row += 1
                        continue

                    for col, month in demand_cols:
                        coord = f"{get_column_letter(col)}{row}"
                        vcore = _numeric(ws.cell(row=row, column=col).value, issues, name, coord)
                        if not vcore:
                            continue
                        months.add(month)
                        count = math.ceil(vcore / size)
                        if abs(vcore - count * size) > 1e-9:
                            issues.append(ImportIssue(
                                "warning", name, coord,
                                f"{coord} 產品 {product} 的 vcore {vcore} 非 VM 規格 {size} "
                                f"的整數倍,已進位為 {count} 台({count * size} vcore)"))
                        vm_demands.append(VmSpecDemand(pool=pool, product=product,
                                                       month=month, vm_size_vcore=size,
                                                       count=count))

                    # Guard D (每台上限):非正整數時以 None 計
                    max_per_machine = None
                    if max_per not in (None, ""):
                        try:
                            max_per_val = int(max_per)
                            if max_per_val <= 0:
                                raise ValueError("非正整數")
                            max_per_machine = max_per_val
                        except (TypeError, ValueError):
                            issues.append(ImportIssue(
                                "warning", name, f"D{row}",
                                f"D{row} 每台上限 '{max_per}' 非正整數,以不限計"))

                    policies.append(ProductPolicy(
                        pool=pool, product=product, vm_size_vcore=size,
                        max_per_machine=max_per_machine,
                        co_residency=_parse_coresidency(coresid_raw)))
                row += 1
            return_cols = _month_columns(ws, 4, 23, parser, name)
            row = 5
            while ws.cell(row=row, column=20).value is not None:
                product = str(ws.cell(row=row, column=20).value)
                group = str(ws.cell(row=row, column=21).value)
                raw_sku = ws.cell(row=row, column=22).value
                if raw_sku is None:
                    issues.append(ImportIssue(
                        "warning", name, f"V{row}",
                        f"V{row} Return 機型空白,以 {first_sku_name} 計"))
                    sku_name = first_sku_name
                else:
                    sku_name = valid_sku(raw_sku, name, row)
                    if sku_name is None:
                        row += 1
                        continue
                pool = Pool(fab=name, bm_group=group)
                for col, month in return_cols:
                    coord = f"{get_column_letter(col)}{row}"
                    cnt = _numeric(ws.cell(row=row, column=col).value, issues, name, coord)
                    if cnt:
                        returns.append(NodeReturn(pool=pool, product=product,
                                                  sku_name=sku_name, month=month,
                                                  count=int(cnt)))
                        months.add(month)
                row += 1
        else:
            # 舊 v2 格式(或 legacy):需求月份自 C(3),Return 月份自 N(14)
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
            return_cols = _month_columns(ws, 4, 14, parser, name)
            row = 5
            while ws.cell(row=row, column=11).value is not None:
                product = str(ws.cell(row=row, column=11).value)
                group = str(ws.cell(row=row, column=12).value)
                raw_sku = ws.cell(row=row, column=13).value
                if raw_sku is None:
                    issues.append(ImportIssue(
                        "warning", name, f"M{row}",
                        f"M{row} Return 機型空白,以 {first_sku_name} 計"))
                    sku_name = first_sku_name
                else:
                    sku_name = valid_sku(raw_sku, name, row)
                    if sku_name is None:
                        row += 1
                        continue
                pool = Pool(fab=name, bm_group=group)
                for col, month in return_cols:
                    coord = f"{get_column_letter(col)}{row}"
                    cnt = _numeric(ws.cell(row=row, column=col).value, issues, name, coord)
                    if cnt:
                        returns.append(NodeReturn(pool=pool, product=product,
                                                  sku_name=sku_name, month=month,
                                                  count=int(cnt)))
                    months.add(month)
                row += 1

    pools = sorted({c.pool for c in currents}
                   | {d.pool for d in demands}
                   | {m.pool for m in moveins}
                   | {v.pool for v in vm_demands}
                   | {r.pool for r in returns}
                   | {p.pool for p in policies},
                   key=lambda p: (p.fab, p.bm_group))
    return PlanInput(skus=skus, months=sorted(months), pools=pools,
                     demands=demands, vm_demands=vm_demands,
                     moveins=moveins, returns=returns, currents=currents,
                     issues=issues, policies=policies)


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


def _parse_coresidency(raw) -> str:
    if raw in (None, ""):
        return "free"
    s = str(raw).strip()
    if s in ("獨佔", "exclusive", "獨占"):
        return "exclusive"
    return s


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
