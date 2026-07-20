"""PlanResult / PlanInput → pandas DataFrame,供 Streamlit UI 使用。純函式。"""
import dataclasses

import pandas as pd

from captool.models import MoveIn, PlanInput, Pool
from captool.months import parse_month
from captool.planner import PlanResult


def _fmt(d: dict[str, int]) -> str:
    return "; ".join(f"{k}×{v}" for k, v in sorted(d.items())) if d else ""


def overview_frame(plan_result: PlanResult, pool: Pool) -> pd.DataFrame:
    rows = {}
    for o in plan_result.for_pool(pool):
        row = {
            "需求 vcore": o.demand_vcore,
            "VM 顆數": sum(c for _, c in o.vm_demand),
            "進機": _fmt(o.movein),
            "退回": _fmt(o.returns),
            "新啟用機台": _fmt(o.machines_opened),
            "月末空機": _fmt(o.stock_empty),
            "剩餘可售 vcore": round(o.free_vcore_total, 1),
            "狀態": "OK" if o.feasible else "缺口",
            "缺口 vcore": round(o.shortfall_vcore, 1) if not o.feasible else 0.0,
        }
        if plan_result.mode == "excel_compat":
            row["in-stock(台)"] = round(o.stock_float, 2)
        rows[o.month] = row
    return pd.DataFrame.from_dict(rows, orient="index")


def placements_frame(plan_result: PlanResult, pool: Pool) -> pd.DataFrame:
    records = []
    for o in plan_result.for_pool(pool):
        sku_names = sorted(set(o.machines_opened) | set(o.suggested_purchases))
        for sku_name in sku_names:
            records.append({
                "月份": o.month,
                "機型": sku_name,
                "新啟用台數": o.machines_opened.get(sku_name, 0),
                "建議採購台數": o.suggested_purchases.get(sku_name, 0),
            })
    return pd.DataFrame(records,
                        columns=["月份", "機型", "新啟用台數", "建議採購台數"])


def movein_frame(plan_input: PlanInput) -> pd.DataFrame:
    return pd.DataFrame(
        [{"fab": m.pool.fab, "bm_group": m.pool.bm_group, "sku": m.sku_name,
          "month": m.month, "count": m.count} for m in plan_input.moveins],
        columns=["fab", "bm_group", "sku", "month", "count"])


def apply_movein_edits(plan_input: PlanInput, df: pd.DataFrame) -> PlanInput:
    moveins: list[MoveIn] = []
    for i, rec in df.iterrows():
        if rec.isna().all():
            continue
        sku_name = str(rec["sku"])
        if sku_name not in plan_input.skus:
            raise ValueError(f"第 {i + 1} 列:機型 '{sku_name}' 不存在")
        month, _ = parse_month(rec["month"])
        if month is None:
            raise ValueError(f"第 {i + 1} 列:月份 '{rec['month']}' 無法解析")
        moveins.append(MoveIn(
            pool=Pool(fab=str(rec["fab"]), bm_group=str(rec["bm_group"])),
            sku_name=sku_name, month=month, count=int(rec["count"] or 0)))
    updated = dataclasses.replace(plan_input, moveins=moveins)
    months = set(updated.months) | {m.month for m in moveins}
    updated.months = sorted(months)
    return updated
