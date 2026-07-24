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


def _ordered(seq):
    out = []
    for x in seq:
        if x not in out:
            out.append(x)
    return out


def summary_frames(plan_result: PlanResult) -> "dict[str, pd.DataFrame]":
    """跨 pool × 月份的總攬,仿 Excel summary 分頁:每個分類一張
    (列=pool、欄=月份)的表。回傳有序 dict {分類名: DataFrame}。"""
    pools = _ordered(o.pool for o in plan_result.outcomes)
    months = _ordered(o.month for o in plan_result.outcomes)
    by = {(o.pool, o.month): o for o in plan_result.outcomes}
    labels = [p.label for p in pools]

    def grid(fn):
        data = {m: [fn(by[(p, m)]) if (p, m) in by else None for p in pools]
                for m in months}
        return pd.DataFrame(data, index=labels, columns=months)

    frames: "dict[str, pd.DataFrame]" = {
        "狀態": grid(lambda o: "OK" if o.feasible else "缺口"),
        "需求 vcore": grid(lambda o: o.demand_vcore),
        "新啟用機台": grid(lambda o: sum(o.machines_opened.values())),
        "進機": grid(lambda o: sum(o.movein.values())),
        "退回": grid(lambda o: sum(o.returns.values())),
    }
    if plan_result.mode == "excel_compat":
        frames["in-stock(台)"] = grid(lambda o: round(o.stock_float, 2))
    else:
        frames["月末剩餘可售 vcore"] = grid(lambda o: round(o.free_vcore_total, 1))
        frames["缺口 vcore"] = grid(
            lambda o: round(o.shortfall_vcore, 1) if not o.feasible else 0.0)
    return frames


def prometheus_ag_placeholder() -> pd.DataFrame:
    """未來 Prometheus 整合的欄位預覽(目前為假資料佔位)。

    每 cluster × AG 的節點現況:已安裝節點數、cordon 節點數、以及該 AG 的節點總數
    (未來由 Prometheus / node label 取得),供 expand / delete 時分配各 AG 的處理量。
    """
    rows = [
        ("c1", "ag1", 8, 1, 20),
        ("c1", "ag2", 7, 0, 20),
        ("c1", "ag3", 6, 2, 18),
        ("c2", "ag1", 5, 0, 20),
        ("c2", "ag2", 5, 1, 20),
        ("c2", "ag3", 4, 0, 18),
    ]
    return pd.DataFrame(rows, columns=[
        "cluster", "AG", "已安裝節點數", "cordon 節點數", "AG 節點總數"])


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
