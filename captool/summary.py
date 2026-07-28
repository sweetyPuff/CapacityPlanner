"""統整 summary(仿 活頁簿1.xlsx summary 分頁,升級成多機型 × 多 AG)。

五個區塊,列的維度視資料而定(cols = 月份):
  1. 需求 vcore          —— Fab/Network(需求無 SKU/AG 維度)
  2. 實體機需求 (台)      —— Fab/Network/AG,solver 求解後 = bm_bought + in_stock_bm_used
  3. 進機 (台)           —— Fab/Network/SKU/AG(HW_MoveIn)
  4. 退還 (台)           —— Fab/Network/SKU(退還無 AG 維度)
  5. 剩餘可用 vcore (月末) —— Fab/Network/AG,solver 回的 in_stock_available

2、5 需要 solver 結果(horizon_result);沒有時該兩塊不產生。
剩餘「庫存」以 vcore 表達而非台數:多機型 + 共居下一台機住多顆 VM 仍是一台,
台數不再良好定義,剩餘可用 vcore 才是有意義的量。
"""
from collections import defaultdict

import pandas as pd


def _df(data: dict, months: list) -> pd.DataFrame:
    idx = sorted(data)
    return pd.DataFrame(
        [[round(data[r].get(m, 0), 1) for m in months] for r in idx],
        index=idx, columns=months)


def capacity_summary(plan_input, horizon_result=None) -> "dict[str, pd.DataFrame]":
    months = list(plan_input.months)
    blocks: "dict[str, pd.DataFrame]" = {}

    # 1. 需求 vcore(coarse + detail)—— Fab/Network
    demand = defaultdict(lambda: defaultdict(float))
    for d in plan_input.demands:
        demand[f"{d.pool.fab}/{d.pool.bm_group}"][d.month] += d.vcore
    for v in plan_input.vm_demands:
        demand[f"{v.pool.fab}/{v.pool.bm_group}"][v.month] += v.vm_size_vcore * v.count
    blocks["1. 需求 vcore(Fab/Network)"] = _df(demand, months)

    # 2. 實體機需求(台)—— Fab/Network/AG(solver)
    if horizon_result is not None:
        req = defaultdict(lambda: defaultdict(int))
        for o in horizon_result.outcomes:
            for c in o.cells:
                req[f"{o.fab}/{c['network']}/{c['ag']}"][o.period] += (
                    c["bm_bought"] + c["in_stock_bm_used"])
        blocks["2. 實體機需求 (台)(Fab/Network/AG,solver)"] = _df(req, months)

    # 3. 進機(台)—— Fab/Network/SKU/AG
    movein = defaultdict(lambda: defaultdict(int))
    for m in plan_input.moveins:
        movein[f"{m.pool.fab}/{m.pool.bm_group}/{m.sku_name}/{m.ag or '-'}"][m.month] += m.count
    blocks["3. 進機 (台)(Fab/Network/SKU/AG)"] = _df(movein, months)

    # 4. 退還(台)—— Fab/Network/SKU/AG
    ret = defaultdict(lambda: defaultdict(int))
    for r in plan_input.returns:
        ret[f"{r.pool.fab}/{r.pool.bm_group}/{r.sku_name}/{r.ag or '-'}"][r.month] += r.count
    blocks["4. 退還 (台)(Fab/Network/SKU/AG)"] = _df(ret, months)

    # 5. 剩餘可用 vcore(月末)—— Fab/Network/AG(solver)
    if horizon_result is not None:
        avail = defaultdict(lambda: defaultdict(float))
        for o in horizon_result.outcomes:
            for c in o.cells:
                avail[f"{o.fab}/{c['network']}/{c['ag']}"][o.period] = c["available_vcore"]
        blocks["5. 剩餘可用 vcore (月末)(Fab/Network/AG,solver)"] = _df(avail, months)

    return blocks
