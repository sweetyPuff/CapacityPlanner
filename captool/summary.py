"""統整 summary(仿 活頁簿1.xlsx summary 分頁,升級成多機型 × 多 AG)。

五個區塊,列的維度視資料而定(cols = 月份):
  1. 需求 vcore          —— Fab/Network(需求無 SKU/AG 維度)
  2. 實體機需求 (台)      —— Fab/Network/AG,solver 求解後 = bm_bought + in_stock_bm_used
  3. 進機 (台)           —— Fab/Network/SKU/AG(HW_MoveIn)
  4. 退還 (台)           —— Fab/Network/SKU(退還無 AG 維度)
  5. 剩餘可用 vcore (月末) —— Fab/Network/AG,solver 回的 in_stock_available
  6. 需求 ↔ 供給機型對照 —— Fab/Network/SKU,附需求產品 + 既有台數 + 每月採購

2、5、6 需要 solver 結果(horizon_result);沒有時該三塊不產生。
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

    # 2. 實體機需求(採購,台)—— Fab/Network/SKU/AG(solver budget_view)
    if horizon_result is not None:
        req = defaultdict(lambda: defaultdict(int))
        for b in horizon_result.budget:
            req[f"{b['fab']}/{b['network']}/{b['sku']}/{b['ag']}"][b["period"]] += b["count"]
        blocks["2. 實體機需求-採購 (台)(Fab/Network/SKU/AG,solver)"] = _df(req, months)

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
                avail[f"{o.fab}/{c['network']}/{c['ag']}"][o.period] = c.get(
                    "available_vcore", 0)
        blocks["5. 剩餘可用 vcore (月末)(Fab/Network/AG,solver)"] = _df(avail, months)

    # 6. 需求 ↔ 供給機型對照 —— Fab/Network/SKU(AG 合併)
    # 誠實粒度:network 是硬分區,同 network 的需求共用機器池;SKU 由 solver 決定。
    # 每列 = 某 network 用某 SKU support;「需求產品」= 該 network 有哪些需求,
    # 「既有台數」= in-stock 現況,月欄 = 該月新採購台數(budget_view)。
    if horizon_result is not None:
        prods = defaultdict(set)
        for d in plan_input.demands:
            if d.vcore:
                prods[(d.pool.fab, d.pool.bm_group)].add(d.product)
        for v in plan_input.vm_demands:
            prods[(v.pool.fab, v.pool.bm_group)].add(v.product)
        instock = defaultdict(int)               # (fab,net,sku) -> 既有台數
        for c in plan_input.currents:
            instock[(c.pool.fab, c.pool.bm_group, c.sku_name)] += c.count
        buy = defaultdict(lambda: defaultdict(int))  # (fab,net,sku) -> month -> 台數
        for b in horizon_result.budget:
            buy[(b["fab"], b["network"], b["sku"])][b["period"]] += b["count"]
        keys = sorted(set(instock) | set(buy))
        idx, data = [], []
        for fab, net, sku in keys:
            idx.append(f"{fab}/{net}/{sku}")
            plist = ", ".join(sorted(prods.get((fab, net), []))) or "—"
            data.append([plist, instock[(fab, net, sku)]]
                        + [buy[(fab, net, sku)].get(m, 0) for m in months])
        blocks["6. 需求 ↔ 供給機型對照 (Fab/Network/SKU,solver)"] = pd.DataFrame(
            data, index=idx, columns=["需求產品", "既有台數"] + months)

    return blocks
