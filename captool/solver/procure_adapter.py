"""執行面 procure adapter:把「某單月已確定需求」翻成 solver 的 ProcurementRequest、
POST 到 `/v1/capacity/procure`,再從 ProcurementResult 的 assignments 組出「需求單」。

與 horizon(規劃面)的差別:
- 規劃面 = 多月 roll-forward、只看整體方向、報表聚合掉逐機落點。
- 執行面 = **單月已確定需求**,走單期 procure 端點 —— 它本來就回 `assignments`
  (逐 VM→BM)+ `bought_type_of`(bought bm→SKU),所以逐需求→SKU 湊得回來,
  **無跨團隊相依**(不需 solver 端改)。

in-stock 起點 = 目標月的淨量:currents + moveins(月≤目標) − returns(月≤目標)。
需求全走 `requirements`(splitter 會拆成 VM 並可落在可買機),VM id = `split-r{idx}-...`,
idx 對回我方送出的 requirements 順序 → cluster 名。
"""
from __future__ import annotations

import math
from collections import defaultdict
from dataclasses import dataclass, field

from captool.models import PlanInput
from captool.solver.horizon_adapter import (DEFAULT_WORKER_VM, _BIG_MEM,
                                            _BIG_STO)


def _net_instock(plan_input: PlanInput, fab: str, month: str) -> "dict[tuple, int]":
    """目標月起點的 in-stock 淨量,per (network, sku, ag)。月字串 ISO,可字典序比較。"""
    net: "dict[tuple, int]" = defaultdict(int)
    for c in plan_input.currents:
        if c.pool.fab == fab:
            net[(c.pool.bm_group, c.sku_name, c.ag)] += c.count
    for m in plan_input.moveins:
        if m.pool.fab == fab and m.month <= month:
            net[(m.pool.bm_group, m.sku_name, m.ag)] += m.count
    for r in plan_input.returns:
        if r.pool.fab != fab or r.month > month:
            continue
        if r.ag:
            net[(r.pool.bm_group, r.sku_name, r.ag)] -= r.count
        else:                       # ag 未指定:從該 (network, sku) 各 ag 依序扣
            rem = r.count
            for key in [k for k in net
                        if k[0] == r.pool.bm_group and k[1] == r.sku_name]:
                take = min(net[key], rem)
                net[key] -= take
                rem -= take
                if rem <= 0:
                    break
    return {k: v for k, v in net.items() if v > 0}


def _allowed_types(skus: dict, vm_size: int) -> "list[str] | None":
    """合身機型:裝得下 vm_size,且剔除『更大但每台裝的 VM 數不變』的機型
    (那種只是浪費,無 packing 好處)。回 None = 不限制(合身集=全部或無)。"""
    if not vm_size:
        return None
    cap = {n: int(math.floor(s.sellable_vcore)) for n, s in skus.items()}
    fit = {n: c for n, c in cap.items() if c >= vm_size}
    if not fit:
        return None                       # 沒機型裝得下 → 交給 solver 回報
    per = {n: c // vm_size for n, c in fit.items()}       # 每台可住幾顆
    keep = [n for n in fit
            if not any(fit[o] < fit[n] and per[o] >= per[n] for o in fit)]
    return sorted(keep) if 0 < len(keep) < len(skus) else None


def build_procurement_request(plan_input: PlanInput, fab: str, month: str,
                              worker_vm: dict = DEFAULT_WORKER_VM,
                              max_solve_time_seconds: float = 10.0):
    """回 (request dict, req_meta, bm_sku)。req_meta[idx] = 該 requirement 的
    {cluster, network, vm_spec_vcore, vm_count};bm_sku[bm_id] = 既有機的 SKU。"""
    requirements: list = []
    req_meta: list = []

    for d in plan_input.demands:
        if d.pool.fab != fab or d.month != month or not d.vcore:
            continue
        cid = d.cluster or d.product            # cluster 為 solver 分組單位
        r = {"total_resources": {"cpu_cores": int(d.vcore), "memory_mib": 0,
                                 "storage_gb": 0},
             "node_role": "worker", "cluster_id": cid, "ip_type": "plan",
             "network": d.pool.bm_group, "vm_specs": [worker_vm]}
        allowed = _allowed_types(plan_input.skus, worker_vm["cpu_cores"])
        if allowed:
            r["allowed_bm_types"] = allowed
        requirements.append(r)
        cpu = worker_vm["cpu_cores"]
        req_meta.append({"product": d.product, "cluster": cid,
                         "tenant": d.tenant or "free",
                         "network": d.pool.bm_group, "vm_spec_vcore": cpu,
                         "vm_count": math.ceil(d.vcore / cpu) if cpu else 0})

    for v in plan_input.vm_demands:
        if v.pool.fab != fab or v.month != month:
            continue
        cid = v.cluster or v.product
        r = {"total_resources": {"cpu_cores": v.vm_size_vcore * v.count,
                                 "memory_mib": 0, "storage_gb": 0},
             "node_role": "worker", "cluster_id": cid, "ip_type": "plan",
             "network": v.pool.bm_group,
             "vm_specs": [{"cpu_cores": v.vm_size_vcore, "memory_mib": 0,
                           "storage_gb": 0}],
             "min_total_vms": v.count, "max_total_vms": v.count}
        allowed = _allowed_types(plan_input.skus, v.vm_size_vcore)
        if allowed:
            r["allowed_bm_types"] = allowed
        requirements.append(r)
        req_meta.append({"product": v.product, "cluster": cid,
                         "tenant": v.tenant or "free",
                         "network": v.pool.bm_group,
                         "vm_spec_vcore": v.vm_size_vcore, "vm_count": v.count})

    in_stock: list = []
    bm_sku: dict = {}
    for (network, sku, ag), cnt in _net_instock(plan_input, fab, month).items():
        cpu = int(math.floor(plan_input.skus[sku].sellable_vcore))
        for k in range(cnt):
            bm_id = f"{fab}~{network}~{sku}~{ag}~{k}"   # ~ 分隔:SKU 名可能含 -
            in_stock.append({
                "id": bm_id,
                "total_capacity": {"cpu_cores": cpu, "memory_mib": _BIG_MEM,
                                   "storage_gb": _BIG_STO},
                "topology": {"ag": ag}, "network": network})
            bm_sku[bm_id] = sku

    procurement_types = [
        {"type_id": name,
         "capacity": {"cpu_cores": int(math.floor(s.sellable_vcore)),
                      "memory_mib": _BIG_MEM, "storage_gb": _BIG_STO}}
        for name, s in plan_input.skus.items()]
    procurement_caps = [
        {"bucket": cap.ag, "max_bm": cap.max_bm, "network": cap.pool.bm_group}
        for cap in plan_input.caps if cap.pool.fab == fab]

    req = {"requirements": requirements, "in_stock": in_stock,
           "procurement_types": procurement_types,
           "procurement_caps": procurement_caps,
           "config": {"procurement_spread_dimension": "ag",
                      "max_solve_time_seconds": max_solve_time_seconds}}
    return req, req_meta, bm_sku


@dataclass
class DemandOrderRow:
    fab: str
    network: str
    product: str
    cluster: str
    month: str
    vm_spec_vcore: int
    vm_count: int
    # sku -> {"vm": 顆數, "bm": 觸及台數(含共用), "new": 其中新採購台數}
    by_sku: dict = field(default_factory=dict)


def _partition_key(meta: dict):
    """租戶 → partition 鍵。exclusive=每 cluster 自成一組;群組名=同名一組;free=一組。
    排序:free(0) → 群組(1) → 獨佔(2),同類再依名稱,結果穩定。"""
    t = meta.get("tenant") or "free"
    if t == "exclusive":
        return (2, meta["cluster"])
    if t == "free":
        return (0, "")
    return (1, t)


def execution_plan(plan_input: PlanInput, month: str, solve_fn):
    """依租戶分 partition、分開求解、既有機用過即從池中移除(下一 partition 拿不到)
    → 不同租戶天生不共住。回 (rows, buys, tree)。

    solver 無原生租戶/獨佔約束(只有 C3/C4/C5),故隔離在我方以「分池序解」達成。
    註:AG 採購上限(caps)目前未跨 partition 遞減,多 partition 同 AG 大量採購時可能超額
    (§ 待辦);sample caps 寬鬆不受影響。
    """
    rows: list = []
    buys: dict = defaultdict(lambda: defaultdict(int))
    tree: dict = {}
    for fab in sorted({p.fab for p in plan_input.pools}):
        req, req_meta, bm_sku = build_procurement_request(plan_input, fab, month)
        if not req["requirements"]:
            continue

        # 每個 requirement 綁上它的 meta 與空的 by_sku 累加器
        entries = [{"req": r, "meta": m,
                    "by": defaultdict(lambda: {"vm": 0, "bm": set(),
                                               "new": set()})}
                   for r, m in zip(req["requirements"], req_meta)]
        parts: dict = defaultdict(list)
        for e in entries:
            parts[_partition_key(e["meta"])].append(e)

        pool = list(req["in_stock"])          # 逐 partition 遞減的既有機池
        for pki, pkey in enumerate(sorted(parts)):
            part = parts[pkey]
            pr = {"requirements": [e["req"] for e in part],
                  "in_stock": pool,
                  "procurement_types": req["procurement_types"],
                  "procurement_caps": req["procurement_caps"],
                  "config": req["config"]}
            res = solve_fn(pr)
            bought_type_of = res.get("bought_type_of", {})
            bought_net = {b.get("id"): b.get("network", "")
                          for b in res.get("bought_bms", [])}
            pool_ids = {bm["id"] for bm in pool}
            used_instock: set = set()

            for a in res.get("assignments", []):
                vid, bm = a.get("vm_id", ""), a.get("baremetal_id", "")
                if not vid.startswith("split-"):
                    continue
                try:
                    ridx = int(vid.split("-")[1][1:])
                except (IndexError, ValueError):
                    continue
                if ridx >= len(part):
                    continue
                e = part[ridx]
                is_new = bm in bought_type_of
                sku = bought_type_of.get(bm) or bm_sku.get(bm, "?")
                cell = e["by"][sku]
                cell["vm"] += 1
                cell["bm"].add(bm)
                if is_new:
                    cell["new"].add(bm)
                else:
                    if bm in pool_ids:
                        used_instock.add(bm)
                # 機櫃樹:bm id 以 partition 前綴命名(避免各 partition 的採購合成 id 相撞)
                ag = a.get("ag", "")
                if is_new:
                    network = bought_net.get(bm, "")
                else:
                    seg = bm.split("~")           # fab~network~sku~ag~k
                    network = seg[1] if len(seg) > 1 else ""
                    if not ag and len(seg) > 3:
                        ag = seg[3]
                node = tree.setdefault(fab, {}).setdefault(
                    network, {}).setdefault(ag or "-", {})
                node.setdefault(f"p{pki}:{bm}",
                                {"sku": sku, "is_new": is_new, "vms": []})[
                    "vms"].append({"p": e["meta"]["cluster"],
                                   "v": e["meta"]["vm_spec_vcore"]})

            for bm in res.get("bought_bms", []):
                buys[(fab, bm.get("network", ""))][
                    bought_type_of.get(bm.get("id"), "?")] += 1
            # 用掉的既有機從池中移除 → 下一 partition 不得再用(不共住、不重複計)
            pool = [bm for bm in pool if bm["id"] not in used_instock]

        # 需求單列:每個 requirement 一列(by_sku 來自其 partition 的累加器)
        for e in entries:
            m = e["meta"]
            by_sku = {sku: {"vm": c["vm"], "bm": len(c["bm"]),
                            "new": len(c["new"])}
                      for sku, c in e["by"].items()}
            rows.append(DemandOrderRow(
                fab=fab, network=m["network"], product=m["product"],
                cluster=m["cluster"], month=month,
                vm_spec_vcore=m["vm_spec_vcore"], vm_count=m["vm_count"],
                by_sku=by_sku))

    # 空 AG 也畫:補上 HW_Caps 宣告、但本月沒放 VM 的 AG(看得出分散空間)
    declared: dict = defaultdict(set)
    for cap in plan_input.caps:
        declared[(cap.pool.fab, cap.pool.bm_group)].add(cap.ag)
    for fab in tree:
        for network in tree[fab]:
            for ag in declared.get((fab, network), ()):
                tree[fab][network].setdefault(ag or "-", {})
    return rows, buys, tree


def demand_order(plan_input: PlanInput, month: str, solve_fn):
    """需求單 (rows, buys) —— execution_plan 的前兩項(相容既有呼叫/測試)。"""
    rows, buys, _ = execution_plan(plan_input, month, solve_fn)
    return rows, buys


def _fmt_by_sku(by_sku: dict) -> str:
    """'big-128×2台  gpu-80×2台(新2)' —— 台數含共用;(新N)=其中新採購。"""
    if not by_sku:
        return "—"
    return "  ".join(
        f"{sku}×{c['bm']}台" + (f"(新{c['new']})" if c["new"] else "")
        for sku, c in sorted(by_sku.items()))


def demand_order_frames(rows: list, buys: dict):
    """(orders_df, buy_df):orders = 每列一個需求;buy = 去重的實際採購清單。"""
    import pandas as pd
    orders = pd.DataFrame([{
        "Fab": r.fab, "Network": r.network, "Demand": r.product,
        "Cluster": r.cluster, "VM 規格": f"{r.vm_spec_vcore}vcore",
        "VM 台數": r.vm_count,
        "建議實體機 (SKU×台;新=新採購,含共用)": _fmt_by_sku(r.by_sku),
        "Due": r.month} for r in rows],
        columns=["Fab", "Network", "Demand", "Cluster", "VM 規格", "VM 台數",
                 "建議實體機 (SKU×台;新=新採購,含共用)", "Due"])
    buy = pd.DataFrame(
        [{"Fab": f, "Network": n, "SKU": s, "採購台數": c}
         for (f, n), d in sorted(buys.items()) for s, c in sorted(d.items())],
        columns=["Fab", "Network", "SKU", "採購台數"])
    return orders, buy
