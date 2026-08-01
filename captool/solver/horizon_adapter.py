"""Horizon adapter:把 PlanInput 翻成同事 solver 的 CapacityPlanRequest、POST 到
`/v1/capacity/plan`、再把 CapacityReport 映射回我方結果結構。

路線 A(見 spec 2026-07-25):走 HTTP,我方不裝 ortools;`solve_fn` 依賴注入
(http_solve_fn 生產 / 假後端測試)。**目前為 worker 路徑**(coarse worker 需求);
control-plane 菜單待 §6.2(role enum / co-tenancy)敲定後再接。

單維 vcore:BM cpu_cores = 機型 sellable(floor);mem/storage 設大值不綁;
worker VM 規格由 DEFAULT_WORKER_VM 決定(§6.5 待定,先用固定值)。
每個 fab 一個 request(single-fab 模式 fab="")、network = bm_group、
procurement_spread_dimension = "ag"。
"""
from __future__ import annotations

import json
import math
import urllib.request
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Callable

from captool.models import PlanInput

SolveFn = Callable[[dict], dict]

DEFAULT_WORKER_VM = {"cpu_cores": 8, "memory_mib": 0, "storage_gb": 0}
_BIG_MEM, _BIG_STO = 10 ** 8, 10 ** 6


# ---------------------------------------------------------------------------
# Request 組裝(PlanInput → CapacityPlanRequest dict,per fab)
# ---------------------------------------------------------------------------

def build_capacity_plan_request(plan_input: PlanInput, fab: str,
                                worker_vm: dict = DEFAULT_WORKER_VM,
                                max_solve_time_seconds: float = 10.0) -> dict:
    """單一 fab 的 CapacityPlanRequest(single-fab 模式 fab="",network=bm_group)。"""
    # 粗粒度 worker:整包 vcore,solver 用預設 worker VM 規格切
    demand_book = [
        {"cluster_id": d.cluster or d.product, "node_role": "worker",
         "period": d.month, "cpu_cores": int(d.vcore),
         "network": d.pool.bm_group, "vm_specs": [worker_vm]}
        for d in plan_input.demands if d.pool.fab == fab and d.vcore
    ]
    # detail worker:固定顆數 × 指定 VM 尺寸(min=max=count)。
    # TODO(§6.2 同步):爆炸半徑(max_per_bm)與 Tenant(共居)尚未下約束,先只送需求。
    for v in plan_input.vm_demands:
        if v.pool.fab != fab:
            continue
        demand_book.append({
            "cluster_id": v.cluster or v.product, "node_role": "worker",
            "period": v.month,
            "cpu_cores": v.vm_size_vcore * v.count, "network": v.pool.bm_group,
            "vm_specs": [{"cpu_cores": v.vm_size_vcore, "memory_mib": 0,
                          "storage_gb": 0}],
            "min_total_vms": v.count, "max_total_vms": v.count,
        })

    in_stock = []
    for c in plan_input.currents:
        if c.pool.fab != fab:
            continue
        cpu = int(math.floor(plan_input.skus[c.sku_name].sellable_vcore))
        for k in range(c.count):
            in_stock.append({
                "id": f"{fab}-{c.pool.bm_group}-{c.sku_name}-{c.ag}-{k}",
                "total_capacity": {"cpu_cores": cpu, "memory_mib": _BIG_MEM,
                                   "storage_gb": _BIG_STO},
                "topology": {"ag": c.ag}, "network": c.pool.bm_group,
            })

    procurement_types = [
        {"type_id": name,
         "capacity": {"cpu_cores": int(math.floor(sku.sellable_vcore)),
                      "memory_mib": _BIG_MEM, "storage_gb": _BIG_STO}}
        for name, sku in plan_input.skus.items()
    ]
    procurement_caps = [
        {"bucket": cap.ag, "max_bm": cap.max_bm, "network": cap.pool.bm_group}
        for cap in plan_input.caps if cap.pool.fab == fab
    ]
    return {
        "demand_book": demand_book,
        "in_stock": in_stock,
        "procurement_types": procurement_types,
        "procurement_caps": procurement_caps,
        "config": {"procurement_spread_dimension": "ag",
                   "max_solve_time_seconds": max_solve_time_seconds,
                   "target_spread": {"ag": 3}},
    }


# ---------------------------------------------------------------------------
# 結果映射(CapacityReport → 我方結構)
# ---------------------------------------------------------------------------

@dataclass
class HorizonOutcome:
    fab: str
    period: str
    feasible: bool
    node_adds: int
    procurement: dict[str, int]        # type_id -> 台數
    balance_after: dict[str, int]      # ag -> 剩餘 cpu
    ag_available: dict[str, int]       # ag -> in_stock_available cpu(月末)
    shortfalls: list[str]
    # 每個 (network, ag) 的硬體數:實體機需求 = in_stock_bm_used + bm_bought
    cells: list[dict] = field(default_factory=list)


@dataclass
class HorizonResult:
    outcomes: list[HorizonOutcome] = field(default_factory=list)
    # 採購明細(來自 CapacityReport.budget_view):每筆 = fab/network/ag/period/sku/count
    budget: list[dict] = field(default_factory=list)

    @property
    def gaps(self) -> list[HorizonOutcome]:
        return [o for o in self.outcomes if not o.feasible]

    def for_fab(self, fab: str) -> list[HorizonOutcome]:
        return [o for o in self.outcomes if o.fab == fab]


def _map_report(fab: str, report: dict) -> list[HorizonOutcome]:
    out = []
    for pf in report.get("by_fab_period", []):
        ag_avail = {c["bucket"]: c.get("in_stock_available", {}).get("cpu_cores", 0)
                    for c in pf.get("cells", [])}
        cells = [{"network": c.get("network", ""), "ag": c["bucket"],
                  "bm_bought": c.get("bm_bought", 0),
                  "in_stock_bm_used": c.get("in_stock_bm_used", 0),
                  "node_adds": c.get("node_adds", 0),
                  "available_vcore": c.get("in_stock_available", {}).get("cpu_cores", 0)}
                 for c in pf.get("cells", [])]
        out.append(HorizonOutcome(
            fab=fab, period=pf["period"], feasible=pf["success"],
            node_adds=pf.get("node_adds_total", 0),
            procurement={d["type_id"]: d["count"] for d in pf.get("procurement", [])},
            balance_after=dict(pf.get("balance_after", {})),
            ag_available=ag_avail,
            shortfalls=[s.get("message", "") for s in pf.get("shortfalls", [])],
            cells=cells,
        ))
    return out


def plan_horizon(plan_input: PlanInput, solve_fn: SolveFn,
                 **req_kwargs) -> HorizonResult:
    """每個 fab 組 request、呼叫 solve_fn、映射並彙整成 HorizonResult。"""
    fabs = sorted({p.fab for p in plan_input.pools})
    result = HorizonResult()
    for fab in fabs:
        req = build_capacity_plan_request(plan_input, fab, **req_kwargs)
        if not req["demand_book"]:
            continue
        report = solve_fn(req)
        result.outcomes.extend(_map_report(fab, report))
        for b in report.get("budget_view", []):
            result.budget.append({
                "fab": fab, "network": b.get("network", ""),
                "ag": b.get("bucket", ""), "period": b.get("period", ""),
                "sku": b.get("type_id", ""), "count": b.get("bm_count", 0)})
    return result


# ---------------------------------------------------------------------------
# solve_fn 工廠
# ---------------------------------------------------------------------------

def http_solve_fn(url: str, timeout: float = 60.0) -> SolveFn:
    """POST 到 solver 的 /v1/capacity/plan。"""
    def _fn(request: dict) -> dict:
        data = json.dumps(request).encode("utf-8")
        req = urllib.request.Request(
            url, data=data, headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    return _fn
