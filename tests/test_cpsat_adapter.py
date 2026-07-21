"""CpSatAdapter 測試。

用一個模擬同事 CP-SAT solver JSON 契約的假後端(partial-placement 最佳適配),
驗證 adapter 兩側翻譯與編排正確 —— 不需 ortools,不依賴她的程式碼可否執行。
"""
import copy

import pytest

from captool.models import Sku
from captool.solver.cpsat_adapter import CpSatAdapter
from captool.solver.interface import DemandBatch, PoolState
from captool.solver.naive import NaiveSolver

SKU64 = Sku(name="default-64", vcore_per_node=64, usable_ratio=0.8)   # 可售 51.2
SKU128 = Sku(name="big-128", vcore_per_node=128, usable_ratio=0.8)    # 可售 102.4


def fake_backend(request: dict) -> dict:
    """模擬她的 solver:partial placement + best-fit-decreasing。回傳 PlacementResult dict。"""
    bms = [{"id": b["id"],
            "avail": b["total_capacity"]["cpu_cores"] - b["used_capacity"].get("cpu_cores", 0)}
           for b in request["baremetals"]]
    vms = sorted(request["vms"], key=lambda v: v["demand"]["cpu_cores"], reverse=True)
    assignments, unplaced = [], []
    for v in vms:
        size = v["demand"]["cpu_cores"]
        candidates = [b for b in bms if b["avail"] >= size]
        if not candidates:
            unplaced.append(v["id"])
            continue
        target = min(candidates, key=lambda b: b["avail"])
        target["avail"] -= size
        assignments.append({"vm_id": v["id"], "baremetal_id": target["id"], "ag": ""})
    return {"success": len(unplaced) == 0, "assignments": assignments,
            "unplaced_vms": unplaced, "solver_status": "OPTIMAL"}


def _state(sku=SKU64, count=2):
    s = PoolState()
    s.add_empty(sku, count)
    return s


def _adapter():
    return CpSatAdapter(solve_fn=fake_backend)


def test_liquid_only_matches_naive():
    state_a = _state(count=3)
    state_b = copy.deepcopy(state_a)
    adapter_res = _adapter().check(state_a, DemandBatch(liquid_vcore=100))
    naive_res = NaiveSolver().check(state_b, DemandBatch(liquid_vcore=100))
    assert adapter_res.feasible == naive_res.feasible
    assert adapter_res.placed_liquid_vcore == pytest.approx(naive_res.placed_liquid_vcore)
    assert adapter_res.machines_opened == naive_res.machines_opened


def test_atomic_fits_via_backend():
    state = _state(count=2)  # 兩台空 SKU64(可用 51)
    result = _adapter().check(state, DemandBatch(atomic_vms=[(32, 2)]))
    assert result.feasible
    assert result.machines_opened == {"default-64": 2}  # 兩顆 32 各佔一台
    assert result.blocked_vms == []
    assert state.free_vcore_total() == pytest.approx(2 * 51.2 - 64)


def test_atomic_blocked_when_too_big():
    state = _state(count=1)  # 一台空 SKU64,floor 可用 51
    result = _adapter().check(state, DemandBatch(atomic_vms=[(60, 1)]))
    assert not result.feasible
    assert result.blocked_vms == [(60, 1)]
    assert result.shortfall_vcore == pytest.approx(60)


def test_atomic_plus_liquid():
    state = _state(count=2)
    result = _adapter().check(state, DemandBatch(liquid_vcore=30, atomic_vms=[(40, 1)]))
    assert result.feasible
    assert result.blocked_vms == []
    assert result.placed_liquid_vcore == pytest.approx(30)
    assert result.machines_opened == {"default-64": 2}


def test_machines_opened_only_counts_newly_used():
    state = _state(count=2)
    _adapter().check(state, DemandBatch(atomic_vms=[(20, 1)]))   # 動用一台
    result = _adapter().check(state, DemandBatch(atomic_vms=[(10, 1)]))  # 應填回同一台
    assert result.machines_opened == {}  # 沒有新的空機被動用


def test_suggest_buys_enough():
    state = PoolState()  # 沒庫存
    sres = _adapter().suggest(state, DemandBatch(atomic_vms=[(32, 3)]), [SKU64])
    assert sres.result.feasible
    assert sres.purchases == {"default-64": 3}
    assert state.total_count() == {"default-64": 3}


def test_suggest_no_purchase_when_enough():
    state = _state(count=5)
    sres = _adapter().suggest(state, DemandBatch(atomic_vms=[(32, 1)]), [SKU64])
    assert sres.purchases == {}
    assert sres.result.feasible


def test_drops_into_planner_unchanged():
    """契約替換:CpSatAdapter 可直接當 AllocationSolver 餵給 planner,零改動。"""
    from captool.models import (CurrentStock, DemandDelta, PlanInput, Pool,
                                VmSpecDemand)
    from captool.planner import run_check, run_suggest

    P = Pool(fab="A", bm_group="network1")
    pi = PlanInput(
        skus={SKU64.name: SKU64}, months=["2026-07"], pools=[P],
        demands=[DemandDelta(pool=P, product="X", month="2026-07", vcore=100)],
        vm_demands=[VmSpecDemand(pool=P, product="Y", month="2026-07",
                                 vm_size_vcore=40, count=2)],
        moveins=[], returns=[],
        currents=[CurrentStock(pool=P, sku_name=SKU64.name, count=5)])
    result = run_check(pi, _adapter(), mode="conservative")
    assert result.for_pool(P)[0].feasible

    pi2 = PlanInput(
        skus={SKU64.name: SKU64}, months=["2026-07"], pools=[P],
        demands=[], vm_demands=[VmSpecDemand(pool=P, product="Y", month="2026-07",
                                             vm_size_vcore=40, count=3)],
        moveins=[], returns=[],
        currents=[CurrentStock(pool=P, sku_name=SKU64.name, count=0)])
    _, suggested = run_suggest(pi2, _adapter(), [SKU64])
    assert sum(m.count for m in suggested) == 3
