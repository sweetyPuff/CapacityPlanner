import pytest

from captool.models import Sku
from captool.solver.interface import DemandBatch, PoolState
from captool.solver.naive import NaiveSolver

SKU64 = Sku(name="default-64", vcore_per_node=64, usable_ratio=0.8)   # 可售 51.2
SKU128 = Sku(name="big-128", vcore_per_node=128, usable_ratio=0.8)    # 可售 102.4


def _state(sku=SKU64, count=10):
    s = PoolState()
    s.add_empty(sku, count)
    return s


def test_liquid_only_single_sku_matches_division():
    # 870 vcore / 51.2 = 16.99…,應開 17 台
    state = _state(count=50)
    result = NaiveSolver().check(state, DemandBatch(liquid_vcore=870))
    assert result.feasible
    assert result.machines_opened == {"default-64": 17}
    assert result.placed_liquid_vcore == 870
    assert state.empty_count() == {"default-64": 33}


def test_liquid_infeasible_partial_place():
    state = _state(count=2)  # 容量 102.4
    result = NaiveSolver().check(state, DemandBatch(liquid_vcore=150))
    assert not result.feasible
    assert result.placed_liquid_vcore == pytest.approx(102.4)
    assert result.unplaced_liquid_vcore == pytest.approx(47.6)
    assert result.shortfall_vcore == pytest.approx(47.6)


def test_atomic_vm_blocked_even_when_total_fits():
    # 總量夠(2 台共 102.4)但單台裝不下 60 vcore 的 VM
    state = _state(count=2)
    result = NaiveSolver().check(state, DemandBatch(atomic_vms=[(60, 1)]))
    assert not result.feasible
    assert result.blocked_vms == [(60, 1)]
    assert result.shortfall_vcore == pytest.approx(60)


def test_atomic_ffd_and_gap_fill_within_batch():
    # 51.2 可售:一台放 32+16,另一台放 32 → 2 台
    state = _state(count=5)
    result = NaiveSolver().check(
        state, DemandBatch(atomic_vms=[(16, 1), (32, 2)]))
    assert result.feasible
    assert result.machines_opened == {"default-64": 2}


def test_liquid_fills_partial_before_opening_empty():
    state = _state(count=2)
    NaiveSolver().check(state, DemandBatch(atomic_vms=[(32, 1)]))  # 剩 19.2 的半台
    result = NaiveSolver().check(state, DemandBatch(liquid_vcore=10))
    assert result.feasible
    assert result.machines_opened == {}  # 填縫,不開新機


def test_drop_partial_leftovers():
    state = _state(count=2)
    NaiveSolver().check(state, DemandBatch(liquid_vcore=60))  # 1 滿 + 1 剩 42.4
    dropped = state.drop_partial_leftovers()
    assert dropped == pytest.approx(42.4)
    assert state.free_vcore_total() == pytest.approx(0)
    assert state.total_count() == {"default-64": 2}


def test_multi_sku_check():
    state = PoolState()
    state.add_empty(SKU64, 1)
    state.add_empty(SKU128, 1)
    result = NaiveSolver().check(state, DemandBatch(atomic_vms=[(100, 1), (40, 1)]))
    assert result.feasible
    assert result.machines_opened == {"big-128": 1, "default-64": 1}
