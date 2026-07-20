import pytest

from captool.models import (CurrentStock, DemandDelta, MoveIn, NodeReturn,
                            PlanInput, Pool, Sku, VmSpecDemand)
from captool.planner import run_check
from captool.solver.naive import NaiveSolver

SKU = Sku(name="default-64", vcore_per_node=64, usable_ratio=0.8)  # 可售 51.2
P = Pool(fab="A", bm_group="network1")


def _pi(demands, moveins=(), returns=(), current=10, vm_demands=(), months=("2026-07", "2026-08")):
    return PlanInput(
        skus={SKU.name: SKU}, months=list(months), pools=[P],
        demands=[DemandDelta(pool=P, product="X", month=m, vcore=v) for m, v in demands],
        vm_demands=[VmSpecDemand(pool=P, product="X", month=m, vm_size_vcore=s, count=c)
                    for m, s, c in vm_demands],
        moveins=[MoveIn(pool=P, sku_name=SKU.name, month=m, count=c) for m, c in moveins],
        returns=[NodeReturn(pool=P, product="X", sku_name=SKU.name, month=m, count=c)
                 for m, c in returns],
        currents=[CurrentStock(pool=P, sku_name=SKU.name, count=current)],
    )


def test_feasible_two_months():
    result = run_check(_pi(demands=[("2026-07", 100), ("2026-08", 100)]), NaiveSolver())
    o7, o8 = result.for_pool(P)
    assert o7.feasible and o8.feasible
    assert o7.machines_opened == {"default-64": 2}   # ceil(100/51.2)
    assert result.gap_months == []


def test_conservative_drops_leftover_across_months():
    # 30 vcore 占 1 台剩 21.2;保守模式下月零頭作廢,8 月的 30 再開新機
    result = run_check(_pi(demands=[("2026-07", 30), ("2026-08", 30)], current=2),
                       NaiveSolver(), mode="conservative")
    o7, o8 = result.for_pool(P)
    assert o7.machines_opened == {"default-64": 1}
    assert o8.machines_opened == {"default-64": 1}
    assert o8.free_vcore_total == pytest.approx(0)   # 兩台的零頭都作廢


def test_gap_fill_reuses_leftover_across_months():
    result = run_check(_pi(demands=[("2026-07", 30), ("2026-08", 20)], current=2),
                       NaiveSolver(), mode="gap_fill")
    o7, o8 = result.for_pool(P)
    assert o8.machines_opened == {}                  # 8 月 20 vcore 填 7 月剩的 21.2
    assert o8.free_vcore_total == pytest.approx(2 * 51.2 - 50)


def test_gap_and_backlog_carries():
    # 7 月需求 120 > 現有 2 台 102.4 → 缺口;8 月進 2 台後補上 backlog
    result = run_check(
        _pi(demands=[("2026-07", 120)], moveins=[("2026-08", 2)], current=2),
        NaiveSolver(), mode="gap_fill")
    o7, o8 = result.for_pool(P)
    assert not o7.feasible
    assert o7.shortfall_vcore == pytest.approx(120 - 102.4)
    assert o8.feasible                               # backlog 17.6 在 8 月被吸收
    assert result.gap_months == [o7]


def test_returns_add_empty_machines():
    result = run_check(
        _pi(demands=[("2026-07", 0), ("2026-08", 100)], returns=[("2026-08", 2)], current=0),
        NaiveSolver())
    o7, o8 = result.for_pool(P)
    assert o8.feasible
    assert o8.stock_total == {"default-64": 2}


def test_excel_compat_matches_formula():
    # stock = 50 + 0 + 0 - 870/51.2 = 33.0078125(對照現行 Excel AB3 邏輯)
    result = run_check(
        _pi(demands=[("2026-07", 870), ("2026-08", 985)], moveins=[("2026-08", 35)],
            returns=[("2026-08", 12)], current=50),
        NaiveSolver(), mode="excel_compat")
    o7, o8 = result.for_pool(P)
    assert o7.stock_float == pytest.approx(50 - 870 / 51.2)
    assert o8.stock_float == pytest.approx(o7.stock_float + 35 + 12 - 985 / 51.2)
    assert o7.feasible and o8.feasible


def test_excel_compat_negative_stock_is_gap():
    result = run_check(_pi(demands=[("2026-07", 600)], current=1, months=("2026-07",)), NaiveSolver(),
                       mode="excel_compat")
    (o7,) = result.for_pool(P)
    assert o7.stock_float < 0
    assert not o7.feasible


def test_excel_compat_requires_single_sku():
    pi = _pi(demands=[("2026-07", 10)])
    pi.skus["big-128"] = Sku(name="big-128", vcore_per_node=128, usable_ratio=0.8)
    with pytest.raises(ValueError):
        run_check(pi, NaiveSolver(), mode="excel_compat")


def test_run_suggest_fills_gap():
    from captool.planner import run_suggest
    pi = _pi(demands=[("2026-07", 120), ("2026-08", 60)], current=0)
    result, suggested = run_suggest(pi, NaiveSolver(), [SKU])
    o7, o8 = result.for_pool(P)
    assert o7.feasible and o8.feasible
    assert o7.suggested_purchases == {"default-64": 3}   # ceil(120/51.2)
    assert result.gap_months == []
    assert [(m.month, m.count) for m in suggested if m.pool == P] == [
        ("2026-07", 3), ("2026-08", 2)]                   # 保守模式:8 月再開 2 台


def test_run_suggest_no_purchase_when_stock_enough():
    from captool.planner import run_suggest
    pi = _pi(demands=[("2026-07", 100)], current=10, months=("2026-07",))
    result, suggested = run_suggest(pi, NaiveSolver(), [SKU])
    assert suggested == []
    assert result.for_pool(P)[0].suggested_purchases == {}


def test_run_suggest_rejects_excel_compat():
    from captool.planner import run_suggest
    with pytest.raises(ValueError):
        run_suggest(_pi(demands=[("2026-07", 1)]), NaiveSolver(), [SKU],
                    mode="excel_compat")
