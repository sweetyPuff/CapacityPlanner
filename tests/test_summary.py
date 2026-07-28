from captool.models import (CurrentStock, DemandDelta, MoveIn, NodeReturn,
                            PlanInput, Pool, Sku)
from captool.solver.horizon_adapter import HorizonOutcome, HorizonResult
from captool.summary import capacity_summary

SKU = Sku(name="std-64", vcore_per_node=64, usable_ratio=0.8)
A1 = Pool(fab="A", bm_group="network1")


def _pi():
    return PlanInput(
        skus={SKU.name: SKU}, months=["2026-07", "2026-08"], pools=[A1],
        demands=[DemandDelta(pool=A1, product="web", month="2026-07", vcore=200)],
        vm_demands=[], returns=[NodeReturn(pool=A1, product="old", sku_name="std-64",
                                           month="2026-08", count=2, ag="ag2")],
        moveins=[MoveIn(pool=A1, sku_name="std-64", month="2026-08", count=3,
                        ag="ag1")],
        currents=[CurrentStock(pool=A1, sku_name="std-64", count=5, ag="ag1")])


def test_input_blocks_without_solver():
    blocks = capacity_summary(_pi(), None)
    assert blocks["1. 需求 vcore(Fab/Network)"].loc["A/network1", "2026-07"] == 200
    assert blocks["3. 進機 (台)(Fab/Network/SKU/AG)"].loc[
        "A/network1/std-64/ag1", "2026-08"] == 3
    assert blocks["4. 退還 (台)(Fab/Network/SKU/AG)"].loc[
        "A/network1/std-64/ag2", "2026-08"] == 2
    # 沒 solver → 無實體機需求 / 剩餘 vcore 區塊
    assert not any("實體機需求" in k for k in blocks)


def test_solver_blocks_with_horizon():
    hr = HorizonResult(
        outcomes=[HorizonOutcome(
            fab="A", period="2026-07", feasible=True, node_adds=5,
            procurement={}, balance_after={}, ag_available={}, shortfalls=[],
            cells=[{"network": "network1", "ag": "ag1", "bm_bought": 1,
                    "in_stock_bm_used": 2, "node_adds": 5,
                    "available_vcore": 100}])],
        budget=[{"fab": "A", "network": "network1", "ag": "ag1",
                 "period": "2026-07", "sku": "std-64", "count": 2}])
    blocks = capacity_summary(_pi(), hr)
    assert blocks["2. 實體機需求-採購 (台)(Fab/Network/SKU/AG,solver)"].loc[
        "A/network1/std-64/ag1", "2026-07"] == 2   # budget_view 採購台數
    assert blocks["5. 剩餘可用 vcore (月末)(Fab/Network/AG,solver)"].loc[
        "A/network1/ag1", "2026-07"] == 100
