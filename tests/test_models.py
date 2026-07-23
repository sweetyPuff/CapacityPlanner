from captool.models import (CurrentStock, DemandDelta, MoveIn, NodeReturn,
                            PlanInput, Pool, Sku, VmSpecDemand)

SKU = Sku(name="default-64", vcore_per_node=64, usable_ratio=0.8)
P = Pool(fab="A", bm_group="network1")


def _plan_input():
    return PlanInput(
        skus={SKU.name: SKU},
        months=["2026-07", "2026-08"],
        pools=[P],
        demands=[
            DemandDelta(pool=P, product="Apple", month="2026-07", vcore=512),
            DemandDelta(pool=P, product="Banana", month="2026-07", vcore=358),
        ],
        vm_demands=[
            VmSpecDemand(pool=P, product="Apple", month="2026-07", vm_size_vcore=32, count=2),
            VmSpecDemand(pool=P, product="Banana", month="2026-07", vm_size_vcore=32, count=1),
        ],
        moveins=[MoveIn(pool=P, sku_name=SKU.name, month="2026-08", count=35)],
        returns=[NodeReturn(pool=P, product="Apple", sku_name=SKU.name, month="2026-08", count=12)],
        currents=[CurrentStock(pool=P, sku_name=SKU.name, count=50)],
    )


def test_sellable_vcore():
    assert SKU.sellable_vcore == 51.2


def test_pool_label_and_hashable():
    assert P.label == "A/network1"
    assert {P: 1}[Pool(fab="A", bm_group="network1")] == 1


def test_demand_vcore_sums_products():
    assert _plan_input().demand_vcore(P, "2026-07") == 870


def test_vm_batch_aggregates_by_size():
    assert _plan_input().vm_batch(P, "2026-07") == [(32, 3)]


def test_lookups():
    pi = _plan_input()
    assert pi.movein_by_sku(P, "2026-08") == {"default-64": 35}
    assert pi.movein_by_sku(P, "2026-07") == {}
    assert pi.return_by_sku(P, "2026-08") == {"default-64": 12}
    assert pi.current_by_sku(P) == {"default-64": 50}


def test_product_policy_and_lookup():
    from captool.models import ProductPolicy
    pol = ProductPolicy(pool=P, product="AI", vm_size_vcore=60,
                        max_per_machine=1, co_residency="exclusive")
    pi = PlanInput(
        skus={SKU.name: SKU}, months=["2026-07"], pools=[P],
        demands=[], vm_demands=[], moveins=[], returns=[], currents=[],
        policies=[pol])
    assert pi.policy_for(P, "AI") is pol
    assert pi.policy_for(P, "nope") is None


def test_plan_input_policies_default_empty():
    pi = PlanInput(skus={}, months=[], pools=[], demands=[], vm_demands=[],
                   moveins=[], returns=[], currents=[])
    assert pi.policies == []
