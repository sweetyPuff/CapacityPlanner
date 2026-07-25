"""horizon_adapter 測試:用假後端(回傳固定 CapacityReport)驗證 request 組裝與結果映射,
不需 live solver。"""
from captool.models import (Cap, CurrentStock, DemandDelta, PlanInput, Pool, Sku)
from captool.solver.horizon_adapter import (build_capacity_plan_request,
                                            plan_horizon)

SKU = Sku(name="std-64", vcore_per_node=64, usable_ratio=0.8)   # sellable 51.2 → 51
A1 = Pool(fab="A", bm_group="network1")


def _plan_input():
    return PlanInput(
        skus={SKU.name: SKU}, months=["2026-07"], pools=[A1],
        demands=[DemandDelta(pool=A1, product="web", month="2026-07", vcore=100)],
        vm_demands=[], moveins=[], returns=[],
        currents=[CurrentStock(pool=A1, sku_name=SKU.name, count=2, ag="ag1"),
                  CurrentStock(pool=A1, sku_name=SKU.name, count=3, ag="ag2")],
        caps=[Cap(pool=A1, ag="ag1", max_bm=20),
              Cap(pool=A1, ag="ag2", max_bm=15)])


def test_build_request_worker_mapping():
    req = build_capacity_plan_request(_plan_input(), "A")
    assert len(req["demand_book"]) == 1
    d = req["demand_book"][0]
    assert d["node_role"] == "worker" and d["cpu_cores"] == 100
    assert d["network"] == "network1" and d["cluster_id"] == "web"
    # in_stock:2 台 ag1 + 3 台 ag2,每台 sellable floor 51
    assert len(req["in_stock"]) == 5
    assert all(bm["total_capacity"]["cpu_cores"] == 51 for bm in req["in_stock"])
    assert sorted(bm["topology"]["ag"] for bm in req["in_stock"]) == \
        ["ag1", "ag1", "ag2", "ag2", "ag2"]
    caps = {c["bucket"]: c["max_bm"] for c in req["procurement_caps"]}
    assert caps == {"ag1": 20, "ag2": 15}
    assert req["config"]["procurement_spread_dimension"] == "ag"


def test_plan_horizon_maps_report():
    report = {"success": True, "by_fab_period": [{
        "period": "2026-07", "success": True, "node_adds_total": 13,
        "procurement": [{"type_id": "std-64", "count": 2}],
        "balance_after": {"ag1": 10, "ag2": 20},
        "cells": [{"bucket": "ag1", "in_stock_available": {"cpu_cores": 10}},
                  {"bucket": "ag2", "in_stock_available": {"cpu_cores": 20}}],
        "shortfalls": []}]}
    captured = {}

    def fake(req):
        captured["req"] = req
        return report

    res = plan_horizon(_plan_input(), fake)
    assert len(res.outcomes) == 1
    o = res.outcomes[0]
    assert o.fab == "A" and o.period == "2026-07" and o.feasible
    assert o.node_adds == 13
    assert o.procurement == {"std-64": 2}
    assert o.balance_after == {"ag1": 10, "ag2": 20}
    assert o.ag_available == {"ag1": 10, "ag2": 20}
    assert res.gaps == []
    # 確認送出的是我方組的 request
    assert captured["req"]["demand_book"][0]["cpu_cores"] == 100


def test_plan_horizon_reports_gap():
    report = {"success": False, "by_fab_period": [{
        "period": "2026-07", "success": False, "node_adds_total": 0,
        "procurement": [], "balance_after": {}, "cells": [],
        "shortfalls": [{"cause": "space", "message": "AG 槽位不足"}]}]}
    res = plan_horizon(_plan_input(), lambda req: report)
    assert len(res.gaps) == 1
    assert "槽位" in res.gaps[0].shortfalls[0]
