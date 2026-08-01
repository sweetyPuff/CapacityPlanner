"""procure_adapter 測試:用假 procure 後端(回固定 ProcurementResult)驗證
request 組裝(單月 in-stock 淨量)與 assignments → 需求單 的映射,不需 live solver。"""
from captool.models import (Cap, CurrentStock, DemandDelta, MoveIn, NodeReturn,
                            PlanInput, Pool, Sku, VmSpecDemand)
from captool.placement_viz import placement_svg
from captool.solver.procure_adapter import (build_procurement_request,
                                            demand_order, execution_plan)

STD = Sku(name="std-64", vcore_per_node=64, usable_ratio=0.8)   # sellable 51
BIG = Sku(name="big-128", vcore_per_node=128, usable_ratio=0.8)
A1 = Pool(fab="A", bm_group="network1")


def _pi():
    return PlanInput(
        skus={STD.name: STD, BIG.name: BIG}, months=["2026-07", "2026-08"],
        pools=[A1],
        demands=[DemandDelta(pool=A1, product="web", month="2026-08", vcore=16)],
        vm_demands=[VmSpecDemand(pool=A1, product="db", month="2026-08",
                                 vm_size_vcore=8, count=1)],
        moveins=[MoveIn(pool=A1, sku_name="std-64", month="2026-07", count=1,
                        ag="ag1")],
        returns=[NodeReturn(pool=A1, product="old", sku_name="std-64",
                            month="2026-07", count=1, ag="ag1")],
        currents=[CurrentStock(pool=A1, sku_name="std-64", count=2, ag="ag1")],
        caps=[Cap(pool=A1, ag="ag1", max_bm=20)])


def test_build_request_single_month():
    req, meta, bm_sku = build_procurement_request(_pi(), "A", "2026-08")
    # 兩個需求 → 兩列 requirements(coarse web + detail db)
    assert len(req["requirements"]) == 2
    web = req["requirements"][0]
    assert web["cluster_id"] == "web" and web["network"] == "network1"
    assert web["node_role"] == "worker" and web["total_resources"]["cpu_cores"] == 16
    db = req["requirements"][1]
    assert db["cluster_id"] == "db" and db["min_total_vms"] == 1
    # req_meta:web = 8vcore worker × ceil(16/8)=2;db = 8vcore × 1
    assert meta[0] == {"product": "web", "cluster": "web", "tenant": "free",
                       "network": "network1", "vm_spec_vcore": 8, "vm_count": 2}
    assert meta[1]["vm_count"] == 1
    # in-stock 淨量:current 2 + movein(≤08) 1 − return(≤08) 1 = 2 台 std-64 ag1
    assert len(req["in_stock"]) == 2
    assert all(b["network"] == "network1" for b in req["in_stock"])
    assert set(bm_sku.values()) == {"std-64"}
    assert len(bm_sku) == 2


def test_demand_order_maps_assignments():
    # web 的 2 顆:1 顆落既有 std-64、1 顆落新採購 big-128;db 的 1 顆也落同一台 big-128
    instock_bm = "A~network1~std-64~ag1~0"
    result = {
        "success": True,
        "bought_type_of": {"buy-big-0": "big-128"},
        "bought_bms": [{"id": "buy-big-0", "network": "network1"}],
        "assignments": [
            {"vm_id": "split-r0-s0-0", "baremetal_id": instock_bm, "ag": "ag1"},
            {"vm_id": "split-r0-s0-1", "baremetal_id": "buy-big-0", "ag": "ag2"},
            {"vm_id": "split-r1-s0-0", "baremetal_id": "buy-big-0", "ag": "ag2"},
        ]}
    rows, buys = demand_order(_pi(), "2026-08", lambda req: result)
    by = {r.product: r for r in rows}
    # web:落 std-64(既有 1 台)+ big-128(新採購 1 台)
    assert by["web"].by_sku["std-64"] == {"vm": 1, "bm": 1, "new": 0}
    assert by["web"].by_sku["big-128"] == {"vm": 1, "bm": 1, "new": 1}
    # db:1 顆落 big-128(與 web 共用同一台 → 各自都算觸及 1 台)
    assert by["db"].by_sku["big-128"] == {"vm": 1, "bm": 1, "new": 1}
    # 實際下單清單:big-128 ×1(去重後的真實採購)
    assert buys[("A", "network1")]["big-128"] == 1


def test_cluster_field_drives_cluster_id():
    pi = PlanInput(
        skus={STD.name: STD, BIG.name: BIG}, months=["2026-08"], pools=[A1],
        demands=[DemandDelta(pool=A1, product="giga", month="2026-08",
                             vcore=16, cluster="stg1")],
        vm_demands=[], moveins=[], returns=[], currents=[],
        caps=[Cap(pool=A1, ag="ag1", max_bm=20)])
    req, meta, _ = build_procurement_request(pi, "A", "2026-08")
    assert req["requirements"][0]["cluster_id"] == "stg1"   # cluster,不是 product
    assert meta[0]["product"] == "giga" and meta[0]["cluster"] == "stg1"


def test_tenant_partitions_isolate_solves():
    # 獨佔自成一組、free 同一組 → solve_fn 應被分開呼叫,requirements 互斥
    pi = PlanInput(
        skus={STD.name: STD}, months=["2026-08"], pools=[A1],
        demands=[
            DemandDelta(pool=A1, product="giga", month="2026-08", vcore=8,
                        cluster="stg1", tenant="exclusive"),
            DemandDelta(pool=A1, product="iso", month="2026-08", vcore=8,
                        cluster="isoc", tenant="free"),
            DemandDelta(pool=A1, product="db", month="2026-08", vcore=8,
                        cluster="dbc", tenant="free"),
        ],
        vm_demands=[], moveins=[], returns=[], currents=[],
        caps=[Cap(pool=A1, ag="ag1", max_bm=20)])
    seen = []

    def fake(req):
        seen.append({r["cluster_id"] for r in req["requirements"]})
        return {"assignments": [], "bought_bms": [], "bought_type_of": {}}

    execution_plan(pi, "2026-08", fake)
    assert {"stg1"} in seen           # 獨佔:自己一組
    assert {"isoc", "dbc"} in seen    # free:同一組
    assert len(seen) == 2             # 剛好兩次求解,互不混合


def _shared_result():
    instock_bm = "A~network1~std-64~ag1~0"
    return {
        "success": True,
        "bought_type_of": {"buy-big-0": "big-128"},
        "bought_bms": [{"id": "buy-big-0", "network": "network1"}],
        "assignments": [
            {"vm_id": "split-r0-s0-0", "baremetal_id": instock_bm, "ag": "ag1"},
            {"vm_id": "split-r0-s0-1", "baremetal_id": "buy-big-0", "ag": "ag2"},
            {"vm_id": "split-r1-s0-0", "baremetal_id": "buy-big-0", "ag": "ag2"},
        ]}


def test_execution_plan_tree_shows_sharing():
    # web + db 皆 free(同 partition)→ 可共住同一台 big-128
    _, _, tree = execution_plan(_pi(), "2026-08", lambda req: _shared_result())
    # tree 的 bm key 以 partition 前綴命名,故用內容找(不寫死 key)
    shared = next(bm for bm in tree["A"]["network1"]["ag2"].values()
                  if bm["sku"] == "big-128")
    assert shared["is_new"] is True
    assert sorted(vm["p"] for vm in shared["vms"]) == ["db", "web"]
    assert all("v" in vm for vm in shared["vms"])
    # 既有 std-64 在 ag1 住 web
    instock = next(bm for bm in tree["A"]["network1"]["ag1"].values()
                   if bm["sku"] == "std-64")
    assert [vm["p"] for vm in instock["vms"]] == ["web"]
    # 空 AG 也在(caps 宣告 ag1)
    assert "ag1" in tree["A"]["network1"]


def test_placement_svg_renders():
    _, _, tree = execution_plan(_pi(), "2026-08", lambda req: _shared_result())
    svg = placement_svg(tree, "2026-08")
    assert svg.startswith("<svg") and svg.rstrip().endswith("</svg>")
    for token in ("web", "db", "big-128", "AG=ag2", "[新]"):
        assert token in svg
