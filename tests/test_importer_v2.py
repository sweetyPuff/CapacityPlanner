from pathlib import Path

from openpyxl import load_workbook

from captool.exporter import generate_v2_template
from captool.importer import import_any, import_legacy, import_v2
from captool.models import NodeReturn, PlanInput, Pool, Sku, VmSpecDemand
from captool.planner import run_check
from captool.solver.naive import NaiveSolver

FIXTURE = Path(__file__).parent / "fixtures" / "sample_legacy.xlsx"


def _v2_file(tmp_path):
    path = tmp_path / "v2.xlsx"
    generate_v2_template(import_legacy(FIXTURE), path)
    return path


def test_roundtrip_preserves_data(tmp_path):
    path = _v2_file(tmp_path)
    original = import_legacy(FIXTURE)
    pi = import_v2(path)
    a1 = Pool(fab="A", bm_group="network1")
    assert pi.demand_vcore(a1, "2026-07") == original.demand_vcore(a1, "2026-07")
    assert pi.return_by_sku(a1, "2026-08") == original.return_by_sku(a1, "2026-08")
    assert pi.current_by_sku(a1) == original.current_by_sku(a1)
    assert pi.movein_by_sku(a1, "2026-08") == original.movein_by_sku(a1, "2026-08")
    assert pi.months == original.months
    # 範本內的示範列不得變成資料
    assert pi.vm_demands == []


def test_import_any_dispatches(tmp_path):
    path = _v2_file(tmp_path)
    assert import_any(path).months == import_v2(path).months


def test_multi_sku_and_vm_spec(tmp_path):
    path = _v2_file(tmp_path)
    wb = load_workbook(path)
    wb["HW_SKU"].append(["big-128", 128, 0.8])
    wb["HW_Current"].append(["A", "network1", "big-128", 3])
    wb["HW_MoveIn"].append(["A", "network1", "big-128", "2026-09", 5])
    # 新格式不產生 VM_Spec,需手動建立以測試向後相容性
    ws = wb.create_sheet("VM_Spec")
    ws.append(["fab", "bm_group", "product", "month", "vm_size_vcore", "count"])
    ws.append(["A", "network1", "Product Apple", "2026-08", 32, 4])
    wb.save(path)
    pi = import_v2(path)
    assert set(pi.skus) == {"default-64", "big-128"}
    a1 = Pool(fab="A", bm_group="network1")
    assert pi.current_by_sku(a1) == {"default-64": 50, "big-128": 3}
    assert pi.movein_by_sku(a1, "2026-09") == {"default-64": 25, "big-128": 5}
    assert pi.vm_batch(a1, "2026-08") == [(32, 4)]


def test_unknown_sku_reported_and_skipped(tmp_path):
    path = _v2_file(tmp_path)
    wb = load_workbook(path)
    wb["HW_MoveIn"].append(["A", "network1", "no-such-sku", "2026-09", 5])
    wb.save(path)
    pi = import_v2(path)
    errors = [i for i in pi.issues if i.severity == "error"]
    assert any("no-such-sku" in i.message for i in errors)
    # 2026-09 default-64 movein 來自 fixture 既有的 summary Q3=25;僅 no-such-sku 該列被跳過
    assert pi.movein_by_sku(Pool(fab="A", bm_group="network1"), "2026-09") == {"default-64": 25}


def test_vm_spec_only_pool_is_planned(tmp_path):
    """I-1: 只出現在 VM_Spec 的 pool 也必須進入 pi.pools 並被排入推演。"""
    path = _v2_file(tmp_path)
    wb = load_workbook(path)
    # 新格式不產生 VM_Spec,需手動建立以測試向後相容性
    ws = wb.create_sheet("VM_Spec")
    ws.append(["fab", "bm_group", "product", "month", "vm_size_vcore", "count"])
    new_pool = Pool(fab="A", bm_group="network9")
    ws.append(["A", "network9", "Product Only-In-VmSpec", "2026-07", 32, 4])
    wb.save(path)
    pi = import_v2(path)
    assert new_pool in pi.pools
    result = run_check(pi, NaiveSolver(), mode="conservative")
    outcomes = result.for_pool(new_pool)
    assert len(outcomes) == len(pi.months)
    assert any(o.vm_demand == [(32, 4)] for o in outcomes)


def test_non_numeric_count_reported_not_raised(tmp_path):
    """I-5: HW_MoveIn 的 count 非數值時應回報錯誤並以 0 計,而非整個 raise 中斷匯入。"""
    path = _v2_file(tmp_path)
    wb = load_workbook(path)
    wb["HW_MoveIn"].append(["A", "network1", "default-64", "2026-09", "abc"])
    wb.save(path)
    pi = import_v2(path)  # 不應 raise
    errors = [i for i in pi.issues if i.severity == "error" and i.sheet == "HW_MoveIn"]
    assert any("應為數值" in i.message and i.cell for i in errors)
    # 非數值列以 0 計,既有 2026-09 default-64 movein(來自 summary Q3=25)不受影響
    assert pi.movein_by_sku(Pool(fab="A", bm_group="network1"), "2026-09") == {"default-64": 25}


def test_multi_sku_returns_roundtrip(tmp_path):
    """I-3: 同一 product/bm_group 但不同機型的 Return 應各自成列,往返後不互相覆蓋。"""
    skus = {"s1": Sku(name="s1", vcore_per_node=64, usable_ratio=0.8),
            "s2": Sku(name="s2", vcore_per_node=128, usable_ratio=0.8)}
    pool = Pool(fab="A", bm_group="network1")
    pi = PlanInput(
        skus=skus, months=["2026-07"], pools=[pool],
        demands=[], vm_demands=[], moveins=[],
        returns=[
            NodeReturn(pool=pool, product="P1", sku_name="s1", month="2026-07", count=5),
            NodeReturn(pool=pool, product="P1", sku_name="s2", month="2026-07", count=7),
        ],
        currents=[])
    path = tmp_path / "v2.xlsx"
    generate_v2_template(pi, path)
    pi2 = import_v2(path)
    assert pi2.return_by_sku(pool, "2026-07") == {"s1": 5, "s2": 7}


def test_vm_demands_roundtrip(tmp_path):
    """I-4: 新格式不產生 VM_Spec,但可手動新增以測試向後相容性。"""
    skus = {"s1": Sku(name="s1", vcore_per_node=64, usable_ratio=0.8)}
    pool1 = Pool(fab="A", bm_group="network1")
    pool2 = Pool(fab="A", bm_group="network2")
    pi = PlanInput(
        skus=skus, months=["2026-07", "2026-08"], pools=[pool1, pool2],
        demands=[],
        vm_demands=[
            VmSpecDemand(pool=pool1, product="Product X", month="2026-07",
                        vm_size_vcore=32, count=4),
            VmSpecDemand(pool=pool2, product="Product Y", month="2026-08",
                        vm_size_vcore=16, count=10),
        ],
        moveins=[], returns=[], currents=[])
    path = tmp_path / "v2.xlsx"
    generate_v2_template(pi, path)
    # 新格式不產生 VM_Spec,需手動建立以測試向後相容性
    wb = load_workbook(path)
    ws = wb.create_sheet("VM_Spec")
    ws.append(["fab", "bm_group", "product", "month", "vm_size_vcore", "count"])
    ws.append(["A", "network1", "Product X", "2026-07", 32, 4])
    ws.append(["A", "network2", "Product Y", "2026-08", 16, 10])
    wb.save(path)
    pi2 = import_v2(path)
    assert pi2.vm_batch(pool1, "2026-07") == [(32, 4)]
    assert pi2.vm_batch(pool2, "2026-08") == [(16, 10)]
