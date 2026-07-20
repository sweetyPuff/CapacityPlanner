from pathlib import Path

from openpyxl import load_workbook

from captool.exporter import generate_v2_template
from captool.importer import import_any, import_legacy, import_v2
from captool.models import Pool, Sku

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
    ws = wb["VM_Spec"]
    ws.delete_rows(2)  # 移除示範列
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
