from pathlib import Path

from captool.importer import DEFAULT_SKU, detect_format, import_legacy
from captool.models import Pool

FIXTURE = Path(__file__).parent / "fixtures" / "sample_legacy.xlsx"


def test_detect_format():
    assert detect_format(FIXTURE) == "legacy"


def test_pools_and_months():
    pi = import_legacy(FIXTURE)
    assert pi.months == ["2026-07", "2026-08", "2026-09",
                         "2026-10", "2026-11", "2026-12"]
    assert Pool(fab="A", bm_group="network1") in pi.pools
    assert len(pi.pools) == 6


def test_single_default_sku():
    pi = import_legacy(FIXTURE)
    assert set(pi.skus) == {DEFAULT_SKU.name}


def test_demand_values():
    pi = import_legacy(FIXTURE)
    # A tab: network1 = Apple 512 + Banana 358 = 870(2026-07)
    assert pi.demand_vcore(Pool(fab="A", bm_group="network1"), "2026-07") == 870
    # C tab: network2 = 331 + 1024 = 1355
    assert pi.demand_vcore(Pool(fab="C", bm_group="network2"), "2026-07") == 1355


def test_returns_and_current_and_movein():
    pi = import_legacy(FIXTURE)
    a1 = Pool(fab="A", bm_group="network1")
    assert pi.return_by_sku(a1, "2026-08") == {DEFAULT_SKU.name: 12}
    assert pi.current_by_sku(a1) == {DEFAULT_SKU.name: 50}
    assert pi.movein_by_sku(a1, "2026-08") == {DEFAULT_SKU.name: 35}


def test_issues_reported():
    pi = import_legacy(FIXTURE)
    messages = " ".join(i.message for i in pi.issues)
    assert "20267'" in messages                 # 月份 typo 警告
    formula_issues = [i for i in pi.issues if "公式" in i.message]
    assert any(i.cell == "T7" for i in formula_issues)   # T7 錯位公式
    assert all(i.severity == "warning" for i in pi.issues)
