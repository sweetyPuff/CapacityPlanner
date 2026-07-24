from pathlib import Path

from captool.importer import import_legacy
from captool.models import Pool
from captool.planner import run_check
from captool.solver.naive import NaiveSolver
from captool.viewmodel import (apply_movein_edits, movein_frame,
                               overview_frame, placements_frame,
                               prometheus_ag_placeholder, summary_frames)

FIXTURE = Path(__file__).parent / "fixtures" / "sample_legacy.xlsx"
B1 = Pool(fab="B", bm_group="network1")


def _result():
    pi = import_legacy(FIXTURE)
    return pi, run_check(pi, NaiveSolver(), mode="conservative")


def test_overview_frame_marks_gap():
    _, result = _result()
    df = overview_frame(result, B1)
    assert list(df.index) == ["2026-07", "2026-08", "2026-09",
                              "2026-10", "2026-11", "2026-12"]
    assert df.loc["2026-07", "狀態"] == "缺口"       # B/network1 期初 0 台
    assert df.loc["2026-07", "缺口 vcore"] > 0


def test_placements_frame():
    _, result = _result()
    df = placements_frame(result, Pool(fab="A", bm_group="network1"))
    assert set(df.columns) == {"月份", "機型", "新啟用台數", "建議採購台數"}
    assert (df["新啟用台數"] > 0).any()


def test_movein_roundtrip():
    pi, _ = _result()
    df = movein_frame(pi)
    df.loc[len(df)] = ["A", "network1", "default-64", "2026-10", 7]
    updated = apply_movein_edits(pi, df)
    assert updated.movein_by_sku(Pool(fab="A", bm_group="network1"),
                                 "2026-10") == {"default-64": 7}
    # 原物件不受影響
    assert pi.movein_by_sku(Pool(fab="A", bm_group="network1"), "2026-10") == {}


def test_summary_frames_shape_and_status():
    _, result = _result()
    frames = summary_frames(result)
    assert "狀態" in frames and "需求 vcore" in frames
    status = frames["狀態"]
    # 列 = 6 個 pool,欄 = 6 個月份
    assert status.shape == (6, 6)
    assert "B/network1" in status.index
    assert list(status.columns) == ["2026-07", "2026-08", "2026-09",
                                    "2026-10", "2026-11", "2026-12"]
    # B/network1 期初 0 台 → 2026-07 缺口
    assert status.loc["B/network1", "2026-07"] == "缺口"
    # 需求 vcore 對得上 overview
    assert frames["需求 vcore"].loc["A/network1", "2026-07"] == 870


def test_summary_frames_excel_compat_has_instock():
    pi = import_legacy(FIXTURE)
    result = run_check(pi, NaiveSolver(), mode="excel_compat")
    frames = summary_frames(result)
    assert "in-stock(台)" in frames
    assert "月末剩餘可售 vcore" not in frames


def test_prometheus_ag_placeholder_columns():
    df = prometheus_ag_placeholder()
    assert list(df.columns) == ["cluster", "AG", "已安裝節點數",
                                "cordon 節點數", "AG 節點總數"]
    assert len(df) > 0


def test_apply_movein_edits_rejects_unknown_sku():
    import pytest
    pi, _ = _result()
    df = movein_frame(pi)
    df.loc[len(df)] = ["A", "network1", "no-such", "2026-10", 7]
    with pytest.raises(ValueError):
        apply_movein_edits(pi, df)
