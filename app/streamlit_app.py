"""容量規劃工具 UI。邏輯全在 captool,本檔僅做狀態管理與呈現。"""
import io
import os
import sys
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from captool.exporter import export_summary, generate_v2_template
from captool.importer import CapacityImportError, import_any
from captool.models import ImportIssue, Pool
from captool.planner import run_check, run_suggest
from captool.solver.naive import NaiveSolver
from captool.viewmodel import (apply_movein_edits, movein_frame,
                               overview_frame, placements_frame,
                               prometheus_ag_placeholder, summary_frames)

st.set_page_config(page_title="容量規劃工具", layout="wide")

MODE_LABELS = {"conservative": "保守(跨月零頭作廢)", "gap_fill": "填縫(零頭跨月可用)"}


def _load(uploaded) -> None:
    st.session_state.pop("suggest", None)
    tmp_file = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp_file.write(uploaded.getvalue())
    tmp_path = tmp_file.name
    tmp_file.close()
    try:
        st.session_state["plan_input"] = import_any(tmp_path)
        st.session_state["source_name"] = uploaded.name
        st.session_state["source_file_id"] = uploaded.file_id
    except CapacityImportError as e:
        st.session_state.pop("plan_input", None)
        st.session_state["fatal_issues"] = e.issues
    except Exception:
        st.session_state.pop("plan_input", None)
        st.session_state["fatal_issues"] = [ImportIssue("error", "", "", "無法讀取檔案:請確認為有效的 Excel (.xlsx) 檔案")]
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


with st.sidebar:
    st.title("容量規劃工具")
    uploaded = st.file_uploader("上傳 Excel(舊格式或 v2)", type=["xlsx"])
    if uploaded is not None and st.session_state.get("source_file_id") != uploaded.file_id:
        _load(uploaded)
    mode = st.radio("推演模式", list(MODE_LABELS), format_func=MODE_LABELS.get)
    page = st.radio("頁面", ["匯入報告", "總表", "總覽", "驗證模式", "回推模式",
                            "配置明細"])

if "fatal_issues" in st.session_state and "plan_input" not in st.session_state:
    st.error("匯入失敗:")
    for issue in st.session_state["fatal_issues"]:
        st.write(f"- [{issue.sheet}!{issue.cell}] {issue.message}")
    st.stop()
if "plan_input" not in st.session_state:
    st.info("請先於左側上傳容量規劃 Excel。")
    st.stop()

plan_input = st.session_state["plan_input"]
solver = NaiveSolver()
result = run_check(plan_input, solver, mode=mode)


def _pool_select(key: str) -> Pool:
    return st.selectbox("Pool", plan_input.pools,
                        format_func=lambda p: p.label, key=key)


def _style_gap(df: pd.DataFrame):
    if "狀態" not in df.columns:
        return df
    return df.style.apply(
        lambda row: ["background-color: #ffc7ce"] * len(row)
        if row["狀態"] == "缺口" else [""] * len(row), axis=1)


if page == "匯入報告":
    st.header("匯入報告")
    st.write(f"檔案:{st.session_state['source_name']}|機型:{len(plan_input.skus)}"
             f"|pool:{len(plan_input.pools)}|月份:{plan_input.months[0]} ~ "
             f"{plan_input.months[-1]}")
    if not plan_input.issues:
        st.success("未發現問題。")
    for issue in plan_input.issues:
        text = f"[{issue.sheet}!{issue.cell}] {issue.message}"
        if issue.severity == "error":
            st.error(text)
        else:
            st.warning(text)

elif page == "總表":
    st.header("總表(所有 pool × 月份總攬)")
    frames = summary_frames(result)
    gaps = result.gap_months
    if gaps:
        st.error(f"共 {len(gaps)} 個缺口(pool×月),見下方紅底格。")
    else:
        st.success("所有 pool 各月皆可行。")

    st.subheader("可行性總攬")
    status = frames["狀態"]
    st.dataframe(
        status.style.map(
            lambda v: "background-color: #ffc7ce" if v == "缺口" else ""),
        use_container_width=True)

    for name, df in frames.items():
        if name == "狀態":
            continue
        st.subheader(name)
        st.dataframe(df, use_container_width=True)

    st.subheader("各 AG 節點現況(Prometheus,假資料)")
    st.caption("未來功能:由 Prometheus 取得每 cluster 已安裝 / cordon 節點數、"
               "各 AG 節點總數,供 expand / delete 時分配各 AG 的處理量。目前為假資料佔位。")
    st.dataframe(prometheus_ag_placeholder(), use_container_width=True,
                 hide_index=True)

    buf = io.BytesIO()
    export_summary(plan_input, result, buf)
    st.download_button("下載 summary Excel", buf.getvalue(), "summary.xlsx")

elif page == "總覽":
    st.header("總覽")
    gaps = result.gap_months
    if gaps:
        st.error("缺口月份:" + "、".join(
            f"{o.pool.label} {o.month}(缺 {o.shortfall_vcore:.0f} vcore)"
            for o in gaps))
    else:
        st.success("所有 pool 各月皆可行。")
    for pool in plan_input.pools:
        st.subheader(pool.label)
        df = overview_frame(result, pool)
        st.dataframe(_style_gap(df), use_container_width=True)
        st.line_chart(df[["剩餘可售 vcore"]])
    buf = io.BytesIO()
    export_summary(plan_input, result, buf)
    st.download_button("下載 summary Excel", buf.getvalue(), "summary.xlsx")
    buf2 = io.BytesIO()
    generate_v2_template(plan_input, buf2)
    st.download_button("下載 v2 範本(帶入目前資料)", buf2.getvalue(),
                       "capacity_v2_template.xlsx")

elif page == "驗證模式":
    st.header("驗證模式:編修進機計畫,即時重算")
    st.caption("編修下表後按「套用」;month 格式 YYYY-MM,sku 須存在於機型目錄。")
    edited = st.data_editor(movein_frame(plan_input), num_rows="dynamic",
                            use_container_width=True)
    if st.button("套用進機計畫"):
        try:
            st.session_state["plan_input"] = apply_movein_edits(plan_input, edited)
            st.rerun()
        except ValueError as e:
            st.error(str(e))
    pool = _pool_select("verify_pool")
    st.dataframe(_style_gap(overview_frame(result, pool)), use_container_width=True)

elif page == "回推模式":
    st.header("回推模式:計算建議進機計畫")
    chosen = st.multiselect("機型 catalog(可複選)", list(plan_input.skus),
                            default=list(plan_input.skus))
    if st.button("計算建議進機"):
        catalog = [plan_input.skus[n] for n in chosen]
        if not catalog:
            st.error("請至少選擇一個機型。")
        else:
            s_result, suggested = run_suggest(plan_input, solver, catalog, mode=mode)
            st.session_state["suggest"] = (s_result, suggested)
    if "suggest" in st.session_state:
        s_result, suggested = st.session_state["suggest"]
        if s_result.gap_months:
            st.error("即使補機仍有缺口(存在單台裝不下的 VM?):" + "、".join(
                f"{o.pool.label} {o.month}" for o in s_result.gap_months))
        df = pd.DataFrame(
            [{"fab": m.pool.fab, "bm_group": m.pool.bm_group, "sku": m.sku_name,
              "month": m.month, "count": m.count} for m in suggested],
            columns=["fab", "bm_group", "sku", "month", "count"])
        st.subheader("建議進機計畫")
        st.dataframe(df, use_container_width=True)
        buf = io.BytesIO()
        export_summary(plan_input, s_result, buf)
        st.download_button("下載回推結果 Excel", buf.getvalue(), "suggest_plan.xlsx")

elif page == "配置明細":
    st.header("配置明細(可行性證明/執行參考)")
    pool = _pool_select("placement_pool")
    st.dataframe(placements_frame(result, pool), use_container_width=True)
    for o in result.for_pool(pool):
        if not o.feasible:
            st.error(f"{o.month}:缺 {o.shortfall_vcore:.0f} vcore"
                     + (f";單台裝不下的 VM:{o.blocked_vms}" if o.blocked_vms else ""))
