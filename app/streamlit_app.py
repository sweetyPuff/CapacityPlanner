"""容量規劃工具 UI:統整 summary(呼叫 solver 求解實體機需求與剩餘 vcore)。"""
import os
import sys
import tempfile
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from captool.importer import CapacityImportError, import_any
from captool.models import ImportIssue
from captool.solver.horizon_adapter import http_solve_fn, plan_horizon
from captool.summary import capacity_summary
from captool.viewmodel import prometheus_ag_placeholder

st.set_page_config(page_title="容量規劃工具", layout="wide")


def _load(uploaded) -> None:
    st.session_state.pop("horizon", None)
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
    except Exception:  # noqa: BLE001
        st.session_state.pop("plan_input", None)
        st.session_state["fatal_issues"] = [ImportIssue(
            "error", "", "", "無法讀取檔案:請確認為有效的 Excel (.xlsx) 檔案")]
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


with st.sidebar:
    st.title("容量規劃工具")
    uploaded = st.file_uploader("上傳 Excel(v2)", type=["xlsx"])
    if uploaded is not None and st.session_state.get("source_file_id") != uploaded.file_id:
        _load(uploaded)
    page = st.radio("頁面", ["總表", "匯入報告"])
    solver_url = st.text_input("Solver 端點",
                               value="http://localhost:50051/v1/capacity/plan")

if "fatal_issues" in st.session_state and "plan_input" not in st.session_state:
    st.error("匯入失敗:")
    for issue in st.session_state["fatal_issues"]:
        st.write(f"- [{issue.sheet}!{issue.cell}] {issue.message}")
    st.stop()
if "plan_input" not in st.session_state:
    st.info("請先於左側上傳容量規劃 Excel。")
    st.stop()

plan_input = st.session_state["plan_input"]

if page == "匯入報告":
    st.header("匯入報告")
    st.write(f"檔案:{st.session_state['source_name']}|機型:{len(plan_input.skus)}"
             f"|pool:{len(plan_input.pools)}|月份:{plan_input.months[0]} ~ "
             f"{plan_input.months[-1]}|new build:{len(plan_input.new_builds)}")
    if not plan_input.issues:
        st.success("未發現問題。")
    for issue in plan_input.issues:
        text = f"[{issue.sheet}!{issue.cell}] {issue.message}"
        (st.error if issue.severity == "error" else st.warning)(text)

elif page == "總表":
    st.header("總表(統整 summary)")
    st.caption("需求 / 進機 / 退還 由輸入計算;實體機需求與剩餘可用 vcore 由 solver 求解後補上。"
               "維度:多機型 × 多 AG。")
    if st.button("呼叫 solver 求解(實體機需求 + 剩餘 vcore)"):
        try:
            st.session_state["horizon"] = plan_horizon(
                plan_input, http_solve_fn(solver_url))
        except Exception as e:  # noqa: BLE001 — 對外呼叫,任何錯都回報
            st.session_state.pop("horizon", None)
            st.error(f"呼叫 solver 失敗:{e}(確認 server 有起在該 URL)")
    horizon = st.session_state.get("horizon")
    if horizon is None:
        st.info("尚未求解 —— 下方僅顯示輸入面(需求 / 進機 / 退還)。"
                "按上方按鈕求解後補上「實體機需求」與「剩餘可用 vcore」。")
    elif horizon.gaps:
        st.warning(f"solver 回報 {len(horizon.gaps)} 個 fab×月缺口。")
    else:
        st.success("solver 求解:所有 fab×月皆可行。")

    for name, df in capacity_summary(plan_input, horizon).items():
        st.subheader(name)
        st.dataframe(df, use_container_width=True)

    st.subheader("6. 各 AG 節點現況(Prometheus,假資料)")
    st.caption("未來由 Prometheus 取得每 cluster 已安裝 / cordon 節點數、各 AG 節點總數,"
               "供 expand / delete 分配各 AG。目前為假資料。")
    st.dataframe(prometheus_ag_placeholder(), use_container_width=True,
                 hide_index=True)
