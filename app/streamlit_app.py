"""容量規劃工具 UI:統整 summary(呼叫 solver 求解實體機需求與剩餘 vcore)。"""
import os
import sys
import tempfile
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import io

from captool.exporter import export_demand_order
from captool.importer import CapacityImportError, import_any
from captool.models import ImportIssue
from captool.solver.horizon_adapter import http_solve_fn, plan_horizon
from captool.solver.procure_adapter import demand_order, demand_order_frames
from captool.summary import capacity_summary
from captool.viewmodel import prometheus_ag_placeholder

st.set_page_config(page_title="容量規劃工具", layout="wide")


def _load(uploaded) -> None:
    st.session_state.pop("horizon", None)
    st.session_state.pop("demand_order", None)
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
    page = st.radio("頁面", ["總表", "執行面需求單", "匯入報告"])
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

_ALL_AGS = ({c.ag for c in plan_input.currents if c.ag}
            | {cap.ag for cap in plan_input.caps}
            | {m.ag for m in plan_input.moveins if m.ag}
            | {r.ag for r in plan_input.returns if r.ag})


def _filter_block(df, fabs, nets, ags):
    """依 row key(fab/network/.../ag)過濾:空選 = 全顯示;AG 過濾只作用於有 AG 維度的列。"""
    keep = []
    for idx in df.index:
        segs = str(idx).split("/")
        if fabs and segs[0] not in fabs:
            continue
        if nets and len(segs) > 1 and segs[1] not in nets:
            continue
        if ags and segs[-1] in _ALL_AGS and segs[-1] not in ags:
            continue
        keep.append(idx)
    return df.loc[keep]


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

    fabs = sorted({p.fab for p in plan_input.pools})
    nets = sorted({p.bm_group for p in plan_input.pools})
    ags = sorted(_ALL_AGS)
    c1, c2, c3 = st.columns(3)
    f_fab = c1.multiselect("Fab", fabs)
    f_net = c2.multiselect("Network", nets)
    f_ag = c3.multiselect("AG", ags)

    for name, df in capacity_summary(plan_input, horizon).items():
        st.subheader(name)
        st.dataframe(_filter_block(df, f_fab, f_net, f_ag),
                     use_container_width=True)

    st.subheader("7. 各 AG 節點現況(Prometheus,假資料)")
    st.caption("未來由 Prometheus 取得每 cluster 已安裝 / cordon 節點數、各 AG 節點總數,"
               "供 expand / delete 分配各 AG。目前為假資料。")
    st.dataframe(prometheus_ag_placeholder(), use_container_width=True,
                 hide_index=True)

elif page == "執行面需求單":
    st.header("執行面需求單(單月已確定需求 → 逐需求實體機)")
    st.caption("月底整理「下個月已確定需求」時用。走 solver 單期 procure 端點,"
               "回傳逐需求的真實落點(哪個 SKU、幾台、其中幾台新採購),非估算分攤。"
               "in-stock 起點 = 現況 + 到目標月的進機 − 退還。")
    month = st.selectbox("目標月(下個月已確定需求)", plan_input.months)
    procure_url = solver_url.replace("capacity/plan", "capacity/procure")
    st.caption(f"procure 端點:{procure_url}")
    if st.button("產生需求單"):
        try:
            rows, buys = demand_order(plan_input, month,
                                      http_solve_fn(procure_url))
            st.session_state["demand_order"] = (month, rows, buys)
        except Exception as e:  # noqa: BLE001 — 對外呼叫,任何錯都回報
            st.session_state.pop("demand_order", None)
            st.error(f"呼叫 procure 失敗:{e}(確認 server 有起、且該月有需求)")
    do = st.session_state.get("demand_order")
    if do and do[0] == month:
        _, rows, buys = do
        orders_df, buy_df = demand_order_frames(rows, buys)
        if orders_df.empty:
            st.warning(f"{month} 沒有需求列。")
        else:
            st.subheader("需求單(每列一個需求)")
            st.caption("「建議實體機」台數含共用(一台機可同住多需求的 VM);"
                       "真正下單以下方去重清單為準。")
            st.dataframe(orders_df, use_container_width=True, hide_index=True)
            st.subheader("本月實際採購清單(去重,下單依據)")
            st.dataframe(buy_df, use_container_width=True, hide_index=True)
            buf = io.BytesIO()
            export_demand_order(orders_df, buy_df, month, buf)
            st.download_button(
                "下載需求單 (xlsx)", buf.getvalue(),
                file_name=f"demand_order_{month}.xlsx",
                mime="application/vnd.openxmlformats-officedocument."
                     "spreadsheetml.sheet")
    else:
        st.info("選擇目標月後按「產生需求單」。")
