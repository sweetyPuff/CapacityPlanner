# CP-SAT Solver 整合說明(cpsat_adapter)

日期:2026-07-21

把同事的 CP-SAT VM placement solver 接進容量規劃工具的 adapter 設計與接線方式。

## 決策(第一版)

- **單維 vcore**:machine 可用 vcore →(`floor` 取整,保守不超賣)baremetal 的
  `cpu_cores`;VM size → `vm.demand.cpu_cores`。mem/disk/gpu 留 0。
- **只有原子 VM 才呼叫 CP-SAT**:純液體 vcore 的格子用內建填縫(除法),不碰 solver。
  多數粗粒度 product 是液體,只有需要 VM 明細的 product 觸發 solver → 呼叫次數最小化。
- **suggest(採購回推)留在我方**:她的 solver 只往既有機器上放,無採購建議概念;
  沿用共用的 `greedy_suggest`(逐輪補機直到可行,內部呼叫 adapter 的 check)。

## 流程:為什麼是「每 pool 每月」呼叫,而非一次大解

planner 逐月時間推進,每 pool 每月呼叫一次 `check`。這是問題本質,不是浪費:

- **pool 之間**硬隔離 → 本來就是獨立問題 → 分開解且可平行。
- **月份之間**狀態遞延(庫存 / 破碎 / backlog),同 pool 內須循序 —— 這才是
  「不可搬遷」的正確語意。一次大解會弄丟缺口出現的月份、且默認可自由重排(違反不可搬遷)。
- 實際呼叫 CP-SAT 的只有「含原子 VM 的格子」,是少數;每次問題小(幾十 VM / 機),
  秒級可解,pool 間可平行。規劃情境用短 timeout + `allow_partial_placement`,只要可行性。

## 接線

`CpSatAdapter(solve_fn, max_solve_time_seconds=3.0)` 實作 `AllocationSolver`,
`solve_fn` 依賴注入(接受 PlacementRequest dict、回傳 PlacementResult dict):

```python
from captool.solver.cpsat_adapter import CpSatAdapter, http_solve_fn, in_process_solve_fn

# 生產(她的 sidecar):
solver = CpSatAdapter(http_solve_fn("http://localhost:50051/v1/placement/solve"))
# 或 in-process(需 ortools + pydantic + 她的 app/ 在 sys.path):
solver = CpSatAdapter(in_process_solve_fn())
```

在 `app/streamlit_app.py` 把 `solver = NaiveSolver()` 換成上面任一即可,planner / UI /
匯入匯出**完全不動**。

## 上線前提(待同事修復)

她的 solver 目前無法實際執行,in-process / HTTP 路徑都會失敗,需先修:

1. `app/server.py`:`_SWAGGER_STATIC_DIR` 未定義 → import 即 `NameError`,sidecar 起不來。
2. `app/solver.py` 的 `solve()`:引用未定義的 `needs_two_phase` / `soft_terms`,且缺
   成功路徑的 `_extract_solution` 呼叫 —— 能跑的邏輯在 `_solve_single` /
   `_solve_two_phase` / `_run_solver`,但 `solve()` 沒接上。

在此之前,adapter 我方邏輯已用假後端(模擬她的 partial-placement JSON 契約)測到綠燈
(`tests/test_cpsat_adapter.py`),她一修好即可切換。

## 後續可考慮

- 升級成多維(mem/disk):把 `Resources` 帶進需求模型,adapter 映射多維而非只有 cpu_cores。
- 利用她的 slot-score(t-shirt sizes)目標項來直接處理「零頭破碎」的最小化。
- pool 間平行呼叫(目前循序)。
