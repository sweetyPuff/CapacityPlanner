# Solver 整合 + Control Plane(new build)設計

日期:2026-07-25
狀態:設計草案,待 review
分支:`integrate-solver`
關聯:延伸自 `2026-07-22-v2-demand-format-evolution-design.md`(co-residency 模型)、
`cpsat-adapter-integration.md`。同事 solver develop 版原始檔在 `solver-develop/`(獨立專案,只讀)。

## 1. 背景與方向(路線 A)

同事的 solver(develop 版)已內建完整容量規劃引擎:
- `solve_capacity_plan`(單廠單期採購解)、`solve_capacity_horizon`(多期逐月 roll-forward、
  節點 sticky、買/committed 變下月在庫、槽位上限遞減);
- `split-and-solve`(給 vcore 預算,自己切成幾顆什麼規格的 VM 再擺放);
- 角色感知(node_role: master/worker/infra/l4lb)、C3 反親和(masters 一台一 AG)、
  C4 max-per-BM、C5 failover。

**決定:路線 A —— 我們的工具收斂成「Excel I/O + UI 層」**,推演/回推改呼叫 solver 的
`solve_capacity_horizon` / `solve_capacity_plan`,不再自行維護 planner / NaiveSolver 的最佳化。
工具的價值集中在:與 PM/硬體 team 的 Excel 溝通、輸入模型、UI 呈現(尤其總表 summary)。

版本維持 **v2**(格式在 v2 內演進,不另立 v3)。

## 2. Control Plane / New Build(菜單)設計

### 決策
- **甲**:菜單只管非 worker 的固定角色(control plane);worker 成長一律走原本的 vcore 需求列,
  兩者分開標示、各自需求。
- **Control plane 是一種獨立 tenant**,與 worker 分開;tenant 內各角色(master / infra(if) /
  l4lb / learner)**可共住同一批 BM**;另有元件需**獨佔 BM**(F5 / Bastion / HA / VT)。
- New build = PM 另外新增一列(如 `ai / new build / c2 / 菜單B`),工具依菜單**自動展開**成
  各角色的需求(role + 台數 + vcore + 共居),掛上該 cluster 的 tenant。
- 菜單為 **2~N 種預設 + 可客製**(該列覆寫台數/規格)。

### Control plane 菜單(使用者提供;vcore 為佔位假值,待填實際規格)

共用底座(A–E 皆含,可共住同一 tenant 的 BM):`master×5、infra×5、l4lb×3、learner×5`

| 菜單 | 共住底座(control-plane tenant) | 額外(各自獨佔 BM) |
|---|---|---|
| A | master×5, infra×5, l4lb×3, learner×5 | — |
| B | 同上 | F5×3(獨佔) |
| C | 同上 | Bastion×1(獨佔) |
| D | 同上 | F5×3、Bastion×1(獨佔) |
| E | 同上 | F5×3、Bastion×1、HA×3、VT×3(獨佔) |

> ⚠️ 每角色的 **vcore 規格尚未提供**,本文件與 sample 先用佔位假值
> (master=16, infra=8, l4lb=8, learner=32, F5=32, Bastion=8, HA=16, VT=16),**待使用者以實際規格取代**。

### 共居語意(延續 v2 co-residency 模型:free / exclusive / 群組)
- 共住底座角色 → co_residency = **該 cluster 的 tenant 群組**(如 `cp-c2`);彼此可共住,
  但與 worker、與其他 cluster 的 control plane 皆分開(群組即分割)。
- 獨佔元件(F5 等) → co_residency = **exclusive**(整台 BM 專用)。

## 3. Excel 格式(v2 內新增兩張長表)

沿用既有 v2:fab 頁 worker 需求列(A/B/C/D/E 欄 + 月份)、HW_SKU/HW_Current/HW_MoveIn。**新增**:

**`Cluster_Menu`(菜單目錄)** —— 欄:`menu, role, count, co_residency, vm_vcore`
- co_residency:`shared`(control-plane tenant)/ `exclusive`(獨佔 BM)。
- 每個菜單列出其全部元件(A 的底座在 B–E 重複列出,直白好讀)。

**`New_Build`(本期新建清單)** —— 欄:`fab, network, cluster, month, menu`
- 一列 = 一次 new build;工具依 `menu` 自 `Cluster_Menu` 展開成該 cluster 的角色需求。
- 客製:允許在此列以額外欄覆寫(未來擴充;v1 先支援選單引用)。

worker 需求維持在 fab 頁需求列;new build 與 worker **分開標示**(決策甲)。

### 拓撲 / AG(失效域)—— solver spread/HA 的必要輸入

masters 一台一 AG 分散、失效域隔離、爆炸半徑,都靠 baremetal 的 **AG** 資訊。沒有 AG,solver
只能回答「總容量夠不夠」,無法驗證「擺得安全」。決策:**失效域維度 = AG**;**在庫機器逐台帶 AG**(細粒度)。

- **`HW_Current` 增一欄 `ag`(放在最後)** —— schema:`fab, bm_group, sku, count, ag`,即
  「某 pool 某機型在某 AG 有幾台」。ag 放最後,現行 importer 讀前四欄自動忽略、整合階段再讀(向前相容)。
- **`HW_Caps`(新表,AG 槽位上限)** —— 欄:`fab, network, ag, max_bm`,即每個 AG 還能放幾台(限制採購)。
  控制 new build 可分散的 AG 數必須 ≥ 菜單最大角色台數(菜單 A master×5 → 該 fab×network 需 ≥5 個 AG)。
- 買的新機器 solver 會給獨立虛擬 rack、可自由分到不同 AG(受 `HW_Caps` 上限);在庫機器用其實際 AG。

## 4. 與 solver 互動(重點章節)

工具把 Excel 輸入翻譯成 solver develop 版的 request 模型,呼叫 `solve_capacity_horizon`,
再把 `CapacityReport` 映射回 UI/summary。以下為欄位級對應。

### 4.1 呼叫方式
- **優先 in-process**:`from app.capacity_planner import solve_capacity_horizon`(需 `solver-develop`
  在 sys.path、且環境裝 `ortools` + `pydantic`)。
- 或 **HTTP sidecar**:目前 README 只暴露 `/v1/placement/solve`、`/split-and-solve`;
  capacity horizon 尚無 HTTP route → 需請同事加一個 `/v1/capacity/horizon` 端點,或走 in-process。
  (待與同事確認 —— 見 §6。)
- 沿用 adapter 的依賴注入精神:`solve_fn` 可切 in-process / HTTP / 測試假後端,工具本體不硬依賴 ortools。

### 4.2 單位對應(我方 → solver)

| 我方概念 | solver 概念 | 備註 |
|---|---|---|
| pool = fab × bm_group | fab(`fab_topology_dimension`)× **network**(BGP zone) | 我方 `bm_group`(network1/2)→ solver `network` |
| SKU(機型) | `procurement_types`(BaremetalType: type_id, capacity, fab) | HW_SKU → 型錄 |
| HW_Current(per pool per SKU per **AG** 台數) | `in_stock`:一台一個 `Baremetal`(topology 帶 fab、**ag**、`network`、capacity) | ag 來自 HW_Current 的 ag 欄 |
| HW_Caps(每 AG 槽位上限) | `procurement_caps`(per (bucket=AG, network) 的 max_bm) | 限制每 AG 可買台數 |
| HW_MoveIn(進機,已定案) | `committed_stock`(近零成本層,帶 type_id/fab/network/bucket/count) | 已決定要進的機器 |
| 失效域 = AG | `config.procurement_spread_dimension = "ag"` | masters 一台一 AG |
| 月份 | `demand_book` 的 `period`;每 (fab, period) 一組 requirements | horizon 逐月 |
| 單維 vcore | `Resources.cpu_cores`(其餘 mem/disk/gpu 先 0) | 維持單維(§v2 決策) |

### 4.3 需求對應(worker 與 control plane)

- **Worker 需求列(vcore 預算)** → `ResourceRequirement`:
  - `node_role=worker`、`total_resources={cpu_cores: 該月 vcore}`、`network=bm_group`、
    `cluster_id`=product(或 worker 所屬 cluster)、`vm_specs`=允許的 worker VM 規格(來自 v2 VM 明細/config)。
  - solver 用 **split-and-solve** 自動把 vcore 預算切成 worker VM 再擺放。
- **New build → 菜單展開** → 每角色一筆 `ResourceRequirement`:
  - `node_role`=master/infra/l4lb/learner/…、`cluster_id`=新 cluster(如 c2)、
    `vm_specs=[該角色 vcore 規格]`、`min_total_vms=max_total_vms=菜單台數`(鎖定台數)、`network`。
  - **共住底座**:同 cluster 的這些角色共享 tenant → 交由 solver 允許同 cluster 混住;
    為與 worker 分離,control-plane 角色走**獨立 candidate/BM 池或 `allowed_bm_types`**(見下)。
  - **獨佔元件(F5/Bastion/HA/VT)**:需整台 BM 專用 → 以**專用 BM type + `allowed_bm_types`
    候選限定**(該角色只能落在其專用型的 BM),或 `max_per_bm=1` + 專池達成。
  - masters 一台一 AG 由 solver 的 `auto_generate_anti_affinity`(依 cluster_id+role 分組)自動處理。

### 4.4 共居 / tenant 對應

| 我方 co-residency | solver 機制 |
|---|---|
| worker(各自 free/群組/exclusive,v2 既有) | worker requirement 的 candidate / max_per_bm / 反親和 |
| control-plane tenant(cluster 內共住) | 同 cluster control-plane 角色共用一組 candidate BM 池;與 worker BM 分離(candidate scoping 或專用 type) |
| exclusive 元件(獨佔 BM) | 專用 BM type + `allowed_bm_types`,或 `max_per_bm=1` 且不與他者共候選 |

> ⚠️ solver 是否已能「乾淨地」表達「這群角色只能彼此共住、且與 worker 分離」需與同事確認 —— 她有
> `max_per_bm_rules`、`anti_affinity_rules`、`allowed_bm_types`、candidate filtering,理論上可組出,
> 但可能需要她補一個「co-tenancy / 專池」的表達或我方以 candidate scoping 達成(見 §6)。

### 4.5 輸出對應(CapacityReport → 我方 UI)

`solve_capacity_horizon` 回傳 `CapacityReport`:
- `by_fab_period: list[PeriodFabReport]` —— 每 (fab, period):`success`、`node_adds_total`、
  `bm_procurement_total`、`committed_bm_used`、`in_stock_bm_used`、`procurement`(買哪型幾台)、
  `split_decisions`(切了幾顆什麼 VM)、`shortfalls`(成因)、`nominal_available`、
  `remaining_node_slots`、`cells: list[BucketMonthCell]`(per (bucket, network) 的在庫 total/used/available)。
- `budget_view: list[BudgetRow]`(fab/bucket/network/period/type/台數)—— 即採購/進機計畫。
- `totals` —— 總量。

映射到我方**總表(summary)**:
- 可行性矩陣 ← `PeriodFabReport.success` / `shortfalls`
- 需求 ← 我方輸入;新啟用/進機/在庫 ← `node_adds_total` / `procurement` / `cells.in_stock_available`
- 缺口成因 ← `shortfalls`(space / capacity / anti_affinity / input_error)
- 採購計畫(回推)← `budget_view` / `procurement`
- control plane 展開結果 ← `split_decisions`(哪些角色各幾顆、落在哪)

## 5. 架構影響

- `captool/planner.py`(自寫逐月推演)+ `captool/solver/naive.py` → **退居備援**(離線/無 ortools 時可用),
  主路徑改呼叫 solver。保留 `AllocationSolver` 契約精神,新增 `captool/solver/horizon_adapter.py`
  負責 Excel 模型 ↔ solver request/response 翻譯 + `solve_fn` 注入。
- `captool/models.py` 增:`Menu` / `MenuComponent`(role, count, co_residency, vm_vcore)、
  `NewBuild`(fab, network, cluster, month, menu);`PlanInput` 帶 menus + new_builds。
- `importer.py` 解析 `Cluster_Menu` / `New_Build`;`exporter.py` 產生這兩張表(含菜單假資料)。
- UI:總表沿用;可加「new build 展開檢視」與採購計畫(budget_view)。

## 6. 待確認 / 待同事協調

1. **capacity horizon 的呼叫介面**:in-process import 還是請同事加 HTTP 端點?(§4.1)
2. **co-tenancy 表達**:control-plane tenant(角色群共住、與 worker 分離)在她的模型怎麼最乾淨地表達 —— candidate scoping / 專用 BM type / 是否需她補一個 co-residency 概念。(§4.4)
3. **各角色 vcore 實際規格**(master/infra/l4lb/learner/F5/Bastion/HA/VT)—— 目前用假值。
4. **ortools 安裝**:我方環境要能裝 ortools + pydantic(in-process 路徑);或改走 HTTP。
5. **worker VM 規格來源**:worker split 的 `vm_specs`(允許的 worker VM 尺寸)從哪來(config / v2 VM 明細)。

## 7. 範圍與順序(建議)

1. 本輪:spec + 新 Excel sample(菜單假資料)+ importer 略過新表不報錯。(整合尚未接)
2. 下一步:`models` 加 Menu/NewBuild → `importer`/`exporter` 解析/產生 → `horizon_adapter`
   翻譯 + 假後端測 → 裝 ortools/接 in-process 實跑 → 總表改吃 CapacityReport。
3. 與同事對 §6 的介面/co-tenancy/規格。
