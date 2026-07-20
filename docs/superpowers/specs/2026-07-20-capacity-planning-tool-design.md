# 容量規劃工具(Capacity Planning Tool)設計文件

日期:2026-07-20
狀態:已與需求方(platform 工程師)逐段確認

## 1. 背景與問題

私有雲 platform 團隊目前以一份 Excel 與 PM、硬體團隊溝通容量規劃:

- 三個廠區(A、B、C)各一個 tab:各 product × BM Group(network1/network2)每月新增的
  vcore 需求、每月可退回的 node 台數
- summary tab:每列為 廠區 × BM Group,橫向展開各月的需求、換算後的機器需求
  (`vcore ÷ (64 × 0.8)`)、進機計畫(手填)、退回、逐月遞推的 in-stock

此模式在**單一硬體機型**時堪用。未來將引入**多種硬體機型**,「vcore → 台數」不再是
一條除法,而是裝箱(bin packing)配置問題,Excel 公式無法承載。同事正在開發名為
solver 的配置引擎(給定機器群與需求,計算最適配置),本工具需與其整合,並在 solver
成熟前以內建演算法先行。

## 2. 目標與非目標

**目標**

1. 讀入現有 Excel(舊格式)即可運作,不強迫 PM 改變填寫習慣
2. 支援多機型:庫存、進機、退回皆帶機型;需求配置交由 solver 介面計算
3. 回答兩類問題:
   - **驗證**:給定進機時程,各 pool 各月是否裝得下?缺口在哪?
   - **回推**:給定需求,各月最少應進哪些機型、各幾台?
4. 呈現 solver 的配置結果(可行性證明 + 執行參考),並支援每月以實況重新校準
5. 定義清楚的 solver 介面契約,作為與 solver 開發者協作的需求文件

**非目標(v1 不做)**

- PM 線上填寫需求(v1 先由本人使用,架構預留擴展)
- 帳號權限系統
- 已上線機器零頭剩餘容量的納入(保守忽略,視為安全邊際)
- 跨 pool 的機器調度建議(pool 為硬隔離)
- 成本/電力/機架空間等採購最佳化維度

## 3. 已確認的領域語意

| 項目 | 語意 |
|---|---|
| 需求數字 | 每月**新增**(delta)vcore,非當月總量 |
| 64 × 0.8 | 每台 64 vcore × 0.8 可用率;多機型後兩者皆為 per 機型參數 |
| BM Group | fab × network group 為**硬隔離**資源池,容量不可跨池互用 |
| 機器需求進位 | 每月**無條件進位**成整數台(工具行為;舊 Excel 為小數) |
| 搬遷假設 | 已上線 VM **不可搬遷**,新需求只能填入剩餘空間 |
| Current | 各機型**完全空置**的可用機台數;規劃範圍外已上線機器的零頭容量 v1 不計 |
| 跨月零頭 | 兩種推演模式可切換:**保守**(月末零頭作廢,當月批次內仍可共用同一台)/**填縫**(機器剩餘空間逐月帶著走,新需求先填縫再開新機) |
| 安全的定義 | 在最新實況下,存在一個不需搬遷的配置方案使各月裝得下 |

## 4. 架構

```
┌─────────────────────────────────────┐
│  UI 層(Streamlit,v1)              │  ← 未來可換 FastAPI + 前端,core 不動
├─────────────────────────────────────┤
│  core 套件(純 Python,無 UI 依賴)  │
│  ├─ models      資料模型             │
│  ├─ importer    Excel 匯入/匯出      │
│  ├─ planner     時間軸推演引擎        │
│  └─ solver      配置介面(可插拔)     │
│      ├─ naive   內建 first-fit 實作   │
│      └─ (未來)  同事的 solver adapter │
└─────────────────────────────────────┘
```

技術選型:Python 3.12 + Streamlit。理由:與 solver 同語言易整合;Streamlit 免寫前端
即有表格/圖表/上傳,內網 `streamlit run` 即可分享;邏輯集中於 core,UI 可低成本汰換。

## 5. 資料模型

| 實體 | 欄位 | 備註 |
|---|---|---|
| Sku | name, vcore_per_node, usable_ratio | 可售 vcore = vcore_per_node × usable_ratio |
| Pool | fab, bm_group | 硬隔離邊界,推演以 pool 為單位 |
| DemandDelta | pool, product, month, vcore | 可切割(液體) |
| VmSpecDemand | pool, product, month, vm_size_vcore, count | 不可切割(原子),僅部分 product 需要 |
| MoveIn | pool, sku, month, count | |
| Return | pool, product, sku, month, count | |
| CurrentStock | pool, sku, count | 期初空機 |
| Machine(推演內部) | sku, free_vcore | planner 逐月維護,破碎狀態不遺失 |

月份一律正規化為 `YYYY-MM`;匯入時容錯解析既有的 `2026'8`、`20267'` 等格式並回報異常。

## 6. Excel 格式

**舊格式(現行檔案)**:直接匯入,全部機器自動視為單一預設機型(參數 64 / 0.8,
可於 UI 調整)。第一天即可用。

**新格式(v2 範本,工具可產生)**:

- 廠區 tab:維持現行配置;Return 區塊增加「機型」欄
- 新增長表格式 tab:`HW_SKU`(機型目錄:名稱、vcore/台、可用率)、`HW_Current`
  (各 pool 期初庫存 per 機型)、`HW_MoveIn`(pool × 機型 × 月份進機計畫)——
  拆成多個長表 tab 以利解析穩定與人工維護
- 新增 `VM_Spec` tab:需要明細的 product 填 VM size × 顆數(pool × product × 月份)
- summary tab 由工具產出,不再要求人工維護公式

匯入驗證:缺 tab/缺欄、月份標頭異常、數值非法、公式錯位皆以中文訊息逐格回報,
不默默吞掉。

## 7. 推演引擎(planner)

每個 pool 獨立推演,逐月執行:

```
1. 庫存更新:加入 MoveIn[機型] 與 Return[機型] 的空機
2. 需求批次:當月新增 vcore(液體)+ 當月新增 VM 明細(原子)
3. 呼叫 solver.check(pool_state, demand_batch)
   - 可行 → 記錄配置、更新各機器 free_vcore,帶入下月
   - 不可行 → 標記缺口月份,記錄缺口量與卡住的原因;未配置的需求
     (backlog)滾入下月批次繼續嘗試
4. 模式處理(§3 跨月零頭):保守模式於月末將非空非滿機器的剩餘空間作廢;
   填縫模式原樣帶入下月
```

回推模式:同樣逐月,改呼叫 `solver.suggest(...)`,將建議採購併入該月庫存後繼續,
輸出各月建議進機計畫。

工具預設每月無條件進位(§3);另提供**相容模式**(不進位,僅供驗證遷移正確性,
不供正式規劃使用)——在單一機型 + 相容模式下,推演結果須與現行 Excel 公式逐格一致
(黃金測試依據),包括重現 B 廠 network1 的負庫存。

## 8. Solver 介面契約(與 solver 開發者協作的需求)

```python
class AllocationSolver(Protocol):

    def check(self, pool_state: PoolState, new_demand: DemandBatch) -> CheckResult:
        """驗證:既有配置固定,新需求只能填入剩餘空間(不可搬遷)"""

    def suggest(self, pool_state: PoolState, new_demand: DemandBatch,
                catalog: list[Sku]) -> SuggestResult:
        """回推:裝不下時,建議自 catalog 補進哪些機型、各幾台"""
```

對 solver 的三項核心需求(請與 solver 開發者對齊):

1. **兩個入口**:check(驗證可行性)與 suggest(建議採購組合)
2. **混合粒度輸入**:同時接受可切割 vcore 總量與不可切割 VM 明細
3. **增量配置模式**:既有配置為既成事實,僅配置新需求,不得假設可搬遷重排

輸出需求:`CheckResult` / `SuggestResult` 必含 placements,粒度需「可對人解釋」——
機型層級彙總為主、逐台明細可展開;不可行時需說明缺口量與受阻需求。

**內建 naive 實作**(solver 到位前的替身):VM 明細以 first-fit-decreasing(大顆先塞),
液體 vcore 填縫;suggest 以貪婪法自 catalog 挑單台可售 vcore 最大者補足。單一機型時
退化為除法 + 進位。

## 9. UI 畫面(Streamlit,五頁)

1. **匯入**:上傳 Excel(舊/新格式自動判別),顯示驗證報告
2. **總覽**:每 pool 一條時間軸(需求、進機、退回、逐月 in-stock:等效 vcore +
   per 機型明細),缺口月份標紅;紅色警示連結至配置明細解釋卡點
3. **驗證模式**:表上直接編修 MoveIn 計畫,即時重算可行性
4. **回推模式**:選擇機型 catalog,產出各月建議進機計畫,可匯出 Excel
5. **配置明細**:每 pool × 月份的 solver 配置結果(哪批需求 → 哪型機器幾台、
   各機器剩餘),兩模式共用;可匯出作為執行參考

## 10. 計畫與現實的校準(re-baseline)

「安全」結論的效力以實際配置跟隨計畫為前提,而現實必然偏離。因此:

- 工具**無內部持久狀態**:每次匯入即以檔案內容為唯一事實重新推演
- 運作節奏(寫入使用手冊):每月初以**實際**庫存/退回數字更新 Excel → 匯入 →
  未來月份全部重推,偏離不累積
- 配置明細可匯出,供執行方參考,縮小計畫與現實的落差

## 11. 測試與正確性

- **黃金測試**:以現行範例檔為 fixture,單一機型、不進位模式下與 Excel 公式逐格
  對值(容忍已知的 T7 公式錯誤,見 §12)
- **單元測試**:planner 逐月遞推(含負庫存、Return/MoveIn 時序)、naive solver
  裝箱(原子 VM 塞不進但總量足夠的案例)、舊/新格式 importer、月份解析容錯
- **錯誤處理**:匯入失敗與資料異常以中文明確訊息定位到 tab/儲存格

## 12. 現行 Excel 已知問題(匯入時回報)

1. summary `T7` 公式錯位:T 欄為 MoveIn,卻填入 Return 公式 `='C'!M5+'C'!M6`
2. 月份標頭 typo:`20267'` 應為 `2026'7`
3. in-stock 僅計算至 2026'9(AD 欄),10–12 月未拉公式

## 13. 未來方向(v2+)

- PM 線上填寫需求、多人協作(屆時評估 FastAPI + 前端)
- 已上線機器零頭容量納入(需平台實際用量資料源)
- 接入同事 solver 的 adapter 與對照模式(naive vs solver 結果比較)
- 採購最佳化維度(成本、交期、機架)
