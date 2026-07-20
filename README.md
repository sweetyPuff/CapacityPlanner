# 容量規劃工具(Capacity Planning Tool)

## 專案目的

私有雲 platform 團隊需要一個工具，將 vcore 需求配置到具體的硬體機型上。本工具支援讀入現有 Excel 檔案(舊/新格式)，以內建的配置演算法或外部 solver 回答兩類問題：
- **驗證**：給定進機計畫，各月份是否裝得下？缺口在哪？
- **回推**：給定需求，各月最少應進哪些機型、各幾台？

設計文件參見：`docs/superpowers/specs/2026-07-20-capacity-planning-tool-design.md`

---

## 安裝

### 建立虛擬環境

使用 Python 3.12 建立虛擬環境：

```powershell
& "C:\Users\pp830\AppData\Local\Programs\Python\Python312\python.exe" -m venv .venv
```

### 安裝依賴

```powershell
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

依賴清單包括：
- `openpyxl` - Excel 檔案解析
- `pandas` - 資料框架與表格操作
- `streamlit` - 網頁 UI 框架
- `pytest` - 測試框架

---

## 啟動 UI

以 Streamlit 啟動網頁介面(預設在 `http://localhost:8501`):

```powershell
.venv\Scripts\python.exe -m streamlit run app\streamlit_app.py
```

工具提供五頁介面：
1. **匯入**：上傳 Excel(舊/新格式自動判別)，顯示驗證報告
2. **總覽**：每個 pool 的時間軸(需求、進機、退回、逐月庫存)，缺口月份標紅
3. **驗證模式**：編修 MoveIn(進機)計畫，即時重算可行性
4. **回推模式**：選擇機型目錄，產出建議進機計畫並匯出
5. **配置明細**：詳細的配置方案，機型層級彙總、可展開機器明細

---

## 每月校準節奏(Monthly Re-baseline)

工具的核心原則是**無內部持久狀態**：每次匯入時，檔案內容為唯一事實，全部月份重新推演。因此，計畫與現實偏離時的校準流程如下：

### 步驟
1. **月初**：根據實際庫存、實際退回數字更新 Excel 檔案
   - 更新 `CurrentStock` 資料表（期初空機庫存 per 機型）
   - 更新實際已退回的機器數(Return)
   - 更新已進機的數字(MoveIn)
   
2. **匯入**：將更新後的 Excel 上傳到工具，工具驗證並載入
   
3. **重新推演**：工具以最新實況為基礎，自動重算未來所有月份的配置
   - 偏離不累積、不外推
   - 每次都是「從現在起往前看」
   
4. **參考執行**：匯出配置明細，供執行方按月參考進機

### 工具保證
- 沒有隱藏的、跨月積累的調整；讀檔即重推
- 任何進月份的變更立即反映到後續所有月份

---

## 兩種推演模式

工具提供兩種推演模式(在 UI 上可切換)，分別應對不同的規劃策略：

### 保守模式（Conservative Mode）
**模式名稱**：`保守(跨月零頭作廢)`

**原理**：每月末時，機器的剩餘空間(未能充分利用的 vcore)全部作廢，不帶入下月。

**適用場景**：
- 採購方要求有**安全邊際**，不依賴零碎空間
- 按月重新規劃進機計畫，每月劃清分界
- 偏保守，易預測，但單機利用率較低

**效果**：
- 推演邏輯：月末 `drop_partial_leftovers()` 清空非空非滿機器的剩餘空間
- 跨月機器數通常較多(因零頭無法利用)
- 推導出的進機數偏高，但確保每月本身可行性

### 填縫模式（Gap-Fill Mode）
**模式名稱**：`填縫(零頭跨月可用)`

**原理**：每月機器的剩餘空間逐月帶著走，新需求優先填補上月的零頭，再開新機。

**適用場景**：
- 需要**真實利用率**，細粒度掌握每台機器的用量
- 機器購置後持續運作，零頭空間確實可被後續需求使用
- 規劃與實際偏離較小，能更準確反映實況

**效果**：
- 推演邏輯：月末不清空零頭，直接帶入下月
- 跨月機器數通常較少(因零頭被充分利用)
- 推導出的進機數偏低，反映無需額外購置即可滿足的狀態

### 選擇建議

| 因素 | 保守 | 填縫 |
|------|------|------|
| 採購安全邊際 | ✓ | 較弱 |
| 反映真實利用 | ✗ | ✓ |
| 進機數預測 | 偏多 | 偏少 |
| 月初校準頻率 | 可低(月邊界清) | 應較高(持續追蹤) |

---

## 舊格式→v2 遷移

目前 Excel 檔案分為**舊格式**(現行)與**新格式(v2 範本)**。工具可同時接受兩種格式。

### 舊格式特性
- 三個廠區 tab(`A`、`B`、`C`)，各含：
  - 需求區塊（product × BM Group，各月新增 vcore）
  - 進機計畫(MoveIn)、退回(Return)
  - Summary tab：逐月推演結果(需求 → 機器台數 → 庫存)
- 所有機器預設為單一預設機型(vcore = 64, 可用率 = 0.8)
- Excel 公式維護負擔重

### v2 格式特性
- 補充機型定義與多機型支援：
  - `HW_SKU` tab：機型目錄(名稱、vcore/台、可用率)
  - `HW_Current` tab：期初庫存 per 機型、per pool
  - `HW_MoveIn` tab：進機計畫拆成長表(pool × 機型 × 月份)
  - `VM_Spec` tab：（選用）某些 product 的 VM 明細(size × 顆數)
- Summary tab 由工具產出，無需人工維護公式

### 遷移流程

1. **取得 v2 範本**
   - 在工具的「總覽」頁，點擊「下載 v2 範本」連結
   - 工具會產生一份空白 v2 Excel 範本（含所有必要 tab 與欄位結構）

2. **硬體團隊補充機型資料**
   - 在 `HW_SKU` tab 中填入實際機型(名稱、vcore、可用率)
   - 在 `HW_Current` tab 中填入各 pool 的期初庫存
   - 在 `HW_MoveIn` tab 中填入進機計畫(pool、機型、月份、台數)

3. **遷移需求與棧位**
   - 將舊格式中的需求(product × vcore)複製到 v2 對應 tab（或使用 `VM_Spec` 填明細）
   - 廠區與 BM Group 結構保持一致，無需重新整理

4. **後續使用**
   - 從此之後上傳 v2 格式的 Excel
   - 工具會自動辨識並按新流程處理

### 文件驗證
工具在匯入時會檢查：
- tab 完整性（缺 tab 回報）
- 欄位完整性（缺欄回報）
- 月份標頭格式(容錯 typo，如 `20267'` → `2026'7`)
- 數值合法性(負數、非數值回報)
- 所有異常以中文訊息逐格定位

---

## Solver 介面契約與協作

### 介面位置

Solver 介面定義在：`captool/solver/interface.py`

該檔案定義了三個核心資料結構與一個協議(Protocol)：

#### 資料結構

1. **Machine** - 單台機器的狀態
   ```python
   @dataclass
   class Machine:
       sku_name: str         # 機型名稱
       sellable_vcore: float # 該機型每台可售vcore (= vcore_per_node × usable_ratio)
       free_vcore: float     # 當前剩餘可用vcore
   ```

2. **PoolState** - 一個 pool 的完整狀態
   ```python
   @dataclass
   class PoolState:
       machines: list[Machine]  # 該pool所有機器(已配置或庫存)
   ```
   - 方法：`add_empty()` 新增空機、`total_count()` 計算各機型總數、`empty_count()` 統計完全空白的機器、`free_vcore_total()` 計算總自由容量

3. **DemandBatch** - 一個月份的新增需求
   ```python
   @dataclass
   class DemandBatch:
       liquid_vcore: float = 0.0              # 可切割的vcore總量
       atomic_vms: list[tuple[int, int]] = [] # 不可切割的VM明細: (vm_size_vcore, count)
   ```

4. **CheckResult** - 驗證結果
   ```python
   @dataclass
   class CheckResult:
       feasible: bool                    # 是否可行
       placed_liquid_vcore: float        # 已配置的液體vcore
       unplaced_liquid_vcore: float      # 未配置的液體vcore(缺口)
       blocked_vms: list[tuple[int, int]] # 無法配置的原子VM: (vm_size, count)
       machines_opened: dict[str, int]   # 本批次新開的機器數 per sku
   ```

5. **SuggestResult** - 回推結果
   ```python
   @dataclass
   class SuggestResult:
       purchases: dict[str, int]  # sku_name -> 建議採購台數
       result: CheckResult        # 採購後的配置結果
   ```

#### AllocationSolver Protocol
```python
class AllocationSolver(Protocol):
    def check(self, pool_state: PoolState, new_demand: DemandBatch) -> CheckResult:
        """驗證:新需求只能填入剩餘空間。就地修改 pool_state 的機器 free_vcore。"""
        ...

    def suggest(self, pool_state: PoolState, new_demand: DemandBatch,
                catalog: list[Sku]) -> SuggestResult:
        """回推:裝不下時建議自 catalog 補機。成功時將採購機器併入 pool_state。"""
        ...
```

### 核心約定

1. **check() 就地修改**
   - `pool_state` 的每台機器的 `free_vcore` 會被實時更新
   - 即使返回 `feasible=False`，已能配置的部分仍已修改狀態(供缺口診斷)

2. **增量配置模式**
   - 既有的配置為既成事實，不可搬遷/重排已配置的 VM
   - 新需求只能填入機器的剩餘空間 `free_vcore`

3. **混合粒度輸入**
   - `DemandBatch.liquid_vcore`：可切割，靈活分配到任何機器
   - `DemandBatch.atomic_vms`：不可切割，必須塞入同一台機器，或整批失敗

4. **可解釋的輸出**
   - `CheckResult.machines_opened`：本批次首次動用多少新空機(per sku)
   - 缺口情況下，需清楚說明 `unplaced_liquid_vcore` 與 `blocked_vms` 各是多少

### 接入自有 Solver 的方法

#### 步驟 1：實作 `AllocationSolver` Protocol

實現一個類別或函數，同時提供 `check()` 與 `suggest()` 方法，遵守上述契約。例如：

```python
# my_solver.py
from captool.solver.interface import AllocationSolver, PoolState, DemandBatch, CheckResult, SuggestResult
from captool.models import Sku

class MySolver:
    def check(self, pool_state: PoolState, new_demand: DemandBatch) -> CheckResult:
        # 實作配置邏輯
        ...

    def suggest(self, pool_state: PoolState, new_demand: DemandBatch,
                catalog: list[Sku]) -> SuggestResult:
        # 實作回推邏輯
        ...
```

#### 步驟 2：在 UI 中替換

開啟 `app/streamlit_app.py`，尋找以下行：

```python
from captool.solver.naive import NaiveSolver
```

改為：

```python
from my_solver import MySolver  # 或你的solver模組路徑
```

再尋找初始化 solver 的地方(通常在推演函數中)，將：

```python
solver = NaiveSolver()
```

改為：

```python
solver = MySolver()
```

#### 步驟 3：測試驗證

執行測試確保整合無誤：

```powershell
.venv\Scripts\python.exe -m pytest
```

工具會自動用你的 solver 執行所有推演。若介面契約遵守得當，現有測試應無損過落。

#### 內建 NaiveSolver 行為參考

工具內建的 `NaiveSolver`（位置：`captool/solver/naive.py`）採用 first-fit-decreasing 策略：
- 原子 VM 按大小降序優先配置
- 液體 vcore 填補剩餘空間或使用新機器
- suggest() 以貪婪法自 catalog 挑可售 vcore 最大的機型補足

你的 solver 可採用更精細的策略(bin packing、整數規劃等)，只要遵守 Protocol 即可。

---

## 測試

### 執行所有測試

```powershell
.venv\Scripts\python.exe -m pytest
```

### 詳細測試輸出

```powershell
.venv\Scripts\python.exe -m pytest -v
```

測試涵蓋：
- **匯入驗證**：舊/新格式 Excel 解析、月份容錯、數值異常檢測
- **推演邏輯**：planner 逐月遞推、負庫存、Return/MoveIn 時序
- **配置演算法**：naive solver 裝箱(原子 VM 塞不進但總量足夠的案例)、suggest 採購推薦
- **黃金測試**：以現行範例檔為 fixture，單一機型、不進位模式下與 Excel 公式逐格對值

測試結果需全數通過才能表示工具正常運作。

---

## 已知限制與未來方向

### 已知 Excel 問題(匯入時回報)

1. Summary `T7` 公式錯位：T 欄為 MoveIn，卻填入 Return 公式
2. 月份標頭 typo：`20267'` 應為 `2026'7`(工具自動容錯)
3. in-stock 未拉滿公式：僅計算至 2026'9 月，10–12 月未拉公式(工具完全推演補充)

### v2+ 方向

- PM 線上填寫需求、多人協作(評估 FastAPI + 前端)
- 已上線機器零頭容量納入(需平台實際用量資料源)
- 接入同事 solver，與 naive 結果對照與效能對比
- 採購最佳化維度(成本、交期、機架)

---

## 常見問題

**Q: 工具會保存我的資料嗎？**  
A: 不。工具無內部持久狀態。每次匯入時重新推演，匯出後關閉瀏覽器，下次再上傳檔案即可。

**Q: 能一邊編修 MoveIn 一邊看效果嗎？**  
A: 可以。在「驗證模式」頁直接編修進機計畫，工具立即重算可行性，無需重新匯入。

**Q: 如何從舊格式遷移到 v2？**  
A: 在「總覽」頁下載 v2 範本，硬體團隊填機型與庫存，之後持續上傳 v2 格式即可。

**Q: 推演結果與老 Excel 不一樣？**  
A: 可能原因：(1) 進位策略差異(工具預設月末無條件進位，可在 UI 切換相容模式)；(2) 模式差異(保守 vs 填縫)；(3) 已知 Excel 問題(T7 公式、月份 typo)。詳見「已知限制」。

---

## 連絡與回饋

工具由 platform 工程師主導開發，設計文件與實作均位於本倉庫。有問題或建議可在 GitHub Issues 反映。
