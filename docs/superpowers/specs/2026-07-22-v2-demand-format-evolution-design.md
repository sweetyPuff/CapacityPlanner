# v2 需求格式演進設計 —— VM 規格與配置政策併入 User Demand

日期:2026-07-22
狀態:設計草案,待使用者 review
關聯:延伸自 `2026-07-20-capacity-planning-tool-design.md`;實作邊界參考 `docs/capabilities-v1.md`、`docs/cpsat-adapter-integration.md`。

## 1. 動機

現行 v2 把 VM 明細放在獨立的 `VM_Spec` 分頁(fab/bm_group/product/month/vm_size/count)。實務討論後發現:

- **VM 規格是 product 的固定屬性,不是逐月會變的東西** —— 同一個 product 不會 day1 用某配置、day2 換另一種。所以它該跟著 product 走,寫在 User Demand 那一列,而不是拆到另一張逐月表(與 Server Return 寫在各廠頁同理)。
- 獨立 `VM_Spec` 表是一個平行資料源,容易 drift(v1 曾出現 roundtrip 掉資料的問題),折進 demand 可移除這個破口。
- PM 只需維護一張表,規格與每月需求量並排,心智模型一致。

## 2. 已確認的領域語意

| 項目 | 語意 |
|---|---|
| 每月數字 | 一律填 **vcore**(全表統一)。VM-detail product 的 vcore 由 PM 心中已換算為「整台 VM 對齊」的倍數(user 要 48 → 依 1:1 規格 60 vcore 填 60) |
| VM 規格 | product 的固定屬性:單顆 VM 的 vcore 尺寸 |
| 1:X 比例 | 一台實體機最多住 X 顆該 product 的 VM。**等同「爆炸半徑上限」** —— 一台機故障影響的 VM 數。1:X 與爆炸半徑是同一個參數 |
| 共居政策 | per-product,三級:**獨佔**(整台專用)/ **群組**(同 tenancy group 才可共用)/ **自由**(預設、多數,與任何不設限者共用) |
| 配置變更 | 若某 product 真的改配置,拆成兩列(如 `A-1:1`、`A-1:2`),保住「一列 = 規格恆定」的不變式 |

共居採「群組(tenancy group)」模型而非任意兩兩白名單:現實多為少數租戶類別可同住,群組用一個欄位即可表達,solver 也好處理。

## 3. 新的 fab 頁 User Demand 版面

每列一個 product,屬性欄接在 Product / BM Group 之後,月份欄右移:

| 欄 | 內容 | 說明 |
|---|---|---|
| A | Product | |
| B | BM Group | |
| C | VM vcore | 空 = 粗粒度液體需求(沿用現行語意);有值 = 原子 VM 尺寸 |
| D | 每台上限 (X) | 1:X 的 X;空 = 不限。等於爆炸半徑上限 |
| E | 共居政策 | 空 = 自由 / `獨佔` / 群組名(如 `teamA`) |
| F.. | 各月 vcore | 月份標頭於此列;數字為 vcore |

- 粗粒度 product:C/D/E 留空,行為與今日完全相同(液體 vcore)。
- VM-detail product:填 C(尺寸),視需要填 D、E。
- 顆數由工具反推:`count = 每月vcore / VM vcore`。

`VM_Spec` 獨立分頁**移除**(見 §7 相容性)。

## 4. 資料模型影響(內部)

- `VmSpecDemand` 保留(pool, product, month, vm_size_vcore, count),count 於匯入時由 vcore ÷ 尺寸推得。
- 新增 per-product 政策(建議獨立小結構,依 pool×product 唯一):
  - `max_per_machine: int | None`(1:X 的 X)
  - `co_residency: "free" | "exclusive" | group_name`
- planner / solver 契約(`AllocationSolver`)本身不變;政策如何被 solver 使用見 §5。
- `DemandBatch` 未來可帶上每顆原子 VM 的政策標記(供 CP-SAT),v1 可先忽略。

## 5. 執行邊界:v1 做什麼、什麼待 CP-SAT

**核心原則:格式一次收齊所有欄位(資料就緒),但強制執行分階段。**

| 屬性 | v1(NaiveSolver,單維 vcore) | 待 CP-SAT |
|---|---|---|
| VM 尺寸(原子裝箱) | ✅ 已支援(FFD) | ✅ 最佳化 |
| 每台上限 / 爆炸半徑 (1:X) | ❌ 不強制(只看 vcore 總量) | ✅ `max_vm_count` |
| 獨佔 | 🟡 可選保守近似:該 VM 佔用整台(浪費零頭) | ✅ dedication |
| 共居群組 | ❌ 不強制 | ✅ 候選過濾 / 同群才配 |
| 自由共用 | ✅ 預設 | ✅ 預設 |

1:X、共居群組屬**擺放約束**,非容量帳,NaiveSolver 無此概念 —— 這正是需要同事 CP-SAT 的具體理由(她的 model 原生有 `max_vm_count` 與反親和 / 候選過濾)。

## 6. 待確認的細節

1. **自由 × 群組能否同住一台?** 建議:群組視為分割,群組機器只給同群;自由 product 自成開放池,不混入群組機器。(待確認)
2. **vcore 非 VM 尺寸整數倍時**:建議「無條件進位成整台 + 警告」,與現行「每月無條件進位」一致(非報錯)。(待確認)
3. **v1 是否現在就做「獨佔 = 整台保留」的保守近似**,還是獨佔也先只收資料、等 CP-SAT?
4. **群組相容性是否只有「同名才可」**,還是需要「群組 A 可與群組 B 相容」的群組間關係?(目前假設只有同名可共用,最簡)

## 7. 相容性與遷移

- 舊 v2(含獨立 `VM_Spec` 分頁)的檔案:建議匯入時**仍相容支援一段時間**(偵測到 `VM_Spec` 分頁就沿用舊解析),同時新範本產出改用新版面。逐步淘汰。
- 新增欄位對現有粗粒度資料**向後相容**(C/D/E 留空 = 舊行為)。

## 8. 範圍評估

這是**收斂的格式演進**:動到 importer(新版面解析 + 政策欄)、exporter(v2 範本產生)、UI(呈現政策欄、配置明細顯示密度/共居)。核心推演引擎與 solver 契約幾乎不動;真正「強制執行密度 / 爆炸半徑 / 共居」的部分,隨 CP-SAT 接入再開。
