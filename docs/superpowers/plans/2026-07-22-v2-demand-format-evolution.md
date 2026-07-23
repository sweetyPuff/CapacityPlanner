# v2 需求格式演進 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 把 VM 規格與配置政策(VM vcore 尺寸、每台上限 1:X、共居政策)併入 fab 頁 User Demand 列、移除獨立 `VM_Spec` 分頁(保留舊 v2 相容匯入),並產出含假資料的新 v2 sample。

**Architecture:** 新 fab 頁版面在 Product/BM Group 後多三個屬性欄(VM vcore / 每台上限 / 共居),月份右移;VM-detail product 的每月 vcore 由匯入器反推顆數(無條件進位 + 警告)成原子 VM。配置政策(max_per_machine、co_residency)為 per-product 常數,**v1 只收資料不強制執行**(等 CP-SAT)。舊 v2(含 VM_Spec 分頁)靠偵測分頁存在與否走舊解析路徑。

**Tech Stack:** Python 3.12、openpyxl、pandas、pytest

**Spec:** `docs/superpowers/specs/2026-07-22-v2-demand-format-evolution-design.md`(必讀,語意見 §2、決策見 §6)

## Global Constraints

- Python:`.venv\Scripts\python.exe`(PATH 無 `python`);指令在 repo 根 `D:\work\CapacityPlanning` 以 PowerShell 執行;測試 `.venv\Scripts\python.exe -m pytest`
- 使用者可見訊息(匯入警告、範本文案)一律繁體中文;月份內部 `YYYY-MM`
- 每個 commit 訊息結尾:`Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>`
- **新 fab 頁欄位常數(importer 與 exporter 必須一致)**:
  - 需求區:A=Product(1)、B=BM Group(2)、C=VM vcore(3)、D=每台上限(4)、E=共居(5)、月份自 F(6) 向右
  - 退回區:Product 於 T(20)、BM Group 於 U(21)、機型 於 V(22)、月份自 W(23) 向右
  - 標頭列 = row 4;資料自 row 5
- **語意(spec §2/§6)**:
  - 每月填 vcore;VM vcore 欄空 = 粗粒度液體(現行語意),有值 = 原子 VM 尺寸
  - detail product 每月顆數 = `ceil(vcore / vm_size)`;非整數倍 → 警告(不報錯)
  - 每台上限 = 1:X 的 X(空=None);共居:空→`"free"` / `獨佔`→`"exclusive"` / 其他→群組名
  - 一個 product 列**非粗即細**(coarse 產生 DemandDelta;detail 產生 VmSpecDemand,不重複計液體)
  - 政策 v1 只收資料,planner / solver 不變、不強制
- **相容**:偵測到 `VM_Spec` 分頁 → 走舊解析;否則走新解析。舊 v2 檔仍可匯入
- 全套測試必須維持綠燈(改到的既有測試一併更新)

## File Structure

```
captool/models.py       + ProductPolicy dataclass;PlanInput 增 policies + policy_for()
captool/importer.py      import_v2 分流新舊版面;新版面解析屬性欄 + 反推顆數 + 政策
captool/exporter.py      generate_v2_template 改新版面 + 示範列;移除 VM_Spec 分頁
examples/gen_v2_sample.py   產生新格式假資料 sample(新檔)
examples/v2_sample_new_format.xlsx  產出的 sample(新檔,commit 進 repo)
tests/test_models.py         + ProductPolicy / policy_for 測試
tests/test_exporter.py       更新:generate_v2_template 新版面斷言
tests/test_importer_v2.py    更新:新版面 roundtrip + 舊版面相容測試
```

---

### Task 1: ProductPolicy 模型

**Files:**
- Modify: `captool/models.py`
- Test: `tests/test_models.py`(附加)

**Interfaces:**
- Produces(後續 task 依賴,簽名照抄):
  - `ProductPolicy(pool: Pool, product: str, vm_size_vcore: int | None, max_per_machine: int | None, co_residency: str)`(frozen dataclass;`co_residency` 為 `"free"` / `"exclusive"` / 群組名)
  - `PlanInput` 新增欄位 `policies: list[ProductPolicy] = field(default_factory=list)`(接在 `issues` 之後,需有預設值)
  - `PlanInput.policy_for(pool: Pool, product: str) -> ProductPolicy | None`

- [ ] **Step 1: 附加失敗測試到 `tests/test_models.py`**

```python
def test_product_policy_and_lookup():
    from captool.models import ProductPolicy
    pol = ProductPolicy(pool=P, product="AI", vm_size_vcore=60,
                        max_per_machine=1, co_residency="exclusive")
    pi = PlanInput(
        skus={SKU.name: SKU}, months=["2026-07"], pools=[P],
        demands=[], vm_demands=[], moveins=[], returns=[], currents=[],
        policies=[pol])
    assert pi.policy_for(P, "AI") is pol
    assert pi.policy_for(P, "nope") is None


def test_plan_input_policies_default_empty():
    pi = PlanInput(skus={}, months=[], pools=[], demands=[], vm_demands=[],
                   moveins=[], returns=[], currents=[])
    assert pi.policies == []
```

(檔案頂端已有 `P`、`SKU`、`PlanInput` 匯入;沿用。)

- [ ] **Step 2: 跑測試確認失敗**

Run: `.venv\Scripts\python.exe -m pytest tests/test_models.py -v`
Expected: FAIL(`ImportError: cannot import name 'ProductPolicy'`)

- [ ] **Step 3: 在 `captool/models.py` 加入 ProductPolicy 並改 PlanInput**

在 `ImportIssue` 之後、`PlanInput` 之前加入:

```python
@dataclass(frozen=True)
class ProductPolicy:
    pool: Pool
    product: str
    vm_size_vcore: int | None      # None = 粗粒度(液體)
    max_per_machine: int | None    # 1:X 的 X;None = 不限
    co_residency: str              # "free" | "exclusive" | 群組名
```

在 `PlanInput` 的 `issues` 欄位之後新增(維持 `field` 匯入已存在):

```python
    policies: list[ProductPolicy] = field(default_factory=list)
```

在 `PlanInput` 類別內新增方法:

```python
    def policy_for(self, pool: Pool, product: str) -> "ProductPolicy | None":
        for p in self.policies:
            if p.pool == pool and p.product == product:
                return p
        return None
```

- [ ] **Step 4: 跑測試確認通過**

Run: `.venv\Scripts\python.exe -m pytest tests/test_models.py -v`
Expected: 全數通過

- [ ] **Step 5: Commit**

```powershell
git add captool/models.py tests/test_models.py
git commit -m "feat: ProductPolicy model and PlanInput.policy_for"
```

---

### Task 2: exporter 新版面 v2 範本

**Files:**
- Modify: `captool/exporter.py`(`generate_v2_template`)
- Test: `tests/test_exporter.py`(更新既有 `test_generate_v2_template`)

**Interfaces:**
- Consumes: `PlanInput`、models
- Produces:`generate_v2_template(plan_input, path) -> None` 輸出**新版面**:
  - 各廠區 tab:需求區標頭 row4 `A4=Product, B4=BM Group, C4=VM vcore, D4=每台上限, E4=共居`,月份自 F4;退回區 `T4=Product, U4=BM Group, V4=機型`,月份自 W4。既有 demand 帶入為粗粒度列(C/D/E 留空);既有 return 帶入(機型填 default sku)。加一列示範 detail 列(product 名含「示範」、淺黃底):`VM vcore=60, 每台上限=1, 共居=teamA`,某月填 120
  - `HW_SKU` / `HW_Current` / `HW_MoveIn`:不變
  - **不再產生 `VM_Spec` 分頁**
  - `README`:更新說明新三欄與「示範列可刪」
- 版面常數同 Global Constraints。

- [ ] **Step 1: 更新 `tests/test_exporter.py` 的 `test_generate_v2_template`(取代整個函式)**

```python
def test_generate_v2_template(tmp_path):
    pi = import_legacy(FIXTURE)
    out = tmp_path / "template_v2.xlsx"
    generate_v2_template(pi, out)
    wb = load_workbook(out)
    for sheet in ("HW_SKU", "HW_Current", "HW_MoveIn", "README", "A", "B", "C"):
        assert sheet in wb.sheetnames, sheet
    assert "VM_Spec" not in wb.sheetnames          # 新格式移除獨立分頁
    wa = wb["A"]
    assert [wa.cell(row=4, column=c).value for c in range(1, 6)] == \
        ["Product", "BM Group", "VM vcore", "每台上限", "共居"]
    assert wa.cell(row=4, column=6).value == "2026-07"   # 月份自 F
    assert wa.cell(row=4, column=20).value == "Product"  # 退回區 Product 於 T
    assert wa.cell(row=4, column=22).value == "機型"
    # 示範 detail 列存在
    found = False
    for r in range(5, wa.max_row + 1):
        if wa.cell(row=r, column=1).value and "示範" in str(wa.cell(row=r, column=1).value):
            assert wa.cell(row=r, column=3).value == 60     # VM vcore
            found = True
    assert found
    from captool.importer import detect_format
    assert detect_format(out) == "v2"
```

- [ ] **Step 2: 跑測試確認失敗**

Run: `.venv\Scripts\python.exe -m pytest tests/test_exporter.py::test_generate_v2_template -v`
Expected: FAIL(舊版面沒有這些欄位 / 仍有 VM_Spec)

- [ ] **Step 3: 取代 `captool/exporter.py` 的 `generate_v2_template`**

```python
def generate_v2_template(plan_input: PlanInput, path) -> None:
    wb = Workbook()

    ws = wb.active
    ws.title = "README"
    lines = [
        "容量規劃 v2 範本填表說明(新版面)",
        "",
        "各廠區 tab —— 需求區(A 欄起):",
        "  A=Product, B=BM Group, C=VM vcore, D=每台上限, E=共居, F 欄起為各月 vcore。",
        "  C(VM vcore)空 = 粗粒度需求(整包 vcore);有值 = 每顆 VM 的 vcore 尺寸。",
        "  D(每台上限)= 一台實體機最多住幾顆該 VM(1:X 的 X,亦即爆炸半徑上限);空=不限。",
        "  E(共居)= 空(自由,可與其他自由 product 混) / 獨佔 / 群組名(同名才可共用)。",
        "  每月一律填 vcore;detail product 的顆數由 vcore ÷ VM vcore 反推(非整數倍會進位並提示)。",
        "  同 product 若配置改變,請拆成兩列(如 A-1:1、A-1:2)。",
        "退回區(T 欄起):T=Product, U=BM Group, V=機型, W 欄起為各月退回台數。",
        "HW_SKU / HW_Current / HW_MoveIn:機型目錄 / 期初庫存 / 進機計畫(硬體 team 維護)。",
        "注意:機型名稱須存在於 HW_SKU;月份一律 YYYY-MM;示範列(product 含『示範』)請刪除。",
    ]
    for line in lines:
        ws.append([line])
    for row in ws.iter_rows():
        for cell in row:
            cell.font = _ARIAL
    ws["A1"].font = _BOLD

    fabs = sorted({p.fab for p in plan_input.pools})
    for fab in fabs:
        wf = wb.create_sheet(fab)
        wf["A4"], wf["B4"] = "Product", "BM Group"
        wf["C4"], wf["D4"], wf["E4"] = "VM vcore", "每台上限", "共居"
        wf.cell(row=4, column=20, value="Product")
        wf.cell(row=4, column=21, value="BM Group")
        wf.cell(row=4, column=22, value="機型")
        for i, month in enumerate(plan_input.months):
            wf.cell(row=4, column=6 + i, value=month)     # 需求月份 F 起
            wf.cell(row=4, column=23 + i, value=month)    # 退回月份 W 起

        # 帶入既有需求(粗粒度)
        fab_demands = [d for d in plan_input.demands if d.pool.fab == fab]
        rows: dict[tuple[str, str], int] = {}
        r = 5
        for d in fab_demands:
            key = (d.product, d.pool.bm_group)
            if key not in rows:
                rows[key] = r
                wf.cell(row=r, column=1, value=d.product)
                wf.cell(row=r, column=2, value=d.pool.bm_group)
                r += 1
            wf.cell(row=rows[key], column=6 + plan_input.months.index(d.month),
                    value=d.vcore)
        # 一列示範 detail
        example_month = plan_input.months[0] if plan_input.months else "2026-07"
        wf.cell(row=r, column=1, value="示範product(可刪)")
        wf.cell(row=r, column=2, value="network1")
        wf.cell(row=r, column=3, value=60)
        wf.cell(row=r, column=4, value=1)
        wf.cell(row=r, column=5, value="teamA")
        wf.cell(row=r, column=6, value=120)
        for c in range(1, 7):
            wf.cell(row=r, column=c).fill = _YELLOW

        # 帶入既有 return(機型填 default)
        fab_returns = [x for x in plan_input.returns if x.pool.fab == fab]
        rrows: dict[tuple[str, str, str], int] = {}
        r = 5
        for x in fab_returns:
            key = (x.product, x.pool.bm_group, x.sku_name)
            if key not in rrows:
                rrows[key] = r
                wf.cell(row=r, column=20, value=x.product)
                wf.cell(row=r, column=21, value=x.pool.bm_group)
                wf.cell(row=r, column=22, value=x.sku_name)
                r += 1
            wf.cell(row=rrows[key], column=23 + plan_input.months.index(x.month),
                    value=x.count)

        for row in wf.iter_rows():
            for cell in row:
                if cell.font is not _BOLD:
                    cell.font = _ARIAL
        for c in range(1, 6):
            wf.cell(row=4, column=c).font = _BOLD
        for c in (20, 21, 22):
            wf.cell(row=4, column=c).font = _BOLD

    ws1 = wb.create_sheet("HW_SKU")
    ws1.append(["name", "vcore_per_node", "usable_ratio"])
    for sku in plan_input.skus.values():
        ws1.append([sku.name, sku.vcore_per_node, sku.usable_ratio])
    ws2 = wb.create_sheet("HW_Current")
    ws2.append(["fab", "bm_group", "sku", "count"])
    for c in plan_input.currents:
        ws2.append([c.pool.fab, c.pool.bm_group, c.sku_name, c.count])
    ws3 = wb.create_sheet("HW_MoveIn")
    ws3.append(["fab", "bm_group", "sku", "month", "count"])
    for m in plan_input.moveins:
        ws3.append([m.pool.fab, m.pool.bm_group, m.sku_name, m.month, m.count])
    for sheet in (ws1, ws2, ws3):
        for row in sheet.iter_rows():
            for cell in row:
                cell.font = _ARIAL
        for cell in sheet[1]:
            cell.font = _BOLD
    wb.save(path)
```

- [ ] **Step 4: 跑測試確認通過**

Run: `.venv\Scripts\python.exe -m pytest tests/test_exporter.py -v`
Expected: `test_generate_v2_template` 通過;`test_export_summary` 系列不受影響仍通過

- [ ] **Step 5: Commit**

```powershell
git add captool/exporter.py tests/test_exporter.py
git commit -m "feat: v2 template new layout (policy columns on demand rows, no VM_Spec sheet)"
```

---

### Task 3: import_v2 新版面解析 + 舊版面相容

**Files:**
- Modify: `captool/importer.py`(`import_v2`、`V2_SHEETS`、`NON_FAB_SHEETS`)
- Test: `tests/test_importer_v2.py`(更新)

**Interfaces:**
- Consumes: Task 1 `ProductPolicy`、Task 2 `generate_v2_template`(新版面,roundtrip 測試用)、既有 `_read_table` / `_MonthParser` / `_numeric` / `_month_columns` / `import_legacy`
- Produces:`import_v2(path)` 分流:
  - 偵測:`"VM_Spec" in wb.sheetnames` → 走舊解析(現行邏輯,保留);否則 → 新解析
  - 新解析 fab 頁:需求列 A/B/C/D/E + 月份自 F;product 名含「示範」跳過;`C` 空 → 粗粒度 `DemandDelta`;`C` 有值 → detail:每月 `count = ceil(vcore / size)`,非整數倍加 warning,建立 `VmSpecDemand`;每個 detail product 建立一筆 `ProductPolicy`(粗粒度 product 不建 policy)。退回列自 T/U/V + 月份自 W(機型驗證同舊,不存在則 error 跳列)
  - `import_any` 對兩種 v2 皆正確分流(既有)
- 產出的 `PlanInput.policies` 帶新解析的政策;舊解析 policies 留空。

- [ ] **Step 1: 更新 `tests/test_importer_v2.py`**

保留檔案頂端 import,新增 `import math`(若無)。將 `_v2_file` 保持用 `generate_v2_template`(現在輸出新版面)。**取代** `test_roundtrip_preserves_data`、`test_multi_sku_and_vm_spec`、`test_unknown_sku_reported_and_skipped`,並新增新測試如下(其餘既有測試若引用 VM_Spec 分頁一併改寫或移除):

```python
import math
from openpyxl import Workbook, load_workbook

from captool.exporter import generate_v2_template
from captool.importer import import_any, import_legacy, import_v2
from captool.models import Pool
from pathlib import Path

FIXTURE = Path(__file__).parent / "fixtures" / "sample_legacy.xlsx"


def _v2_file(tmp_path):
    path = tmp_path / "v2.xlsx"
    generate_v2_template(import_legacy(FIXTURE), path)
    return path


def test_roundtrip_coarse_demand(tmp_path):
    path = _v2_file(tmp_path)
    original = import_legacy(FIXTURE)
    pi = import_v2(path)
    a1 = Pool(fab="A", bm_group="network1")
    assert pi.demand_vcore(a1, "2026-07") == original.demand_vcore(a1, "2026-07")
    assert pi.return_by_sku(a1, "2026-08") == original.return_by_sku(a1, "2026-08")
    assert pi.months == original.months
    # 示範列不得成為資料
    assert all("示範" not in d.product for d in pi.demands)
    assert pi.vm_demands == []            # 示範 detail 列被跳過


def test_detail_product_derives_count_and_policy(tmp_path):
    path = _v2_file(tmp_path)
    wb = load_workbook(path)
    wa = wb["A"]
    # 找一個空列加入 detail product:VM vcore=60, 每台上限=2, 共居=teamB, 2026-07=180
    r = wa.max_row + 1
    wa.cell(row=r, column=1, value="db")
    wa.cell(row=r, column=2, value="network1")
    wa.cell(row=r, column=3, value=60)
    wa.cell(row=r, column=4, value=2)
    wa.cell(row=r, column=5, value="teamB")
    wa.cell(row=r, column=6, value=180)     # 180/60 = 3 台
    wb.save(path)
    pi = import_v2(path)
    a1 = Pool(fab="A", bm_group="network1")
    assert pi.vm_batch(a1, "2026-07") == [(60, 3)]
    pol = pi.policy_for(a1, "db")
    assert pol is not None
    assert pol.vm_size_vcore == 60 and pol.max_per_machine == 2
    assert pol.co_residency == "teamB"


def test_detail_non_multiple_warns_and_rounds_up(tmp_path):
    path = _v2_file(tmp_path)
    wb = load_workbook(path)
    wa = wb["A"]
    r = wa.max_row + 1
    wa.cell(row=r, column=1, value="web")
    wa.cell(row=r, column=2, value="network1")
    wa.cell(row=r, column=3, value=60)
    wa.cell(row=r, column=6, value=100)     # 100/60 → 進位 2 台
    wb.save(path)
    pi = import_v2(path)
    a1 = Pool(fab="A", bm_group="network1")
    assert pi.vm_batch(a1, "2026-07") == [(60, 2)]
    assert any("整數倍" in i.message and i.severity == "warning" for i in pi.issues)


def test_exclusive_and_free_policy(tmp_path):
    path = _v2_file(tmp_path)
    wb = load_workbook(path)
    wa = wb["A"]
    r = wa.max_row + 1
    wa.cell(row=r, column=1, value="iso"); wa.cell(row=r, column=2, value="network1")
    wa.cell(row=r, column=3, value=51); wa.cell(row=r, column=5, value="獨佔")
    wa.cell(row=r, column=6, value=51)
    wb.save(path)
    pi = import_v2(path)
    pol = pi.policy_for(Pool(fab="A", bm_group="network1"), "iso")
    assert pol.co_residency == "exclusive"
    assert pol.max_per_machine is None      # D 空 = None


def test_import_any_dispatches_new(tmp_path):
    path = _v2_file(tmp_path)
    assert import_any(path).months == import_v2(path).months


def test_old_format_still_imports(tmp_path):
    # 手工建一個含 VM_Spec 分頁的舊 v2,確認相容路徑仍可解析
    path = tmp_path / "old_v2.xlsx"
    wb = Workbook()
    wb.remove(wb.active)
    a = wb.create_sheet("A")
    a["C3"] = "User Demand (vcore)"
    a["A4"], a["B4"], a["C4"] = "Product", "BM Group", "2026-07"
    a["A5"], a["B5"], a["C5"] = "svc", "network1", 100
    a["K4"], a["L4"], a["M4"], a["N4"] = "Product", "BM Group", "機型", "2026-07"
    hs = wb.create_sheet("HW_SKU"); hs.append(["name", "vcore_per_node", "usable_ratio"]); hs.append(["std-64", 64, 0.8])
    hc = wb.create_sheet("HW_Current"); hc.append(["fab", "bm_group", "sku", "count"]); hc.append(["A", "network1", "std-64", 5])
    hm = wb.create_sheet("HW_MoveIn"); hm.append(["fab", "bm_group", "sku", "month", "count"])
    vs = wb.create_sheet("VM_Spec"); vs.append(["fab", "bm_group", "product", "month", "vm_size_vcore", "count"]); vs.append(["A", "network1", "ai", "2026-07", 32, 2])
    wb.save(path)
    pi = import_v2(path)
    a1 = Pool(fab="A", bm_group="network1")
    assert pi.demand_vcore(a1, "2026-07") == 100
    assert pi.vm_batch(a1, "2026-07") == [(32, 2)]   # 舊 VM_Spec 分頁仍解析
```

- [ ] **Step 2: 跑測試確認失敗**

Run: `.venv\Scripts\python.exe -m pytest tests/test_importer_v2.py -v`
Expected: 新測試 FAIL(新解析未實作)

- [ ] **Step 3: 改 `captool/importer.py`**

在檔案頂端 import 區加入 `import math` 與 `ProductPolicy`:

```python
import math
from captool.models import (CurrentStock, DemandDelta, ImportIssue, MoveIn,
                            NodeReturn, PlanInput, Pool, ProductPolicy, Sku,
                            VmSpecDemand)
```

`V2_SHEETS` 改為只要求三張(VM_Spec 不再必需):

```python
V2_SHEETS = {"HW_SKU", "HW_Current", "HW_MoveIn"}
NON_FAB_SHEETS = V2_SHEETS | {"VM_Spec", "summary", "README"}
```

把現行 `import_v2` 改名為 `_import_v2_legacy_layout`(內容不變,即現行舊分頁解析),並新增分流入口與新解析:

```python
def import_v2(path) -> PlanInput:
    wb = load_workbook(path, read_only=True)
    try:
        old_layout = "VM_Spec" in wb.sheetnames
    finally:
        wb.close()
    return _import_v2_old_layout(path) if old_layout else _import_v2_new_layout(path)
```

(把現行 `import_v2` 函式主體改名成 `_import_v2_old_layout`;其內對必需分頁的檢查仍檢查 HW_SKU/HW_Current/HW_MoveIn/VM_Spec 四張——因為它是舊格式路徑,VM_Spec 本就存在。)

新增新解析:

```python
def _import_v2_new_layout(path) -> PlanInput:
    issues: list[ImportIssue] = []
    parser = _MonthParser(issues)
    wb_v = load_workbook(path, data_only=True)

    for required in V2_SHEETS:
        if required not in wb_v.sheetnames:
            raise CapacityImportError(
                [ImportIssue("error", required, "", f"找不到 {required} sheet")])

    skus = _read_skus(wb_v, issues)          # 見下(自 _import_v2_old_layout 抽出共用)
    first_sku_name = next(iter(skus))
    months: set[str] = set()

    currents = _read_currents(wb_v, skus, issues)      # 共用
    moveins = _read_moveins(wb_v, skus, issues, months)  # 共用

    demands: list[DemandDelta] = []
    vm_demands: list[VmSpecDemand] = []
    policies: list[ProductPolicy] = []
    returns: list[NodeReturn] = []

    fab_sheets = [n for n in wb_v.sheetnames if n not in NON_FAB_SHEETS]
    for name in fab_sheets:
        ws = wb_v[name]
        demand_cols = _month_columns(ws, 4, 6, parser, name)   # 月份自 F(6)
        row = 5
        while ws.cell(row=row, column=1).value is not None:
            product = str(ws.cell(row=row, column=1).value)
            if "示範" in product:
                row += 1
                continue
            group = str(ws.cell(row=row, column=2).value)
            pool = Pool(fab=name, bm_group=group)
            vm_size = ws.cell(row=row, column=3).value
            max_per = ws.cell(row=row, column=4).value
            coresid_raw = ws.cell(row=row, column=5).value
            if vm_size in (None, ""):
                # 粗粒度
                for col, month in demand_cols:
                    coord = f"{get_column_letter(col)}{row}"
                    vcore = _numeric(ws.cell(row=row, column=col).value, issues, name, coord)
                    if vcore:
                        demands.append(DemandDelta(pool=pool, product=product,
                                                   month=month, vcore=vcore))
                    months.add(month)
            else:
                size = int(vm_size)
                for col, month in demand_cols:
                    coord = f"{get_column_letter(col)}{row}"
                    vcore = _numeric(ws.cell(row=row, column=col).value, issues, name, coord)
                    months.add(month)
                    if vcore <= 0:
                        continue
                    count = math.ceil(vcore / size)
                    if abs(vcore - count * size) > 1e-9 and vcore % size != 0:
                        issues.append(ImportIssue(
                            "warning", name, coord,
                            f"{coord} 產品 {product} 的 vcore {vcore} 非 VM 規格 {size} "
                            f"的整數倍,已進位為 {count} 台({count*size} vcore)"))
                    vm_demands.append(VmSpecDemand(pool=pool, product=product,
                                                   month=month, vm_size_vcore=size,
                                                   count=count))
                policies.append(ProductPolicy(
                    pool=pool, product=product, vm_size_vcore=size,
                    max_per_machine=int(max_per) if max_per not in (None, "") else None,
                    co_residency=_parse_coresidency(coresid_raw)))
            row += 1

        # 退回:T/U/V + 月份自 W(23)
        return_cols = _month_columns(ws, 4, 23, parser, name)
        row = 5
        while ws.cell(row=row, column=20).value is not None:
            product = str(ws.cell(row=row, column=20).value)
            group = str(ws.cell(row=row, column=21).value)
            raw_sku = ws.cell(row=row, column=22).value
            if raw_sku is None:
                issues.append(ImportIssue("warning", name, f"V{row}",
                                          f"V{row} 退回機型空白,以 {first_sku_name} 計"))
                sku_name = first_sku_name
            elif str(raw_sku) not in skus:
                issues.append(ImportIssue("error", name, f"V{row}",
                                          f"機型 '{raw_sku}' 不存在於 HW_SKU,該列已跳過"))
                row += 1
                continue
            else:
                sku_name = str(raw_sku)
            pool = Pool(fab=name, bm_group=group)
            for col, month in return_cols:
                coord = f"{get_column_letter(col)}{row}"
                cnt = _numeric(ws.cell(row=row, column=col).value, issues, name, coord)
                if cnt:
                    returns.append(NodeReturn(pool=pool, product=product,
                                              sku_name=sku_name, month=month,
                                              count=int(cnt)))
                months.add(month)
            row += 1

    pools = sorted({c.pool for c in currents} | {d.pool for d in demands}
                   | {v.pool for v in vm_demands} | {m.pool for m in moveins}
                   | {r.pool for r in returns},
                   key=lambda p: (p.fab, p.bm_group))
    return PlanInput(skus=skus, months=sorted(months), pools=pools,
                     demands=demands, vm_demands=vm_demands, moveins=moveins,
                     returns=returns, currents=currents, issues=issues,
                     policies=policies)


def _parse_coresidency(raw) -> str:
    if raw in (None, ""):
        return "free"
    s = str(raw).strip()
    if s in ("獨佔", "exclusive", "獨占"):
        return "exclusive"
    return s
```

為讓新舊路徑共用讀取,將現行 `_import_v2_old_layout`(即原 import_v2)中「讀 HW_SKU / HW_Current / HW_MoveIn」的區塊抽成模組函式 `_read_skus(wb_v, issues) -> dict[str, Sku]`、`_read_currents(wb_v, skus, issues) -> list[CurrentStock]`、`_read_moveins(wb_v, skus, issues, months) -> list[MoveIn]`,兩條路徑都呼叫(行為與原本一致:重複 sku → error 取第一筆、參數非數值 → error 跳過、引用不存在 sku → error 跳列、month parse 失敗跳列)。`_read_skus` 為空丟 `CapacityImportError`。

- [ ] **Step 4: 跑全套測試確認通過**

Run: `.venv\Scripts\python.exe -m pytest -v`
Expected: 全數通過(含既有測試不回歸;`test_importer_v2.py` 新舊測試皆綠;golden / legacy / exporter / adapter 不受影響)

- [ ] **Step 5: Commit**

```powershell
git add captool/importer.py tests/test_importer_v2.py
git commit -m "feat: v2 new-layout import (policy columns, count derivation) with old-format compat"
```

---

### Task 4: 新 v2 sample + 假資料

**Files:**
- Create: `examples/gen_v2_sample.py`、`examples/v2_sample_new_format.xlsx`
- (無獨立測試;以匯入 + 推演驗證取代)

**Interfaces:**
- Consumes: `import_v2`、`run_check`、`run_suggest`、`NaiveSolver`、models、openpyxl

**目的:** 產一份新版面 sample,含粗粒度 + 三種 detail(自由 / 獨佔 / 群組)product,讓使用者直接看到新欄位與假資料;並驗證能匯入、推演可行。

- [ ] **Step 1: 建立 `examples/gen_v2_sample.py`**

```python
"""產生新版面 v2 sample(含假資料),並驗證可匯入 + 推演。

用法:.venv\\Scripts\\python.exe examples\\gen_v2_sample.py
"""
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

BOLD = Font(bold=True)
YELLOW = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
MONTHS = ["2026-07", "2026-08", "2026-09"]

SKUS = [("std-64", 64, 0.8), ("big-128", 128, 0.8), ("gpu-80", 80, 0.8)]
# 需求列:(fab, product, group, vm_size|None, max_per|None, coresid, {month: vcore})
DEMAND_ROWS = [
    ("A", "web-svc", "network1", None, None, "", {"2026-07": 200, "2026-08": 120}),           # 粗粒度
    ("A", "AI-train", "network1", 96, 1, "獨佔", {"2026-07": 192}),                            # detail 獨佔,2 顆
    ("A", "cache", "network1", 60, 2, "teamA", {"2026-07": 180, "2026-08": 120}),              # detail 群組
    ("B", "db", "network1", 48, 2, "", {"2026-07": 144, "2026-08": 96, "2026-09": 48}),        # detail 自由
]
CURRENT = [("A", "network1", "std-64", 6), ("A", "network1", "big-128", 3),
           ("B", "network1", "std-64", 8)]
MOVEIN = [("B", "network1", "std-64", "2026-08", 2)]
RETURNS = [("B", "network1", "db", "std-64", "2026-09", 1)]


def build(path):
    wb = Workbook(); wb.remove(wb.active)
    fabs = sorted({r[0] for r in DEMAND_ROWS} | {c[0] for c in CURRENT})
    for fab in fabs:
        wf = wb.create_sheet(fab)
        wf["A4"], wf["B4"] = "Product", "BM Group"
        wf["C4"], wf["D4"], wf["E4"] = "VM vcore", "每台上限", "共居"
        wf.cell(row=4, column=20, value="Product")
        wf.cell(row=4, column=21, value="BM Group")
        wf.cell(row=4, column=22, value="機型")
        for i, m in enumerate(MONTHS):
            wf.cell(row=4, column=6 + i, value=m)
            wf.cell(row=4, column=23 + i, value=m)
        r = 5
        for (f, prod, grp, size, mx, cor, mvals) in DEMAND_ROWS:
            if f != fab:
                continue
            wf.cell(row=r, column=1, value=prod)
            wf.cell(row=r, column=2, value=grp)
            if size is not None:
                wf.cell(row=r, column=3, value=size)
                if mx is not None:
                    wf.cell(row=r, column=4, value=mx)
                if cor:
                    wf.cell(row=r, column=5, value=cor)
                for c in range(1, 6):
                    wf.cell(row=r, column=c).fill = YELLOW
            for i, m in enumerate(MONTHS):
                if m in mvals:
                    wf.cell(row=r, column=6 + i, value=mvals[m])
            r += 1
        r = 5
        for (f, grp, prod, sku, month, cnt) in RETURNS:
            if f != fab:
                continue
            wf.cell(row=r, column=20, value=prod)
            wf.cell(row=r, column=21, value=grp)
            wf.cell(row=r, column=22, value=sku)
            wf.cell(row=r, column=23 + MONTHS.index(month), value=cnt)
            r += 1
        for c in range(1, 6):
            wf.cell(row=4, column=c).font = BOLD
        for c in (20, 21, 22):
            wf.cell(row=4, column=c).font = BOLD

    hs = wb.create_sheet("HW_SKU"); hs.append(["name", "vcore_per_node", "usable_ratio"])
    for s in SKUS:
        hs.append(list(s))
    hc = wb.create_sheet("HW_Current"); hc.append(["fab", "bm_group", "sku", "count"])
    for row in CURRENT:
        hc.append(list(row))
    hm = wb.create_sheet("HW_MoveIn"); hm.append(["fab", "bm_group", "sku", "month", "count"])
    for row in MOVEIN:
        hm.append(list(row))
    for sheet in (hs, hc, hm):
        for cell in sheet[1]:
            cell.font = BOLD
    wb.save(path)


def verify(path):
    from captool.importer import import_v2
    from captool.planner import run_check
    from captool.solver.naive import NaiveSolver
    pi = import_v2(path)
    print("機型:", list(pi.skus), "pools:", [p.label for p in pi.pools], "月份:", pi.months)
    print("政策:")
    for pol in pi.policies:
        print(f"  {pol.pool.label} {pol.product}: VM {pol.vm_size_vcore}v "
              f"每台上限={pol.max_per_machine} 共居={pol.co_residency}")
    print("警告:", [i.message for i in pi.issues] or "無")
    res = run_check(pi, NaiveSolver(), mode="conservative")
    for o in res.outcomes:
        print(f"  [{'OK  ' if o.feasible else '缺口'}] {o.pool.label} {o.month}: "
              f"液體={o.demand_vcore} VM={o.vm_demand} 新啟用={o.machines_opened}")


if __name__ == "__main__":
    out = ROOT / "examples" / "v2_sample_new_format.xlsx"
    build(out)
    print("wrote", out, "\n")
    verify(out)
```

- [ ] **Step 2: 執行產生並驗證**

Run: `.venv\Scripts\python.exe examples\gen_v2_sample.py`
Expected:輸出機型 3 / pools 含 A/network1、B/network1;政策列出 AI-train(獨佔)、cache(teamA)、db(free);`web-svc` 無政策(粗粒度);警告應為「無」(vcore 皆為 size 整數倍);run_check 各月多為 OK(如有缺口屬正常,只需匯入與推演不報錯)。若有非預期例外 → 用 systematic-debugging 修正 importer,不改期望值遷就。

- [ ] **Step 3: Commit**

```powershell
git add examples/gen_v2_sample.py examples/v2_sample_new_format.xlsx
git commit -m "docs: new-format v2 sample with fake data (coarse + detail free/exclusive/group)"
```

---

## Self-Review 紀錄

- **Spec coverage:** §3 新版面 → Task 2/3(欄位常數一致);§2 語意(填 vcore、非整數倍進位警告、粗/細二選一)→ Task 3;§4 資料模型(ProductPolicy)→ Task 1;§5 政策 v1 只收不強制 → planner/solver 不動(Task 3 只建 policy,不改推演);§6 決策(獨佔只收資料、群組同名、進位警告)→ Task 3 + `_parse_coresidency`;§7 相容(偵測 VM_Spec 走舊路徑)→ Task 3 分流;sample + 假資料 → Task 4。
- **Placeholder scan:** 無 TBD;測試與實作皆含完整程式碼。抽出 `_read_skus/_read_currents/_read_moveins` 為明確重構指示(自現行 import_v2 主體搬移,行為不變)。
- **Type consistency:** `ProductPolicy` 欄位(vm_size_vcore/max_per_machine/co_residency)於 Task 1 定義,Task 3 建構時一致;`co_residency` 值域 `"free"/"exclusive"/群組名` 由 `_parse_coresidency` 保證;欄位常數(demand F=6、return T=20/V=22/W=23)在 Task 2 與 Task 3 一致。
- **UI:** 政策為 capture-only,planner 用 demand_vcore/vm_batch(Task 3 正確填充),UI 無需改動即可繼續運作(總覽/配置明細/回推照舊);政策的 UI 呈現列為後續,不在本計畫。
