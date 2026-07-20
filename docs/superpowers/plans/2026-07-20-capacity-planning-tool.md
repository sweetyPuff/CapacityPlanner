# 容量規劃工具 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 建立多機型容量規劃工具:core Python 套件(匯入 Excel、逐月推演、可插拔 solver)+ Streamlit UI,取代單機型 Excel 公式。

**Architecture:** 純 Python core 套件 `captool`(models / months / importer / planner / solver / exporter / viewmodel),UI 層 `app/streamlit_app.py` 僅做狀態管理與呈現。Solver 以 Protocol 定義契約,v1 內建 NaiveSolver(FFD 裝箱),未來替換為同事的 solver adapter。

**Tech Stack:** Python 3.12、openpyxl、pandas、Streamlit、pytest

**Spec:** `docs/superpowers/specs/2026-07-20-capacity-planning-tool-design.md`(領域語意見 §3,必讀)

## Global Constraints

- Python 執行檔:`C:\Users\pp830\AppData\Local\Programs\Python\Python312\python.exe`(PATH 上沒有 `python`,一律用全路徑或 venv 內路徑)
- 所有指令在 repo 根目錄 `D:\work\CapacityPlanning` 以 PowerShell 執行;venv 建於 `.venv`,測試一律 `\.venv\Scripts\python.exe -m pytest`
- 使用者可見訊息(匯入警告、UI 文案)一律繁體中文
- 月份內部表示一律 `YYYY-MM` 字串
- 浮點比較容差 `1e-9`;機器「空/滿」判斷用 `math.isclose`
- 匯出的 Excel 只寫值不寫公式(數值為 solver 計算結果,無對應公式);不得使用 XLOOKUP/FILTER 等新式函數
- 每個 commit 訊息結尾加:`Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>`
- 領域語意(勿違反):需求為每月新增 delta;pool(fab × bm_group)硬隔離;已上線 VM 不可搬遷;保守模式月末零頭作廢、填縫模式零頭跨月帶走;Current = 空機台數

## File Structure

```
captool/
├─ __init__.py            (空)
├─ months.py              月份解析容錯
├─ models.py              領域資料模型 + PlanInput
├─ solver/
│  ├─ __init__.py         (空)
│  ├─ interface.py        Protocol + PoolState/DemandBatch/CheckResult/SuggestResult
│  └─ naive.py            NaiveSolver(FFD + 貪婪 suggest)
├─ planner.py             逐月推演引擎(run_check / run_suggest,三模式)
├─ importer.py            legacy/v2 匯入 + 驗證報告
├─ exporter.py            summary 匯出 + v2 範本產生
└─ viewmodel.py           PlanResult → DataFrame(供 UI)
app/
└─ streamlit_app.py       五頁 UI
tests/
├─ fixtures/sample_legacy.xlsx   (自 活頁簿1.xlsx 複製)
├─ test_months.py
├─ test_models.py
├─ test_solver_naive.py
├─ test_planner.py
├─ test_importer_legacy.py
├─ test_golden.py
├─ test_exporter.py
├─ test_importer_v2.py
└─ test_viewmodel.py
conftest.py               (repo 根,空檔,讓 captool 可 import)
requirements.txt
pytest.ini
```

---

### Task 1: 專案骨架 + 月份解析

**Files:**
- Create: `requirements.txt`, `pytest.ini`, `conftest.py`, `.gitignore`, `captool/__init__.py`, `captool/months.py`
- Test: `tests/test_months.py`

**Interfaces:**
- Produces: `captool.months.parse_month(raw) -> tuple[str | None, str | None]` — 回傳 `(正規化 "YYYY-MM" 或 None, 警告訊息或 None)`。合法輸入無警告;可救回的異常格式(如 `20267'`)回傳正規化值 + 警告;救不回回傳 `(None, 錯誤訊息)`。

- [ ] **Step 1: 建立 venv 與安裝依賴**

```powershell
& "C:\Users\pp830\AppData\Local\Programs\Python\Python312\python.exe" -m venv .venv
.venv\Scripts\python.exe -m pip install --quiet openpyxl pandas streamlit pytest
```

- [ ] **Step 2: 建立骨架檔案**

`requirements.txt`:
```
openpyxl
pandas
streamlit
pytest
```

`pytest.ini`:
```ini
[pytest]
testpaths = tests
```

`.gitignore`:
```
.venv/
__pycache__/
*.pyc
~$*.xlsx
```

`conftest.py`(repo 根)與 `captool/__init__.py`:空檔案。

- [ ] **Step 3: 寫失敗測試**

`tests/test_months.py`:
```python
from captool.months import parse_month


def test_normal_formats():
    assert parse_month("2026'8") == ("2026-08", None)
    assert parse_month("2026-08") == ("2026-08", None)
    assert parse_month("2026/12") == ("2026-12", None)
    assert parse_month("202608") == ("2026-08", None)


def test_typo_header_recovered_with_warning():
    value, warning = parse_month("20267'")
    assert value == "2026-07"
    assert warning is not None and "20267'" in warning


def test_datetime_input():
    from datetime import datetime
    assert parse_month(datetime(2026, 9, 1)) == ("2026-09", None)


def test_unparseable():
    value, warning = parse_month("N/A")
    assert value is None
    assert warning is not None


def test_none():
    value, warning = parse_month(None)
    assert value is None
    assert warning is not None
```

- [ ] **Step 4: 跑測試確認失敗**

Run: `.venv\Scripts\python.exe -m pytest tests/test_months.py -v`
Expected: FAIL(`ModuleNotFoundError: No module named 'captool.months'`)

- [ ] **Step 5: 實作 `captool/months.py`**

```python
"""月份標頭解析:容錯處理現行 Excel 的各種格式與 typo。"""
import re
from datetime import datetime

_PATTERNS = [
    re.compile(r"^(?P<y>\d{4})['./\-](?P<m>\d{1,2})$"),      # 2026'8, 2026-08, 2026/8
    re.compile(r"^(?P<y>\d{4})(?P<m>0[1-9]|1[0-2])$"),        # 202608
]


def parse_month(raw):
    """回傳 (正規化 'YYYY-MM' 或 None, 警告訊息或 None)。"""
    if raw is None:
        return None, "月份標頭為空"
    if isinstance(raw, datetime):
        return f"{raw.year:04d}-{raw.month:02d}", None
    s = str(raw).strip()
    for pat in _PATTERNS:
        m = pat.match(s)
        if m:
            month = int(m.group("m"))
            if 1 <= month <= 12:
                return f"{int(m.group('y')):04d}-{month:02d}", None
    # 容錯:如 "20267'" 之類的 typo(去除非數字後為 4 位年 + 1~2 位月)
    digits = re.sub(r"\D", "", s)
    if len(digits) in (5, 6):
        y, mo = int(digits[:4]), int(digits[4:])
        if 2000 <= y <= 2100 and 1 <= mo <= 12:
            return f"{y:04d}-{mo:02d}", f"月份標頭 '{s}' 格式異常,已解讀為 {y:04d}-{mo:02d},請確認"
    return None, f"無法解析月份標頭 '{s}'"
```

- [ ] **Step 6: 跑測試確認通過**

Run: `.venv\Scripts\python.exe -m pytest tests/test_months.py -v`
Expected: 5 passed

- [ ] **Step 7: Commit**

```powershell
git add requirements.txt pytest.ini conftest.py .gitignore captool tests/test_months.py
git commit -m "feat: project scaffold and month header parsing"
```

---

### Task 2: 資料模型與 PlanInput

**Files:**
- Create: `captool/models.py`
- Test: `tests/test_models.py`

**Interfaces:**
- Produces(後續所有 task 依賴,簽名照抄):
  - `Sku(name: str, vcore_per_node: int, usable_ratio: float)`,property `sellable_vcore -> float`
  - `Pool(fab: str, bm_group: str)`(frozen,可作 dict key),property `label -> str`(`"A/network1"`)
  - `DemandDelta(pool, product, month, vcore)`
  - `VmSpecDemand(pool, product, month, vm_size_vcore: int, count: int)`
  - `MoveIn(pool, sku_name: str, month: str, count: int)`
  - `NodeReturn(pool, product: str, sku_name: str, month: str, count: int)`
  - `CurrentStock(pool, sku_name: str, count: int)`
  - `ImportIssue(severity: str, sheet: str, cell: str, message: str)`(severity 為 `"error"` 或 `"warning"`)
  - `PlanInput(skus: dict[str, Sku], months: list[str], pools: list[Pool], demands, vm_demands, moveins, returns, currents, issues)`,方法:
    - `demand_vcore(pool, month) -> float`
    - `vm_batch(pool, month) -> list[tuple[int, int]]`(同 size 聚合)
    - `movein_by_sku(pool, month) -> dict[str, int]`
    - `return_by_sku(pool, month) -> dict[str, int]`
    - `current_by_sku(pool) -> dict[str, int]`

- [ ] **Step 1: 寫失敗測試**

`tests/test_models.py`:
```python
from captool.models import (CurrentStock, DemandDelta, MoveIn, NodeReturn,
                            PlanInput, Pool, Sku, VmSpecDemand)

SKU = Sku(name="default-64", vcore_per_node=64, usable_ratio=0.8)
P = Pool(fab="A", bm_group="network1")


def _plan_input():
    return PlanInput(
        skus={SKU.name: SKU},
        months=["2026-07", "2026-08"],
        pools=[P],
        demands=[
            DemandDelta(pool=P, product="Apple", month="2026-07", vcore=512),
            DemandDelta(pool=P, product="Banana", month="2026-07", vcore=358),
        ],
        vm_demands=[
            VmSpecDemand(pool=P, product="Apple", month="2026-07", vm_size_vcore=32, count=2),
            VmSpecDemand(pool=P, product="Banana", month="2026-07", vm_size_vcore=32, count=1),
        ],
        moveins=[MoveIn(pool=P, sku_name=SKU.name, month="2026-08", count=35)],
        returns=[NodeReturn(pool=P, product="Apple", sku_name=SKU.name, month="2026-08", count=12)],
        currents=[CurrentStock(pool=P, sku_name=SKU.name, count=50)],
    )


def test_sellable_vcore():
    assert SKU.sellable_vcore == 51.2


def test_pool_label_and_hashable():
    assert P.label == "A/network1"
    assert {P: 1}[Pool(fab="A", bm_group="network1")] == 1


def test_demand_vcore_sums_products():
    assert _plan_input().demand_vcore(P, "2026-07") == 870


def test_vm_batch_aggregates_by_size():
    assert _plan_input().vm_batch(P, "2026-07") == [(32, 3)]


def test_lookups():
    pi = _plan_input()
    assert pi.movein_by_sku(P, "2026-08") == {"default-64": 35}
    assert pi.movein_by_sku(P, "2026-07") == {}
    assert pi.return_by_sku(P, "2026-08") == {"default-64": 12}
    assert pi.current_by_sku(P) == {"default-64": 50}
```

- [ ] **Step 2: 跑測試確認失敗**

Run: `.venv\Scripts\python.exe -m pytest tests/test_models.py -v`
Expected: FAIL(ModuleNotFoundError)

- [ ] **Step 3: 實作 `captool/models.py`**

```python
"""領域資料模型。語意見 spec §3:需求為每月新增 delta,pool 硬隔離。"""
from collections import defaultdict
from dataclasses import dataclass, field


@dataclass(frozen=True)
class Sku:
    name: str
    vcore_per_node: int
    usable_ratio: float

    @property
    def sellable_vcore(self) -> float:
        return self.vcore_per_node * self.usable_ratio


@dataclass(frozen=True)
class Pool:
    fab: str
    bm_group: str

    @property
    def label(self) -> str:
        return f"{self.fab}/{self.bm_group}"


@dataclass(frozen=True)
class DemandDelta:
    pool: Pool
    product: str
    month: str
    vcore: float


@dataclass(frozen=True)
class VmSpecDemand:
    pool: Pool
    product: str
    month: str
    vm_size_vcore: int
    count: int


@dataclass(frozen=True)
class MoveIn:
    pool: Pool
    sku_name: str
    month: str
    count: int


@dataclass(frozen=True)
class NodeReturn:
    pool: Pool
    product: str
    sku_name: str
    month: str
    count: int


@dataclass(frozen=True)
class CurrentStock:
    pool: Pool
    sku_name: str
    count: int


@dataclass(frozen=True)
class ImportIssue:
    severity: str  # "error" | "warning"
    sheet: str
    cell: str      # 例如 "T7";非特定儲存格用 ""
    message: str


@dataclass
class PlanInput:
    skus: dict[str, Sku]
    months: list[str]
    pools: list[Pool]
    demands: list[DemandDelta]
    vm_demands: list[VmSpecDemand]
    moveins: list[MoveIn]
    returns: list[NodeReturn]
    currents: list[CurrentStock]
    issues: list[ImportIssue] = field(default_factory=list)

    def demand_vcore(self, pool: Pool, month: str) -> float:
        return sum(d.vcore for d in self.demands if d.pool == pool and d.month == month)

    def vm_batch(self, pool: Pool, month: str) -> list[tuple[int, int]]:
        agg: dict[int, int] = defaultdict(int)
        for v in self.vm_demands:
            if v.pool == pool and v.month == month:
                agg[v.vm_size_vcore] += v.count
        return sorted(agg.items())

    def movein_by_sku(self, pool: Pool, month: str) -> dict[str, int]:
        agg: dict[str, int] = defaultdict(int)
        for m in self.moveins:
            if m.pool == pool and m.month == month:
                agg[m.sku_name] += m.count
        return dict(agg)

    def return_by_sku(self, pool: Pool, month: str) -> dict[str, int]:
        agg: dict[str, int] = defaultdict(int)
        for r in self.returns:
            if r.pool == pool and r.month == month:
                agg[r.sku_name] += r.count
        return dict(agg)

    def current_by_sku(self, pool: Pool) -> dict[str, int]:
        agg: dict[str, int] = defaultdict(int)
        for c in self.currents:
            if c.pool == pool:
                agg[c.sku_name] += c.count
        return dict(agg)
```

- [ ] **Step 4: 跑測試確認通過**

Run: `.venv\Scripts\python.exe -m pytest tests/test_models.py -v`
Expected: 5 passed

- [ ] **Step 5: Commit**

```powershell
git add captool/models.py tests/test_models.py
git commit -m "feat: domain models and PlanInput accessors"
```

---

### Task 3: Solver 介面 + NaiveSolver.check

**Files:**
- Create: `captool/solver/__init__.py`(空)、`captool/solver/interface.py`、`captool/solver/naive.py`
- Test: `tests/test_solver_naive.py`

**Interfaces:**
- Consumes: `captool.models.Sku`
- Produces(planner 與 UI 依賴):
  - `Machine(sku_name: str, sellable_vcore: float, free_vcore: float)`
  - `PoolState(machines: list[Machine])`,方法 `add_empty(sku: Sku, count: int)`、`total_count() -> dict[str, int]`、`empty_count() -> dict[str, int]`、`free_vcore_total() -> float`、`drop_partial_leftovers() -> float`
  - `DemandBatch(liquid_vcore: float = 0.0, atomic_vms: list[tuple[int, int]] = [])`,方法 `is_empty() -> bool`
  - `CheckResult(feasible, placed_liquid_vcore, unplaced_liquid_vcore, blocked_vms, machines_opened)`,property `shortfall_vcore`
  - `SuggestResult(purchases: dict[str, int], result: CheckResult)`
  - `AllocationSolver`(Protocol):`check(pool_state, new_demand) -> CheckResult`(**就地修改** pool_state,不可行時已放得下的部分仍配置)、`suggest(pool_state, new_demand, catalog) -> SuggestResult`
  - `NaiveSolver`(實作 Protocol;本 task 先實作 check,suggest 於 Task 4)

- [ ] **Step 1: 寫失敗測試**

`tests/test_solver_naive.py`:
```python
import pytest

from captool.models import Sku
from captool.solver.interface import DemandBatch, PoolState
from captool.solver.naive import NaiveSolver

SKU64 = Sku(name="default-64", vcore_per_node=64, usable_ratio=0.8)   # 可售 51.2
SKU128 = Sku(name="big-128", vcore_per_node=128, usable_ratio=0.8)    # 可售 102.4


def _state(sku=SKU64, count=10):
    s = PoolState()
    s.add_empty(sku, count)
    return s


def test_liquid_only_single_sku_matches_division():
    # 870 vcore / 51.2 = 16.99…,應開 17 台
    state = _state(count=50)
    result = NaiveSolver().check(state, DemandBatch(liquid_vcore=870))
    assert result.feasible
    assert result.machines_opened == {"default-64": 17}
    assert result.placed_liquid_vcore == 870
    assert state.empty_count() == {"default-64": 33}


def test_liquid_infeasible_partial_place():
    state = _state(count=2)  # 容量 102.4
    result = NaiveSolver().check(state, DemandBatch(liquid_vcore=150))
    assert not result.feasible
    assert result.placed_liquid_vcore == pytest.approx(102.4)
    assert result.unplaced_liquid_vcore == pytest.approx(47.6)
    assert result.shortfall_vcore == pytest.approx(47.6)


def test_atomic_vm_blocked_even_when_total_fits():
    # 總量夠(2 台共 102.4)但單台裝不下 60 vcore 的 VM
    state = _state(count=2)
    result = NaiveSolver().check(state, DemandBatch(atomic_vms=[(60, 1)]))
    assert not result.feasible
    assert result.blocked_vms == [(60, 1)]
    assert result.shortfall_vcore == pytest.approx(60)


def test_atomic_ffd_and_gap_fill_within_batch():
    # 51.2 可售:一台放 32+16,另一台放 32 → 2 台
    state = _state(count=5)
    result = NaiveSolver().check(
        state, DemandBatch(atomic_vms=[(16, 1), (32, 2)]))
    assert result.feasible
    assert result.machines_opened == {"default-64": 2}


def test_liquid_fills_partial_before_opening_empty():
    state = _state(count=2)
    NaiveSolver().check(state, DemandBatch(atomic_vms=[(32, 1)]))  # 剩 19.2 的半台
    result = NaiveSolver().check(state, DemandBatch(liquid_vcore=10))
    assert result.feasible
    assert result.machines_opened == {}  # 填縫,不開新機


def test_drop_partial_leftovers():
    state = _state(count=2)
    NaiveSolver().check(state, DemandBatch(liquid_vcore=60))  # 1 滿 + 1 剩 42.4
    dropped = state.drop_partial_leftovers()
    assert dropped == pytest.approx(42.4)
    assert state.free_vcore_total() == pytest.approx(0)
    assert state.total_count() == {"default-64": 2}


def test_multi_sku_check():
    state = PoolState()
    state.add_empty(SKU64, 1)
    state.add_empty(SKU128, 1)
    result = NaiveSolver().check(state, DemandBatch(atomic_vms=[(100, 1), (40, 1)]))
    assert result.feasible
    assert result.machines_opened == {"big-128": 1, "default-64": 1}
```

- [ ] **Step 2: 跑測試確認失敗**

Run: `.venv\Scripts\python.exe -m pytest tests/test_solver_naive.py -v`
Expected: FAIL(ModuleNotFoundError)

- [ ] **Step 3: 實作 `captool/solver/interface.py`**

```python
"""Solver 介面契約(spec §8)。任何 solver 實作(內建/外部 adapter)都遵守此檔。

核心約定:
- check() 就地修改 pool_state(不可行時已放得下的部分仍配置,供缺口診斷與 backlog)
- 增量配置:既有配置為既成事實,不可搬遷
- 混合粒度:liquid_vcore 可切割,atomic_vms 不可切割
"""
import math
from dataclasses import dataclass, field
from typing import Protocol

from captool.models import Sku

_EPS = 1e-9


@dataclass
class Machine:
    sku_name: str
    sellable_vcore: float
    free_vcore: float

    @property
    def is_empty(self) -> bool:
        return math.isclose(self.free_vcore, self.sellable_vcore)

    @property
    def is_full(self) -> bool:
        return self.free_vcore <= _EPS


@dataclass
class PoolState:
    machines: list[Machine] = field(default_factory=list)

    def add_empty(self, sku: Sku, count: int) -> None:
        for _ in range(count):
            self.machines.append(
                Machine(sku.name, sku.sellable_vcore, sku.sellable_vcore))

    def total_count(self) -> dict[str, int]:
        agg: dict[str, int] = {}
        for m in self.machines:
            agg[m.sku_name] = agg.get(m.sku_name, 0) + 1
        return agg

    def empty_count(self) -> dict[str, int]:
        agg: dict[str, int] = {}
        for m in self.machines:
            if m.is_empty:
                agg[m.sku_name] = agg.get(m.sku_name, 0) + 1
        return agg

    def free_vcore_total(self) -> float:
        return sum(m.free_vcore for m in self.machines)

    def drop_partial_leftovers(self) -> float:
        """保守模式月末呼叫:非空非滿機器的剩餘空間作廢。回傳作廢的 vcore 量。"""
        dropped = 0.0
        for m in self.machines:
            if not m.is_empty and not m.is_full:
                dropped += m.free_vcore
                m.free_vcore = 0.0
        return dropped


@dataclass
class DemandBatch:
    liquid_vcore: float = 0.0
    atomic_vms: list[tuple[int, int]] = field(default_factory=list)  # (vm_size, count)

    def is_empty(self) -> bool:
        return self.liquid_vcore <= _EPS and not self.atomic_vms


@dataclass
class CheckResult:
    feasible: bool
    placed_liquid_vcore: float
    unplaced_liquid_vcore: float
    blocked_vms: list[tuple[int, int]]
    machines_opened: dict[str, int]  # 本批次首次動用的空機數 per sku

    @property
    def shortfall_vcore(self) -> float:
        return self.unplaced_liquid_vcore + sum(s * c for s, c in self.blocked_vms)


@dataclass
class SuggestResult:
    purchases: dict[str, int]  # sku_name -> 建議採購台數
    result: CheckResult


class AllocationSolver(Protocol):
    def check(self, pool_state: PoolState, new_demand: DemandBatch) -> CheckResult:
        """驗證:新需求只能填入剩餘空間。就地修改 pool_state。"""
        ...

    def suggest(self, pool_state: PoolState, new_demand: DemandBatch,
                catalog: list[Sku]) -> SuggestResult:
        """回推:裝不下時建議自 catalog 補機。成功時將採購機器併入 pool_state。"""
        ...
```

- [ ] **Step 4: 實作 `captool/solver/naive.py`(check;suggest 先放 NotImplementedError)**

```python
"""內建 naive solver:同事 solver 到位前的替身。
原子 VM 用 first-fit-decreasing(填縫優先),液體 vcore 填縫後開空機。
單一機型情境退化為「除法 + 無條件進位」。
"""
from captool.models import Sku
from captool.solver.interface import (CheckResult, DemandBatch, PoolState,
                                      SuggestResult)

_EPS = 1e-9


class NaiveSolver:

    def check(self, pool_state: PoolState, new_demand: DemandBatch) -> CheckResult:
        opened: dict[str, int] = {}

        def open_if_empty(machine):
            if machine.is_empty:
                opened[machine.sku_name] = opened.get(machine.sku_name, 0) + 1

        # 原子 VM:大顆先塞;優先填已部分使用的機器,塞不下才開空機
        blocked: list[tuple[int, int]] = []
        for size, count in sorted(new_demand.atomic_vms, reverse=True):
            miss = 0
            for _ in range(count):
                partials = [m for m in pool_state.machines
                            if not m.is_empty and m.free_vcore >= size - _EPS]
                target = min(partials, key=lambda m: m.free_vcore, default=None)
                if target is None:
                    target = next((m for m in pool_state.machines
                                   if m.is_empty and m.free_vcore >= size - _EPS), None)
                if target is None:
                    miss += 1
                    continue
                open_if_empty(target)
                target.free_vcore -= size
            if miss:
                blocked.append((size, miss))

        # 液體 vcore:先填縫(剩餘小者優先),再開空機
        remaining = new_demand.liquid_vcore
        partials = sorted((m for m in pool_state.machines
                           if not m.is_empty and not m.is_full),
                          key=lambda m: m.free_vcore)
        empties = [m for m in pool_state.machines if m.is_empty]
        for m in partials + empties:
            if remaining <= _EPS:
                break
            take = min(m.free_vcore, remaining)
            if take <= _EPS:
                continue
            open_if_empty(m)
            m.free_vcore -= take
            remaining -= take
        remaining = max(remaining, 0.0)
        placed = new_demand.liquid_vcore - remaining

        return CheckResult(
            feasible=(remaining <= _EPS and not blocked),
            placed_liquid_vcore=placed,
            unplaced_liquid_vcore=remaining,
            blocked_vms=blocked,
            machines_opened=opened,
        )

    def suggest(self, pool_state: PoolState, new_demand: DemandBatch,
                catalog: list[Sku]) -> SuggestResult:
        raise NotImplementedError  # Task 4
```

- [ ] **Step 5: 跑測試確認通過**

Run: `.venv\Scripts\python.exe -m pytest tests/test_solver_naive.py -v`
Expected: 7 passed

- [ ] **Step 6: Commit**

```powershell
git add captool/solver tests/test_solver_naive.py
git commit -m "feat: solver interface contract and NaiveSolver.check"
```

---

### Task 4: NaiveSolver.suggest

**Files:**
- Modify: `captool/solver/naive.py`(補 suggest)
- Test: `tests/test_solver_naive.py`(附加)

**Interfaces:**
- Consumes: Task 3 全部型別
- Produces: `NaiveSolver.suggest(pool_state, new_demand, catalog) -> SuggestResult` — 貪婪:每輪以缺口量 ÷ catalog 中單台可售 vcore 最大機型估算加購台數,重試至可行;可行後把採購機器併入 pool_state 並完成配置。VM 大於任何機型單台可售容量時,回傳不可行結果與空採購(不無窮迴圈)。`catalog` 為空丟 `ValueError`。

- [ ] **Step 1: 附加失敗測試到 `tests/test_solver_naive.py`**

```python
def test_suggest_single_sku_buys_ceiling():
    state = PoolState()  # 沒庫存
    sres = NaiveSolver().suggest(state, DemandBatch(liquid_vcore=870), [SKU64])
    assert sres.result.feasible
    assert sres.purchases == {"default-64": 17}  # ceil(870/51.2)
    assert state.total_count() == {"default-64": 17}


def test_suggest_no_purchase_needed():
    state = _state(count=20)
    sres = NaiveSolver().suggest(state, DemandBatch(liquid_vcore=100), [SKU64])
    assert sres.purchases == {}
    assert sres.result.feasible


def test_suggest_picks_biggest_sellable_sku():
    state = PoolState()
    sres = NaiveSolver().suggest(state, DemandBatch(liquid_vcore=100), [SKU64, SKU128])
    assert sres.result.feasible
    assert sres.purchases == {"big-128": 1}


def test_suggest_atomic_converges():
    # (32,3)=96 vcore:1 台 51.2 只裝 1 顆,貪婪要迭代補到 3 台
    state = PoolState()
    sres = NaiveSolver().suggest(state, DemandBatch(atomic_vms=[(32, 3)]), [SKU64])
    assert sres.result.feasible
    assert sres.purchases == {"default-64": 3}


def test_suggest_unfittable_vm_returns_infeasible():
    state = PoolState()
    sres = NaiveSolver().suggest(state, DemandBatch(atomic_vms=[(200, 1)]), [SKU64, SKU128])
    assert not sres.result.feasible
    assert sres.purchases == {}
    assert sres.result.blocked_vms == [(200, 1)]


def test_suggest_empty_catalog_raises():
    with pytest.raises(ValueError):
        NaiveSolver().suggest(PoolState(), DemandBatch(liquid_vcore=1), [])
```

- [ ] **Step 2: 跑測試確認失敗**

Run: `.venv\Scripts\python.exe -m pytest tests/test_solver_naive.py -v`
Expected: 新增 6 個測試 FAIL(NotImplementedError / ValueError 未拋)

- [ ] **Step 3: 實作 suggest(取代 naive.py 中的 NotImplementedError 版本)**

```python
    MAX_TOTAL_PURCHASE = 100_000

    def suggest(self, pool_state: PoolState, new_demand: DemandBatch,
                catalog: list[Sku]) -> SuggestResult:
        import copy
        import math

        if not catalog:
            raise ValueError("catalog 不可為空")
        best = max(catalog, key=lambda s: s.sellable_vcore)
        max_vm = max((size for size, _ in new_demand.atomic_vms), default=0)
        if max_vm > best.sellable_vcore + _EPS:
            # 任何機型單台都裝不下的 VM:直接回報不可行(不採購)
            result = self.check(pool_state, new_demand)
            return SuggestResult(purchases={}, result=result)

        purchases: dict[str, int] = {}
        while True:
            trial = copy.deepcopy(pool_state)
            for name, cnt in purchases.items():
                sku = next(s for s in catalog if s.name == name)
                trial.add_empty(sku, cnt)
            result = self.check(trial, new_demand)
            if result.feasible:
                pool_state.machines[:] = trial.machines
                return SuggestResult(purchases=purchases, result=result)
            need = max(1, math.ceil(result.shortfall_vcore / best.sellable_vcore))
            purchases[best.name] = purchases.get(best.name, 0) + need
            if sum(purchases.values()) > self.MAX_TOTAL_PURCHASE:
                raise RuntimeError("suggest 未收斂:採購量超過上限")
```

注意:`_EPS` 已在檔案頂端定義;`import copy, math` 放檔案頂端亦可。

- [ ] **Step 4: 跑測試確認通過**

Run: `.venv\Scripts\python.exe -m pytest tests/test_solver_naive.py -v`
Expected: 13 passed

- [ ] **Step 5: Commit**

```powershell
git add captool/solver/naive.py tests/test_solver_naive.py
git commit -m "feat: NaiveSolver.suggest greedy purchase"
```

---

### Task 5: 推演引擎 run_check(三模式)

**Files:**
- Create: `captool/planner.py`
- Test: `tests/test_planner.py`

**Interfaces:**
- Consumes: `PlanInput`、solver 介面型別、`NaiveSolver`
- Produces(UI 與 exporter 依賴,簽名照抄):
  - `Mode = Literal["conservative", "gap_fill", "excel_compat"]`
  - `MonthOutcome`:欄位 `pool: Pool, month: str, demand_vcore: float, vm_demand: list[tuple[int,int]], movein: dict[str,int], returns: dict[str,int], feasible: bool, shortfall_vcore: float, blocked_vms: list[tuple[int,int]], machines_opened: dict[str,int], stock_total: dict[str,int], stock_empty: dict[str,int], free_vcore_total: float, suggested_purchases: dict[str,int]（預設 {}）, stock_float: float | None（預設 None,excel_compat 專用）`
  - `PlanResult(mode: str, outcomes: list[MonthOutcome])`,方法 `for_pool(pool) -> list[MonthOutcome]`、property `gap_months -> list[MonthOutcome]`
  - `run_check(plan_input, solver, mode="conservative") -> PlanResult`
- 行為約定:
  - 每 pool 獨立、逐月:先加 MoveIn/Return 空機 → 組批次(當月 delta + 上月 backlog)→ `solver.check` → 記錄 → 保守模式月末 `drop_partial_leftovers()`
  - backlog:`unplaced_liquid_vcore` 與 `blocked_vms` 滾入下月批次
  - `excel_compat`:不經 solver,聚合小數台遞推 `stock = prev + movein + return − demand/sellable`;要求恰好單一機型,否則 `ValueError`;`feasible = stock >= -1e-9`

- [ ] **Step 1: 寫失敗測試**

`tests/test_planner.py`:
```python
import pytest

from captool.models import (CurrentStock, DemandDelta, MoveIn, NodeReturn,
                            PlanInput, Pool, Sku, VmSpecDemand)
from captool.planner import run_check
from captool.solver.naive import NaiveSolver

SKU = Sku(name="default-64", vcore_per_node=64, usable_ratio=0.8)  # 可售 51.2
P = Pool(fab="A", bm_group="network1")


def _pi(demands, moveins=(), returns=(), current=10, vm_demands=(), months=("2026-07", "2026-08")):
    return PlanInput(
        skus={SKU.name: SKU}, months=list(months), pools=[P],
        demands=[DemandDelta(pool=P, product="X", month=m, vcore=v) for m, v in demands],
        vm_demands=[VmSpecDemand(pool=P, product="X", month=m, vm_size_vcore=s, count=c)
                    for m, s, c in vm_demands],
        moveins=[MoveIn(pool=P, sku_name=SKU.name, month=m, count=c) for m, c in moveins],
        returns=[NodeReturn(pool=P, product="X", sku_name=SKU.name, month=m, count=c)
                 for m, c in returns],
        currents=[CurrentStock(pool=P, sku_name=SKU.name, count=current)],
    )


def test_feasible_two_months():
    result = run_check(_pi(demands=[("2026-07", 100), ("2026-08", 100)]), NaiveSolver())
    o7, o8 = result.for_pool(P)
    assert o7.feasible and o8.feasible
    assert o7.machines_opened == {"default-64": 2}   # ceil(100/51.2)
    assert result.gap_months == []


def test_conservative_drops_leftover_across_months():
    # 30 vcore 占 1 台剩 21.2;保守模式下月零頭作廢,8 月的 30 再開新機
    result = run_check(_pi(demands=[("2026-07", 30), ("2026-08", 30)], current=2),
                       NaiveSolver(), mode="conservative")
    o7, o8 = result.for_pool(P)
    assert o7.machines_opened == {"default-64": 1}
    assert o8.machines_opened == {"default-64": 1}
    assert o8.free_vcore_total == pytest.approx(0)   # 兩台的零頭都作廢


def test_gap_fill_reuses_leftover_across_months():
    result = run_check(_pi(demands=[("2026-07", 30), ("2026-08", 20)], current=2),
                       NaiveSolver(), mode="gap_fill")
    o7, o8 = result.for_pool(P)
    assert o8.machines_opened == {}                  # 8 月 20 vcore 填 7 月剩的 21.2
    assert o8.free_vcore_total == pytest.approx(2 * 51.2 - 50)


def test_gap_and_backlog_carries():
    # 7 月需求 120 > 現有 2 台 102.4 → 缺口;8 月進 2 台後補上 backlog
    result = run_check(
        _pi(demands=[("2026-07", 120)], moveins=[("2026-08", 2)], current=2),
        NaiveSolver(), mode="gap_fill")
    o7, o8 = result.for_pool(P)
    assert not o7.feasible
    assert o7.shortfall_vcore == pytest.approx(120 - 102.4)
    assert o8.feasible                               # backlog 17.6 在 8 月被吸收
    assert result.gap_months == [o7]


def test_returns_add_empty_machines():
    result = run_check(
        _pi(demands=[("2026-07", 0), ("2026-08", 100)], returns=[("2026-08", 2)], current=0),
        NaiveSolver())
    o7, o8 = result.for_pool(P)
    assert o8.feasible
    assert o8.stock_total == {"default-64": 2}


def test_excel_compat_matches_formula():
    # stock = 50 + 0 + 0 - 870/51.2 = 33.0078125(對照現行 Excel AB3 邏輯)
    result = run_check(
        _pi(demands=[("2026-07", 870), ("2026-08", 985)], moveins=[("2026-08", 35)],
            returns=[("2026-08", 12)], current=50),
        NaiveSolver(), mode="excel_compat")
    o7, o8 = result.for_pool(P)
    assert o7.stock_float == pytest.approx(50 - 870 / 51.2)
    assert o8.stock_float == pytest.approx(o7.stock_float + 35 + 12 - 985 / 51.2)
    assert o7.feasible and o8.feasible


def test_excel_compat_negative_stock_is_gap():
    result = run_check(_pi(demands=[("2026-07", 600)], current=1), NaiveSolver(),
                       mode="excel_compat")
    (o7,) = result.for_pool(P)
    assert o7.stock_float < 0
    assert not o7.feasible


def test_excel_compat_requires_single_sku():
    pi = _pi(demands=[("2026-07", 10)])
    pi.skus["big-128"] = Sku(name="big-128", vcore_per_node=128, usable_ratio=0.8)
    with pytest.raises(ValueError):
        run_check(pi, NaiveSolver(), mode="excel_compat")
```

- [ ] **Step 2: 跑測試確認失敗**

Run: `.venv\Scripts\python.exe -m pytest tests/test_planner.py -v`
Expected: FAIL(ModuleNotFoundError)

- [ ] **Step 3: 實作 `captool/planner.py`**

```python
"""逐月推演引擎(spec §7)。每 pool 獨立;backlog 滾入下月;三種模式。"""
from dataclasses import dataclass, field
from typing import Literal

from captool.models import MoveIn, PlanInput, Pool, Sku
from captool.solver.interface import (AllocationSolver, DemandBatch, PoolState)

Mode = Literal["conservative", "gap_fill", "excel_compat"]

_EPS = 1e-9


@dataclass
class MonthOutcome:
    pool: Pool
    month: str
    demand_vcore: float
    vm_demand: list[tuple[int, int]]
    movein: dict[str, int]
    returns: dict[str, int]
    feasible: bool
    shortfall_vcore: float
    blocked_vms: list[tuple[int, int]]
    machines_opened: dict[str, int]
    stock_total: dict[str, int]
    stock_empty: dict[str, int]
    free_vcore_total: float
    suggested_purchases: dict[str, int] = field(default_factory=dict)
    stock_float: float | None = None  # excel_compat 專用


@dataclass
class PlanResult:
    mode: str
    outcomes: list[MonthOutcome]

    def for_pool(self, pool: Pool) -> list[MonthOutcome]:
        return [o for o in self.outcomes if o.pool == pool]

    @property
    def gap_months(self) -> list[MonthOutcome]:
        return [o for o in self.outcomes if not o.feasible]


def _merge_vms(a: list[tuple[int, int]], b: list[tuple[int, int]]) -> list[tuple[int, int]]:
    agg: dict[int, int] = {}
    for size, count in list(a) + list(b):
        agg[size] = agg.get(size, 0) + count
    return sorted(agg.items())


def _initial_state(plan_input: PlanInput, pool: Pool) -> PoolState:
    state = PoolState()
    for sku_name, cnt in plan_input.current_by_sku(pool).items():
        state.add_empty(plan_input.skus[sku_name], cnt)
    return state


def _single_sku(plan_input: PlanInput) -> Sku:
    if len(plan_input.skus) != 1:
        raise ValueError("excel_compat 模式僅支援單一機型")
    return next(iter(plan_input.skus.values()))


def run_check(plan_input: PlanInput, solver: AllocationSolver,
              mode: Mode = "conservative") -> PlanResult:
    if mode == "excel_compat":
        return _run_excel_compat(plan_input)
    outcomes: list[MonthOutcome] = []
    for pool in plan_input.pools:
        state = _initial_state(plan_input, pool)
        backlog = DemandBatch()
        for month in plan_input.months:
            movein = plan_input.movein_by_sku(pool, month)
            returns = plan_input.return_by_sku(pool, month)
            for sku_name, cnt in movein.items():
                state.add_empty(plan_input.skus[sku_name], cnt)
            for sku_name, cnt in returns.items():
                state.add_empty(plan_input.skus[sku_name], cnt)
            demand = plan_input.demand_vcore(pool, month)
            batch = DemandBatch(
                liquid_vcore=demand + backlog.liquid_vcore,
                atomic_vms=_merge_vms(backlog.atomic_vms,
                                      plan_input.vm_batch(pool, month)))
            result = solver.check(state, batch)
            backlog = DemandBatch(liquid_vcore=result.unplaced_liquid_vcore,
                                  atomic_vms=list(result.blocked_vms))
            if mode == "conservative":
                state.drop_partial_leftovers()
            outcomes.append(MonthOutcome(
                pool=pool, month=month, demand_vcore=demand,
                vm_demand=plan_input.vm_batch(pool, month),
                movein=movein, returns=returns,
                feasible=result.feasible,
                shortfall_vcore=result.shortfall_vcore,
                blocked_vms=list(result.blocked_vms),
                machines_opened=dict(result.machines_opened),
                stock_total=state.total_count(),
                stock_empty=state.empty_count(),
                free_vcore_total=state.free_vcore_total()))
    return PlanResult(mode=mode, outcomes=outcomes)


def _run_excel_compat(plan_input: PlanInput) -> PlanResult:
    sku = _single_sku(plan_input)
    outcomes: list[MonthOutcome] = []
    for pool in plan_input.pools:
        stock = float(sum(plan_input.current_by_sku(pool).values()))
        for month in plan_input.months:
            movein = plan_input.movein_by_sku(pool, month)
            returns = plan_input.return_by_sku(pool, month)
            demand = plan_input.demand_vcore(pool, month)
            stock = (stock + sum(movein.values()) + sum(returns.values())
                     - demand / sku.sellable_vcore)
            outcomes.append(MonthOutcome(
                pool=pool, month=month, demand_vcore=demand,
                vm_demand=[], movein=movein, returns=returns,
                feasible=stock >= -_EPS,
                shortfall_vcore=max(0.0, -stock) * sku.sellable_vcore,
                blocked_vms=[], machines_opened={},
                stock_total={}, stock_empty={}, free_vcore_total=0.0,
                stock_float=stock))
    return PlanResult(mode="excel_compat", outcomes=outcomes)
```

- [ ] **Step 4: 跑測試確認通過**

Run: `.venv\Scripts\python.exe -m pytest tests/test_planner.py -v`
Expected: 8 passed

- [ ] **Step 5: Commit**

```powershell
git add captool/planner.py tests/test_planner.py
git commit -m "feat: planner run_check with conservative/gap_fill/excel_compat modes"
```

---

### Task 6: 推演引擎 run_suggest(回推模式)

**Files:**
- Modify: `captool/planner.py`
- Test: `tests/test_planner.py`(附加)

**Interfaces:**
- Consumes: Task 5 全部 + `SuggestResult`
- Produces: `run_suggest(plan_input, solver, catalog: list[Sku], mode="conservative") -> tuple[PlanResult, list[MoveIn]]` — 逐月改呼叫 `solver.suggest`;建議採購記入 `MonthOutcome.suggested_purchases` 並彙整成 `list[MoveIn]` 回傳(pool、月份、機型、台數)。`mode="excel_compat"` 丟 `ValueError`。

- [ ] **Step 1: 附加失敗測試到 `tests/test_planner.py`**

```python
def test_run_suggest_fills_gap():
    from captool.planner import run_suggest
    pi = _pi(demands=[("2026-07", 120), ("2026-08", 60)], current=0)
    result, suggested = run_suggest(pi, NaiveSolver(), [SKU])
    o7, o8 = result.for_pool(P)
    assert o7.feasible and o8.feasible
    assert o7.suggested_purchases == {"default-64": 3}   # ceil(120/51.2)
    assert result.gap_months == []
    assert [(m.month, m.count) for m in suggested if m.pool == P] == [
        ("2026-07", 3), ("2026-08", 2)]                   # 保守模式:8 月再開 2 台


def test_run_suggest_no_purchase_when_stock_enough():
    from captool.planner import run_suggest
    pi = _pi(demands=[("2026-07", 100)], current=10, months=("2026-07",))
    result, suggested = run_suggest(pi, NaiveSolver(), [SKU])
    assert suggested == []
    assert result.for_pool(P)[0].suggested_purchases == {}


def test_run_suggest_rejects_excel_compat():
    from captool.planner import run_suggest
    with pytest.raises(ValueError):
        run_suggest(_pi(demands=[("2026-07", 1)]), NaiveSolver(), [SKU],
                    mode="excel_compat")
```

- [ ] **Step 2: 跑測試確認失敗**

Run: `.venv\Scripts\python.exe -m pytest tests/test_planner.py -v`
Expected: 新增 3 個 FAIL(ImportError: cannot import name 'run_suggest')

- [ ] **Step 3: 在 `captool/planner.py` 加入 run_suggest**

```python
def run_suggest(plan_input: PlanInput, solver: AllocationSolver,
                catalog: list[Sku], mode: Mode = "conservative"
                ) -> tuple[PlanResult, list[MoveIn]]:
    if mode == "excel_compat":
        raise ValueError("回推模式不支援 excel_compat")
    outcomes: list[MonthOutcome] = []
    suggested: list[MoveIn] = []
    for pool in plan_input.pools:
        state = _initial_state(plan_input, pool)
        backlog = DemandBatch()
        for month in plan_input.months:
            movein = plan_input.movein_by_sku(pool, month)
            returns = plan_input.return_by_sku(pool, month)
            for sku_name, cnt in movein.items():
                state.add_empty(plan_input.skus[sku_name], cnt)
            for sku_name, cnt in returns.items():
                state.add_empty(plan_input.skus[sku_name], cnt)
            demand = plan_input.demand_vcore(pool, month)
            batch = DemandBatch(
                liquid_vcore=demand + backlog.liquid_vcore,
                atomic_vms=_merge_vms(backlog.atomic_vms,
                                      plan_input.vm_batch(pool, month)))
            sres = solver.suggest(state, batch, catalog)
            result = sres.result
            for sku_name, cnt in sres.purchases.items():
                suggested.append(MoveIn(pool=pool, sku_name=sku_name,
                                        month=month, count=cnt))
            backlog = DemandBatch(liquid_vcore=result.unplaced_liquid_vcore,
                                  atomic_vms=list(result.blocked_vms))
            if mode == "conservative":
                state.drop_partial_leftovers()
            outcomes.append(MonthOutcome(
                pool=pool, month=month, demand_vcore=demand,
                vm_demand=plan_input.vm_batch(pool, month),
                movein=movein, returns=returns,
                feasible=result.feasible,
                shortfall_vcore=result.shortfall_vcore,
                blocked_vms=list(result.blocked_vms),
                machines_opened=dict(result.machines_opened),
                stock_total=state.total_count(),
                stock_empty=state.empty_count(),
                free_vcore_total=state.free_vcore_total(),
                suggested_purchases=dict(sres.purchases)))
    return PlanResult(mode=mode, outcomes=outcomes), suggested
```

- [ ] **Step 4: 跑測試確認通過**

Run: `.venv\Scripts\python.exe -m pytest tests/test_planner.py -v`
Expected: 11 passed

- [ ] **Step 5: Commit**

```powershell
git add captool/planner.py tests/test_planner.py
git commit -m "feat: planner run_suggest for purchase recommendation"
```

---

### Task 7: Legacy Excel 匯入

**Files:**
- Create: `captool/importer.py`、`tests/fixtures/sample_legacy.xlsx`
- Test: `tests/test_importer_legacy.py`

**Interfaces:**
- Consumes: `models`、`months.parse_month`
- Produces:
  - `DEFAULT_SKU = Sku(name="default-64", vcore_per_node=64, usable_ratio=0.8)`
  - `class CapacityImportError(Exception)`:屬性 `issues: list[ImportIssue]`
  - `import_legacy(path, default_sku=DEFAULT_SKU) -> PlanInput`(結構性致命問題丟 CapacityImportError;可續行問題進 `PlanInput.issues`)
  - `detect_format(path) -> str`(`"v2"` 若含 `HW_SKU` sheet,否則 `"legacy"`)
  - `import_any(path) -> PlanInput`(依 detect_format 分派;v2 於 Task 10 前先丟 NotImplementedError)
- Legacy 檔案結構約定(來自 `活頁簿1.xlsx` 實測):
  - 廠區 tab = 除 `summary` 外的所有 sheet。需求區:`C3` 起為 `User Demand (vcore)` 合併標頭、月份在 row 4 自 C 欄向右至空白、資料自 row 5(A=product、B=bm_group)向下至 A 欄空白。Return 區:`M3` 標頭、月份 row 4 自 M 欄向右、資料自 row 5(K=product、L=bm_group)向下至 K 欄空白
  - summary tab:資料自 row 3 至 A 欄空白(A=fab、B=bm_group);`Current` 欄 = row 2 中值為 `Current` 的欄;MoveIn 月份欄 = row 1 中值以 `Server MoveIn` 開頭之合併範圍所涵蓋的欄,月份在 row 2
  - MoveIn 與 Current 區若含公式(用非 data_only 的 workbook 檢查)→ warning issue(捕捉現行 T7 錯位)
  - 月份 parse 警告 → warning issue(每個相異原始字串只報一次)

- [ ] **Step 1: 複製 fixture**

```powershell
New-Item -ItemType Directory -Force tests\fixtures
Copy-Item "活頁簿1.xlsx" tests\fixtures\sample_legacy.xlsx
```

- [ ] **Step 2: 寫失敗測試**

`tests/test_importer_legacy.py`:
```python
from pathlib import Path

from captool.importer import DEFAULT_SKU, detect_format, import_legacy
from captool.models import Pool

FIXTURE = Path(__file__).parent / "fixtures" / "sample_legacy.xlsx"


def test_detect_format():
    assert detect_format(FIXTURE) == "legacy"


def test_pools_and_months():
    pi = import_legacy(FIXTURE)
    assert pi.months == ["2026-07", "2026-08", "2026-09",
                         "2026-10", "2026-11", "2026-12"]
    assert Pool(fab="A", bm_group="network1") in pi.pools
    assert len(pi.pools) == 6


def test_single_default_sku():
    pi = import_legacy(FIXTURE)
    assert set(pi.skus) == {DEFAULT_SKU.name}


def test_demand_values():
    pi = import_legacy(FIXTURE)
    # A tab: network1 = Apple 512 + Banana 358 = 870(2026-07)
    assert pi.demand_vcore(Pool(fab="A", bm_group="network1"), "2026-07") == 870
    # C tab: network2 = 331 + 1024 = 1355
    assert pi.demand_vcore(Pool(fab="C", bm_group="network2"), "2026-07") == 1355


def test_returns_and_current_and_movein():
    pi = import_legacy(FIXTURE)
    a1 = Pool(fab="A", bm_group="network1")
    assert pi.return_by_sku(a1, "2026-08") == {DEFAULT_SKU.name: 12}
    assert pi.current_by_sku(a1) == {DEFAULT_SKU.name: 50}
    assert pi.movein_by_sku(a1, "2026-08") == {DEFAULT_SKU.name: 35}


def test_issues_reported():
    pi = import_legacy(FIXTURE)
    messages = " ".join(i.message for i in pi.issues)
    assert "20267'" in messages                 # 月份 typo 警告
    formula_issues = [i for i in pi.issues if "公式" in i.message]
    assert any(i.cell == "T7" for i in formula_issues)   # T7 錯位公式
    assert all(i.severity == "warning" for i in pi.issues)
```

- [ ] **Step 3: 跑測試確認失敗**

Run: `.venv\Scripts\python.exe -m pytest tests/test_importer_legacy.py -v`
Expected: FAIL(ModuleNotFoundError)

- [ ] **Step 4: 實作 `captool/importer.py`**

```python
"""Excel 匯入:legacy(現行單機型格式)與 v2(多機型格式,Task 10)。"""
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from captool.models import (CurrentStock, DemandDelta, ImportIssue, MoveIn,
                            NodeReturn, PlanInput, Pool, Sku)
from captool.months import parse_month

DEFAULT_SKU = Sku(name="default-64", vcore_per_node=64, usable_ratio=0.8)

V2_SHEETS = {"HW_SKU", "HW_Current", "HW_MoveIn", "VM_Spec"}
NON_FAB_SHEETS = V2_SHEETS | {"summary", "README"}


class CapacityImportError(Exception):
    def __init__(self, issues: list[ImportIssue]):
        self.issues = issues
        super().__init__("; ".join(i.message for i in issues))


def detect_format(path) -> str:
    wb = load_workbook(path, read_only=True)
    try:
        return "v2" if "HW_SKU" in wb.sheetnames else "legacy"
    finally:
        wb.close()


def import_any(path) -> PlanInput:
    if detect_format(path) == "v2":
        return import_v2(path)  # Task 10
    return import_legacy(path)


def import_v2(path) -> PlanInput:
    raise NotImplementedError("v2 匯入於 Task 10 實作")


class _MonthParser:
    """包裝 parse_month:相同原始字串的警告只回報一次。"""

    def __init__(self, issues: list[ImportIssue]):
        self.issues = issues
        self._seen: set[str] = set()

    def parse(self, raw, sheet: str, cell: str):
        value, warning = parse_month(raw)
        if warning and str(raw) not in self._seen:
            self._seen.add(str(raw))
            self.issues.append(ImportIssue("warning", sheet, cell, warning))
        return value


def _numeric(value, issues, sheet, cell) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    issues.append(ImportIssue(
        "error", sheet, cell, f"儲存格 {cell} 應為數值,實際為 {value!r},已以 0 計"))
    return 0.0


def _month_columns(ws, header_row: int, start_col: int, parser, sheet: str):
    """自 start_col 向右讀月份標頭至空白,回傳 [(col_index, 'YYYY-MM')]。"""
    cols = []
    col = start_col
    while True:
        raw = ws.cell(row=header_row, column=col).value
        if raw is None:
            break
        coord = f"{get_column_letter(col)}{header_row}"
        month = parser.parse(raw, sheet, coord)
        if month:
            cols.append((col, month))
        col += 1
    return cols


def import_legacy(path, default_sku: Sku = DEFAULT_SKU) -> PlanInput:
    issues: list[ImportIssue] = []
    parser = _MonthParser(issues)
    wb_v = load_workbook(path, data_only=True)
    wb_f = load_workbook(path)

    if "summary" not in wb_v.sheetnames:
        raise CapacityImportError(
            [ImportIssue("error", "summary", "", "找不到 summary sheet")])

    demands: list[DemandDelta] = []
    returns: list[NodeReturn] = []
    months: set[str] = set()

    fab_sheets = [n for n in wb_v.sheetnames if n not in NON_FAB_SHEETS]
    for name in fab_sheets:
        ws = wb_v[name]
        # 需求區:月份 row 4 自 C(3) 向右;資料自 row 5,A=product、B=group
        demand_cols = _month_columns(ws, 4, 3, parser, name)
        row = 5
        while ws.cell(row=row, column=1).value is not None:
            product = str(ws.cell(row=row, column=1).value)
            group = str(ws.cell(row=row, column=2).value)
            pool = Pool(fab=name, bm_group=group)
            for col, month in demand_cols:
                coord = f"{get_column_letter(col)}{row}"
                vcore = _numeric(ws.cell(row=row, column=col).value, issues, name, coord)
                if vcore:
                    demands.append(DemandDelta(pool=pool, product=product,
                                               month=month, vcore=vcore))
                months.add(month)
            row += 1
        # Return 區:月份 row 4 自 M(13) 向右;資料自 row 5,K=product、L=group
        return_cols = _month_columns(ws, 4, 13, parser, name)
        row = 5
        while ws.cell(row=row, column=11).value is not None:
            product = str(ws.cell(row=row, column=11).value)
            group = str(ws.cell(row=row, column=12).value)
            pool = Pool(fab=name, bm_group=group)
            for col, month in return_cols:
                coord = f"{get_column_letter(col)}{row}"
                cnt = _numeric(ws.cell(row=row, column=col).value, issues, name, coord)
                if cnt:
                    returns.append(NodeReturn(pool=pool, product=product,
                                              sku_name=default_sku.name,
                                              month=month, count=int(cnt)))
                months.add(month)
            row += 1

    # summary:pools、Current、MoveIn
    ws = wb_v["summary"]
    ws_f = wb_f["summary"]
    current_col = None
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=2, column=col).value == "Current":
            current_col = col
            break
    if current_col is None:
        raise CapacityImportError(
            [ImportIssue("error", "summary", "", "summary 找不到 Current 欄")])
    movein_cols = []
    for rng in ws_f.merged_cells.ranges:
        anchor = ws_f.cell(row=rng.min_row, column=rng.min_col).value
        if rng.min_row == 1 and isinstance(anchor, str) and anchor.startswith("Server MoveIn"):
            movein_cols = [(c, parser.parse(ws.cell(row=2, column=c).value,
                                            "summary", f"{get_column_letter(c)}2"))
                           for c in range(rng.min_col, rng.max_col + 1)]
            movein_cols = [(c, m) for c, m in movein_cols if m]
            break
    if not movein_cols:
        issues.append(ImportIssue("warning", "summary", "",
                                  "summary 找不到 Server MoveIn 區塊,進機計畫視為 0"))

    pools: list[Pool] = []
    currents: list[CurrentStock] = []
    moveins: list[MoveIn] = []
    row = 3
    while ws.cell(row=row, column=1).value is not None:
        pool = Pool(fab=str(ws.cell(row=row, column=1).value),
                    bm_group=str(ws.cell(row=row, column=2).value))
        pools.append(pool)
        cur_coord = f"{get_column_letter(current_col)}{row}"
        cur = _numeric(ws.cell(row=row, column=current_col).value, issues,
                       "summary", cur_coord)
        currents.append(CurrentStock(pool=pool, sku_name=default_sku.name,
                                     count=int(cur)))
        _warn_if_formula(ws_f, row, current_col, issues)
        for col, month in movein_cols:
            coord = f"{get_column_letter(col)}{row}"
            cnt = _numeric(ws.cell(row=row, column=col).value, issues, "summary", coord)
            if cnt:
                moveins.append(MoveIn(pool=pool, sku_name=default_sku.name,
                                      month=month, count=int(cnt)))
            months.add(month)
            _warn_if_formula(ws_f, row, col, issues)
        row += 1

    return PlanInput(
        skus={default_sku.name: default_sku},
        months=sorted(months),
        pools=pools,
        demands=demands, vm_demands=[],
        moveins=moveins, returns=returns, currents=currents,
        issues=issues)


def _warn_if_formula(ws_f, row: int, col: int, issues: list[ImportIssue]) -> None:
    value = ws_f.cell(row=row, column=col).value
    if isinstance(value, str) and value.startswith("="):
        coord = f"{get_column_letter(col)}{row}"
        issues.append(ImportIssue(
            "warning", ws_f.title, coord,
            f"{coord} 為手填區卻含公式 {value},請確認是否錯位(已採計算後數值)"))
```

- [ ] **Step 5: 跑測試確認通過**

Run: `.venv\Scripts\python.exe -m pytest tests/test_importer_legacy.py -v`
Expected: 6 passed

- [ ] **Step 6: Commit**

```powershell
git add captool/importer.py tests/fixtures/sample_legacy.xlsx tests/test_importer_legacy.py
git commit -m "feat: legacy Excel importer with validation issues"
```

---

### Task 8: 黃金測試(對照現行 Excel)

**Files:**
- Test: `tests/test_golden.py`

**Interfaces:**
- Consumes: `import_legacy`、`run_check(mode="excel_compat")`、`NaiveSolver`

- [ ] **Step 1: 寫黃金測試**

`tests/test_golden.py`:
```python
"""黃金測試:excel_compat 模式必須與現行 Excel 公式逐格一致(遷移正確性)。"""
from pathlib import Path

import pytest
from openpyxl import load_workbook

from captool.importer import import_legacy
from captool.models import Pool
from captool.planner import run_check
from captool.solver.naive import NaiveSolver

FIXTURE = Path(__file__).parent / "fixtures" / "sample_legacy.xlsx"

# Excel 只算到 2026-09(AD 欄);AB=2026-07 起
MONTH_COLS = {"2026-07": "AB", "2026-08": "AC", "2026-09": "AD"}


def test_instock_matches_excel_cached_values():
    plan_input = import_legacy(FIXTURE)
    result = run_check(plan_input, NaiveSolver(), mode="excel_compat")
    wb = load_workbook(FIXTURE, data_only=True)
    ws = wb["summary"]
    checked = 0
    for row in range(3, 9):
        pool = Pool(fab=str(ws[f"A{row}"].value), bm_group=str(ws[f"B{row}"].value))
        for month, col in MONTH_COLS.items():
            expected = ws[f"{col}{row}"].value
            outcome = next(o for o in result.outcomes
                           if o.pool == pool and o.month == month)
            assert outcome.stock_float == pytest.approx(expected), (
                f"{pool.label} {month}: 工具 {outcome.stock_float} != Excel {expected}")
            checked += 1
    assert checked == 18


def test_reproduces_negative_stock_gap():
    plan_input = import_legacy(FIXTURE)
    result = run_check(plan_input, NaiveSolver(), mode="excel_compat")
    b1 = Pool(fab="B", bm_group="network1")
    o7 = next(o for o in result.for_pool(b1) if o.month == "2026-07")
    assert o7.stock_float == pytest.approx(-7.36328125)
    assert not o7.feasible
```

- [ ] **Step 2: 跑測試**

Run: `.venv\Scripts\python.exe -m pytest tests/test_golden.py -v`
Expected: 2 passed(若 FAIL,是 importer 或 planner 的 bug — 用 systematic-debugging 查,禁止直接改期望值遷就實作)

- [ ] **Step 3: 跑全部測試**

Run: `.venv\Scripts\python.exe -m pytest -v`
Expected: 全部通過

- [ ] **Step 4: Commit**

```powershell
git add tests/test_golden.py
git commit -m "test: golden test against legacy Excel cached values"
```

---

### Task 9: 匯出(summary 報表 + v2 範本)

**Files:**
- Create: `captool/exporter.py`
- Test: `tests/test_exporter.py`

**Interfaces:**
- Consumes: `PlanInput`、`PlanResult`、`MonthOutcome`、models
- Produces:
  - `export_summary(plan_input, plan_result, path) -> None`:寫出 `summary` sheet(長表:Fab、BM Group、月份、需求 vcore、VM 顆數、進機、退回、新啟用機台、建議採購、月末總機台、月末空機、剩餘可售 vcore、狀態、缺口 vcore;dict 欄位以 `"sku×n; …"` 字串呈現;缺口列 `狀態="缺口"` 並整列紅底)與 `placements` sheet(Fab、BM Group、月份、機型、新啟用台數)。全檔 Arial 字型、首列粗體、只寫值不寫公式,附註列說明「本表由容量規劃工具產出,數值為 solver 計算結果」
  - `generate_v2_template(plan_input, path) -> None`:自 legacy `PlanInput` 產生 v2 範本,sheet 與版面(v2 匯入的解析依據,Task 10 必須照此讀):
    - 各廠區 tab(以 pool 的 fab 值命名):需求區同 legacy(`C3` 標頭、月份 row 4 自 C 欄、資料 row 5 起 A=product、B=group);Return 區 `M3` 標頭改為含機型 — `K`=product、`L`=group、`M`=機型(row 4 標 `機型`)、月份 row 4 自 `N`(14) 欄向右;既有 return 資料帶入且機型填 default sku 名
    - `HW_SKU`:row 1 標頭 `name, vcore_per_node, usable_ratio`,row 2 起資料(含 default sku)
    - `HW_Current`:標頭 `fab, bm_group, sku, count`
    - `HW_MoveIn`:標頭 `fab, bm_group, sku, month, count`(month 為 `YYYY-MM` 字串)
    - `VM_Spec`:標頭 `fab, bm_group, product, month, vm_size_vcore, count` + 一列示範資料(淺黃底,附註為範例可刪)
    - `README`:填表說明(哪些區塊誰維護、月份格式、機型須存在於 HW_SKU)

- [ ] **Step 1: 寫失敗測試**

`tests/test_exporter.py`:
```python
from pathlib import Path

from openpyxl import load_workbook

from captool.exporter import export_summary, generate_v2_template
from captool.importer import import_legacy
from captool.planner import run_check
from captool.solver.naive import NaiveSolver

FIXTURE = Path(__file__).parent / "fixtures" / "sample_legacy.xlsx"


def test_export_summary(tmp_path):
    pi = import_legacy(FIXTURE)
    result = run_check(pi, NaiveSolver(), mode="conservative")
    out = tmp_path / "out.xlsx"
    export_summary(pi, result, out)
    wb = load_workbook(out)
    assert "summary" in wb.sheetnames and "placements" in wb.sheetnames
    ws = wb["summary"]
    headers = [c.value for c in ws[1]]
    assert "狀態" in headers and "缺口 vcore" in headers
    # 6 pools × 6 months = 36 資料列(+1 標頭 +1 附註)
    data_rows = [r for r in ws.iter_rows(min_row=2) if r[0].value in ("A", "B", "C")]
    assert len(data_rows) == 36


def test_export_summary_marks_gap(tmp_path):
    pi = import_legacy(FIXTURE)
    result = run_check(pi, NaiveSolver(), mode="conservative")
    out = tmp_path / "out.xlsx"
    export_summary(pi, result, out)
    ws = load_workbook(out)["summary"]
    statuses = {(r[0].value, r[1].value, r[2].value): r for r in ws.iter_rows(min_row=2)
                if r[0].value in ("A", "B", "C")}
    header = [c.value for c in ws[1]]
    status_idx = header.index("狀態")
    # B/network1 2026-07 在 Excel 中即為負庫存 → 工具也應標缺口
    assert statuses[("B", "network1", "2026-07")][status_idx].value == "缺口"


def test_generate_v2_template(tmp_path):
    pi = import_legacy(FIXTURE)
    out = tmp_path / "template_v2.xlsx"
    generate_v2_template(pi, out)
    wb = load_workbook(out)
    for sheet in ("HW_SKU", "HW_Current", "HW_MoveIn", "VM_Spec", "README", "A", "B", "C"):
        assert sheet in wb.sheetnames, sheet
    ws = wb["HW_SKU"]
    assert [c.value for c in ws[1]] == ["name", "vcore_per_node", "usable_ratio"]
    assert ws["A2"].value == "default-64"
    # 廠區 tab Return 區含機型欄
    wa = wb["A"]
    assert wa["M4"].value == "機型"
    assert wa["M5"].value == "default-64"
    # v2 範本可被 detect_format 認出
    from captool.importer import detect_format
    assert detect_format(out) == "v2"
```

- [ ] **Step 2: 跑測試確認失敗**

Run: `.venv\Scripts\python.exe -m pytest tests/test_exporter.py -v`
Expected: FAIL(ModuleNotFoundError)

- [ ] **Step 3: 實作 `captool/exporter.py`**

```python
"""匯出:summary 報表(工具計算結果)與 v2 範本(多機型格式)。"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from captool.models import PlanInput
from captool.planner import PlanResult

_ARIAL = Font(name="Arial")
_BOLD = Font(name="Arial", bold=True)
_RED = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
_YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

_SUMMARY_HEADERS = ["Fab", "BM Group", "月份", "需求 vcore", "VM 顆數", "進機",
                    "退回", "新啟用機台", "建議採購", "月末總機台", "月末空機",
                    "剩餘可售 vcore", "狀態", "缺口 vcore"]


def _fmt(d: dict[str, int]) -> str:
    return "; ".join(f"{k}×{v}" for k, v in sorted(d.items())) if d else ""


def _style_row(ws, row_idx: int, fill=None) -> None:
    for cell in ws[row_idx]:
        cell.font = _ARIAL
        if fill is not None:
            cell.fill = fill


def export_summary(plan_input: PlanInput, plan_result: PlanResult, path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "summary"
    ws.append(_SUMMARY_HEADERS)
    _style_row(ws, 1)
    for cell in ws[1]:
        cell.font = _BOLD
    for o in plan_result.outcomes:
        ws.append([
            o.pool.fab, o.pool.bm_group, o.month, o.demand_vcore,
            sum(c for _, c in o.vm_demand),
            _fmt(o.movein), _fmt(o.returns), _fmt(o.machines_opened),
            _fmt(o.suggested_purchases), _fmt(o.stock_total), _fmt(o.stock_empty),
            round(o.free_vcore_total, 2),
            "OK" if o.feasible else "缺口",
            round(o.shortfall_vcore, 2) if not o.feasible else 0,
        ])
        _style_row(ws, ws.max_row, fill=None if o.feasible else _RED)
    ws.append([])
    ws.append([f"本表由容量規劃工具產出(模式:{plan_result.mode}),數值為 solver 計算結果,非公式。"])
    _style_row(ws, ws.max_row)

    wp = wb.create_sheet("placements")
    wp.append(["Fab", "BM Group", "月份", "機型", "新啟用台數"])
    _style_row(wp, 1)
    for cell in wp[1]:
        cell.font = _BOLD
    for o in plan_result.outcomes:
        for sku_name, cnt in sorted(o.machines_opened.items()):
            wp.append([o.pool.fab, o.pool.bm_group, o.month, sku_name, cnt])
            _style_row(wp, wp.max_row)
    wb.save(path)


def generate_v2_template(plan_input: PlanInput, path) -> None:
    default_sku_name = next(iter(plan_input.skus))
    wb = Workbook()

    # README(第一個 sheet)
    ws = wb.active
    ws.title = "README"
    lines = [
        "容量規劃 v2 範本填表說明",
        "",
        "各廠區 tab:PM 維護需求區(C 欄起,每月新增 vcore);Return 區(K 欄起)含機型欄。",
        "HW_SKU:機型目錄(平台/硬體 team 維護)。",
        "HW_Current:各 pool 期初完全空置機台數(per 機型)。",
        "HW_MoveIn:進機計畫(硬體 team 維護),月份格式 YYYY-MM。",
        "VM_Spec:需要 VM 規格明細的 product 才填(size 為單顆 vcore 數)。",
        "注意:機型名稱必須存在於 HW_SKU;月份一律 YYYY-MM。",
    ]
    for line in lines:
        ws.append([line])
    for row in ws.iter_rows():
        for cell in row:
            cell.font = _ARIAL
    ws["A1"].font = _BOLD

    # 廠區 tabs
    fabs = sorted({p.fab for p in plan_input.pools})
    for fab in fabs:
        wf = wb.create_sheet(fab)
        wf["C3"] = "User Demand (vcore)"
        wf["M3"] = "Server Return"
        wf["A4"], wf["B4"] = "Product", "BM Group"
        wf["K4"], wf["L4"], wf["M4"] = "Product", "BM Group", "機型"
        for i, month in enumerate(plan_input.months):
            wf.cell(row=4, column=3 + i, value=month)       # 需求月份 C4 起
            wf.cell(row=4, column=14 + i, value=month)      # Return 月份 N4 起
        # 帶入既有需求
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
            col = 3 + plan_input.months.index(d.month)
            wf.cell(row=rows[key], column=col, value=d.vcore)
        # 帶入既有 Return(機型填 default)
        fab_returns = [x for x in plan_input.returns if x.pool.fab == fab]
        rrows: dict[tuple[str, str], int] = {}
        r = 5
        for x in fab_returns:
            key = (x.product, x.pool.bm_group)
            if key not in rrows:
                rrows[key] = r
                wf.cell(row=r, column=11, value=x.product)
                wf.cell(row=r, column=12, value=x.pool.bm_group)
                wf.cell(row=r, column=13, value=x.sku_name)
                r += 1
            col = 14 + plan_input.months.index(x.month)
            wf.cell(row=rrows[key], column=col, value=x.count)
        for row in wf.iter_rows():
            for cell in row:
                cell.font = _ARIAL
        wf["C3"].font = _BOLD
        wf["M3"].font = _BOLD

    # HW_SKU
    ws1 = wb.create_sheet("HW_SKU")
    ws1.append(["name", "vcore_per_node", "usable_ratio"])
    for sku in plan_input.skus.values():
        ws1.append([sku.name, sku.vcore_per_node, sku.usable_ratio])
    # HW_Current
    ws2 = wb.create_sheet("HW_Current")
    ws2.append(["fab", "bm_group", "sku", "count"])
    for c in plan_input.currents:
        ws2.append([c.pool.fab, c.pool.bm_group, c.sku_name, c.count])
    # HW_MoveIn
    ws3 = wb.create_sheet("HW_MoveIn")
    ws3.append(["fab", "bm_group", "sku", "month", "count"])
    for m in plan_input.moveins:
        ws3.append([m.pool.fab, m.pool.bm_group, m.sku_name, m.month, m.count])
    # VM_Spec(含示範列)
    ws4 = wb.create_sheet("VM_Spec")
    ws4.append(["fab", "bm_group", "product", "month", "vm_size_vcore", "count"])
    example_fab = fabs[0] if fabs else "A"
    example_month = plan_input.months[0] if plan_input.months else "2026-07"
    ws4.append([example_fab, "network1", "(示範列,請刪除後填入實際資料)",
                example_month, 32, 2])
    for cell in ws4[2]:
        cell.fill = _YELLOW
    for sheet in (ws1, ws2, ws3, ws4):
        for row in sheet.iter_rows():
            for cell in row:
                cell.font = _ARIAL
        for cell in sheet[1]:
            cell.font = _BOLD
    wb.save(path)
```

- [ ] **Step 4: 跑測試確認通過**

Run: `.venv\Scripts\python.exe -m pytest tests/test_exporter.py -v`
Expected: 3 passed

- [ ] **Step 5: Commit**

```powershell
git add captool/exporter.py tests/test_exporter.py
git commit -m "feat: summary export and v2 template generation"
```

---

### Task 10: v2 Excel 匯入

**Files:**
- Modify: `captool/importer.py`(實作 `import_v2`)
- Test: `tests/test_importer_v2.py`

**Interfaces:**
- Consumes: Task 9 的 `generate_v2_template`(測試用它產 fixture,確保匯出/匯入互為反函數)、Task 7 的工具函式
- Produces: `import_v2(path) -> PlanInput`;`import_any` 對 v2 檔案自動分派。解析規則照 Task 9 範本版面:
  - `HW_SKU` → `skus`(name 重複 → error issue 取第一筆;vcore/ratio 非數值 → error)
  - `HW_Current` / `HW_MoveIn` → `currents` / `moveins`;引用不存在的 sku → **error** issue 並跳過該列
  - `VM_Spec` → `vm_demands`;product 欄含「示範列」字樣的列跳過
  - 廠區 tab:需求區同 legacy;Return 區 K=product、L=group、M=機型、月份自 N 欄;機型空白 → warning 並以 HW_SKU 第一個機型計
  - months = 需求區 ∪ Return 區 ∪ HW_MoveIn ∪ VM_Spec 的月份聯集排序

- [ ] **Step 1: 寫失敗測試**

`tests/test_importer_v2.py`:
```python
from pathlib import Path

from openpyxl import load_workbook

from captool.exporter import generate_v2_template
from captool.importer import import_any, import_legacy, import_v2
from captool.models import Pool, Sku

FIXTURE = Path(__file__).parent / "fixtures" / "sample_legacy.xlsx"


def _v2_file(tmp_path):
    path = tmp_path / "v2.xlsx"
    generate_v2_template(import_legacy(FIXTURE), path)
    return path


def test_roundtrip_preserves_data(tmp_path):
    path = _v2_file(tmp_path)
    original = import_legacy(FIXTURE)
    pi = import_v2(path)
    a1 = Pool(fab="A", bm_group="network1")
    assert pi.demand_vcore(a1, "2026-07") == original.demand_vcore(a1, "2026-07")
    assert pi.return_by_sku(a1, "2026-08") == original.return_by_sku(a1, "2026-08")
    assert pi.current_by_sku(a1) == original.current_by_sku(a1)
    assert pi.movein_by_sku(a1, "2026-08") == original.movein_by_sku(a1, "2026-08")
    assert pi.months == original.months
    # 範本內的示範列不得變成資料
    assert pi.vm_demands == []


def test_import_any_dispatches(tmp_path):
    path = _v2_file(tmp_path)
    assert import_any(path).months == import_v2(path).months


def test_multi_sku_and_vm_spec(tmp_path):
    path = _v2_file(tmp_path)
    wb = load_workbook(path)
    wb["HW_SKU"].append(["big-128", 128, 0.8])
    wb["HW_Current"].append(["A", "network1", "big-128", 3])
    wb["HW_MoveIn"].append(["A", "network1", "big-128", "2026-09", 5])
    ws = wb["VM_Spec"]
    ws.delete_rows(2)  # 移除示範列
    ws.append(["A", "network1", "Product Apple", "2026-08", 32, 4])
    wb.save(path)
    pi = import_v2(path)
    assert set(pi.skus) == {"default-64", "big-128"}
    a1 = Pool(fab="A", bm_group="network1")
    assert pi.current_by_sku(a1) == {"default-64": 50, "big-128": 3}
    assert pi.movein_by_sku(a1, "2026-09") == {"big-128": 5}
    assert pi.vm_batch(a1, "2026-08") == [(32, 4)]


def test_unknown_sku_reported_and_skipped(tmp_path):
    path = _v2_file(tmp_path)
    wb = load_workbook(path)
    wb["HW_MoveIn"].append(["A", "network1", "no-such-sku", "2026-09", 5])
    wb.save(path)
    pi = import_v2(path)
    errors = [i for i in pi.issues if i.severity == "error"]
    assert any("no-such-sku" in i.message for i in errors)
    assert pi.movein_by_sku(Pool(fab="A", bm_group="network1"), "2026-09") == {}
```

- [ ] **Step 2: 跑測試確認失敗**

Run: `.venv\Scripts\python.exe -m pytest tests/test_importer_v2.py -v`
Expected: FAIL(NotImplementedError)

- [ ] **Step 3: 實作 `import_v2`(取代 importer.py 的 stub)**

```python
def _read_table(ws, headers: list[str], issues, sheet: str):
    """讀長表:row 1 為標頭,row 2 起至全空列。回傳 list[dict]。"""
    actual = [c.value for c in ws[1][:len(headers)]]
    if actual != headers:
        issues.append(ImportIssue(
            "error", sheet, "A1", f"{sheet} 標頭應為 {headers},實際為 {actual}"))
        return []
    rows = []
    for r in range(2, ws.max_row + 1):
        values = [ws.cell(row=r, column=c).value for c in range(1, len(headers) + 1)]
        if all(v is None for v in values):
            continue
        rows.append({"_row": r, **dict(zip(headers, values))})
    return rows


def import_v2(path) -> PlanInput:
    issues: list[ImportIssue] = []
    parser = _MonthParser(issues)
    wb_v = load_workbook(path, data_only=True)

    for required in V2_SHEETS:
        if required not in wb_v.sheetnames:
            raise CapacityImportError(
                [ImportIssue("error", required, "", f"找不到 {required} sheet")])

    # HW_SKU
    skus: dict[str, Sku] = {}
    for rec in _read_table(wb_v["HW_SKU"], ["name", "vcore_per_node", "usable_ratio"],
                           issues, "HW_SKU"):
        name = str(rec["name"])
        if name in skus:
            issues.append(ImportIssue("error", "HW_SKU", f"A{rec['_row']}",
                                      f"機型 {name} 重複定義,僅取第一筆"))
            continue
        try:
            skus[name] = Sku(name=name, vcore_per_node=int(rec["vcore_per_node"]),
                             usable_ratio=float(rec["usable_ratio"]))
        except (TypeError, ValueError):
            issues.append(ImportIssue("error", "HW_SKU", f"A{rec['_row']}",
                                      f"機型 {name} 的參數非數值,已跳過"))
    if not skus:
        raise CapacityImportError(
            [ImportIssue("error", "HW_SKU", "", "HW_SKU 無有效機型")])
    first_sku_name = next(iter(skus))

    def valid_sku(name, sheet, row) -> str | None:
        if name is None or str(name) not in skus:
            issues.append(ImportIssue(
                "error", sheet, f"C{row}",
                f"機型 '{name}' 不存在於 HW_SKU,該列已跳過"))
            return None
        return str(name)

    months: set[str] = set()

    # HW_Current / HW_MoveIn
    currents: list[CurrentStock] = []
    for rec in _read_table(wb_v["HW_Current"], ["fab", "bm_group", "sku", "count"],
                           issues, "HW_Current"):
        sku_name = valid_sku(rec["sku"], "HW_Current", rec["_row"])
        if sku_name is None:
            continue
        pool = Pool(fab=str(rec["fab"]), bm_group=str(rec["bm_group"]))
        currents.append(CurrentStock(pool=pool, sku_name=sku_name,
                                     count=int(rec["count"] or 0)))
    moveins: list[MoveIn] = []
    for rec in _read_table(wb_v["HW_MoveIn"],
                           ["fab", "bm_group", "sku", "month", "count"],
                           issues, "HW_MoveIn"):
        sku_name = valid_sku(rec["sku"], "HW_MoveIn", rec["_row"])
        if sku_name is None:
            continue
        month = parser.parse(rec["month"], "HW_MoveIn", f"D{rec['_row']}")
        if month is None:
            continue
        months.add(month)
        pool = Pool(fab=str(rec["fab"]), bm_group=str(rec["bm_group"]))
        moveins.append(MoveIn(pool=pool, sku_name=sku_name, month=month,
                              count=int(rec["count"] or 0)))

    # VM_Spec
    vm_demands: list[VmSpecDemand] = []
    for rec in _read_table(wb_v["VM_Spec"],
                           ["fab", "bm_group", "product", "month",
                            "vm_size_vcore", "count"], issues, "VM_Spec"):
        product = str(rec["product"])
        if "示範列" in product:
            continue
        month = parser.parse(rec["month"], "VM_Spec", f"D{rec['_row']}")
        if month is None:
            continue
        months.add(month)
        pool = Pool(fab=str(rec["fab"]), bm_group=str(rec["bm_group"]))
        vm_demands.append(VmSpecDemand(
            pool=pool, product=product, month=month,
            vm_size_vcore=int(rec["vm_size_vcore"]), count=int(rec["count"] or 0)))

    # 廠區 tabs:需求區同 legacy;Return 區 K/L/M + 月份自 N(14)
    demands: list[DemandDelta] = []
    returns: list[NodeReturn] = []
    fab_sheets = [n for n in wb_v.sheetnames if n not in NON_FAB_SHEETS]
    for name in fab_sheets:
        ws = wb_v[name]
        demand_cols = _month_columns(ws, 4, 3, parser, name)
        row = 5
        while ws.cell(row=row, column=1).value is not None:
            product = str(ws.cell(row=row, column=1).value)
            group = str(ws.cell(row=row, column=2).value)
            pool = Pool(fab=name, bm_group=group)
            for col, month in demand_cols:
                coord = f"{get_column_letter(col)}{row}"
                vcore = _numeric(ws.cell(row=row, column=col).value, issues, name, coord)
                if vcore:
                    demands.append(DemandDelta(pool=pool, product=product,
                                               month=month, vcore=vcore))
                months.add(month)
            row += 1
        return_cols = _month_columns(ws, 4, 14, parser, name)
        row = 5
        while ws.cell(row=row, column=11).value is not None:
            product = str(ws.cell(row=row, column=11).value)
            group = str(ws.cell(row=row, column=12).value)
            raw_sku = ws.cell(row=row, column=13).value
            if raw_sku is None:
                issues.append(ImportIssue(
                    "warning", name, f"M{row}",
                    f"M{row} Return 機型空白,以 {first_sku_name} 計"))
                sku_name = first_sku_name
            else:
                sku_name = valid_sku(raw_sku, name, row)
                if sku_name is None:
                    row += 1
                    continue
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

    pools = sorted({c.pool for c in currents}
                   | {d.pool for d in demands}
                   | {m.pool for m in moveins},
                   key=lambda p: (p.fab, p.bm_group))
    return PlanInput(skus=skus, months=sorted(months), pools=pools,
                     demands=demands, vm_demands=vm_demands,
                     moveins=moveins, returns=returns, currents=currents,
                     issues=issues)
```

同時把檔案頂端 import 補上 `VmSpecDemand`。

- [ ] **Step 4: 跑全部測試確認通過**

Run: `.venv\Scripts\python.exe -m pytest -v`
Expected: 全部通過(含既有測試不回歸)

- [ ] **Step 5: Commit**

```powershell
git add captool/importer.py tests/test_importer_v2.py
git commit -m "feat: v2 multi-SKU Excel importer"
```

---

### Task 11: Viewmodel + Streamlit UI

**Files:**
- Create: `captool/viewmodel.py`、`app/streamlit_app.py`、`.claude/launch.json`
- Test: `tests/test_viewmodel.py`

**Interfaces:**
- Consumes: 全部 core 模組
- Produces:
  - `captool/viewmodel.py`:
    - `overview_frame(plan_result, pool) -> pandas.DataFrame`(index=月份,欄:`需求 vcore`、`VM 顆數`、`進機`、`退回`、`新啟用機台`、`月末空機`、`剩餘可售 vcore`、`狀態`(OK/缺口)、`缺口 vcore`;excel_compat 時改含 `in-stock(台)`)
    - `placements_frame(plan_result, pool) -> DataFrame`(欄:`月份`、`機型`、`新啟用台數`、`建議採購台數`)
    - `movein_frame(plan_input) -> DataFrame`(欄:`fab`、`bm_group`、`sku`、`month`、`count`;供 st.data_editor 編輯)
    - `apply_movein_edits(plan_input, df) -> PlanInput`(以 df 全量取代 moveins,回傳新 PlanInput,不改原物件;sku 不存在或月份無法解析的列丟 `ValueError`,訊息含列號)
  - `app/streamlit_app.py`:sidebar = 檔案上傳 + 模式選擇(保守/填縫)+ 頁面 radio(匯入報告/總覽/驗證模式/回推模式/配置明細);`st.session_state["plan_input"]` 保存匯入結果

- [ ] **Step 1: 寫 viewmodel 失敗測試**

`tests/test_viewmodel.py`:
```python
from pathlib import Path

from captool.importer import import_legacy
from captool.models import Pool
from captool.planner import run_check
from captool.solver.naive import NaiveSolver
from captool.viewmodel import (apply_movein_edits, movein_frame,
                               overview_frame, placements_frame)

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


def test_apply_movein_edits_rejects_unknown_sku():
    import pytest
    pi, _ = _result()
    df = movein_frame(pi)
    df.loc[len(df)] = ["A", "network1", "no-such", "2026-10", 7]
    with pytest.raises(ValueError):
        apply_movein_edits(pi, df)
```

- [ ] **Step 2: 跑測試確認失敗**

Run: `.venv\Scripts\python.exe -m pytest tests/test_viewmodel.py -v`
Expected: FAIL(ModuleNotFoundError)

- [ ] **Step 3: 實作 `captool/viewmodel.py`**

```python
"""PlanResult / PlanInput → pandas DataFrame,供 Streamlit UI 使用。純函式。"""
import dataclasses

import pandas as pd

from captool.models import MoveIn, PlanInput, Pool
from captool.months import parse_month
from captool.planner import PlanResult


def _fmt(d: dict[str, int]) -> str:
    return "; ".join(f"{k}×{v}" for k, v in sorted(d.items())) if d else ""


def overview_frame(plan_result: PlanResult, pool: Pool) -> pd.DataFrame:
    rows = {}
    for o in plan_result.for_pool(pool):
        row = {
            "需求 vcore": o.demand_vcore,
            "VM 顆數": sum(c for _, c in o.vm_demand),
            "進機": _fmt(o.movein),
            "退回": _fmt(o.returns),
            "新啟用機台": _fmt(o.machines_opened),
            "月末空機": _fmt(o.stock_empty),
            "剩餘可售 vcore": round(o.free_vcore_total, 1),
            "狀態": "OK" if o.feasible else "缺口",
            "缺口 vcore": round(o.shortfall_vcore, 1) if not o.feasible else 0.0,
        }
        if plan_result.mode == "excel_compat":
            row["in-stock(台)"] = round(o.stock_float, 2)
        rows[o.month] = row
    return pd.DataFrame.from_dict(rows, orient="index")


def placements_frame(plan_result: PlanResult, pool: Pool) -> pd.DataFrame:
    records = []
    for o in plan_result.for_pool(pool):
        sku_names = sorted(set(o.machines_opened) | set(o.suggested_purchases))
        for sku_name in sku_names:
            records.append({
                "月份": o.month,
                "機型": sku_name,
                "新啟用台數": o.machines_opened.get(sku_name, 0),
                "建議採購台數": o.suggested_purchases.get(sku_name, 0),
            })
    return pd.DataFrame(records,
                        columns=["月份", "機型", "新啟用台數", "建議採購台數"])


def movein_frame(plan_input: PlanInput) -> pd.DataFrame:
    return pd.DataFrame(
        [{"fab": m.pool.fab, "bm_group": m.pool.bm_group, "sku": m.sku_name,
          "month": m.month, "count": m.count} for m in plan_input.moveins],
        columns=["fab", "bm_group", "sku", "month", "count"])


def apply_movein_edits(plan_input: PlanInput, df: pd.DataFrame) -> PlanInput:
    moveins: list[MoveIn] = []
    for i, rec in df.iterrows():
        if rec.isna().all():
            continue
        sku_name = str(rec["sku"])
        if sku_name not in plan_input.skus:
            raise ValueError(f"第 {i + 1} 列:機型 '{sku_name}' 不存在")
        month, _ = parse_month(rec["month"])
        if month is None:
            raise ValueError(f"第 {i + 1} 列:月份 '{rec['month']}' 無法解析")
        moveins.append(MoveIn(
            pool=Pool(fab=str(rec["fab"]), bm_group=str(rec["bm_group"])),
            sku_name=sku_name, month=month, count=int(rec["count"] or 0)))
    updated = dataclasses.replace(plan_input, moveins=moveins)
    months = set(updated.months) | {m.month for m in moveins}
    updated.months = sorted(months)
    return updated
```

- [ ] **Step 4: 跑測試確認通過**

Run: `.venv\Scripts\python.exe -m pytest tests/test_viewmodel.py -v`
Expected: 4 passed

- [ ] **Step 5: 實作 `app/streamlit_app.py`**

```python
"""容量規劃工具 UI。邏輯全在 captool,本檔僅做狀態管理與呈現。"""
import io
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from captool.exporter import export_summary, generate_v2_template
from captool.importer import CapacityImportError, import_any
from captool.models import Pool, Sku
from captool.planner import run_check, run_suggest
from captool.solver.naive import NaiveSolver
from captool.viewmodel import (apply_movein_edits, movein_frame,
                               overview_frame, placements_frame)

st.set_page_config(page_title="容量規劃工具", layout="wide")

MODE_LABELS = {"conservative": "保守(跨月零頭作廢)", "gap_fill": "填縫(零頭跨月可用)"}


def _load(uploaded) -> None:
    tmp = Path(st.session_state.get("_tmpdir", ".")) / "_uploaded.xlsx"
    tmp.write_bytes(uploaded.getvalue())
    try:
        st.session_state["plan_input"] = import_any(tmp)
        st.session_state["source_name"] = uploaded.name
    except CapacityImportError as e:
        st.session_state.pop("plan_input", None)
        st.session_state["fatal_issues"] = e.issues


with st.sidebar:
    st.title("容量規劃工具")
    uploaded = st.file_uploader("上傳 Excel(舊格式或 v2)", type=["xlsx"])
    if uploaded is not None and st.session_state.get("source_name") != uploaded.name:
        _load(uploaded)
    mode = st.radio("推演模式", list(MODE_LABELS), format_func=MODE_LABELS.get)
    page = st.radio("頁面", ["匯入報告", "總覽", "驗證模式", "回推模式", "配置明細"])

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
        st.error(text) if issue.severity == "error" else st.warning(text)

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
```

`.claude/launch.json`:
```json
{
  "version": "0.0.1",
  "configurations": [
    {
      "name": "capacity-ui",
      "runtimeExecutable": ".venv\\Scripts\\python.exe",
      "runtimeArgs": ["-m", "streamlit", "run", "app/streamlit_app.py",
                      "--server.headless", "true", "--server.port", "8501"],
      "port": 8501
    }
  ]
}
```

- [ ] **Step 6: 手動冒煙測試(用 Browser 工具)**

1. 以 preview_start 啟動 `capacity-ui`,開 `http://localhost:8501`
2. 上傳 `tests/fixtures/sample_legacy.xlsx`
3. 逐頁檢查:匯入報告列出 `20267'` 與 `T7` 警告;總覽出現 6 個 pool、B/network1 的 2026-07 標紅;驗證模式新增一列 movein(B, network1, default-64, 2026-07, 10)套用後缺口消失;回推模式按鈕產出建議計畫;配置明細有資料
4. 下載按鈕實際點擊,確認產出檔案可開啟

Expected: 全部行為符合;console 無紅字錯誤

- [ ] **Step 7: 跑全部測試 + Commit**

Run: `.venv\Scripts\python.exe -m pytest -v`
Expected: 全部通過

```powershell
git add captool/viewmodel.py app tests/test_viewmodel.py .claude/launch.json
git commit -m "feat: viewmodel and Streamlit five-page UI"
```

---

### Task 12: README 使用手冊 + 收尾

**Files:**
- Create: `README.md`

**Interfaces:**
- Consumes: 全部

- [ ] **Step 1: 撰寫 `README.md`**

內容必須涵蓋(用實際指令,不可含 TBD):
1. 專案目的(一段)與 spec 連結
2. 安裝:venv 建立 + `pip install -r requirements.txt`(用全路徑 Python)
3. 啟動 UI:`.venv\Scripts\python.exe -m streamlit run app\streamlit_app.py`
4. **每月校準節奏**(spec §10):月初以實際庫存/退回更新 Excel → 匯入 → 全部重推;工具無持久狀態,匯入即事實
5. 兩種推演模式的差異與適用場景(保守=採購安全邊際;填縫=真實利用率)
6. 舊格式 → v2 遷移:總覽頁下載 v2 範本 → 硬體 team 補機型資料 → 之後匯 v2
7. Solver 介面契約位置(`captool/solver/interface.py`)與同事 solver 的接入方式(實作 `AllocationSolver` Protocol,在 UI 替換 `NaiveSolver()`)
8. 測試:`.venv\Scripts\python.exe -m pytest`

- [ ] **Step 2: 最終驗證**

```powershell
.venv\Scripts\python.exe -m pytest -v
```
Expected: 全部通過

依 superpowers:verification-before-completion skill 逐項核對 spec §2 目標 1~5 是否皆有對應可運作功能。

- [ ] **Step 3: Commit**

```powershell
git add README.md
git commit -m "docs: README with usage and monthly re-baseline workflow"
```

---

## Self-Review 紀錄

- **Spec coverage:** §3 語意 → Task 2/3/5(constants 與模式);§5 資料模型 → Task 2;§6 兩種格式 → Task 7/9/10;§7 推演 → Task 5/6;§8 契約 → Task 3/4;§9 五頁 UI → Task 11;§10 校準 → Task 12 README + 無狀態架構;§11 測試 → Task 8 黃金測試 + 各 task TDD;§12 已知問題回報 → Task 7(T7 公式警告、月份 typo)。in-stock 未拉滿公式(§12-3)不需工具處理 — 工具自行推演全部月份,天然解決。
- **Placeholder scan:** 無 TBD/TODO;所有測試與實作皆含完整程式碼。
- **Type consistency:** `machines_opened`/`suggested_purchases`/`stock_float` 等欄位名稱已於 Task 3/5 定義並在 9/11 沿用;`import_v2` 依 Task 9 範本版面解析(roundtrip 測試鎖定一致性)。
