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
