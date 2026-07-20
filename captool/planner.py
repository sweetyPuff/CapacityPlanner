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
