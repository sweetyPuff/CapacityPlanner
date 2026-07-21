"""內建 naive solver:同事 solver 到位前的替身。
原子 VM 用 first-fit-decreasing(填縫優先),液體 vcore 填縫後開空機。
單一機型情境退化為「除法 + 無條件進位」。
"""
from captool.models import Sku
from captool.solver.greedy import MAX_TOTAL_PURCHASE, greedy_suggest
from captool.solver.interface import (CheckResult, DemandBatch, PoolState,
                                      SuggestResult)

_EPS = 1e-9


class NaiveSolver:

    MAX_TOTAL_PURCHASE = MAX_TOTAL_PURCHASE

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
        return greedy_suggest(self.check, pool_state, new_demand, catalog,
                              self.MAX_TOTAL_PURCHASE)
