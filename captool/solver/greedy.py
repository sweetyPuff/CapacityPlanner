"""共用的貪婪採購邏輯(suggest)。

抽出成獨立函式,讓 NaiveSolver 與 CpSatAdapter 共用同一套「逐輪補機直到可行」的
回推流程 —— 兩者只差在 check 的實作(內建 FFD vs 委派 CP-SAT),貪婪外殼相同。
"""
import copy
import math

from captool.models import Sku
from captool.solver.interface import (CheckResult, DemandBatch, PoolState,
                                      SuggestResult)

_EPS = 1e-9
MAX_TOTAL_PURCHASE = 100_000


def greedy_suggest(check_fn, pool_state: PoolState, new_demand: DemandBatch,
                   catalog: list[Sku], max_total: int = MAX_TOTAL_PURCHASE
                   ) -> SuggestResult:
    """以缺口量 ÷ 最大機型可售 vcore 估算加購台數,重試 check_fn 直到可行。

    check_fn(pool_state, demand) -> CheckResult:就地配置的可行性檢查。
    可行後把採購機器併入 pool_state 並回傳配置結果。
    VM 大於任何機型單台可售容量時回傳不可行 + 空採購(不無窮迴圈);catalog 空丟 ValueError。
    """
    if not catalog:
        raise ValueError("catalog 不可為空")
    best = max(catalog, key=lambda s: s.sellable_vcore)
    max_vm = max((size for size, _ in new_demand.atomic_vms), default=0)
    if max_vm > best.sellable_vcore + _EPS:
        result = check_fn(pool_state, new_demand)
        return SuggestResult(purchases={}, result=result)

    purchases: dict[str, int] = {}
    while True:
        trial = copy.deepcopy(pool_state)
        for name, cnt in purchases.items():
            sku = next(s for s in catalog if s.name == name)
            trial.add_empty(sku, cnt)
        result: CheckResult = check_fn(trial, new_demand)
        if result.feasible:
            pool_state.machines[:] = trial.machines
            return SuggestResult(purchases=purchases, result=result)
        need = max(1, math.ceil(result.shortfall_vcore / best.sellable_vcore))
        purchases[best.name] = purchases.get(best.name, 0) + need
        if sum(purchases.values()) > max_total:
            raise RuntimeError("suggest 未收斂:採購量超過上限")
