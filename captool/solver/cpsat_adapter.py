"""CP-SAT solver adapter:把原子 VM 的配置委派給同事的 VM placement solver。

實作 AllocationSolver 契約(check / suggest),讓 planner / UI 完全不用改就能改用
CP-SAT 求解。設計決策(2026-07,單維 vcore):
  - 只有含原子 VM 的批次才呼叫 CP-SAT;純液體 vcore 沿用內建填縫(除法),不碰 solver。
  - 單維映射:machine 可用 vcore →(floor 取整,保守)baremetal 的 cpu_cores;
    VM size → vm.demand.cpu_cores。其餘維度(mem/disk/gpu)留 0。
  - 液體 vcore 與採購貪婪(suggest)仍由內建 / 共用邏輯處理 —— 她的 solver 無此概念。

依賴注入:solve_fn 接受 PlacementRequest dict、回傳 PlacementResult dict。這讓本模組
不硬依賴 ortools 或她的程式碼。生產環境用 http_solve_fn 或 in_process_solve_fn,測試用
假後端。上線前提:她的 solver 需能實際執行(見 README 記載的 solve()/server.py 待修項)。
"""
from __future__ import annotations

import math
from collections import defaultdict
from typing import Callable

from captool.solver.greedy import greedy_suggest
from captool.solver.interface import (CheckResult, DemandBatch, PoolState)
from captool.solver.naive import NaiveSolver

_EPS = 1e-9

SolveFn = Callable[[dict], dict]


class CpSatAdapter:
    """AllocationSolver 實作,原子 VM 配置委派給注入的 CP-SAT 後端。"""

    def __init__(self, solve_fn: SolveFn, max_solve_time_seconds: float = 3.0):
        self._solve_fn = solve_fn
        self._max_solve_time = max_solve_time_seconds
        self._naive = NaiveSolver()  # 液體 vcore 沿用內建填縫

    def check(self, pool_state: PoolState, new_demand: DemandBatch) -> CheckResult:
        # 快照:哪些機器在配置前為空(供 machines_opened 計算,涵蓋原子 + 液體)
        empty_before = [m.is_empty for m in pool_state.machines]

        blocked: list[tuple[int, int]] = []
        if new_demand.atomic_vms:
            blocked = self._place_atomic(pool_state, new_demand.atomic_vms)

        # 液體 vcore:在原子配置後的狀態上,沿用內建填縫邏輯
        liquid_result = self._naive.check(
            pool_state, DemandBatch(liquid_vcore=new_demand.liquid_vcore))

        opened: dict[str, int] = defaultdict(int)
        for was_empty, machine in zip(empty_before, pool_state.machines):
            if was_empty and not machine.is_empty:
                opened[machine.sku_name] += 1

        return CheckResult(
            feasible=(not blocked and liquid_result.unplaced_liquid_vcore <= _EPS),
            placed_liquid_vcore=liquid_result.placed_liquid_vcore,
            unplaced_liquid_vcore=liquid_result.unplaced_liquid_vcore,
            blocked_vms=blocked,
            machines_opened=dict(opened),
        )

    def suggest(self, pool_state, new_demand, catalog):
        return greedy_suggest(self.check, pool_state, new_demand, catalog)

    # ------------------------------------------------------------------

    def _place_atomic(self, pool_state: PoolState,
                      atomic_vms: list[tuple[int, int]]) -> list[tuple[int, int]]:
        """把原子 VM 丟給 CP-SAT 後端配置,就地扣減 free_vcore,回傳 blocked (size,count)。"""
        machines = pool_state.machines
        baremetals = []
        for i, m in enumerate(machines):
            # floor:整數裝箱下,零頭小數對整數 VM 無用,取整較保守(不超賣)
            avail = int(math.floor(m.free_vcore + _EPS))
            baremetals.append({
                "id": f"m{i}",
                "total_capacity": {"cpu_cores": avail},
                "used_capacity": {"cpu_cores": 0},
                "topology": {"ag": ""},
            })

        vms = []
        vm_size: dict[str, int] = {}
        k = 0
        for size, count in atomic_vms:
            for _ in range(count):
                vid = f"v{k}"
                k += 1
                vms.append({"id": vid, "demand": {"cpu_cores": int(size)}})
                vm_size[vid] = int(size)

        request = {
            "vms": vms,
            "baremetals": baremetals,
            "config": {
                "allow_partial_placement": True,     # 規劃只要可行性,盡量塞
                "auto_generate_anti_affinity": False,  # 容量規劃不建模 AG
                "max_solve_time_seconds": self._max_solve_time,
            },
        }
        result = self._solve_fn(request)

        # 套用配置:每個 assignment 對應機器扣減
        idx_of = {f"m{i}": i for i in range(len(machines))}
        for a in result.get("assignments", []):
            machines[idx_of[a["baremetal_id"]]].free_vcore -= vm_size[a["vm_id"]]

        # unplaced → blocked，依 size 聚合
        agg: dict[int, int] = defaultdict(int)
        for vid in result.get("unplaced_vms", []):
            agg[vm_size[vid]] += 1
        return sorted(agg.items())


# ---------------------------------------------------------------------------
# solve_fn 工廠:生產環境接線(import 延遲到呼叫時,模組載入不需 ortools)
# ---------------------------------------------------------------------------

def http_solve_fn(url: str, timeout: float = 30.0) -> SolveFn:
    """POST 到她的 FastAPI sidecar(/v1/placement/solve)。"""
    import json
    import urllib.request

    def _fn(request: dict) -> dict:
        data = json.dumps(request).encode("utf-8")
        req = urllib.request.Request(
            url, data=data, headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))

    return _fn


def in_process_solve_fn() -> SolveFn:
    """直接 import 她的 VMPlacementSolver(需 ortools、pydantic 及其 app/ 套件在 sys.path)。

    注意:目前她的 solve() 有未修的 NameError(needs_two_phase / soft_terms),
    server.py 也有 _SWAGGER_STATIC_DIR NameError —— 這兩者修好前此路徑無法運作。
    """
    def _fn(request: dict) -> dict:
        from app.models import PlacementRequest
        from app.solver import VMPlacementSolver
        req = PlacementRequest.model_validate(request)
        return VMPlacementSolver(req).solve().model_dump()

    return _fn
