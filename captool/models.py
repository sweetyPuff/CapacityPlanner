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
    cluster: str = ""          # 執行面需求單指定的 cluster;空=沿用 product


@dataclass(frozen=True)
class VmSpecDemand:
    pool: Pool
    product: str
    month: str
    vm_size_vcore: int
    count: int
    cluster: str = ""          # 執行面需求單指定的 cluster;空=沿用 product


@dataclass(frozen=True)
class MoveIn:
    pool: Pool
    sku_name: str
    month: str
    count: int
    ag: str = ""            # 選配:進機落點 AG;"" = 未定(solver 自選)


@dataclass(frozen=True)
class NodeReturn:
    pool: Pool
    product: str
    sku_name: str
    month: str
    count: int
    ag: str = ""            # 選配:退還機器所屬 AG


@dataclass(frozen=True)
class CurrentStock:
    pool: Pool
    sku_name: str
    count: int
    ag: str = ""            # 失效域(availability group);"" = 未指定


@dataclass(frozen=True)
class ImportIssue:
    severity: str  # "error" | "warning"
    sheet: str
    cell: str      # 例如 "T7";非特定儲存格用 ""
    message: str


@dataclass(frozen=True)
class ProductPolicy:
    pool: Pool
    product: str
    vm_size_vcore: int | None      # None = 粗粒度(液體)
    max_per_machine: int | None    # 1:X 的 X;None = 不限
    co_residency: str              # "free" | "exclusive" | 群組名


@dataclass(frozen=True)
class Cap:
    """每個 AG 的採購槽位上限(來自 HW_Caps)。pool = fab × bm_group(network)。"""
    pool: Pool
    ag: str
    max_bm: int


@dataclass(frozen=True)
class NewBuild:
    """fab 頁上 Menu 欄非 Worker 的列:某月在某 pool 用某菜單建幾個 cluster。
    cluster = 該列的 Product;menu 對應 Cluster_Menu 的菜單名。"""
    pool: Pool
    cluster: str
    menu: str
    month: str
    count: int


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
    policies: list[ProductPolicy] = field(default_factory=list)
    caps: list["Cap"] = field(default_factory=list)
    new_builds: list["NewBuild"] = field(default_factory=list)
    # 菜單目錄:menu 名 -> [(role, count, co_residency, vm_vcore)]
    menus: dict[str, list] = field(default_factory=dict)

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

    def policy_for(self, pool: Pool, product: str) -> "ProductPolicy | None":
        for p in self.policies:
            if p.pool == pool and p.product == product:
                return p
        return None
