"""把 execution_plan 的 placement tree 畫成「AG 機櫃 → 實體機 → VM(標 product)」SVG。

一台機住多個 product 的 VM 時,框內會出現多種顏色 → 共用一眼可見。
tree[fab][network][ag][bm_id] = {sku, is_new, vms:[{"p":cluster,"v":vcore}, ...]}。
空 AG(有宣告但本月沒 VM)也會畫成空機櫃。純函式,不碰 solver。
"""
from __future__ import annotations

import html
import math

_PALETTE = ["#4e79a7", "#f28e2b", "#59a14f", "#e15759", "#b07aa1",
            "#76b7b2", "#edc948", "#ff9da7", "#9c755f", "#bab0ac",
            "#86bcb6", "#d37295", "#a0cbe8", "#8cd17d", "#b6992d"]

# 版面常數
_VMW, _VMH, _VMG, _VPR = 60, 26, 4, 4           # VM 格寬/高/間距/每列幾格
_BM_PAD, _BM_HEAD = 6, 16                        # 實體機內距 / 標頭高
_AG_PAD, _AG_HEAD, _AG_MINW = 8, 22, 96          # 機櫃內距 / 標頭高 / 空機櫃最小寬
_COL_GAP, _SEC_GAP, _SEC_HEAD = 16, 18, 26
_MARGIN, _LEGEND_H = 14, 30


def _esc(s) -> str:
    return html.escape(str(s))


def _trunc(s: str, n: int = 8) -> str:
    return s if len(s) <= n else s[: n - 1] + "…"


def _bm_size(bm: dict) -> "tuple[int, int]":
    rows = max(1, math.ceil(len(bm["vms"]) / _VPR))
    ncol = min(_VPR, max(1, len(bm["vms"])))
    w = _BM_PAD * 2 + ncol * _VMW + (ncol - 1) * _VMG
    h = _BM_HEAD + _BM_PAD + rows * (_VMH + _VMG)
    return w, h


def _draw_bm(x: int, y: int, bm: dict, color: dict) -> "tuple[str, int, int]":
    w, h = _bm_size(bm)
    stroke = "#e8873a" if bm["is_new"] else "#8a8a8a"
    tag = "新" if bm["is_new"] else "既有"
    parts = [f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="4" '
             f'fill="#ffffff" stroke="{stroke}" '
             f'stroke-width="{2 if bm["is_new"] else 1}"/>',
             f'<text x="{x + _BM_PAD}" y="{y + 12}" font-size="11" '
             f'font-weight="bold" fill="#333">{_esc(bm["sku"])} '
             f'<tspan fill="{stroke}">[{tag}]</tspan></text>']
    vx0, vy = x + _BM_PAD, y + _BM_HEAD + _BM_PAD
    for i, vm in enumerate(bm["vms"]):
        prod, vcore = vm["p"], vm["v"]
        cx = vx0 + (i % _VPR) * (_VMW + _VMG)
        cy = vy + (i // _VPR) * (_VMH + _VMG)
        mid = cx + _VMW / 2
        parts.append(
            f'<rect x="{cx}" y="{cy}" width="{_VMW}" height="{_VMH}" rx="2" '
            f'fill="{color.get(prod, "#ccc")}">'
            f'<title>{_esc(prod)} · {_esc(vcore)}vcore</title></rect>'
            f'<text x="{mid:.0f}" y="{cy + 12}" font-size="10" '
            f'text-anchor="middle" fill="#fff" font-weight="bold">'
            f'{_esc(_trunc(prod))}<title>{_esc(prod)}</title></text>'
            f'<text x="{mid:.0f}" y="{cy + 22}" font-size="8" '
            f'text-anchor="middle" fill="#ffffffcc">{_esc(vcore)}vcore</text>')
    return "\n".join(parts), w, h


def _draw_ag(x: int, y: int, ag: str, bms: dict, color: dict) -> "tuple[str, int, int]":
    frags, inner_w, cy = [], 0, y + _AG_HEAD
    for bm in bms.values():
        f, w, h = _draw_bm(x + _AG_PAD, cy, bm, color)
        frags.append(f)
        inner_w = max(inner_w, w)
        cy += h + _AG_PAD
    empty = not bms
    if empty:
        inner_w = _AG_MINW
        cy = y + _AG_HEAD + 30
    w = inner_w + _AG_PAD * 2
    h = cy - y
    nvm = sum(len(b["vms"]) for b in bms.values())
    head = (f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="6" '
            f'fill="{"#fafbfc" if empty else "#f4f5f7"}" '
            f'stroke="{"#dfe3e8" if empty else "#c2c7d0"}" '
            f'{"stroke-dasharray=\"4 3\" " if empty else ""}/>'
            f'<text x="{x + _AG_PAD}" y="{y + 15}" font-size="12" '
            f'font-weight="bold" fill="#333">AG={_esc(ag)} '
            f'<tspan font-weight="normal" fill="#888">'
            f'({nvm} VM / {len(bms)} 台)</tspan></text>')
    if empty:
        head += (f'<text x="{x + w / 2:.0f}" y="{y + _AG_HEAD + 20}" '
                 f'font-size="10" text-anchor="middle" fill="#b0b6be">空</text>')
    return head + "\n" + "\n".join(frags), w, h


def placement_svg(tree: dict, month: str = "") -> str:
    """回單一 <svg>:各 (fab, network) 一個區塊,內含 AG 機櫃並排(含空 AG)。"""
    products = sorted({vm["p"] for nets in tree.values() for ags in nets.values()
                       for bms in ags.values() for bm in bms.values()
                       for vm in bm["vms"]})
    color = {p: _PALETTE[i % len(_PALETTE)] for i, p in enumerate(products)}

    body, max_w, y = [], 600, _MARGIN
    legend = [f'<text x="{_MARGIN}" y="{y + 14}" font-size="12" '
              f'font-weight="bold" fill="#333">需求單機櫃圖 {_esc(month)}</text>']
    lx, ly = _MARGIN, y + _LEGEND_H
    for p in products:
        legend.append(
            f'<rect x="{lx}" y="{ly - 11}" width="13" height="13" rx="2" '
            f'fill="{color[p]}"/>'
            f'<text x="{lx + 18}" y="{ly}" font-size="11" fill="#333">'
            f'{_esc(p)}</text>')
        lx += 20 + 8 * (len(p) + 1) + 12
    body.append("\n".join(legend))
    y = ly + 10

    for fab in sorted(tree):
        for network in sorted(tree[fab]):
            ags = tree[fab][network]
            body.append(f'<text x="{_MARGIN}" y="{y + 16}" font-size="14" '
                        f'font-weight="bold" fill="#1f2d3d">'
                        f'Fab {_esc(fab)} / {_esc(network)}</text>')
            y += _SEC_HEAD
            x, row_h = _MARGIN, 0
            for ag in sorted(ags):
                f, w, h = _draw_ag(x, y, ag, ags[ag], color)
                body.append(f)
                x += w + _COL_GAP
                row_h = max(row_h, h)
            max_w = max(max_w, x)
            y += row_h + _SEC_GAP

    if not products:
        body.append(f'<text x="{_MARGIN}" y="60" font-size="13" fill="#888">'
                    f'（無 placement 資料）</text>')
    W, H = max_w + _MARGIN, y + _MARGIN
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" '
            f'viewBox="0 0 {W} {H}" font-family="Arial, sans-serif">'
            f'<rect width="{W}" height="{H}" fill="#ffffff"/>'
            + "\n".join(body) + "</svg>")
