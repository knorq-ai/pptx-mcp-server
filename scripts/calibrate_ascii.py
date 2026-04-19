"""ASCII 印字可能文字の advance width を Liberation Sans で実測しバケット化する.

Issue #71 のための再較正ツールである (#69/#70 の後続)。出力を
`text_metrics.py` の多段バケット定数に反映する。

使い方:
    python scripts/calibrate_ascii.py [font_path]

font_path を省略すると Liberation Sans / Arial 互換フォントを自動検索する。
`tests/calibration_helpers.py::advance_width_inches` を流用する。

#71 対応:
- 全 ASCII printable (0x20–0x7E) の advance を測定した上で、
  「worst-case per-char 相対誤差を最小化する」バケット分け探索を行う。
- 3 バケット分割と 4 バケット分割を両方試し、
  worst-case が ±20% 未満に収まる最小バケット数を採用する。
- 各バケットの代表値は単純平均ではなく、そのバケットに属する文字の
  worst-case 相対誤差を最小化する値 (= (min + max) / 2 の調和平均的な補正)
  を採用する。厳密には worst-case 誤差を最小化する代表値は
  sqrt(min*max) ではなく (2*min*max) / (min+max) (調和平均) で与えられる
  ため、それを使う。
"""

from __future__ import annotations

import os
import sys
from itertools import combinations
from statistics import mean

# tests/ の calibration_helpers を import できるように sys.path を調整する。
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

from tests.calibration_helpers import advance_width_inches  # noqa: E402

_ARIAL_CANDIDATES = [
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    "/Users/yuyamorita/Projects/grants/node_modules/pdfjs-dist/standard_fonts/LiberationSans-Regular.ttf",
    "/Library/Fonts/Arial.ttf",
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "/Windows/Fonts/arial.ttf",
]


def _find_font() -> str:
    for p in _ARIAL_CANDIDATES:
        if os.path.exists(p):
            return p
    raise SystemExit("Liberation Sans / Arial not found. Install fonts-liberation.")


def _harmonic_mid(lo: float, hi: float) -> float:
    """2 値 [lo, hi] に対して worst-case 相対誤差を最小化する代表値を返す.

    err_at(c) = max(|c-lo|/lo, |c-hi|/hi)
    を最小化する c を閉形式で求めると c = 2*lo*hi/(lo+hi) (調和平均) となる。
    """
    if lo <= 0 or hi <= 0:
        return (lo + hi) / 2.0
    return 2.0 * lo * hi / (lo + hi)


def _worst_rel_err(repr_val: float, ws: list[float]) -> float:
    return max(abs(repr_val - w) / w for w in ws) if ws else 0.0


def _partition(widths: list[tuple[str, float]], cuts: list[float]) -> list[list[tuple[str, float]]]:
    """``cuts`` (昇順) を境界として ``widths`` を (k+1) バケットに分ける."""
    buckets: list[list[tuple[str, float]]] = [[] for _ in range(len(cuts) + 1)]
    for c, w in widths:
        idx = 0
        for cut in cuts:
            if w >= cut:
                idx += 1
            else:
                break
        buckets[idx].append((c, w))
    return buckets


def _search_best_partition(
    widths: list[tuple[str, float]], k: int
) -> tuple[list[list[tuple[str, float]]], list[float], list[float], float]:
    """k バケットの境界を全探索して worst-case rel_err を最小化する.

    境界候補: 実測 advance のソート済み一意値の隣接中点。
    (文字数 ~95、k=3 の場合 ~C(94,2) ≈ 4400 通り、k=4 でも ~1.3e5 通りと十分小さい。)
    """
    if k < 2:
        raise ValueError("k must be >= 2")
    sorted_ws = sorted({w for _, w in widths})
    # 隣接中点候補
    candidates = [
        (sorted_ws[i] + sorted_ws[i + 1]) / 2.0 for i in range(len(sorted_ws) - 1)
    ]
    best_worst = float("inf")
    best_cuts: list[float] = []
    best_buckets: list[list[tuple[str, float]]] = []
    best_reprs: list[float] = []
    for combo in combinations(candidates, k - 1):
        cuts = list(combo)
        buckets = _partition(widths, cuts)
        if any(not b for b in buckets):
            continue
        reprs: list[float] = []
        worst = 0.0
        for b in buckets:
            ws = [w for _, w in b]
            r = _harmonic_mid(min(ws), max(ws))
            reprs.append(r)
            worst = max(worst, _worst_rel_err(r, ws))
        if worst < best_worst:
            best_worst = worst
            best_cuts = cuts
            best_buckets = buckets
            best_reprs = reprs
    return best_buckets, best_cuts, best_reprs, best_worst


def _print_bucket(name: str, items: list[tuple[str, float]], repr_val: float) -> None:
    if not items:
        print(f"  {name}: (empty)")
        return
    chars = "".join(c for c, _ in items)
    ws = [w for _, w in items]
    worst = _worst_rel_err(repr_val, ws)
    print(f"  --- {name} ({len(items)} chars) ---")
    print(f"    members: {chars!r}")
    print(
        f"    min={min(ws):.5f}  mean={mean(ws):.5f}  max={max(ws):.5f}  in/pt"
    )
    print(f"    repr (harmonic mid) = {repr_val:.5f}  worst rel_err = {worst * 100:.2f}%")


def _print_char_table(
    widths: list[tuple[str, float]],
    buckets: list[list[tuple[str, float]]],
    reprs: list[float],
    size_pt: float = 12.0,
) -> None:
    """全文字 → 割当バケット → 相対誤差 の表を印字する."""
    # 文字 → バケット idx のマップを作る
    char_to_idx: dict[str, int] = {}
    for idx, b in enumerate(buckets):
        for c, _ in b:
            char_to_idx[c] = idx
    print(f"\nPer-char assignment table (at {size_pt}pt, measured vs bucket repr):")
    print(f"  {'char':>5}  {'code':>6}  {'measured':>9}  {'bucket':>6}  {'repr':>8}  {'rel_err':>8}")
    for c, w in sorted(widths, key=lambda t: t[1]):
        idx = char_to_idx[c]
        r = reprs[idx]
        measured_at_pt = w * size_pt
        est_at_pt = r * size_pt
        rel_err = abs(est_at_pt - measured_at_pt) / measured_at_pt
        print(
            f"  {c!r:>5}  U+{ord(c):04X}  {measured_at_pt:9.5f}  {idx:>6}  {est_at_pt:8.5f}  {rel_err * 100:7.2f}%"
        )


def main() -> None:
    font_path = sys.argv[1] if len(sys.argv) > 1 else _find_font()
    print(f"Font: {font_path}\n")

    # ASCII printable 0x20–0x7E を測る。advance(1pt) = inches/pt.
    widths: list[tuple[str, float]] = []
    for code in range(0x20, 0x7F):
        ch = chr(code)
        try:
            w = advance_width_inches(font_path, ch, 1.0)
        except KeyError:
            continue
        widths.append((ch, w))

    print(f"Measured {len(widths)} ASCII printable chars.\n")
    print("Raw sorted advance (in/pt):")
    for c, w in sorted(widths, key=lambda t: t[1]):
        print(f"  {c!r:>5}  U+{ord(c):04X}  {w:.5f}")
    print()

    # 3-bucket と 4-bucket を両方試す。
    for k in (3, 4):
        print(f"\n===== Best {k}-bucket partition (minimize worst per-char rel_err) =====")
        buckets, cuts, reprs, worst = _search_best_partition(widths, k)
        print(f"  cuts: {[f'{c:.5f}' for c in cuts]}")
        print(f"  worst-case rel_err across all chars: {worst * 100:.2f}%")
        bucket_names = ["VERY_NARROW", "NARROW", "NORMAL", "WIDE"]
        if k == 3:
            bucket_names = ["NARROW", "NORMAL", "WIDE"]
        for name, b, r in zip(bucket_names, buckets, reprs, strict=False):
            _print_bucket(name, b, r)
        _print_char_table(widths, buckets, reprs)

        if worst <= 0.20:
            print(f"\n  --> {k}-bucket partition satisfies <=20% per-char target.")
        else:
            print(f"\n  --> {k}-bucket partition does NOT satisfy <=20% target.")

        # Python リテラル出力。
        print("\nSuggested constants:")
        for name, b, r in zip(bucket_names, buckets, reprs, strict=False):
            chars = "".join(c for c, _ in b)
            # sort for readability
            chars_sorted = "".join(sorted(chars))
            print(f"  _ASCII_{name}_WIDTH_PER_PT: float = {r:.5f}")
            print(f"  _ASCII_{name}_CHARS = frozenset({chars_sorted!r})")


if __name__ == "__main__":
    main()
