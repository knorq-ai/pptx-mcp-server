"""ASCII 印字可能文字の advance width を Liberation Sans で実測しバケット化する.

Issue #69 のための一回限りの較正ツールである。出力を `text_metrics.py` の
3-tier 定数に反映する。

使い方:
    python scripts/calibrate_ascii.py [font_path]

font_path を省略すると Liberation Sans / Arial 互換フォントを自動検索する。
`tests/calibration_helpers.py::advance_width_inches` を流用する。
"""

from __future__ import annotations

import os
import sys
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

    # 閾値
    NARROW_MAX = 0.0055
    WIDE_MIN = 0.009

    narrow = [(c, w) for c, w in widths if w < NARROW_MAX]
    normal = [(c, w) for c, w in widths if NARROW_MAX <= w < WIDE_MIN]
    wide = [(c, w) for c, w in widths if w >= WIDE_MIN]

    def fmt_bucket(name: str, items: list[tuple[str, float]]) -> None:
        if not items:
            print(f"{name}: (empty)")
            return
        chars = "".join(c for c, _ in items)
        ws = [w for _, w in items]
        print(f"--- {name} bucket ({len(items)} chars) ---")
        print(f"  members: {chars!r}")
        print(f"  min={min(ws):.5f}  mean={mean(ws):.5f}  max={max(ws):.5f}  in/pt")
        print()

    print("All measurements (advance width in inches/pt):")
    for c, w in sorted(widths, key=lambda t: t[1]):
        print(f"  {c!r:>5}  U+{ord(c):04X}  {w:.5f}")
    print()

    fmt_bucket("NARROW (<0.0055)", narrow)
    fmt_bucket("NORMAL [0.0055, 0.009)", normal)
    fmt_bucket("WIDE (>=0.009)", wide)

    # Python リテラルとして転記しやすい形で出力する。
    def as_frozenset(items: list[tuple[str, float]]) -> str:
        chars = "".join(c for c, _ in items)
        # バックスラッシュ・引用符はエスケープする
        return repr(chars)

    print("Suggested constants:")
    if narrow:
        print(f"  _ASCII_NARROW_WIDTH_PER_PT: float = {mean([w for _, w in narrow]):.5f}")
        print(f"  _ASCII_NARROW_CHARS = frozenset({as_frozenset(narrow)})")
    if normal:
        print(f"  _ASCII_NORMAL_WIDTH_PER_PT: float = {mean([w for _, w in normal]):.5f}")
    if wide:
        print(f"  _ASCII_WIDE_WIDTH_PER_PT:   float = {mean([w for _, w in wide]):.5f}")
        print(f"  _ASCII_WIDE_CHARS = frozenset({as_frozenset(wide)})")


if __name__ == "__main__":
    main()
