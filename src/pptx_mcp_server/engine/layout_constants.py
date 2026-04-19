"""Shared layout constants for sizing and padding.

Centralized so that text-height estimation, auto-fit shrinking, card rendering,
and overflow validation all use the same numeric assumptions. Drift between
these call sites has historically caused false-positive overflow findings and
clipped text in generated decks.
"""

from __future__ import annotations

# 2026-04: 各 textbox の内側左右の padding (inches)。python-pptx / PowerPoint が
# 標準の textbox に適用する既定マージンに合わせて calibrate した値である。
TEXTBOX_INNER_PADDING_PER_SIDE: float = 0.05

# Convenience: 左右合計 (width calculations で頻繁に使うため事前計算してある)。
TEXTBOX_INNER_PADDING_TOTAL: float = 2 * TEXTBOX_INNER_PADDING_PER_SIDE
