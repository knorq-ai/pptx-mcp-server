"""KPIRow block component (Issue #133, v0.6.0).

Renders N KPI cells evenly distributed across a horizontal ``width`` with a
fixed ``gap`` between cells. Each cell stacks a short label (9pt), a large
value (26pt, bold, wrap=False auto-fit) and an optional detail line (8pt).

Typical consulting / IR use:

    with begin_container(slide, name="kpi_row", left=L, top=T, width=W, height=H):
        add_kpi_row(slide, kpis, left=L, top=T, width=W, height=H, theme="ir")

The component wraps its own ``begin_container(name="kpi_row")`` so child
shapes are validator-visible via ``check_containment``.

Theme tokens are resolved via a thin local helper so a caller passing
``theme=None`` still gets sensible defaults (matches the v0.6.0 atomic-
primitive convention established in #131).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List, Optional


__all__ = ["KPISpec", "add_kpi_row"]


@dataclass
class KPISpec:
    """Specification for a single KPI cell in a KPIRow.

    Attributes:
        label: Short descriptor above the value (e.g. ``"Revenue"``).
        value: Large-format numeric / text value (e.g. ``"107.8M"``,
            ``"+12.3%"``, ``"プレミアムモルツ +156%"``). Single-line with
            auto-fit shrink + ellipsis fallback.
        detail: Optional smaller caption below the value (e.g.
            ``"+12% QoQ"``). When empty, the detail textbox is not created
            and no vertical space is reclaimed — cells remain uniform.
        value_color: Theme token (e.g. ``"primary"``, ``"positive"``) or
            hex (e.g. ``"051C2C"``) for the value text.
        delta_color: Reserved for a future delta / arrow variant. Not
            consumed by the MVP renderer; kept on the dataclass so callers
            can pre-wire it without breaking a later minor bump.
    """

    label: str
    value: str
    detail: str = ""
    value_color: str = "primary"
    delta_color: str = "primary"  # reserved; unused in MVP


# Theme-token fallback map: used when ``theme=None`` so hardcoded token names
# inside the renderer still resolve to reasonable hex colors without a theme
# registry lookup. Keys mirror the v0.6.0 standard tokens (primary,
# text_secondary, highlight_row, rule_subtle). Values are ``#``-less 6-hex.
_FALLBACK_TOKENS = {
    "primary": "051C2C",
    "text_secondary": "666666",
    "highlight_row": "F8F9F5",  # cream
    "rule_subtle": "E0E0E0",
}


def _resolve_color(token_or_hex: str, theme: Optional[str]) -> str:
    """Resolve a theme token / raw hex to ``#``-less 6-hex.

    If ``theme`` is provided, delegate to the shared
    ``pptx_mcp_server.theme.resolve_theme_color`` so the caller's chosen
    palette is honored. Otherwise fall back to the built-in token map
    (primary, text_secondary, highlight_row, rule_subtle) so a no-theme
    caller still gets sensible defaults. Raw hex inputs pass through
    unchanged.

    The helper is defined locally (rather than pulled from theme.py) so
    the component's fallback behavior stays unit-test-stable even as the
    shared theme module evolves.
    """
    if theme:
        # Imported lazily to keep component-level import cost low and avoid
        # any theory-of-circularity issues during engine init.
        from pptx_mcp_server.theme import resolve_theme_color

        return resolve_theme_color(token_or_hex, theme)
    return _FALLBACK_TOKENS.get(token_or_hex, token_or_hex)


# Layout constants (inches) — tuned for a 0.95" cell height default. Y
# offsets inside a cell:
#
#   label  : [0.00 .. 0.22)     9pt  text_secondary   wrap=True
#   (gap)  : [0.22 .. 0.24)     (2 pt visual breathing room)
#   value  : [0.24 .. 0.74)     26pt bold value_color wrap=False auto-fit
#   detail : [0.76 .. 0.94)     8pt  text_secondary   wrap=True (optional)
#
# If the caller passes a taller ``height``, the stacked blocks sit at the
# top of the cell and the extra room appears below the detail line. This
# mirrors how existing card_row components behave (top-aligned stack).
_LABEL_H = 0.22
_LABEL_VALUE_GAP = 0.02
_VALUE_H = 0.50
_DETAIL_TOP_OFFSET = 0.76
_DETAIL_H = 0.18


def add_kpi_row(
    slide,
    kpis: List[KPISpec],
    *,
    left: float,
    top: float,
    width: float,
    height: float = 0.95,
    gap: float = 0.15,
    theme: Optional[str] = None,
    card_fill: Optional[str] = None,
    card_border: Optional[str] = None,
) -> Dict[str, Any]:
    """Render ``kpis`` as a horizontal row of N evenly distributed cells.

    Each cell width is ``(width - (n-1) * gap) / n``. When ``n == 1`` the
    single cell consumes the full ``width`` and ``gap`` is ignored.

    Args:
        slide: python-pptx slide object.
        kpis: list of ``KPISpec`` instances. Must be non-empty.
        left, top, width, height: cell-row bbox in inches.
        gap: horizontal spacing between cells in inches.
        theme: optional theme-registry key (``"ir"``, ``"mckinsey"``, …).
            When ``None``, hardcoded fallback colors are used (see
            ``_FALLBACK_TOKENS``).
        card_fill: optional theme token / hex for a per-cell rectangle
            fill. ``None`` means no fill. Typical value: ``"highlight_row"``.
        card_border: optional theme token / hex for a per-cell rectangle
            border. ``None`` means no border. Typical value: ``"rule_subtle"``.

    Returns:
        dict with keys:
            - ``cells``: per-cell dicts with ``bounds``, ``label_shape``,
              ``value_shape``, ``detail_shape`` (or ``None``),
              ``card_shape`` (or ``None``), and ``value_actual_font_size``
              (float; < 26 when auto-fit shrank the value).
            - ``consumed_height`` / ``consumed_width``: mirrors the input
              so callers composing vertical stacks know how much space
              the row consumed.

    Notes:
        The function wraps its children in ``begin_container(name="kpi_row")``
        so ``check_containment`` can verify nothing escapes the row bbox.
        All atomic-primitive shape calls pass ``theme=None`` because color
        tokens have already been resolved by ``_resolve_color``; letting
        the primitive resolve again would either double-resolve or fail.
    """
    # 循環 import 回避: components/ 配下は engine.shapes を lazy import する
    # (#135/#136 で確立した pattern を踏襲)。engine/__init__ の読み込み順に
    # 対する defensive measure で、直接的な循環は現状発生しないが、将来の
    # components/ 内相互依存 (#134 など) で顕在化しうるため予防する。
    from ..shapes import _add_shape, add_auto_fit_textbox
    from .container import begin_container

    n = len(kpis)
    if n <= 0:
        # Empty input: container still declared so container-aware callers
        # don't blow up, but no cells are rendered.
        with begin_container(
            slide, name="kpi_row",
            left=left, top=top, width=width, height=height,
        ):
            pass
        return {"cells": [], "consumed_height": height, "consumed_width": width}

    if n == 1:
        cell_w = width
    else:
        cell_w = (width - (n - 1) * gap) / n

    # Pre-resolve frame colors once; they're identical across cells.
    card_fill_hex = _resolve_color(card_fill, theme) if card_fill else None
    card_border_hex = _resolve_color(card_border, theme) if card_border else None

    cells: List[Dict[str, Any]] = []

    with begin_container(
        slide, name="kpi_row",
        left=left, top=top, width=width, height=height,
    ):
        for i, spec in enumerate(kpis):
            cell_x = left + i * (cell_w + gap) if n > 1 else left

            # 1) Card rectangle first, so text sits above it in z-order.
            card_shape = None
            if card_fill_hex or card_border_hex:
                idx = _add_shape(
                    slide,
                    "rectangle",
                    cell_x,
                    top,
                    cell_w,
                    height,
                    fill_color=card_fill_hex,
                    line_color=card_border_hex,
                    no_line=(card_border_hex is None),
                    # テーマは既に local _resolve_color で解決済み。ここで
                    # 再度 theme を渡すと resolve_theme_color が二重に走って
                    # hex をトークンとして誤解釈するリスクがあるため None.
                    theme=None,
                )
                card_shape = slide.shapes[idx]

            # 2) Label (9pt, text_secondary).
            label_shape, _ = add_auto_fit_textbox(
                slide,
                spec.label,
                cell_x,
                top,
                cell_w,
                _LABEL_H,
                font_size_pt=9,
                min_size_pt=7,
                color_hex=_resolve_color("text_secondary", theme),
                align="left",
                wrap=True,
                theme=None,
            )

            # 3) Value (26pt bold; single-line auto-fit via wrap=False).
            value_top = top + _LABEL_H + _LABEL_VALUE_GAP
            value_shape, value_size = add_auto_fit_textbox(
                slide,
                spec.value,
                cell_x,
                value_top,
                cell_w,
                _VALUE_H,
                font_size_pt=26,
                min_size_pt=14,
                bold=True,
                color_hex=_resolve_color(spec.value_color, theme),
                align="left",
                wrap=False,
                theme=None,
            )

            # 4) Detail (optional, 8pt).
            detail_shape = None
            if spec.detail:
                detail_shape, _ = add_auto_fit_textbox(
                    slide,
                    spec.detail,
                    cell_x,
                    top + _DETAIL_TOP_OFFSET,
                    cell_w,
                    _DETAIL_H,
                    font_size_pt=8,
                    min_size_pt=6,
                    color_hex=_resolve_color("text_secondary", theme),
                    align="left",
                    wrap=True,
                    theme=None,
                )

            cells.append({
                "bounds": {
                    "left": cell_x,
                    "top": top,
                    "width": cell_w,
                    "height": height,
                },
                "label_shape": label_shape,
                "value_shape": value_shape,
                "detail_shape": detail_shape,
                "card_shape": card_shape,
                "value_actual_font_size": float(value_size),
            })

    return {
        "cells": cells,
        "consumed_height": height,
        "consumed_width": width,
    }
