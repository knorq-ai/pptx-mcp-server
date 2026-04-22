"""NumberedList block component — stacked numbered items with rules (Issue #134).

Renders N numbered items vertically inside a declared container. Each item
displays:

1. A number + caption row (e.g. ``01 ／ 主要ターゲット``)
2. A bold title summary
3. A wrapped body paragraph (optional — empty bodies skip the textbox)
4. A thin horizontal rule separating this item from the next (unless the
   item is last, or ``rule_between=False``)

Item heights are distributed **equally** across the declared ``height`` (MVP
interpretation — variable-per-item heights are out of scope for v0.6.0).
Rendering is wrapped in a ``begin_container`` so ``check_containment`` can
validate all children stay inside the bounds.

Color tokens (``rule_subtle``, ``primary``, ``text_secondary``) resolve via
the theme registry when ``theme`` is supplied; otherwise the shared
``engine.components._util._FALLBACK_TOKENS`` table provides sensible
neutral defaults so the component renders cleanly without a theme.

Layout (per item, top to bottom):

- Number + caption row: 0.28" tall, two side-by-side textboxes
  (number in left 0.7", caption in the remainder)
- 0.04" gap
- Title row: 0.45" tall, bold
- 0.08" gap
- Body row: remaining item height (wrapped, auto-fit + ellipsis)
- Optional rule (0.008" tall rectangle) inside a 0.04" rule row gap
  between this item and the next

Circular-import avoidance: imports from ``..shapes`` are deferred to
function scope (mirrors ``timeline.py`` / ``cards.py`` pattern).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional

from ._util import resolve_component_color as _resolve_color


# --------------------------------------------------------------------------
# Data model
# --------------------------------------------------------------------------


@dataclass
class NumberedItem:
    """A single entry in a numbered list.

    Fields are caller-formatted — the component does not pad/align numbers
    nor compose separators between number and caption. Pass exactly the
    string you want rendered.

    Attributes:
        number: Short number label, e.g. ``"01"``, ``"02"``, ... Rendered
            in the left-most position of the number row.
        caption: Descriptor after the number, e.g. ``"／ 主要ターゲット"``.
            Rendered immediately to the right of ``number`` on the same row.
        title: Bold summary line.
        body: Wrapped explanation paragraph. Empty string → body textbox is
            skipped (cleaner output than a zero-height empty frame).
    """

    number: str
    caption: str
    title: str
    body: str


# Theme-token resolution is delegated to the shared
# ``engine.components._util.resolve_component_color`` helper (imported above
# as ``_resolve_color``) to keep every block component's fallback palette
# and precedence rules identical (Task D, v0.6.0 consolidation).


# --------------------------------------------------------------------------
# Per-item layout constants (inches)
# --------------------------------------------------------------------------

_NUMBER_ROW_H = 0.28       # height of the number+caption row
_NUMBER_COL_W = 0.70       # width of the number textbox (caption fills remainder)
_GAP_NUMBER_TITLE = 0.04   # gap between number row and title row
_TITLE_ROW_H = 0.45        # height of the bold title row
_GAP_TITLE_BODY = 0.08     # gap between title row and body
_RULE_ROW_H = 0.04         # vertical allocation for the rule between items
_RULE_THICKNESS = 0.008    # actual thickness of the rule rectangle


# --------------------------------------------------------------------------
# Public API
# --------------------------------------------------------------------------


def add_numbered_list(
    slide,
    items: List[NumberedItem],
    *,
    left: float,
    top: float,
    width: float,
    height: float,
    rule_between: bool = True,
    rule_color: str = "rule_subtle",
    number_color: str = "text_secondary",
    caption_color: str = "text_secondary",
    title_color: str = "primary",
    body_color: str = "text_secondary",
    number_size_pt: float = 10,
    caption_size_pt: float = 10,
    title_size_pt: float = 16,
    body_size_pt: float = 10,
    theme: Optional[str] = None,
) -> dict:
    """Render ``items`` as a vertical numbered list within the declared bounds.

    Each item is laid out as number+caption → title → body, optionally
    followed by a thin rule. Item heights divide ``height`` equally (minus
    the combined rule-row allocation when ``rule_between=True``). The last
    item never has a trailing rule.

    Args:
        slide: python-pptx slide to draw on.
        items: Non-empty list of ``NumberedItem`` entries.
        left, top, width, height: Outer bounding box in inches. All child
            shapes are guaranteed to lie within these bounds (validated
            via ``check_containment``).
        rule_between: When ``True`` (default), draw a thin horizontal rule
            between consecutive items. Last item never gets a rule.
        rule_color, number_color, caption_color, title_color, body_color:
            Theme tokens or raw 6-hex strings. See ``_resolve_color``.
        number_size_pt, caption_size_pt, title_size_pt, body_size_pt:
            Font sizes in points. Body/title use auto-fit ellipsis if the
            content is too long for the allocated space.
        theme: Optional theme name for token resolution. ``None`` → use
            ``engine.components._util._FALLBACK_TOKENS`` defaults.

    Returns:
        A dict with::

            {
                "items": [
                    {
                        "bounds": {"left", "top", "width", "height"},
                        "number_shape": shape,
                        "caption_shape": shape,
                        "title_shape": shape,
                        "body_shape": shape | None,   # None when body empty
                        "rule_shape": shape | None,   # None for last item
                                                       # or when rule_between=False
                    },
                    ...
                ],
                "consumed_height": height,
                "consumed_width": width,
            }

    Notes:
        - Body textbox is omitted entirely for empty ``item.body`` (not
          rendered as a zero-height placeholder).
        - Long body text is handled by ``add_auto_fit_textbox``'s ellipsis
          shrinking; no error is raised when content exceeds the allotted
          space.
    """
    # Deferred import: avoids circular dependency between
    # ``engine.components`` and ``engine.shapes`` (shapes.py registers
    # shapes with the container stack defined here).
    from ..shapes import _add_shape, add_auto_fit_textbox
    from .container import begin_container

    n = len(items)
    if n == 0:
        return {"items": [], "consumed_height": 0.0, "consumed_width": width}

    # Resolve all colors once up-front (cheaper + easier to reason about).
    rule_hex = _resolve_color(rule_color, theme)
    number_hex = _resolve_color(number_color, theme)
    caption_hex = _resolve_color(caption_color, theme)
    title_hex = _resolve_color(title_color, theme)
    body_hex = _resolve_color(body_color, theme)

    # Allocate per-item height evenly. Rule rows steal (n-1) * _RULE_ROW_H
    # from the total height so per-item content rectangles remain equal.
    draw_rules = bool(rule_between) and n >= 2
    rule_allocation = (n - 1) * _RULE_ROW_H if draw_rules else 0.0
    item_height = (height - rule_allocation) / n

    results: list[dict] = []

    with begin_container(
        slide,
        name="numbered_list",
        left=left,
        top=top,
        width=width,
        height=height,
    ):
        cursor_top = top
        for idx, item in enumerate(items):
            item_top = cursor_top
            item_bounds = {
                "left": left,
                "top": item_top,
                "width": width,
                "height": item_height,
            }

            # --- Number row (left 0.7") -----------------------------------
            number_shape, _ = add_auto_fit_textbox(
                slide,
                item.number,
                left,
                item_top,
                _NUMBER_COL_W,
                _NUMBER_ROW_H,
                font_size_pt=number_size_pt,
                min_size_pt=max(1, int(number_size_pt) - 2),
                color_hex=number_hex,
                align="left",
                wrap=False,
                truncate_with_ellipsis=True,
            )

            # --- Caption row (right of number) ----------------------------
            caption_left = left + _NUMBER_COL_W
            caption_width = max(0.1, width - _NUMBER_COL_W)
            caption_shape, _ = add_auto_fit_textbox(
                slide,
                item.caption,
                caption_left,
                item_top,
                caption_width,
                _NUMBER_ROW_H,
                font_size_pt=caption_size_pt,
                min_size_pt=max(1, int(caption_size_pt) - 2),
                color_hex=caption_hex,
                align="left",
                wrap=False,
                truncate_with_ellipsis=True,
            )

            # --- Title row ------------------------------------------------
            title_top = item_top + _NUMBER_ROW_H + _GAP_NUMBER_TITLE
            title_shape, _ = add_auto_fit_textbox(
                slide,
                item.title,
                left,
                title_top,
                width,
                _TITLE_ROW_H,
                font_size_pt=title_size_pt,
                min_size_pt=max(1, int(title_size_pt) - 4),
                color_hex=title_hex,
                bold=True,
                align="left",
                wrap=False,
                truncate_with_ellipsis=True,
            )

            # --- Body row -------------------------------------------------
            body_top = title_top + _TITLE_ROW_H + _GAP_TITLE_BODY
            # Remaining vertical space inside this item's allocated rect.
            body_height = max(
                0.0,
                (item_top + item_height) - body_top,
            )
            body_shape = None
            if item.body and body_height > 0:
                body_shape, _ = add_auto_fit_textbox(
                    slide,
                    item.body,
                    left,
                    body_top,
                    width,
                    body_height,
                    font_size_pt=body_size_pt,
                    min_size_pt=max(1, int(body_size_pt) - 2),
                    color_hex=body_hex,
                    align="left",
                    wrap=True,
                    truncate_with_ellipsis=True,
                )

            # --- Optional rule between this item and the next -------------
            rule_shape = None
            is_last = idx == n - 1
            if draw_rules and not is_last:
                # Place the rule centered within the rule-row gap.
                rule_row_top = item_top + item_height
                rule_top_pos = rule_row_top + (_RULE_ROW_H - _RULE_THICKNESS) / 2
                rule_idx = _add_shape(
                    slide,
                    "rectangle",
                    left,
                    rule_top_pos,
                    width,
                    _RULE_THICKNESS,
                    fill_color=rule_hex,
                    no_line=True,
                )
                rule_shape = slide.shapes[rule_idx]

            results.append(
                {
                    "bounds": item_bounds,
                    "number_shape": number_shape,
                    "caption_shape": caption_shape,
                    "title_shape": title_shape,
                    "body_shape": body_shape,
                    "rule_shape": rule_shape,
                }
            )

            # Advance cursor past this item (+ rule row if applicable).
            cursor_top = item_top + item_height
            if draw_rules and not is_last:
                cursor_top += _RULE_ROW_H

    return {
        "items": results,
        "consumed_height": height,
        "consumed_width": width,
    }
