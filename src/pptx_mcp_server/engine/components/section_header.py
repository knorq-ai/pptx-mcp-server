"""SectionHeader block component (Issue #136, v0.6.0).

Renders a vertically-stacked section header composed of:

1. A **title** textbox (single-line, auto-fit by width, bold).
2. An optional **subtitle** textbox (single-line, auto-fit by width).
3. A thin **divider** rectangle spanning the full width.

The component is wrapped in ``begin_container`` so ``check_containment``
can verify that no child shape escapes the declared bounds. Text auto-fit
uses ``wrap=False`` + ``truncate_with_ellipsis=True`` — this mirrors the
v0.3.1 ``_title_bar`` behavior and guarantees long strings never wrap into
a second line.

The function returns a ``dict`` that includes ``consumed_height`` so the
caller can place body content directly below without re-deriving the
vertical math. This is the primary contract with callers.

Design notes:
- Layout constants (``_TITLE_H`` etc.) live at module top for easy tuning
  and so tests can pin the exact bounds math.
- Colors are passed through as tokens (``"primary"``, ``"text_secondary"``,
  ``"rule_subtle"``); ``add_auto_fit_textbox`` and ``_add_shape`` resolve
  them against ``theme`` internally — we don't resolve here to avoid
  double-resolving.
- Divider is a filled rectangle (``no_line=True``) rather than a line
  connector, matching ``tables_grid`` hairline rule style.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, Optional

from ...theme import resolve_theme_color
from .container import begin_container

# NOTE: ``..shapes`` (engine.shapes) imports ``begin_container`` from this
# components sub-package at module load, so importing ``_add_shape`` /
# ``add_auto_fit_textbox`` at the top of this file would form a circular
# import (engine.shapes ↔ engine.components.__init__). We defer those
# imports to the body of :func:`add_section_header` instead.

# Fallback hex values used when ``theme`` is ``None`` — matches the
# MCKINSEY palette which is the historical default across the engine
# (cards.py, timeline.py follow the same convention, #125).
_TOKEN_FALLBACK_HEX: dict[str, str] = {
    "primary": "051C2C",
    "text_secondary": "666666",
    "rule_subtle": "E0E0E0",
}


def _resolve_color(token_or_hex: str, theme_name: Optional[str]) -> str:
    """Resolve a theme token / raw hex to a 6-hex value primitives accept.

    When ``theme_name`` is provided, delegates to the central resolver
    (``resolve_theme_color``). When no theme is in play and the input is
    not recognizable as hex, falls back to the MCKINSEY-equivalent default
    hex value so atomic primitives never receive an unresolved token like
    ``"primary"`` (which would raise ``Invalid color`` in ``_parse_color``).
    """
    if not token_or_hex:
        return ""
    if theme_name:
        return resolve_theme_color(token_or_hex, theme_name)
    stripped = token_or_hex.lstrip("#")
    # Heuristic: treat a 6-char all-hex string as a raw color.
    if len(stripped) == 6 and all(
        c in "0123456789abcdefABCDEF" for c in stripped
    ):
        return stripped
    return _TOKEN_FALLBACK_HEX.get(token_or_hex, stripped)

# ---------------------------------------------------------------------------
# Layout constants (inches)
# ---------------------------------------------------------------------------
# Title strip height — accommodates a 32pt bold line with vertical padding
# on both sides so auto-fit's single-line render does not clip.
_TITLE_H = 0.55
# Gap between title and subtitle (or title and divider if no subtitle).
_INTRA_GAP_TITLE = 0.08
# Subtitle strip height — accommodates a 14pt line with padding.
_SUBTITLE_H = 0.30
# Gap between subtitle (or title if no subtitle) and the divider rule.
_INTRA_GAP_SUBTITLE = 0.12


@dataclass
class SectionHeaderSpec:
    """Declarative spec for :func:`add_section_header`.

    Attributes:
        title: Section title text (required, non-empty in practice).
        subtitle: Optional subtitle text; empty string omits the subtitle
            strip and reduces ``consumed_height``.
        title_color: Theme token or 6-hex for the title text color.
        subtitle_color: Theme token or 6-hex for the subtitle text color.
        title_size_pt: Starting title font size (pt); auto-fit may shrink
            down to ``title_min_size_pt``.
        title_min_size_pt: Minimum title font size (pt) before truncation.
        subtitle_size_pt: Starting subtitle font size (pt).
        subtitle_min_size_pt: Minimum subtitle font size (pt).
        divider_color: Theme token or 6-hex for the divider fill.
        divider_thickness: Divider rectangle height (inches). Defaults to
            ~0.008" ≈ 0.75pt hairline.
    """

    title: str
    subtitle: str = ""
    title_color: str = "primary"
    subtitle_color: str = "text_secondary"
    title_size_pt: float = 32
    title_min_size_pt: float = 22
    subtitle_size_pt: float = 14
    subtitle_min_size_pt: float = 10
    divider_color: str = "rule_subtle"
    divider_thickness: float = 0.008


def _compute_consumed_height(spec: SectionHeaderSpec) -> float:
    """Return total vertical footprint of the header in inches.

    Broken out so tests and callers can predict layout without re-running
    the full render.
    """
    h = _TITLE_H + _INTRA_GAP_TITLE
    if spec.subtitle:
        h += _SUBTITLE_H
    h += _INTRA_GAP_SUBTITLE + float(spec.divider_thickness)
    return h


def add_section_header(
    slide,
    spec: SectionHeaderSpec,
    *,
    left: float,
    top: float,
    width: float,
    theme: Optional[str] = None,
) -> Dict[str, Any]:
    """Render a section header (title + optional subtitle + divider rule).

    The full stack is wrapped in a ``begin_container`` with name
    ``"section_header"`` so ``check_containment`` verifies children stay
    inside the declared bounds.

    Args:
        slide: Target python-pptx slide object.
        spec: Declarative :class:`SectionHeaderSpec`.
        left, top, width: Outer geometry in inches. The component consumes
            height dynamically based on whether a subtitle is present; see
            the returned ``consumed_height`` field.
        theme: Optional theme name (e.g. ``"ir"``, ``"mckinsey"``).
            Propagated to atomic primitives so color tokens resolve
            against the registered theme palette.

    Returns:
        A dict with:

        - ``title_bounds``: ``{left, top, width, height}`` for the title
          textbox.
        - ``subtitle_bounds``: Same shape if a subtitle was rendered,
          else ``None``.
        - ``divider_bounds``: ``{left, top, width, height}`` for the
          divider rectangle.
        - ``consumed_height`` (float): Total vertical footprint — the
          caller contract for placing body content.
        - ``title_actual_font_size`` (float): The post-auto-fit title
          font size; useful for tests.
        - ``subtitle_actual_font_size`` (float | None): Post-auto-fit
          subtitle size, or ``None`` if no subtitle.
    """
    # Deferred imports to avoid a circular dependency: ``engine.shapes``
    # imports ``begin_container`` from this package at load time.
    from ..shapes import _add_shape, add_auto_fit_textbox

    # Resolve theme tokens upfront so atomic primitives always receive raw
    # hex. This also keeps ``theme=None`` (unit-test) calls working when
    # the spec uses the default tokens ``"primary"`` / ``"text_secondary"``
    # / ``"rule_subtle"``.
    title_color_hex = _resolve_color(spec.title_color, theme)
    subtitle_color_hex = _resolve_color(spec.subtitle_color, theme)
    divider_color_hex = _resolve_color(spec.divider_color, theme)

    consumed_height = _compute_consumed_height(spec)

    title_top = float(top)
    # If no subtitle, the divider sits right after the title + a single gap.
    if spec.subtitle:
        subtitle_top: Optional[float] = title_top + _TITLE_H + _INTRA_GAP_TITLE
        divider_top = (
            subtitle_top + _SUBTITLE_H + _INTRA_GAP_SUBTITLE
        )
    else:
        subtitle_top = None
        divider_top = title_top + _TITLE_H + _INTRA_GAP_SUBTITLE

    with begin_container(
        slide,
        name="section_header",
        left=float(left),
        top=float(top),
        width=float(width),
        height=consumed_height,
    ):
        # 1. Title — bold, single-line, auto-fit by width.
        _, title_actual_font_size = add_auto_fit_textbox(
            slide,
            spec.title,
            left=float(left),
            top=title_top,
            width=float(width),
            height=_TITLE_H,
            font_size_pt=spec.title_size_pt,
            min_size_pt=spec.title_min_size_pt,
            bold=True,
            color_hex=title_color_hex,
            align="left",
            vertical_anchor="top",
            truncate_with_ellipsis=True,
            wrap=False,
            theme=theme,
        )

        # 2. Optional subtitle — regular weight, single-line.
        subtitle_actual_font_size: Optional[float] = None
        subtitle_bounds: Optional[Dict[str, float]] = None
        if spec.subtitle and subtitle_top is not None:
            _, subtitle_actual_font_size = add_auto_fit_textbox(
                slide,
                spec.subtitle,
                left=float(left),
                top=subtitle_top,
                width=float(width),
                height=_SUBTITLE_H,
                font_size_pt=spec.subtitle_size_pt,
                min_size_pt=spec.subtitle_min_size_pt,
                bold=False,
                color_hex=subtitle_color_hex,
                align="left",
                vertical_anchor="top",
                truncate_with_ellipsis=True,
                wrap=False,
                theme=theme,
            )
            subtitle_bounds = {
                "left": float(left),
                "top": float(subtitle_top),
                "width": float(width),
                "height": _SUBTITLE_H,
            }

        # 3. Divider rule — hairline filled rectangle spanning full width.
        _add_shape(
            slide,
            "rectangle",
            float(left),
            float(divider_top),
            float(width),
            float(spec.divider_thickness),
            fill_color=divider_color_hex,
            no_line=True,
            theme=theme,
        )

    return {
        "title_bounds": {
            "left": float(left),
            "top": float(title_top),
            "width": float(width),
            "height": _TITLE_H,
        },
        "subtitle_bounds": subtitle_bounds,
        "divider_bounds": {
            "left": float(left),
            "top": float(divider_top),
            "width": float(width),
            "height": float(spec.divider_thickness),
        },
        "consumed_height": consumed_height,
        "title_actual_font_size": float(title_actual_font_size),
        "subtitle_actual_font_size": (
            float(subtitle_actual_font_size)
            if subtitle_actual_font_size is not None
            else None
        ),
    }
