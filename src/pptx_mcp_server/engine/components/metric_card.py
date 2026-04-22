"""MetricCard block component (Issue #132).

A single metric card renders, top-to-bottom, a small uppercase ``label``,
a ``title``, a chart (image or placeholder rectangle), and an optional row
of ``(metric.label, metric.value)`` cells. Sibling cards of equal height
can be laid out horizontally with :func:`add_metric_card_row`.

Design notes
------------
- Uses :func:`begin_container` to declare a validator-checkable bounding box
  per card. Child shapes auto-register via the thread-local stack maintained
  by the atomic primitives.
- Deferred imports are used for :mod:`..shapes` symbols to avoid circular
  imports: ``engine.shapes`` already imports ``components.container``, and
  ``components/__init__`` re-exports this module. Mirrors the pattern used
  by sibling components landing on the same foundation branch.
- Color fields on :class:`MetricCardSpec` default to theme tokens
  (``"primary"``, ``"text_secondary"``, ``"rule_subtle"``). When the caller
  passes ``theme=None``, the module-local :func:`_resolve_color` falls back
  to MCKINSEY-equivalent hexes so token defaults never reach the paint
  layer unresolved (mirror of ``engine/timeline.py:_style_color`` /
  ``engine/components/section_header.py:_resolve_color``).
- Metric value textboxes use ``wrap=False`` so large numbers scale by width
  rather than wrapping mid-figure.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from ..pptx_io import EngineError, ErrorCode
from .container import begin_container

# Vertical allocation inside the inner (padded) box.
_LABEL_H: float = 0.25
_LABEL_TITLE_GAP: float = 0.05
_TITLE_H: float = 0.40
_TITLE_CHART_GAP: float = 0.20
_CHART_METRICS_GAP: float = 0.20
_METRICS_ROW_H: float = 0.80
_METRIC_LABEL_CELL_H: float = 0.30
_METRIC_VALUE_CELL_H: float = 0.50
_MIN_CHART_H: float = 0.5


# Hardcoded fallbacks for the default theme tokens used by this component.
# Covers every token appearing in MetricCardSpec defaults + the background
# hex used as chart-placeholder fill.
_FALLBACK_COLORS = {
    "primary": "051C2C",
    "text_secondary": "666666",
    "rule_subtle": "E0E0E0",
}


@dataclass
class MetricEntry:
    """A single (label, value) cell in the metrics row."""

    label: str
    value: str


@dataclass
class MetricCardSpec:
    """Content + style definition for one metric card.

    Attributes:
        label: Small uppercase caption above the title (e.g. "KPI").
        title: Main heading text for the card.
        chart_image_path: Optional path to a chart PNG; when ``None``, a
            blank placeholder rectangle fills the chart area.
        metrics: Optional list of :class:`MetricEntry` for the bottom row.
            Empty list → chart area expands to fill the remaining height.
        fill_color: Card background fill (hex or theme token).
        border_color: Card stroke color (hex or theme token).
        border_width: Card stroke width in points.
        label_color: Color for the ``label`` text.
        title_color: Color for the ``title`` text.
        metric_label_color: Color for per-metric labels.
        metric_value_color: Color for per-metric values.
        padding: Inner inset (inches) applied equally to all 4 sides.
        title_size_pt: Font size for the title.
        label_size_pt: Font size for the label.
        metric_label_size_pt: Font size for per-metric labels.
        metric_value_size_pt: Font size for per-metric values.
    """

    label: str = ""
    title: str = ""
    chart_image_path: Optional[str] = None
    metrics: list[MetricEntry] = field(default_factory=list)
    fill_color: str = "F8F9F5"
    border_color: str = "rule_subtle"
    border_width: float = 0.012
    label_color: str = "text_secondary"
    title_color: str = "primary"
    metric_label_color: str = "text_secondary"
    metric_value_color: str = "primary"
    padding: float = 0.3
    title_size_pt: float = 14
    label_size_pt: float = 9
    metric_label_size_pt: float = 10
    metric_value_size_pt: float = 22


def _resolve_color(token_or_hex: str, theme: str | None) -> str:
    """Resolve a token/hex to 6-hex (no ``#``) even when ``theme`` is None.

    When a theme is provided, delegates to :func:`resolve_theme_color`.
    Otherwise maps the well-known default tokens to hardcoded MCKINSEY hex
    equivalents; raw hex passes through (with an optional leading ``#``
    stripped).
    """
    if not token_or_hex:
        return ""
    if theme:
        from ...theme import resolve_theme_color

        return resolve_theme_color(token_or_hex, theme)
    if token_or_hex in _FALLBACK_COLORS:
        return _FALLBACK_COLORS[token_or_hex]
    # Raw hex passthrough (strip optional ``#``).
    return token_or_hex.lstrip("#")


def add_metric_card(
    slide,
    spec: MetricCardSpec,
    *,
    left: float,
    top: float,
    width: float,
    height: float,
    theme: str | None = None,
    container_name: str = "metric_card",
) -> dict:
    """Render a single metric card inside ``(left, top, width, height)``.

    Lays out (top-to-bottom) label → title → chart → metrics row inside an
    inner box shrunk by :attr:`MetricCardSpec.padding` on all sides.

    Returns:
        Dict with the rendered bounds and each child shape:
        ``{"bounds": {...}, "frame_shape", "label_shape", "title_shape",
        "chart_shape", "metric_shapes": [(label_shape, value_shape), ...]}``.

    Raises:
        EngineError: ``INVALID_PARAMETER`` when the inner height is too
            small to satisfy the label/title/chart/metrics allocation.
    """
    # Deferred imports to avoid circular import with ``engine.shapes``
    # (which imports ``components.container``).
    from ..shapes import _add_image, _add_shape, add_auto_fit_textbox

    # Validate optional chart image path up front so a missing file fails
    # before any shapes are appended (avoids partial-card leaks on error).
    if spec.chart_image_path is not None:
        import os

        if not os.path.isfile(spec.chart_image_path):
            raise EngineError(
                ErrorCode.INVALID_PARAMETER,
                f"MetricCard chart_image_path does not exist: {spec.chart_image_path!r}",
            )

    # Resolve colors once at entry. Raw hex / theme tokens both normalize
    # to 6-hex (no ``#``) so downstream primitives receive uniform input.
    fill_color = _resolve_color(spec.fill_color, theme) or "F8F9F5"
    border_color = _resolve_color(spec.border_color, theme) or "E0E0E0"
    label_color = _resolve_color(spec.label_color, theme) or "666666"
    title_color = _resolve_color(spec.title_color, theme) or "051C2C"
    metric_label_color = _resolve_color(spec.metric_label_color, theme) or "666666"
    metric_value_color = _resolve_color(spec.metric_value_color, theme) or "051C2C"

    pad = float(spec.padding)
    inner_w = width - 2 * pad
    inner_h = height - 2 * pad

    has_metrics = bool(spec.metrics)
    metrics_row_h = _METRICS_ROW_H if has_metrics else 0.0
    metrics_gap = _CHART_METRICS_GAP if has_metrics else 0.0

    required = (
        _LABEL_H
        + _LABEL_TITLE_GAP
        + _TITLE_H
        + _TITLE_CHART_GAP
        + _MIN_CHART_H
        + metrics_gap
        + metrics_row_h
    )
    if inner_h < required:
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            (
                "MetricCard height too small: need >= "
                f"{required + 2 * pad:.2f}\" "
                f"for padding/label/title/chart/metrics; got {height:.2f}\"."
            ),
        )

    with begin_container(
        slide,
        name=container_name,
        left=float(left),
        top=float(top),
        width=float(width),
        height=float(height),
    ):
        # 1) Card background + border frame.
        bg_idx = _add_shape(
            slide,
            "rectangle",
            float(left),
            float(top),
            float(width),
            float(height),
            fill_color=fill_color,
            line_color=border_color,
            line_width=float(spec.border_width),
        )

        # 2) Inner layout cursor.
        inner_left = float(left) + pad
        cursor_y = float(top) + pad

        # 3) Label row.
        label_shape, _ = add_auto_fit_textbox(
            slide,
            spec.label,
            left=inner_left,
            top=cursor_y,
            width=inner_w,
            height=_LABEL_H,
            font_size_pt=float(spec.label_size_pt),
            min_size_pt=max(float(spec.label_size_pt) - 2, 6),
            color_hex=label_color,
            align="left",
            vertical_anchor="top",
        )
        cursor_y += _LABEL_H + _LABEL_TITLE_GAP

        # 4) Title row.
        title_shape, _ = add_auto_fit_textbox(
            slide,
            spec.title,
            left=inner_left,
            top=cursor_y,
            width=inner_w,
            height=_TITLE_H,
            font_size_pt=float(spec.title_size_pt),
            min_size_pt=max(float(spec.title_size_pt) - 4, 8),
            bold=True,
            color_hex=title_color,
            align="left",
            vertical_anchor="top",
        )
        cursor_y += _TITLE_H + _TITLE_CHART_GAP

        # 5) Chart area: fills remaining space above the metrics row.
        chart_bottom = float(top) + pad + inner_h - metrics_row_h - metrics_gap
        chart_h = chart_bottom - cursor_y

        if spec.chart_image_path:
            chart_idx = _add_image(
                slide,
                spec.chart_image_path,
                left=inner_left,
                top=cursor_y,
                width=inner_w,
                height=chart_h,
            )
        else:
            chart_idx = _add_shape(
                slide,
                "rectangle",
                inner_left,
                cursor_y,
                inner_w,
                chart_h,
                fill_color=fill_color,
                line_color=None,
                no_line=True,
            )
        chart_shape = slide.shapes[chart_idx]

        # 6) Metrics row: N cells split evenly across inner width.
        metric_shapes: list[tuple[object, object]] = []
        if has_metrics:
            metrics_top = cursor_y + chart_h + metrics_gap
            n = len(spec.metrics)
            cell_w = inner_w / n
            for i, metric in enumerate(spec.metrics):
                cell_x = inner_left + i * cell_w
                m_label_shape, _ = add_auto_fit_textbox(
                    slide,
                    metric.label,
                    left=cell_x,
                    top=metrics_top,
                    width=cell_w,
                    height=_METRIC_LABEL_CELL_H,
                    font_size_pt=float(spec.metric_label_size_pt),
                    min_size_pt=max(float(spec.metric_label_size_pt) - 2, 6),
                    color_hex=metric_label_color,
                    align="left",
                    vertical_anchor="top",
                )
                m_value_shape, _ = add_auto_fit_textbox(
                    slide,
                    metric.value,
                    left=cell_x,
                    top=metrics_top + _METRIC_LABEL_CELL_H,
                    width=cell_w,
                    height=_METRIC_VALUE_CELL_H,
                    font_size_pt=float(spec.metric_value_size_pt),
                    min_size_pt=max(float(spec.metric_value_size_pt) - 12, 10),
                    bold=True,
                    color_hex=metric_value_color,
                    align="left",
                    vertical_anchor="top",
                    wrap=False,
                )
                metric_shapes.append((m_label_shape, m_value_shape))

    # The background rectangle index captured pre-children; re-resolve to
    # a live shape reference for the return payload. (Additional shapes
    # were added after it, but `bg_idx` remains valid because indexes are
    # append-only for in-memory slides.)
    bg_shape = slide.shapes[bg_idx]

    return {
        "bounds": {
            "left": float(left),
            "top": float(top),
            "width": float(width),
            "height": float(height),
        },
        "frame_shape": bg_shape,
        "label_shape": label_shape,
        "title_shape": title_shape,
        "chart_shape": chart_shape,
        "metric_shapes": metric_shapes,
    }


def add_metric_card_row(
    slide,
    specs: list[MetricCardSpec],
    *,
    left: float,
    top: float,
    width: float,
    height: float,
    gap: float = 0.3,
    theme: str | None = None,
) -> dict:
    """Render N metric cards side-by-side with equal widths.

    Each card occupies ``(width - (n - 1) * gap) / n`` inches of horizontal
    space and the full ``height``. Single-card rows ignore ``gap``.

    Returns:
        ``{"cards": [bounds_dict, ...], "consumed_height": height,
        "consumed_width": width}``.
    """
    n = len(specs)
    if n == 0:
        return {"cards": [], "consumed_height": 0.0, "consumed_width": 0.0}

    effective_gap = gap if n > 1 else 0.0
    card_w = (width - effective_gap * (n - 1)) / n

    placements: list[dict] = []
    for i, spec in enumerate(specs):
        x = left + i * (card_w + effective_gap)
        result = add_metric_card(
            slide,
            spec,
            left=x,
            top=top,
            width=card_w,
            height=height,
            theme=theme,
            container_name=f"metric_card_{i}",
        )
        placements.append(result["bounds"])

    return {
        "cards": placements,
        "consumed_height": float(height),
        "consumed_width": float(width),
    }
