"""
Theme definitions for presentation styling.

Provides a Theme dataclass, built-in themes (``mckinsey``, ``deloitte``,
``neutral``, ``ir``), a theme registry for lookup by name, and color
utilities (tint/shade).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass
class Theme:
    """Presentation theme with colors, fonts, sizes, and layout defaults."""

    name: str = "custom"
    colors: Dict[str, str] = field(default_factory=dict)
    fonts: Dict[str, str] = field(default_factory=dict)
    sizes: Dict[str, float] = field(default_factory=dict)
    slide: Dict[str, float] = field(default_factory=dict)
    margins: Dict[str, float] = field(default_factory=dict)
    layout: Dict[str, float] = field(default_factory=dict)
    table: Dict[str, object] = field(default_factory=dict)
    chart_colors: List[str] = field(default_factory=list)
    connector: Dict[str, object] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Theme registry
# ---------------------------------------------------------------------------

_THEME_REGISTRY: Dict[str, Theme] = {}


def register_theme(theme: Theme) -> None:
    """Register a theme in the global registry."""
    _THEME_REGISTRY[theme.name] = theme


def get_theme(name: str) -> Optional[Theme]:
    """Look up a theme by name.  Returns None if not found."""
    return _THEME_REGISTRY.get(name)


def list_themes() -> List[str]:
    """Return names of all registered themes."""
    return list(_THEME_REGISTRY.keys())


# ---------------------------------------------------------------------------
# Color utilities
# ---------------------------------------------------------------------------

def resolve_color(theme: Theme, token_or_hex: str) -> str:
    """Look up a color name from theme.colors or return the hex directly.

    Supports "series_N" tokens that resolve from theme.chart_colors.

    Examples:
        resolve_color(theme, "primary")   -> "#051C2C"
        resolve_color(theme, "#FF0000")   -> "#FF0000"
        resolve_color(theme, "series_0")  -> first chart color
    """
    if theme and token_or_hex in theme.colors:
        return theme.colors[token_or_hex]
    if theme and token_or_hex.startswith("series_") and theme.chart_colors:
        try:
            idx = int(token_or_hex.split("_", 1)[1])
            return theme.chart_colors[idx % len(theme.chart_colors)]
        except (ValueError, IndexError):
            pass
    return token_or_hex


def tint_color(hex_color: str, factor: float) -> str:
    """Generate a tint (lighter) of a hex color.

    factor: 0.0 = original, 1.0 = white.
    Formula per channel: channel + (255 - channel) * factor
    Matches PowerPoint's native tint algorithm.
    """
    hex_color = hex_color.lstrip("#")
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    r = min(255, int(r + (255 - r) * factor))
    g = min(255, int(g + (255 - g) * factor))
    b = min(255, int(b + (255 - b) * factor))
    return f"#{r:02X}{g:02X}{b:02X}"


def shade_color(hex_color: str, factor: float) -> str:
    """Generate a shade (darker) of a hex color.

    factor: 0.0 = original, 1.0 = black.
    Formula per channel: channel * (1 - factor)
    Matches PowerPoint's native shade algorithm.
    """
    hex_color = hex_color.lstrip("#")
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    r = max(0, int(r * (1 - factor)))
    g = max(0, int(g * (1 - factor)))
    b = max(0, int(b * (1 - factor)))
    return f"#{r:02X}{g:02X}{b:02X}"


def get_chart_color(theme: Theme, series_index: int) -> str:
    """Get the color for a chart series by index, cycling through chart_colors."""
    if theme and theme.chart_colors:
        return theme.chart_colors[series_index % len(theme.chart_colors)]
    return "#2251FF"


def resolve_theme_color(
    token_or_hex: str,
    theme_name: Optional[str] = None,
) -> str:
    """theme トークン / hex を primitive が期待する 6-hex (no ``#`` prefix) に解決する.

    Primitive 向けの thin wrapper。``theme_name`` が指定されていればテーマの
    ``colors`` を先に参照し、一致しなければ入力を hex として扱う。戻り値は
    `_parse_color` の期待に合わせて先頭の ``#`` を必ず除く。

    - ``resolve_theme_color("primary", "mckinsey")`` → ``"051C2C"``
    - ``resolve_theme_color("#051C2C")`` → ``"051C2C"``
    - ``resolve_theme_color("051C2C")`` → ``"051C2C"`` (passthrough)
    - ``resolve_theme_color("", "ir")`` → ``""`` (empty for "disable")

    Args:
        token_or_hex: theme token (e.g. ``"primary"``) or raw hex. Empty
            string is returned as-is (callers use ``""`` to mean "disable").
        theme_name: theme registry key (e.g. ``"mckinsey"``); if ``None`` or
            unregistered, only hex passthrough happens.

    Returns:
        6-hex string without ``#`` prefix, or the empty string if input was
        empty.
    """
    if not token_or_hex:
        return ""
    if theme_name:
        theme = get_theme(theme_name)
        if theme is not None:
            resolved = resolve_color(theme, token_or_hex)
            return resolved.lstrip("#")
    return token_or_hex.lstrip("#")


# ---------------------------------------------------------------------------
# Built-in themes
# ---------------------------------------------------------------------------

MCKINSEY = Theme(
    name="mckinsey",
    colors={
        "primary": "#051C2C",
        "accent": "#2251FF",
        "white": "#FFFFFF",
        "text_secondary": "#666666",
        "footnote": "#A2AAAD",
        "positive": "#2E7D32",
        "negative": "#C62828",
        "bg_alt": "#F5F5F5",
        "border": "#D0D0D0",
    },
    fonts={
        "title": "Arial",
        "body": "Arial",
        # 明示的に east-asian typeface を指定し、非 JP Windows での中国語
        # フォントへの自動 fallback を防ぐ (issue #40)。
        "east_asian": "Yu Gothic",
    },
    sizes={
        "title": 22,
        "subtitle": 16,
        "body": 12,
        "table": 10,
        "caption": 9,
        "footnote": 8,
    },
    slide={
        "width": 13.333,
        "height": 7.5,
    },
    margins={
        "left": 0.9,
        "right": 0.9,
        "top": 0.5,
    },
    layout={
        "title_top": 0.45,
        "title_height": 0.5,
        "divider_top": 0.95,
        "body_top": 1.15,
        "footer_top": 6.65,
    },
    table={
        "header_bg": "primary",
        "header_fg": "white",
        "alt_row_bg": "bg_alt",
        "border_color": "border",
        "no_vertical_borders": True,
        "row_height": 0.30,
        "border_width": 0.5,
    },
    chart_colors=[
        "#2251FF",  # accent blue
        "#051C2C",  # primary navy
        "#A2AAAD",  # gray
        "#2E7D32",  # positive green
        "#C62828",  # negative red
        "#666666",  # secondary text
    ],
    connector={
        "color": "accent",
        "width": 1.5,
        "arrow_end": "triangle",
    },
)
register_theme(MCKINSEY)


DELOITTE = Theme(
    name="deloitte",
    colors={
        "primary": "#002776",
        "accent": "#81BC00",
        "white": "#FFFFFF",
        "text_secondary": "#575757",
        "footnote": "#8C8C8C",
        "positive": "#3C8A2E",
        "negative": "#C62828",
        "bg_alt": "#F5F5F5",
        "border": "#DCDCDC",
        "mid_blue": "#00A1DE",
        "light_blue": "#72C7E7",
        "light_green": "#BDD203",
        "dark_green": "#3C8A2E",
        "gray_1": "#DCDCDC",
        "gray_2": "#B4B4B4",
        "gray_3": "#8C8C8C",
        "gray_4": "#575757",
        "gray_5": "#313131",
    },
    fonts={
        "title": "Arial",
        "body": "Arial",
        # 明示的に east-asian typeface を指定し、非 JP Windows での中国語
        # フォントへの自動 fallback を防ぐ (issue #40)。
        "east_asian": "Yu Gothic",
    },
    sizes={
        "title": 22,
        "subtitle": 16,
        "body": 12,
        "table": 10,
        "caption": 9,
        "footnote": 8,
    },
    slide={
        "width": 13.333,
        "height": 7.5,
    },
    margins={
        "left": 0.9,
        "right": 0.9,
        "top": 0.5,
    },
    layout={
        "title_top": 0.45,
        "title_height": 0.5,
        "divider_top": 0.95,
        "body_top": 1.15,
        "footer_top": 6.65,
    },
    table={
        "header_bg": "primary",
        "header_fg": "white",
        "alt_row_bg": "bg_alt",
        "border_color": "border",
        "no_vertical_borders": True,
        "row_height": 0.30,
        "border_width": 0.5,
    },
    chart_colors=[
        "#002776",  # Navy
        "#81BC00",  # Green
        "#00A1DE",  # Mid Blue
        "#3C8A2E",  # Dark Green
        "#72C7E7",  # Light Blue
        "#BDD203",  # Light Green
        "#575757",  # Dark Gray
        "#B4B4B4",  # Mid Gray
    ],
    connector={
        "color": "primary",
        "width": 1.5,
        "arrow_end": "triangle",
    },
)
register_theme(DELOITTE)


NEUTRAL = Theme(
    name="neutral",
    colors={
        "primary": "#333333",
        "accent": "#4A90D9",
        "white": "#FFFFFF",
        "text_secondary": "#666666",
        "footnote": "#999999",
        "positive": "#27AE60",
        "negative": "#E74C3C",
        "bg_alt": "#F8F8F8",
        "border": "#E0E0E0",
    },
    fonts={
        "title": "Arial",
        "body": "Arial",
        # 明示的に east-asian typeface を指定し、非 JP Windows での中国語
        # フォントへの自動 fallback を防ぐ (issue #40)。
        "east_asian": "Yu Gothic",
    },
    sizes={
        "title": 22,
        "subtitle": 16,
        "body": 12,
        "table": 10,
        "caption": 9,
        "footnote": 8,
    },
    slide={
        "width": 13.333,
        "height": 7.5,
    },
    margins={
        "left": 0.9,
        "right": 0.9,
        "top": 0.5,
    },
    layout={
        "title_top": 0.45,
        "title_height": 0.5,
        "divider_top": 0.95,
        "body_top": 1.15,
        "footer_top": 6.65,
    },
    table={
        "header_bg": "primary",
        "header_fg": "white",
        "alt_row_bg": "bg_alt",
        "border_color": "border",
        "no_vertical_borders": True,
        "row_height": 0.30,
        "border_width": 0.5,
    },
    chart_colors=[
        "#4A90D9",  # Blue
        "#E67E22",  # Orange
        "#27AE60",  # Green
        "#9B59B6",  # Purple
        "#E74C3C",  # Red
        "#1ABC9C",  # Teal
        "#F39C12",  # Yellow
        "#95A5A6",  # Gray
    ],
    connector={
        "color": "accent",
        "width": 1.0,
        "arrow_end": "triangle",
    },
)
register_theme(NEUTRAL)


# 日本の上場企業の IR (四半期決算) 資料向けプリセット。
# クリーム色の背景 (#F8F9F5) とディープネイビー (#0A2540) を基調とし、
# HD ワイド (20.0 x 11.25 インチ) で Yu Gothic を東アジア字形に固定する。
IR = Theme(
    name="ir",
    colors={
        "primary": "#0A2540",          # deep navy
        "accent": "#1E3A8A",           # blue accent
        "background": "#F8F9F5",       # cream
        "text_primary": "#1A1A1A",
        "text_secondary": "#6B7280",
        "rule_strong": "#C0C0C0",
        "rule_subtle": "#E0E0E0",
        "highlight_row": "#F0F0F0",    # for data tables
        "positive": "#059669",         # green for +% figures
        "negative": "#DC2626",
    },
    fonts={
        "title": "Arial",
        "body": "Arial",
        # 明示的に east-asian typeface を指定し、非 JP Windows での中国語
        # フォントへの自動 fallback を防ぐ (issue #40)。
        "east_asian": "Yu Gothic",
    },
    sizes={
        "title": 28,
        "subtitle": 14,
        "body": 11,
        "table": 10,
        "caption": 9,
        "footnote": 8,
    },
    slide={
        "width": 20.0,   # HD widescreen (IR spec)
        "height": 11.25,
    },
    margins={
        "left": 1.0,
        "right": 1.0,
        "top": 0.6,
    },
    layout={
        "title_top": 0.6,
        "title_height": 0.8,
        "divider_top": 1.7,
        "body_top": 1.85,
        "footer_top": 10.3,
    },
    table={
        "header_bg": "background",
        "header_fg": "text_primary",
        "alt_row_bg": "highlight_row",
        "border_color": "rule_subtle",
        "no_vertical_borders": True,
        "row_height": 0.5,
        "border_width": 0.5,
    },
    chart_colors=[
        "#1E3A8A",  # primary navy
        "#3B82F6",  # lighter blue
        "#059669",  # green
        "#DC2626",  # red
        "#6B7280",  # gray
        "#C0C0C0",  # rule gray
    ],
    connector={
        "color": "rule_strong",
        "width": 1.0,
        "arrow_end": "none",
    },
)
register_theme(IR)
