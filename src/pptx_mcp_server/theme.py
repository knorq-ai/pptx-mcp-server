"""
Theme definitions for presentation styling.

Provides a Theme dataclass, built-in themes (McKinsey, Deloitte, Neutral),
a theme registry for lookup by name, and color utilities (tint/shade).
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
