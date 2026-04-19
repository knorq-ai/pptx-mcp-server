"""
Unit tests for theme definitions and resolve_color.
"""

from __future__ import annotations

import pytest

from pptx_mcp_server.theme import MCKINSEY, Theme, resolve_color


class TestMcKinseyThemeStructure:
    """MCKINSEY theme must expose all required configuration sections."""

    def test_has_colors(self):
        assert isinstance(MCKINSEY.colors, dict)
        assert len(MCKINSEY.colors) > 0

    def test_has_fonts(self):
        assert isinstance(MCKINSEY.fonts, dict)
        assert "title" in MCKINSEY.fonts
        assert "body" in MCKINSEY.fonts

    def test_has_sizes(self):
        assert isinstance(MCKINSEY.sizes, dict)
        assert "title" in MCKINSEY.sizes
        assert "body" in MCKINSEY.sizes

    def test_has_slide(self):
        assert isinstance(MCKINSEY.slide, dict)
        assert "width" in MCKINSEY.slide
        assert "height" in MCKINSEY.slide

    def test_has_margins(self):
        assert isinstance(MCKINSEY.margins, dict)
        assert "left" in MCKINSEY.margins
        assert "right" in MCKINSEY.margins

    def test_has_layout(self):
        assert isinstance(MCKINSEY.layout, dict)
        assert "title_top" in MCKINSEY.layout
        assert "body_top" in MCKINSEY.layout

    def test_has_table(self):
        assert isinstance(MCKINSEY.table, dict)
        assert "header_bg" in MCKINSEY.table
        assert "header_fg" in MCKINSEY.table

    def test_all_required_keys_present(self):
        required = {"colors", "fonts", "sizes", "slide", "margins", "layout", "table"}
        actual = {f.name for f in MCKINSEY.__dataclass_fields__.values()}
        assert required.issubset(actual)


class TestResolveColor:
    """resolve_color must map token names via theme and pass-through raw hex."""

    def test_returns_hex_for_known_token(self):
        assert resolve_color(MCKINSEY, "primary") == "#051C2C"

    def test_returns_hex_for_accent_token(self):
        assert resolve_color(MCKINSEY, "accent") == "#2251FF"

    def test_passes_through_hex_with_hash(self):
        assert resolve_color(MCKINSEY, "#FF0000") == "#FF0000"

    def test_passes_through_hex_without_hash(self):
        assert resolve_color(MCKINSEY, "FF0000") == "FF0000"

    def test_unknown_token_passes_through(self):
        # resolve_color does NOT raise; it returns the raw string for unknown tokens
        result = resolve_color(MCKINSEY, "nonexistent_token")
        assert result == "nonexistent_token"

    def test_none_theme_passes_through(self):
        assert resolve_color(None, "#ABCDEF") == "#ABCDEF"
