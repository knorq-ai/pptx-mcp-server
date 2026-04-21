"""
Unit tests for theme definitions and resolve_color.
"""

from __future__ import annotations

import pytest

import importlib

from pptx_mcp_server.theme import IR, MCKINSEY, Theme, get_theme, list_themes, resolve_color


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


class TestIRTheme:
    """IR theme: クリーム背景 + ネイビーの和製コーポレート IR プリセット。"""

    def test_get_theme_returns_ir(self):
        theme = get_theme("ir")
        assert theme is not None
        assert theme is IR
        # 必須色キーが揃っていること
        expected_color_keys = {
            "primary",
            "accent",
            "background",
            "rule_strong",
            "rule_subtle",
            "highlight_row",
            "positive",
            "negative",
            "text_primary",
            "text_secondary",
        }
        assert expected_color_keys.issubset(theme.colors.keys())

    def test_slide_dimensions_hd_widescreen(self):
        # IR spec: 20.0 x 11.25 インチ (HD ワイド)。既定の 13.333 x 7.5 ではない。
        assert IR.slide["width"] == 20.0
        assert IR.slide["height"] == 11.25

    def test_east_asian_font_is_yu_gothic(self):
        assert IR.fonts["east_asian"] == "Yu Gothic"

    def test_chart_colors_has_six_entries(self):
        assert len(IR.chart_colors) == 6

    def test_ir_in_list_themes(self):
        assert "ir" in list_themes()

    def test_registration_is_idempotent(self):
        # theme.py を再 import しても重複登録されない (register_theme は上書き)。
        before = len(list_themes())
        import pptx_mcp_server.theme as theme_module

        importlib.reload(theme_module)
        # reload 後も "ir" は 1 回だけ登録されている。
        from pptx_mcp_server.theme import list_themes as list_themes_reloaded

        after = len(list_themes_reloaded())
        assert after == before
        assert list_themes_reloaded().count("ir") == 1


# ---------------------------------------------------------------------------
# resolve_theme_color (v0.5.0 central resolver — #123/#124/#125)
# ---------------------------------------------------------------------------


class TestResolveThemeColor:
    def test_theme_token_resolves_to_hex(self):
        from pptx_mcp_server.theme import resolve_theme_color
        # IR theme's "rule_subtle" = "#E0E0E0"
        assert resolve_theme_color("rule_subtle", "ir") == "E0E0E0"
        assert resolve_theme_color("primary", "ir") == "0A2540"
        assert resolve_theme_color("highlight_row", "ir") == "F0F0F0"

    def test_raw_hex_passthrough_with_hash_stripped(self):
        from pptx_mcp_server.theme import resolve_theme_color
        assert resolve_theme_color("#FF0000", "ir") == "FF0000"
        assert resolve_theme_color("051C2C", "mckinsey") == "051C2C"

    def test_unknown_token_passthrough(self):
        from pptx_mcp_server.theme import resolve_theme_color
        # Unknown token with theme — no resolution match, returns as-is (stripped)
        assert resolve_theme_color("no_such_token", "ir") == "no_such_token"

    def test_empty_string_returns_empty(self):
        from pptx_mcp_server.theme import resolve_theme_color
        # Empty string is caller's "disable" signal
        assert resolve_theme_color("", "ir") == ""
        assert resolve_theme_color("", None) == ""

    def test_no_theme_passthrough_strips_hash(self):
        from pptx_mcp_server.theme import resolve_theme_color
        assert resolve_theme_color("#ABCDEF", None) == "ABCDEF"
        # No theme + unknown token → token returned as-is (caller deals)
        assert resolve_theme_color("primary", None) == "primary"

    def test_unregistered_theme_name_falls_back(self):
        from pptx_mcp_server.theme import resolve_theme_color
        # Theme doesn't exist → behaves like theme=None
        assert resolve_theme_color("#ABCDEF", "nonexistent") == "ABCDEF"
