"""
Integration tests for MCP server tool functions.

These tests call the tool functions directly (they are regular Python functions
wrapped by the MCP decorator). All file-based operations use tmp_path.
"""

from __future__ import annotations

import json
import os

import pytest
from pptx import Presentation

from pptx_mcp_server.engine.pptx_io import save_pptx, open_pptx
from pptx_mcp_server.server import (
    mcp,
    pptx_create,
    pptx_get_info,
    pptx_read_slide,
    pptx_list_shapes,
    pptx_add_slide,
    pptx_delete_slide,
    pptx_duplicate_slide,
    pptx_set_slide_background,
    pptx_set_dimensions,
    pptx_add_textbox,
    pptx_edit_text,
    pptx_add_paragraph,
    pptx_add_shape,
    pptx_delete_shape,
    pptx_format_shape,
    pptx_add_table,
    pptx_edit_table_cell,
    pptx_edit_table_cells,
    pptx_format_table,
    pptx_add_content_slide,
    pptx_add_section_divider,
    pptx_add_kpi_row,
    pptx_add_bullet_block,
    pptx_add_image,
    pptx_render_slide,
    pptx_check_layout,
    pptx_add_flex_container,
    pptx_add_responsive_card_row,
    _format_check_layout_summary,
)


@pytest.fixture
def deck(tmp_path):
    """Create a blank deck with one slide, return file path."""
    path = str(tmp_path / "deck.pptx")
    pptx_create(path)
    pptx_add_slide(path, layout_index=6)
    return path


class TestToolRegistration:
    """All 25 tools must be registered on the MCP server."""

    def test_all_tools_registered(self):
        # FastMCP stores tools internally; list them via the _tool_manager
        tool_names = list(mcp._tool_manager._tools.keys())
        expected = [
            "pptx_create",
            "pptx_get_info",
            "pptx_read_slide",
            "pptx_list_shapes",
            "pptx_add_slide",
            "pptx_delete_slide",
            "pptx_duplicate_slide",
            "pptx_set_slide_background",
            "pptx_set_dimensions",
            "pptx_add_textbox",
            "pptx_edit_text",
            "pptx_add_paragraph",
            "pptx_add_shape",
            "pptx_add_image",
            "pptx_delete_shape",
            "pptx_format_shape",
            "pptx_add_table",
            "pptx_edit_table_cell",
            "pptx_edit_table_cells",
            "pptx_format_table",
            "pptx_add_content_slide",
            "pptx_add_section_divider",
            "pptx_add_kpi_row",
            "pptx_add_bullet_block",
            "pptx_render_slide",
        ]
        for name in expected:
            assert name in tool_names, f"Tool '{name}' not registered"
        assert len(expected) == 25


class TestCreatePptx:
    """pptx_create tool creates a valid file."""

    def test_creates_file(self, tmp_path):
        path = str(tmp_path / "new.pptx")
        result = pptx_create(path)
        assert os.path.exists(path)
        assert "Created" in result


class TestFileBased:
    """File-based tool calls modify the underlying PPTX correctly."""

    def test_add_slide(self, deck):
        result = pptx_add_slide(deck)
        assert "Added slide" in result
        prs = open_pptx(deck)
        assert len(prs.slides) == 2

    def test_add_textbox(self, deck):
        result = pptx_add_textbox(deck, 0, 1, 1, 4, 0.5, text="Hello")
        assert "Added textbox" in result
        prs = open_pptx(deck)
        slide = prs.slides[0]
        texts = [s.text_frame.text for s in slide.shapes if s.has_text_frame]
        assert "Hello" in texts

    def test_add_table(self, deck):
        rows_json = json.dumps([["Name", "Score"], ["Alice", "95"]])
        result = pptx_add_table(deck, 0, rows_json, 1, 1, 5)
        assert "Added table" in result
        prs = open_pptx(deck)
        slide = prs.slides[0]
        table_shapes = [s for s in slide.shapes if s.has_table]
        assert len(table_shapes) == 1


class TestCompositeTools:
    """Composite tool calls produce correct output."""

    def test_add_content_slide(self, deck):
        result = pptx_add_content_slide(deck, "Revenue Analysis")
        assert "Added content slide" in result
        prs = open_pptx(deck)
        assert len(prs.slides) == 2  # original + content

    def test_add_section_divider(self, deck):
        result = pptx_add_section_divider(deck, "Q1 Results", subtitle="FY2024")
        assert "Added section divider" in result
        prs = open_pptx(deck)
        assert len(prs.slides) == 2


class TestErrorCases:
    """Error cases must return formatted error strings, not raise exceptions."""

    def test_open_nonexistent_returns_error_string(self, tmp_path):
        result = pptx_get_info(str(tmp_path / "nope.pptx"))
        assert "FILE_NOT_FOUND" in result

    def test_invalid_slide_returns_error_string(self, deck):
        result = pptx_read_slide(deck, 99)
        assert "SLIDE_NOT_FOUND" in result

    def test_invalid_shape_type_returns_error_string(self, deck):
        result = pptx_add_shape(deck, 0, "nonexistent", 1, 1, 2, 2)
        assert "INVALID_PARAMETER" in result


class TestJsonParsing:
    """JSON-based tools must parse input correctly and reject invalid JSON."""

    def test_add_table_with_valid_json(self, deck):
        rows = json.dumps([["X", "Y"], ["1", "2"]])
        result = pptx_add_table(deck, 0, rows, 1, 1, 5)
        assert "Added table" in result

    def test_add_kpi_row_with_valid_json(self, deck):
        kpis = json.dumps([{"value": "99", "label": "Score"}])
        result = pptx_add_kpi_row(deck, 0, kpis, 2.0)
        assert "Added 1 KPI" in result

    def test_invalid_json_returns_error(self, deck):
        result = pptx_add_table(deck, 0, "not valid json", 1, 1, 5)
        assert "INTERNAL_ERROR" in result or "Error" in result or "error" in result.lower()

    def test_add_bullet_block_with_valid_json(self, deck):
        items = json.dumps(["Point A", "Point B"])
        result = pptx_add_bullet_block(deck, 0, items, 1, 2, 5, 3)
        assert "Added bullet block" in result


# ── Issue #33: pptx_check_layout back-compat ────────────────────────

class TestCheckLayoutBackCompat:
    """``pptx_check_layout`` が legacy string / detailed JSON 両方を返すこと."""

    def test_clean_deck_returns_legacy_string(self, deck):
        # 何も追加していない blank deck は clean (あるいは inconsistent_gaps のみ)
        # である想定だが、少なくとも string を返す契約は不変。
        result = pptx_check_layout(deck)
        assert isinstance(result, str)
        # 空 deck の場合 "All slides clean" が返るはず。
        assert result.startswith("All slides clean") or result.startswith("Found")

    def test_clean_deck_wording_exact(self, deck):
        result = pptx_check_layout(deck)
        if result.startswith("All slides clean"):
            assert result == (
                "All slides clean — no overlaps, out-of-bounds, text overflow, "
                "or readability issues detected."
            )

    def test_deck_with_overlap_returns_found_string(self, deck):
        # 2 つの重なる shape を追加して overlap を誘発する。
        pptx_add_shape(deck, 0, "rectangle", 1.0, 1.0, 3.0, 2.0, fill_color="2251FF")
        pptx_add_shape(deck, 0, "rectangle", 2.0, 1.5, 3.0, 2.0, fill_color="FF0000")
        result = pptx_check_layout(deck)
        assert isinstance(result, str)
        assert result.startswith("Found")
        # overlap 検出メッセージは legacy string で出るはず。
        assert "overlap" in result.lower()

    def test_detailed_returns_json(self, deck):
        result = pptx_check_layout(deck, detailed=True)
        assert isinstance(result, str)
        parsed = json.loads(result)
        assert "slides" in parsed
        assert "summary" in parsed
        assert isinstance(parsed["slides"], list)
        assert set(parsed["summary"].keys()) >= {"errors", "warnings", "infos"}

    def test_format_helper_clean(self):
        empty = {"slides": [], "summary": {"errors": 0, "warnings": 0, "infos": 0}}
        assert _format_check_layout_summary(empty).startswith("All slides clean")

    def test_format_helper_problems(self):
        result = {
            "slides": [
                {
                    "index": 1,
                    "overlaps": ["shape A overlaps shape B"],
                    "out_of_bounds": [],
                    "text_overflow": [
                        {
                            "severity": "error",
                            "slide_index": 1,
                            "shape_name": "tb_1",
                            "category": "text_overflow",
                            "message": "text overflows",
                        }
                    ],
                    "unreadable_text": [],
                    "divider_collision": [],
                    "inconsistent_gaps": [],
                }
            ],
            "summary": {"errors": 2, "warnings": 0, "infos": 0},
        }
        s = _format_check_layout_summary(result)
        assert s.startswith("Found 2 layout issues:")
        assert "Slide 1 [error] overlap:" in s
        assert "Slide 1 [error] text_overflow:" in s


# ── Issue #34: card_row save order ─────────────────────────────────

class TestResponsiveCardRowSaveOrder:
    """``pptx_add_responsive_card_row`` が save を最後に行うこと."""

    def test_save_not_called_when_post_processing_raises(self, deck, monkeypatch):
        """placement serialize 段階で例外が起きたら save は呼ばれないこと."""
        from pptx_mcp_server import server as server_mod

        save_called = {"count": 0}
        original_save = server_mod.save_pptx

        def tracking_save(prs, path):
            save_called["count"] += 1
            return original_save(prs, path)

        monkeypatch.setattr(server_mod, "save_pptx", tracking_save)

        # CardPlacement.left 取得時に発火する boom を注入する。
        # card_row の後段で placements を消費する dict 構築前に save が済んで
        # いないことを確認したいので、add_responsive_card_row をパッチして
        # 異常な object を返させる。
        class Boom:
            @property
            def left(self):
                raise RuntimeError("boom")
            top = 0.0
            width = 0.0
            height = 0.0

        def fake_add_row(slide, cards, **kwargs):
            return [Boom()], 0.0

        monkeypatch.setattr(server_mod, "add_responsive_card_row", fake_add_row)

        cards_json = json.dumps([{"title": "A", "body": "a"}])
        result = server_mod.pptx_add_responsive_card_row(
            file_path=deck,
            slide_index=0,
            cards_json=cards_json,
            left=0.5, top=0.5, width=10.0, max_height=3.0,
        )
        # error が返る。そして save は 0 回。
        assert "boom" in result or "INTERNAL_ERROR" in result
        assert save_called["count"] == 0, (
            f"save_pptx should not have been called; was called {save_called['count']} times"
        )

    def test_happy_path_saves_once_and_returns_placements(self, deck):
        from pptx_mcp_server.server import pptx_add_responsive_card_row

        cards_json = json.dumps(
            [{"title": "A", "body": "aa"}, {"title": "B", "body": "bb"}]
        )
        result = pptx_add_responsive_card_row(
            file_path=deck,
            slide_index=0,
            cards_json=cards_json,
            left=0.5, top=0.5, width=10.0, max_height=3.0,
        )
        # 先頭は JSON (preview 情報が付く場合あり、なければそのまま)
        first_line = result.split("\n")[0]
        parsed = json.loads(first_line)
        assert "cards" in parsed
        assert len(parsed["cards"]) == 2
        assert {"left", "top", "width", "height"} <= set(parsed["cards"][0].keys())
        assert "consumed_height" in parsed
        # ファイルが実際に保存されている (開けること) を確認する。
        prs = open_pptx(deck)
        assert len(prs.slides) >= 1


# ── Issue #35: flex_container items_json ───────────────────────────

class TestFlexContainerItemsJson:
    """``pptx_add_flex_container`` が items_json (JSON 文字列) を受け取る."""

    def test_items_json_string_works(self, deck):
        items_json = json.dumps(
            [{"sizing": "fixed", "size": 2.0, "type": "rectangle", "fill_color": "2251FF"}]
        )
        result = pptx_add_flex_container(
            file_path=deck,
            slide_index=0,
            items_json=items_json,
            left=0.5, top=0.5, width=10.0, height=1.0,
        )
        parsed = json.loads(result)
        assert "allocations" in parsed

    def test_raw_list_rejected(self, deck):
        # 生の Python list を渡すと INVALID_PARAMETER エラー
        result = pptx_add_flex_container(
            file_path=deck,
            slide_index=0,
            items_json=[{"sizing": "fixed", "size": 2.0, "type": "rectangle"}],  # type: ignore[arg-type]
            left=0.5, top=0.5, width=10.0, height=1.0,
        )
        assert "INVALID_PARAMETER" in result

    def test_invalid_json_rejected(self, deck):
        result = pptx_add_flex_container(
            file_path=deck,
            slide_index=0,
            items_json="not valid json",
            left=0.5, top=0.5, width=10.0, height=1.0,
        )
        assert "INVALID_PARAMETER" in result

    def test_json_decoded_object_rejected(self, deck):
        # JSON ではあるが array でない場合
        result = pptx_add_flex_container(
            file_path=deck,
            slide_index=0,
            items_json='{"not": "array"}',
            left=0.5, top=0.5, width=10.0, height=1.0,
        )
        assert "INVALID_PARAMETER" in result


# ── #43: CardSpec JSON strict unknown-key validation ───────────────

class TestResponsiveCardRowStrictKeys:
    """``pptx_add_responsive_card_row`` が未知キーを弾く (#43)."""

    def test_unknown_key_rejected_with_index_and_name(self, deck):
        """``typo`` など CardSpec に無いキーは INVALID_PARAMETER。
        エラーメッセージにインデックスとキー名が含まれる。"""
        cards_json = json.dumps([{"title": "t", "body": "b", "typo": "x"}])
        result = pptx_add_responsive_card_row(
            file_path=deck,
            slide_index=0,
            cards_json=cards_json,
            left=0.5, top=0.5, width=10.0, max_height=3.0,
        )
        assert "INVALID_PARAMETER" in result
        assert "typo" in result
        # どのカードが原因かを示すため、index プレフィックスも含む
        assert "card[0]" in result

    def test_unknown_key_index_points_to_bad_card(self, deck):
        """不良カードが配列の途中にある場合、index が正しく反映される。"""
        cards_json = json.dumps(
            [
                {"title": "ok", "body": "ok"},
                {"title": "bad", "body": "b", "oops": 1},
            ]
        )
        result = pptx_add_responsive_card_row(
            file_path=deck,
            slide_index=0,
            cards_json=cards_json,
            left=0.5, top=0.5, width=10.0, max_height=3.0,
        )
        assert "INVALID_PARAMETER" in result
        assert "card[1]" in result
        assert "oops" in result

    def test_valid_cards_ok_regression(self, deck):
        """既知キーのみのカードは従来どおり成功する (回帰防止)。"""
        cards_json = json.dumps(
            [
                {"title": "A", "body": "aa", "accent_color": "2251FF"},
                {"title": "B", "body": "bb", "padding": 0.25},
            ]
        )
        result = pptx_add_responsive_card_row(
            file_path=deck,
            slide_index=0,
            cards_json=cards_json,
            left=0.5, top=0.5, width=10.0, max_height=3.0,
        )
        # 先頭行は成功時の JSON
        first_line = result.split("\n")[0]
        parsed = json.loads(first_line)
        assert "cards" in parsed
        assert len(parsed["cards"]) == 2
