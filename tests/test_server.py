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
