"""
Integration tests for MCP server tool functions.

These tests call the tool functions directly (they are regular Python functions
wrapped by the MCP decorator). All file-based operations use tmp_path.

v0.2.0 response contract (issue #88):
    - Success: json.dumps({"ok": true, "result": <legacy payload>})
    - Error:   json.dumps({"ok": false, "error": {"code", "message", ...}})
Tests unwrap via ``_unwrap_ok`` / ``_unwrap_err`` helpers below.
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
    INSTRUCTIONS,
)


# ── Response-shape helpers (v0.2.0, issue #88) ──────────────────────────

def _unwrap_ok(payload: str):
    """Parse a tool response and assert it is a success; return the result."""
    parsed = json.loads(payload)
    assert isinstance(parsed, dict), f"response must be a JSON object, got {type(parsed)}"
    assert parsed.get("ok") is True, f"expected ok=true, got {parsed!r}"
    assert "result" in parsed, f"success payload must include 'result': {parsed!r}"
    return parsed["result"]


def _unwrap_err(payload: str) -> dict:
    """Parse a tool response and assert it is an error; return the error dict."""
    parsed = json.loads(payload)
    assert isinstance(parsed, dict)
    assert parsed.get("ok") is False, f"expected ok=false, got {parsed!r}"
    err = parsed.get("error")
    assert isinstance(err, dict), f"error must be an object, got {err!r}"
    assert "code" in err and "message" in err, f"error missing required keys: {err!r}"
    return err


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
        msg = _unwrap_ok(result)
        assert "Created" in msg


class TestFileBased:
    """File-based tool calls modify the underlying PPTX correctly."""

    def test_add_slide(self, deck):
        result = pptx_add_slide(deck)
        assert "Added slide" in _unwrap_ok(result)
        prs = open_pptx(deck)
        assert len(prs.slides) == 2

    def test_add_textbox(self, deck):
        result = pptx_add_textbox(deck, 0, 1, 1, 4, 0.5, text="Hello")
        assert "Added textbox" in _unwrap_ok(result)
        prs = open_pptx(deck)
        slide = prs.slides[0]
        texts = [s.text_frame.text for s in slide.shapes if s.has_text_frame]
        assert "Hello" in texts

    def test_add_table(self, deck):
        rows_json = json.dumps([["Name", "Score"], ["Alice", "95"]])
        result = pptx_add_table(deck, 0, rows_json, 1, 1, 5)
        assert "Added table" in _unwrap_ok(result)
        prs = open_pptx(deck)
        slide = prs.slides[0]
        table_shapes = [s for s in slide.shapes if s.has_table]
        assert len(table_shapes) == 1


class TestCompositeTools:
    """Composite tool calls produce correct output."""

    def test_add_content_slide(self, deck):
        result = pptx_add_content_slide(deck, "Revenue Analysis")
        # auto-render disabled by default → result is plain string payload
        payload = _unwrap_ok(result)
        assert "Added content slide" in payload
        prs = open_pptx(deck)
        assert len(prs.slides) == 2  # original + content

    def test_add_section_divider(self, deck):
        result = pptx_add_section_divider(deck, "Q1 Results", subtitle="FY2024")
        payload = _unwrap_ok(result)
        assert "Added section divider" in payload
        prs = open_pptx(deck)
        assert len(prs.slides) == 2


class TestErrorCases:
    """Error cases must return structured JSON error payloads (issue #88)."""

    def test_open_nonexistent_returns_structured_error(self, tmp_path):
        result = pptx_get_info(str(tmp_path / "nope.pptx"))
        err = _unwrap_err(result)
        assert err["code"] == "FILE_NOT_FOUND"
        assert "message" in err

    def test_invalid_slide_returns_structured_error(self, deck):
        result = pptx_read_slide(deck, 99)
        err = _unwrap_err(result)
        assert err["code"] == "SLIDE_NOT_FOUND"

    def test_invalid_shape_type_returns_structured_error(self, deck):
        result = pptx_add_shape(deck, 0, "nonexistent", 1, 1, 2, 2)
        err = _unwrap_err(result)
        assert err["code"] == "INVALID_PARAMETER"


class TestJsonParsing:
    """JSON-based tools must parse input correctly and reject invalid JSON."""

    def test_add_table_with_valid_json(self, deck):
        rows = json.dumps([["X", "Y"], ["1", "2"]])
        result = pptx_add_table(deck, 0, rows, 1, 1, 5)
        assert "Added table" in _unwrap_ok(result)

    def test_add_kpi_row_with_valid_json(self, deck):
        kpis = json.dumps([{"value": "99", "label": "Score"}])
        result = pptx_add_kpi_row(deck, 0, kpis, 2.0)
        payload = _unwrap_ok(result)
        # payload may be either the plain legacy string (render disabled) or
        # a dict with {"value": ..., "render_warning"/"preview_path"} when enabled.
        if isinstance(payload, dict):
            payload = payload.get("value", "")
        assert "Added 1 KPI" in payload

    def test_invalid_json_returns_error(self, deck):
        result = pptx_add_table(deck, 0, "not valid json", 1, 1, 5)
        err = _unwrap_err(result)
        # json.loads on gibberish raises JSONDecodeError → INTERNAL_ERROR wrap
        assert err["code"] in {"INVALID_PARAMETER", "INTERNAL_ERROR"}

    def test_add_bullet_block_with_valid_json(self, deck):
        items = json.dumps(["Point A", "Point B"])
        result = pptx_add_bullet_block(deck, 0, items, 1, 2, 5, 3)
        payload = _unwrap_ok(result)
        if isinstance(payload, dict):
            payload = payload.get("value", "")
        assert "Added bullet block" in payload


# ── Issue #33: pptx_check_layout back-compat ────────────────────────

class TestCheckLayoutBackCompat:
    """``pptx_check_layout`` wraps the legacy string in the new ``result`` field.

    Issue #33 の legacy 文字列は ``result`` フィールドに入ったまま保持される。
    v0.2.0 breaking: caller は ``_unwrap_ok`` で result を取り出す必要がある。
    """

    def test_clean_deck_returns_legacy_string(self, deck):
        result = pptx_check_layout(deck)
        payload = _unwrap_ok(result)
        assert isinstance(payload, str)
        assert payload.startswith("All slides clean") or payload.startswith("Found")

    def test_clean_deck_wording_exact(self, deck):
        payload = _unwrap_ok(pptx_check_layout(deck))
        if payload.startswith("All slides clean"):
            assert payload == (
                "All slides clean — no overlaps, out-of-bounds, text overflow, "
                "or readability issues detected."
            )

    def test_deck_with_overlap_returns_found_string(self, deck):
        pptx_add_shape(deck, 0, "rectangle", 1.0, 1.0, 3.0, 2.0, fill_color="2251FF")
        pptx_add_shape(deck, 0, "rectangle", 2.0, 1.5, 3.0, 2.0, fill_color="FF0000")
        payload = _unwrap_ok(pptx_check_layout(deck))
        assert isinstance(payload, str)
        assert payload.startswith("Found")
        assert "overlap" in payload.lower()

    def test_detailed_returns_json(self, deck):
        payload = _unwrap_ok(pptx_check_layout(deck, detailed=True))
        # payload is itself a JSON string (legacy behaviour of detailed=True)
        assert isinstance(payload, str)
        parsed = json.loads(payload)
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
        err = _unwrap_err(result)
        # error code should be INTERNAL_ERROR (boom is a RuntimeError).
        assert err["code"] == "INTERNAL_ERROR"
        assert "boom" in err["message"]
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
        payload = _unwrap_ok(result)
        # render disabled by default → payload is the plain cards dict.
        assert "cards" in payload
        assert len(payload["cards"]) == 2
        assert {"left", "top", "width", "height"} <= set(payload["cards"][0].keys())
        assert "consumed_height" in payload
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
        payload = _unwrap_ok(result)
        assert "allocations" in payload

    def test_raw_list_rejected(self, deck):
        result = pptx_add_flex_container(
            file_path=deck,
            slide_index=0,
            items_json=[{"sizing": "fixed", "size": 2.0, "type": "rectangle"}],  # type: ignore[arg-type]
            left=0.5, top=0.5, width=10.0, height=1.0,
        )
        err = _unwrap_err(result)
        assert err["code"] == "INVALID_PARAMETER"
        assert err.get("parameter") == "items_json"

    def test_invalid_json_rejected(self, deck):
        result = pptx_add_flex_container(
            file_path=deck,
            slide_index=0,
            items_json="not valid json",
            left=0.5, top=0.5, width=10.0, height=1.0,
        )
        err = _unwrap_err(result)
        assert err["code"] == "INVALID_PARAMETER"
        assert err.get("parameter") == "items_json"

    def test_json_decoded_object_rejected(self, deck):
        result = pptx_add_flex_container(
            file_path=deck,
            slide_index=0,
            items_json='{"not": "array"}',
            left=0.5, top=0.5, width=10.0, height=1.0,
        )
        err = _unwrap_err(result)
        assert err["code"] == "INVALID_PARAMETER"


# ── #43: CardSpec JSON strict unknown-key validation ───────────────

class TestResponsiveCardRowStrictKeys:
    """``pptx_add_responsive_card_row`` が未知キーを弾く (#43)."""

    def test_unknown_key_rejected_with_index_and_name(self, deck):
        cards_json = json.dumps([{"title": "t", "body": "b", "typo": "x"}])
        result = pptx_add_responsive_card_row(
            file_path=deck,
            slide_index=0,
            cards_json=cards_json,
            left=0.5, top=0.5, width=10.0, max_height=3.0,
        )
        err = _unwrap_err(result)
        assert err["code"] == "INVALID_PARAMETER"
        assert "typo" in err["message"]
        assert "card[0]" in err["message"]

    def test_unknown_key_index_points_to_bad_card(self, deck):
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
        err = _unwrap_err(result)
        assert err["code"] == "INVALID_PARAMETER"
        assert "card[1]" in err["message"]
        assert "oops" in err["message"]

    def test_valid_cards_ok_regression(self, deck):
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
        payload = _unwrap_ok(result)
        assert "cards" in payload
        assert len(payload["cards"]) == 2


# ── Issue #88: structured response contract ────────────────────────

class TestStructuredResponseContract:
    """All tool returns follow the v0.2.0 ``{ok, result|error}`` contract."""

    def test_success_shape_has_ok_and_result(self, tmp_path):
        path = str(tmp_path / "s.pptx")
        response = pptx_create(path)
        parsed = json.loads(response)
        assert parsed == {"ok": True, "result": parsed["result"]}
        assert isinstance(parsed["result"], str)

    def test_caller_can_discriminate_on_ok(self, deck, tmp_path):
        ok_response = pptx_get_info(deck)
        err_response = pptx_get_info(str(tmp_path / "missing.pptx"))
        assert json.loads(ok_response)["ok"] is True
        assert json.loads(err_response)["ok"] is False

    def test_error_payload_has_code_and_message(self, tmp_path):
        err = _unwrap_err(pptx_get_info(str(tmp_path / "nope.pptx")))
        assert err["code"] == "FILE_NOT_FOUND"
        assert isinstance(err["message"], str) and err["message"]

    def test_error_engine_code_preserved(self, deck):
        err = _unwrap_err(pptx_read_slide(deck, 99))
        assert err["code"] == "SLIDE_NOT_FOUND"

    def test_error_can_include_parameter_and_hint(self, deck):
        """flex_container raw-list rejection sets parameter + hint + issue."""
        err = _unwrap_err(pptx_add_flex_container(
            file_path=deck,
            slide_index=0,
            items_json=[],  # type: ignore[arg-type]
            left=0.5, top=0.5, width=10.0, height=1.0,
        ))
        assert err["parameter"] == "items_json"
        assert "hint" in err
        assert err.get("issue") == 35

    def test_all_error_paths_return_valid_json(self, tmp_path):
        """Every error path must round-trip through json.loads cleanly."""
        # A handful of representative error sources.
        responses = [
            pptx_get_info(str(tmp_path / "x.pptx")),
            pptx_read_slide(str(tmp_path / "x.pptx"), 0),
            pptx_list_shapes(str(tmp_path / "x.pptx"), 0),
        ]
        for r in responses:
            parsed = json.loads(r)
            assert parsed["ok"] is False
            assert parsed["error"]["code"]
            assert parsed["error"]["message"]


# ── Issue #86: auto_render opt-in + timeout ────────────────────────

class TestAutoRenderOptIn:
    """Auto-render is OFF by default; opt-in via PPTX_MCP_AUTO_RENDER.

    See issue #86. All tests install monkey-patched ``render_slide`` so no
    real LibreOffice subprocess ever fires.
    """

    def test_default_does_not_invoke_renderer(self, deck, monkeypatch):
        """With PPTX_MCP_AUTO_RENDER unset, renderer is NOT called."""
        from pptx_mcp_server import server as server_mod

        monkeypatch.delenv("PPTX_MCP_AUTO_RENDER", raising=False)
        called = {"count": 0}

        def fake_render(*a, **kw):
            called["count"] += 1
            return "/tmp/fake.png"

        monkeypatch.setattr(server_mod, "render_slide", fake_render)

        result = server_mod.pptx_add_content_slide(deck, "Title")
        _unwrap_ok(result)
        assert called["count"] == 0, "renderer should not fire when auto-render disabled"

    def test_env_var_enables_renderer(self, deck, monkeypatch):
        """With PPTX_MCP_AUTO_RENDER=1, renderer IS called."""
        from pptx_mcp_server import server as server_mod

        monkeypatch.setenv("PPTX_MCP_AUTO_RENDER", "1")
        called = {"count": 0}

        def fake_render(file_path, slide_index=-1, dpi=100):
            called["count"] += 1
            return "/tmp/fake.png"

        monkeypatch.setattr(server_mod, "render_slide", fake_render)

        result = server_mod.pptx_add_content_slide(deck, "Title")
        payload = _unwrap_ok(result)
        assert called["count"] == 1, "renderer should fire when auto-render enabled"
        # result is wrapped as {"value": <legacy>, "preview_path": ...}
        assert isinstance(payload, dict)
        assert payload["preview_path"] == "/tmp/fake.png"
        assert "Added content slide" in payload["value"]

    def test_env_var_various_truthy_values(self, monkeypatch):
        from pptx_mcp_server import server as server_mod
        for v in ("1", "true", "TRUE", "yes", "on"):
            monkeypatch.setenv("PPTX_MCP_AUTO_RENDER", v)
            assert server_mod._auto_render_enabled(), f"{v!r} should enable"
        for v in ("0", "false", "no", "off", ""):
            monkeypatch.setenv("PPTX_MCP_AUTO_RENDER", v)
            assert not server_mod._auto_render_enabled(), f"{v!r} should disable"

    def test_render_timeout_returns_warning_quickly(self, deck, monkeypatch):
        """Slow renderer + tight timeout → tool returns within ~timeout+buffer
        with a render_warning (not a failure)."""
        import time

        from pptx_mcp_server import server as server_mod

        monkeypatch.setenv("PPTX_MCP_AUTO_RENDER", "1")
        monkeypatch.setenv("PPTX_MCP_RENDER_TIMEOUT", "1")

        def slow_render(*a, **kw):
            time.sleep(20)
            return "/tmp/never.png"

        monkeypatch.setattr(server_mod, "render_slide", slow_render)

        t0 = time.time()
        result = server_mod.pptx_add_content_slide(deck, "Title")
        elapsed = time.time() - t0

        # Must return well before the 20s the renderer would take.
        assert elapsed < 5.0, f"timeout did not kick in: {elapsed:.1f}s"
        payload = _unwrap_ok(result)
        assert isinstance(payload, dict)
        assert "render_warning" in payload
        assert payload["render_warning"]["reason"] == "timeout"
        # primary action still succeeded.
        assert "Added content slide" in payload["value"]

    def test_render_failure_does_not_fail_tool(self, deck, monkeypatch):
        """Renderer raising an exception → tool still returns ok=true."""
        from pptx_mcp_server import server as server_mod

        monkeypatch.setenv("PPTX_MCP_AUTO_RENDER", "1")

        def broken_render(*a, **kw):
            raise RuntimeError("libreoffice exploded")

        monkeypatch.setattr(server_mod, "render_slide", broken_render)

        result = server_mod.pptx_add_content_slide(deck, "Title")
        payload = _unwrap_ok(result)  # primary action still succeeds
        assert isinstance(payload, dict)
        assert "render_warning" in payload
        assert payload["render_warning"]["reason"] == "failed"
        assert "libreoffice exploded" in payload["render_warning"]["error"]

    def test_primary_action_error_skips_renderer(self, deck, monkeypatch):
        """If primary action fails, renderer must not be invoked."""
        from pptx_mcp_server import server as server_mod

        monkeypatch.setenv("PPTX_MCP_AUTO_RENDER", "1")
        called = {"count": 0}

        def fake_render(*a, **kw):
            called["count"] += 1
            return "/tmp/x.png"

        monkeypatch.setattr(server_mod, "render_slide", fake_render)

        # invalid slide_index → add_kpi_row raises before _auto_render runs.
        result = server_mod.pptx_add_kpi_row(
            deck, 999, json.dumps([{"value": "1", "label": "x"}]), 2.0
        )
        _unwrap_err(result)
        assert called["count"] == 0, "renderer fired despite primary-action error"

    def test_render_timeout_env_var_parsing(self, monkeypatch):
        """Invalid / missing PPTX_MCP_RENDER_TIMEOUT falls back to default."""
        from pptx_mcp_server import server as server_mod

        monkeypatch.delenv("PPTX_MCP_RENDER_TIMEOUT", raising=False)
        assert server_mod._auto_render_timeout() == 10.0

        monkeypatch.setenv("PPTX_MCP_RENDER_TIMEOUT", "garbage")
        assert server_mod._auto_render_timeout() == 10.0

        monkeypatch.setenv("PPTX_MCP_RENDER_TIMEOUT", "0")
        assert server_mod._auto_render_timeout() == 10.0

        monkeypatch.setenv("PPTX_MCP_RENDER_TIMEOUT", "2.5")
        assert server_mod._auto_render_timeout() == 2.5


# ── Issue #90: INSTRUCTIONS trimmed of product policy ──────────────

class TestInstructionsContent:
    """INSTRUCTIONS prose should describe capability, not prescribe UX."""

    def test_no_ask_user_prompt(self):
        lower = INSTRUCTIONS.lower()
        # The banned phrases from the pre-#90 INSTRUCTIONS.
        assert "always ask" not in lower, (
            "INSTRUCTIONS must not tell agents to 'always ask' the user "
            "about product decisions (issue #90)."
        )
        assert "ask the user which color palette" not in lower
        assert "ask user for color palette" not in lower
        assert "ask for color palette" not in lower

    def test_themes_enumeration_present(self):
        # Neutral themes enumeration replaces the ask-first directive.
        assert "mckinsey" in INSTRUCTIONS.lower()
        assert "deloitte" in INSTRUCTIONS.lower()
        assert "neutral" in INSTRUCTIONS.lower()
        assert "Available Themes" in INSTRUCTIONS or "Available themes" in INSTRUCTIONS

    def test_response_shape_documented(self):
        # Caller-facing response shape must be discoverable in the prompt.
        assert '"ok": true' in INSTRUCTIONS
        assert '"ok": false' in INSTRUCTIONS

    def test_auto_render_opt_in_documented(self):
        assert "PPTX_MCP_AUTO_RENDER" in INSTRUCTIONS
