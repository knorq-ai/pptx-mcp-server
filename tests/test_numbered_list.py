"""Tests for the ``add_numbered_list`` block component (Issue #134).

Covers:
- 4-item even height distribution inside the declared bounds
- Empty ``body`` → ``body_shape`` is None (body textbox skipped)
- ``rule_between=False`` → no rule shapes for any item
- Long body auto-fits via truncate-with-ellipsis (no exception)
- Theme='ir' → colors resolve to IR theme hex values
- Single item → no rule drawn
- ``check_containment`` emits zero findings (all children inside bounds)
"""

from __future__ import annotations

import pytest

from pptx_mcp_server.engine.components.numbered_list import (
    NumberedItem,
    add_numbered_list,
    _NUMBER_ROW_H,
    _RULE_ROW_H,
)
from pptx_mcp_server.engine.components.container import clear_container_registry
from pptx_mcp_server.engine.validation import check_containment


EMU_PER_INCH = 914400


def _in(emu: int) -> float:
    return emu / EMU_PER_INCH


@pytest.fixture(autouse=True)
def _isolated_registry():
    """Each test starts with a clean container registry."""
    clear_container_registry()
    yield
    clear_container_registry()


def _four_items() -> list[NumberedItem]:
    return [
        NumberedItem(
            number=f"0{i + 1}",
            caption=f"／ caption {i + 1}",
            title=f"Title {i + 1}",
            body=f"Body copy number {i + 1}.",
        )
        for i in range(4)
    ]


# ---------------------------------------------------------------------------
# Layout
# ---------------------------------------------------------------------------


def test_four_items_evenly_distributed_8in(slide):
    """4 items, height=8" → item tops advance by the item slot height.

    With ``rule_between=True`` the 3 rule rows (0.04" each) eat into
    ``height`` so ``item_height = (8 - 3*0.04) / 4 = 1.97"``. Each item
    advances by ``item_height + 0.04`` except the last.
    """
    result = add_numbered_list(
        slide, _four_items(),
        left=0.5, top=0.5, width=6.0, height=8.0,
        rule_between=True,
    )

    assert len(result["items"]) == 4
    expected_item_h = (8.0 - 3 * _RULE_ROW_H) / 4
    # Each item's bounds.top should advance by (item_h + rule_row_h) per step
    # (except the advance past the last item, which doesn't exist here).
    stride = expected_item_h + _RULE_ROW_H
    for i, entry in enumerate(result["items"]):
        expected_top = 0.5 + i * stride
        assert entry["bounds"]["top"] == pytest.approx(expected_top, abs=1e-4)
        assert entry["bounds"]["height"] == pytest.approx(expected_item_h, abs=1e-4)
        assert entry["bounds"]["left"] == pytest.approx(0.5, abs=1e-6)
        assert entry["bounds"]["width"] == pytest.approx(6.0, abs=1e-6)

    assert result["consumed_height"] == pytest.approx(8.0, abs=1e-6)
    assert result["consumed_width"] == pytest.approx(6.0, abs=1e-6)


def test_empty_body_body_shape_is_none(slide):
    """Item with empty body → ``body_shape`` is None (textbox omitted)."""
    items = [
        NumberedItem("01", "／ a", "Title A", "Body A"),
        NumberedItem("02", "／ b", "Title B", ""),   # empty body
        NumberedItem("03", "／ c", "Title C", "Body C"),
    ]
    result = add_numbered_list(
        slide, items,
        left=0.5, top=0.5, width=6.0, height=6.0,
    )

    assert result["items"][0]["body_shape"] is not None
    assert result["items"][1]["body_shape"] is None
    assert result["items"][2]["body_shape"] is not None


# ---------------------------------------------------------------------------
# Rule behavior
# ---------------------------------------------------------------------------


def test_rule_between_false_no_rules(slide):
    """``rule_between=False`` → every item has ``rule_shape is None``."""
    result = add_numbered_list(
        slide, _four_items(),
        left=0.5, top=0.5, width=6.0, height=8.0,
        rule_between=False,
    )

    for entry in result["items"]:
        assert entry["rule_shape"] is None


def test_single_item_no_rule(slide):
    """n == 1 → no rule drawn regardless of ``rule_between``."""
    items = [NumberedItem("01", "／ solo", "Solo title", "Body of the only item.")]
    result = add_numbered_list(
        slide, items,
        left=1.0, top=1.0, width=5.0, height=3.0,
        rule_between=True,
    )

    assert len(result["items"]) == 1
    assert result["items"][0]["rule_shape"] is None


def test_rule_between_true_last_has_no_rule(slide):
    """Regression: the last item never gets a trailing rule."""
    result = add_numbered_list(
        slide, _four_items(),
        left=0.5, top=0.5, width=6.0, height=8.0,
        rule_between=True,
    )

    # First 3 items have rules, last does not.
    for i, entry in enumerate(result["items"]):
        if i < 3:
            assert entry["rule_shape"] is not None, f"item {i} missing rule"
        else:
            assert entry["rule_shape"] is None


# ---------------------------------------------------------------------------
# Auto-fit
# ---------------------------------------------------------------------------


def test_long_body_auto_fits_with_wrap(slide):
    """A very long body renders without error and gets a live body_shape."""
    long_body = "This is a very long body text that should wrap across multiple lines. " * 10
    items = [
        NumberedItem("01", "／ long", "Long body item", long_body),
    ]
    result = add_numbered_list(
        slide, items,
        left=0.5, top=0.5, width=5.0, height=3.0,
    )

    body_shape = result["items"][0]["body_shape"]
    assert body_shape is not None
    # Body shape width should match the declared width in EMU.
    assert _in(body_shape.width) == pytest.approx(5.0, abs=1e-3)
    # word_wrap must be True for wrapped body rendering.
    assert body_shape.text_frame.word_wrap is True


# ---------------------------------------------------------------------------
# Theme resolution
# ---------------------------------------------------------------------------


def test_theme_ir_resolves_colors(slide):
    """theme='ir' resolves number/caption to text_secondary, title to primary."""
    # IR theme (see theme.py): primary=#0A2540, text_secondary=#6B7280,
    # rule_subtle=#E0E0E0.
    result = add_numbered_list(
        slide,
        [NumberedItem("01", "／ c", "Title 1", "Body 1"),
         NumberedItem("02", "／ c", "Title 2", "Body 2")],
        left=0.5, top=0.5, width=6.0, height=4.0,
        theme="ir",
    )

    number_shape = result["items"][0]["number_shape"]
    caption_shape = result["items"][0]["caption_shape"]
    title_shape = result["items"][0]["title_shape"]
    rule_shape = result["items"][0]["rule_shape"]  # between item 0 and 1

    # Number uses text_secondary (6B7280)
    number_run = number_shape.text_frame.paragraphs[0].runs[0]
    assert str(number_run.font.color.rgb) == "6B7280"

    # Caption uses text_secondary (6B7280)
    caption_run = caption_shape.text_frame.paragraphs[0].runs[0]
    assert str(caption_run.font.color.rgb) == "6B7280"

    # Title uses primary (0A2540)
    title_run = title_shape.text_frame.paragraphs[0].runs[0]
    assert str(title_run.font.color.rgb) == "0A2540"

    # Rule uses rule_subtle (E0E0E0)
    assert rule_shape is not None
    assert str(rule_shape.fill.fore_color.rgb) == "E0E0E0"


# ---------------------------------------------------------------------------
# Containment
# ---------------------------------------------------------------------------


def test_check_containment_zero_findings(one_slide_prs):
    """All rendered children must lie inside the declared bounds."""
    slide = one_slide_prs.slides[0]
    add_numbered_list(
        slide, _four_items(),
        left=0.5, top=0.5, width=6.0, height=8.0,
        rule_between=True,
    )

    findings = check_containment(one_slide_prs)
    assert findings == []


def test_check_containment_zero_findings_no_rules(one_slide_prs):
    """Same but with rule_between=False — still zero findings."""
    slide = one_slide_prs.slides[0]
    add_numbered_list(
        slide, _four_items(),
        left=0.5, top=0.5, width=6.0, height=8.0,
        rule_between=False,
    )

    findings = check_containment(one_slide_prs)
    assert findings == []


# ---------------------------------------------------------------------------
# MCP tool boundary
# ---------------------------------------------------------------------------


def test_mcp_tool_strict_key_validation(pptx_file):
    """Unknown keys on the items dict are rejected with INVALID_PARAMETER."""
    from pptx_mcp_server.server import pptx_add_numbered_list

    out = pptx_add_numbered_list(
        file_path=pptx_file,
        slide_index=0,
        items=[{"number": "01", "caption": "c", "title": "t", "body": "b", "bogus": "x"}],
        left=0.5, top=0.5, width=6.0, height=3.0,
    )
    assert "INVALID_PARAMETER" in out or "unknown" in out


def test_mcp_tool_success_path(pptx_file):
    """Happy path through the MCP tool returns a success envelope."""
    import json

    from pptx_mcp_server.server import pptx_add_numbered_list

    out = pptx_add_numbered_list(
        file_path=pptx_file,
        slide_index=0,
        items=[
            {"number": "01", "caption": "／ a", "title": "T1", "body": "B1"},
            {"number": "02", "caption": "／ b", "title": "T2", "body": "B2"},
        ],
        left=0.5, top=0.5, width=6.0, height=4.0,
    )
    parsed = json.loads(out)
    assert parsed.get("ok") is True
    assert len(parsed["result"]["items"]) == 2
    assert parsed["result"]["consumed_height"] == pytest.approx(4.0, abs=1e-6)
