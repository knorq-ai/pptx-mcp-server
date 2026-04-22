"""MCP tool tiering (v0.6.0, issue #137).

The ``_tool_registry.mcp_tool`` decorator hides advanced-tier tools from
FastMCP's registry unless ``PPTX_MCP_INCLUDE_ADVANCED`` is set at import
time. These tests verify:

1. Default surface excludes advanced tools.
2. Opting in via env var registers all tools.
3. Advanced tool functions remain importable from the server module so
   library-mode callers and existing tests keep working untouched.
4. The default tier has the expected count (regression guard).
5. Classification is exhaustive — every ``@mcp_tool(...)``-decorated
   tool in ``server.py`` is assigned to exactly one tier.
"""

from __future__ import annotations

import importlib
import re
import sys
from pathlib import Path

import pytest

from pptx_mcp_server import _tool_registry, server


# Snapshot of the expected default tier at v0.6.0. Update this set when
# a new default-tier tool is added; advanced additions should NOT touch
# this fixture (keep the regression guard meaningful).
EXPECTED_DEFAULT_TIER = frozenset({
    # Setup / inspection
    "pptx_create", "pptx_get_info", "pptx_read_slide", "pptx_add_slide",
    # Block components (v0.6.0)
    "pptx_add_section_header", "pptx_add_metric_card_row",
    "pptx_add_kpi_row", "pptx_add_numbered_list",
    "pptx_add_page_marker", "pptx_add_slide_footer",
    # v0.5.0 high-level primitives
    "pptx_add_data_table", "pptx_add_responsive_card_row",
    "pptx_add_milestone_timeline", "pptx_add_flex_container",
    "pptx_add_chart", "pptx_add_image",
    # Batch build
    "pptx_build_slide", "pptx_build_deck",
    # Validation / rendering
    "pptx_check_layout", "pptx_render_slide",
})

EXPECTED_ADVANCED_TIER = frozenset({
    # Shape primitives
    "pptx_add_textbox", "pptx_add_shape", "pptx_add_auto_fit_textbox",
    # Edit ops
    "pptx_edit_text", "pptx_add_paragraph", "pptx_delete_shape",
    "pptx_list_shapes", "pptx_format_shape",
    # Connectors / callouts / icons
    "pptx_add_connector", "pptx_add_callout",
    "pptx_add_icon", "pptx_list_icons",
    # Slide-level editing
    "pptx_set_dimensions", "pptx_set_slide_background",
    "pptx_move_slide", "pptx_delete_slide", "pptx_duplicate_slide",
    # Table editing
    "pptx_add_table", "pptx_edit_table_cell",
    "pptx_edit_table_cells", "pptx_format_table",
    # Composite helpers
    "pptx_add_section_divider", "pptx_add_content_slide",
    "pptx_add_bullet_block",
    # Legacy (deprecated in v0.7.0)
    "pptx_add_kpi_row_legacy",
})


def _registered_tool_names() -> set[str]:
    """Live FastMCP registry of the currently-imported server module."""
    return set(server.mcp._tool_manager._tools.keys())


def _reload_with_env(monkeypatch: pytest.MonkeyPatch, value: str | None):
    """Reload ``_tool_registry`` + ``server`` under a specific env state.

    The env var is read once at module import, so flipping the gate
    mid-test requires re-importing both modules. Returns the reloaded
    ``server`` module so callers can introspect it.
    """
    if value is None:
        monkeypatch.delenv("PPTX_MCP_INCLUDE_ADVANCED", raising=False)
    else:
        monkeypatch.setenv("PPTX_MCP_INCLUDE_ADVANCED", value)
    # Drop cached modules so reload re-runs module-body code fresh.
    sys.modules.pop("pptx_mcp_server.server", None)
    sys.modules.pop("pptx_mcp_server._tool_registry", None)
    import pptx_mcp_server._tool_registry as reg  # noqa: F401
    import pptx_mcp_server.server as srv
    return srv


# ---------------------------------------------------------------------------
# Regression fixture — reset modules after each test so the live imports at
# top of this file see the "unset env" state for other tests / collections.
# ---------------------------------------------------------------------------

@pytest.fixture(autouse=True)
def _restore_default_registry():
    yield
    # After any test that mutated env + reloaded, restore the import with
    # no env var so the rest of the suite sees the default tier.
    sys.modules.pop("pptx_mcp_server.server", None)
    sys.modules.pop("pptx_mcp_server._tool_registry", None)
    # Re-import so ``server`` / ``_tool_registry`` at module-scope of
    # OTHER test files remain consistent.
    import pptx_mcp_server._tool_registry  # noqa: F401
    import pptx_mcp_server.server  # noqa: F401


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

def test_default_surface_excludes_advanced_tools(monkeypatch):
    """Without the env var, no advanced tool appears in the registry."""
    srv = _reload_with_env(monkeypatch, None)
    registered = set(srv.mcp._tool_manager._tools.keys())
    leaked = registered & EXPECTED_ADVANCED_TIER
    assert not leaked, f"Advanced tools leaked into default surface: {sorted(leaked)}"
    # And every default tool is there.
    missing = EXPECTED_DEFAULT_TIER - registered
    assert not missing, f"Default tools missing: {sorted(missing)}"


def test_include_advanced_env_registers_all(monkeypatch):
    """With PPTX_MCP_INCLUDE_ADVANCED=1, all 45 tools register."""
    srv = _reload_with_env(monkeypatch, "1")
    registered = set(srv.mcp._tool_manager._tools.keys())
    expected = EXPECTED_DEFAULT_TIER | EXPECTED_ADVANCED_TIER
    missing = expected - registered
    assert not missing, f"Tools missing after opt-in: {sorted(missing)}"
    assert len(registered) == len(expected), (
        f"Expected {len(expected)} tools with advanced on, "
        f"got {len(registered)}: unexpected extras "
        f"{sorted(registered - expected)}"
    )


@pytest.mark.parametrize("truthy", ["1", "true", "yes", "TRUE", "Yes"])
def test_env_var_truthy_values(monkeypatch, truthy):
    """The env var accepts case-insensitive ``1``/``true``/``yes``."""
    srv = _reload_with_env(monkeypatch, truthy)
    registered = set(srv.mcp._tool_manager._tools.keys())
    assert EXPECTED_ADVANCED_TIER <= registered, (
        f"Advanced tier should register when env={truthy!r}"
    )


def test_advanced_functions_still_importable(monkeypatch):
    """Advanced tool *functions* remain importable even when gated off.

    This is critical: existing tests (test_server.py, etc.) call tool
    functions directly as library functions. Gating only removes the
    FastMCP registration, not the Python symbol.
    """
    srv = _reload_with_env(monkeypatch, None)
    # Sample advanced functions from different sub-categories.
    assert callable(srv.pptx_add_textbox)
    assert callable(srv.pptx_add_shape)
    assert callable(srv.pptx_edit_text)
    assert callable(srv.pptx_list_shapes)
    assert callable(srv.pptx_add_kpi_row_legacy)
    assert callable(srv.pptx_format_table)
    # And none of these are in the default tier registry.
    registered = set(srv.mcp._tool_manager._tools.keys())
    for name in (
        "pptx_add_textbox", "pptx_add_shape", "pptx_edit_text",
        "pptx_list_shapes", "pptx_add_kpi_row_legacy", "pptx_format_table",
    ):
        assert name not in registered, (
            f"{name} should be hidden from the default MCP registry"
        )


def test_default_tier_count_matches_expectation(monkeypatch):
    """Freeze the default-tier count as a regression guard.

    Update ``EXPECTED_DEFAULT_TIER`` and this count together when
    promoting a new tool into the default surface.

    Uses ``_reload_with_env(None)`` so the assertion is independent of
    the ambient ``PPTX_MCP_INCLUDE_ADVANCED`` env state — without this,
    running under ``PPTX_MCP_INCLUDE_ADVANCED=1`` would surface 45 tools
    in the live registry and fail the count.
    """
    srv = _reload_with_env(monkeypatch, None)
    registered = set(srv.mcp._tool_manager._tools.keys())
    default_count = len(EXPECTED_DEFAULT_TIER)
    assert default_count == 20, (
        f"Default tier expected size 20, got {default_count}. "
        "Update the test + EXPECTED_DEFAULT_TIER together."
    )
    assert len(registered) == default_count, (
        f"Registered {len(registered)} tools, expected "
        f"{default_count} in default tier."
    )


def test_classification_is_exhaustive():
    """Every ``@mcp_tool(...)``-decorated tool in server.py is classified.

    Parses the server source to find every decorator + function pair,
    then asserts the name is either in the default or advanced tier. If
    this fails, a new tool was added without a tier decision.
    """
    src = Path(server.__file__).read_text()
    # Match "@mcp_tool(mcp...)\n[async ]def pptx_x(":
    # capture the tier flag and the function name separately.
    pattern = re.compile(
        r"^@mcp_tool\(mcp(?P<args>[^)]*)\)\n"
        r"(?:async\s+)?def\s+(?P<name>pptx_[A-Za-z0-9_]+)\(",
        re.MULTILINE,
    )
    found = {m.group("name"): ("advanced=True" in m.group("args"))
             for m in pattern.finditer(src)}
    assert found, "No decorated tools found — regex broken?"
    assert len(found) == 45, (
        f"Expected 45 decorated tools in server.py, found {len(found)}. "
        "Did the regex miss a new decorator form?"
    )
    classified = EXPECTED_DEFAULT_TIER | EXPECTED_ADVANCED_TIER
    unclassified = set(found) - classified
    assert not unclassified, (
        f"Tools in server.py not in either tier fixture: {sorted(unclassified)}. "
        "Add them to EXPECTED_DEFAULT_TIER or EXPECTED_ADVANCED_TIER."
    )
    # And the decorator flag matches the fixture.
    for name, is_advanced in found.items():
        if is_advanced:
            assert name in EXPECTED_ADVANCED_TIER, (
                f"{name} decorated advanced=True but not in advanced fixture"
            )
        else:
            assert name in EXPECTED_DEFAULT_TIER, (
                f"{name} decorated without advanced=True but in advanced fixture"
            )


def test_advanced_tool_names_helper_matches_fixture(monkeypatch):
    """The ``advanced_tool_names()`` introspection helper agrees with the
    decorator calls found in server.py."""
    srv = _reload_with_env(monkeypatch, None)
    tracked = set(srv.advanced_tool_names())
    assert tracked == EXPECTED_ADVANCED_TIER, (
        f"advanced_tool_names() mismatch.\n"
        f"  extra={sorted(tracked - EXPECTED_ADVANCED_TIER)}\n"
        f"  missing={sorted(EXPECTED_ADVANCED_TIER - tracked)}"
    )


def test_is_advanced_enabled_reflects_env(monkeypatch):
    srv_off = _reload_with_env(monkeypatch, None)
    assert srv_off.is_advanced_enabled() is False
    srv_on = _reload_with_env(monkeypatch, "1")
    assert srv_on.is_advanced_enabled() is True
