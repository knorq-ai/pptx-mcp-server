"""Tool-tier gating for FastMCP registration (v0.6.0, #137).

FastMCP doesn't natively support tool tiering. This module provides a
custom :func:`mcp_tool` decorator that gates registration by env var,
keeping the default agent-facing surface focused on the productive
block-component + batch layer while hiding atomic raw-shape primitives
behind ``PPTX_MCP_INCLUDE_ADVANCED=1``.

Design notes
------------
* Advanced tools still exist as callable Python functions even when the
  env var is unset — the decorator simply skips FastMCP registration in
  that case. Existing library-mode tests that import tool functions
  directly (``from pptx_mcp_server.server import pptx_add_textbox``)
  keep working untouched.
* The env var is read **once at module import**. Tests that flip the
  gate mid-run must ``importlib.reload`` this module *and* ``server``
  so the registration decisions are re-evaluated.
* Classification is tracked in a module-level set so tests can assert
  exhaustiveness / count stability.
"""

from __future__ import annotations

import os

__all__ = [
    "mcp_tool",
    "is_advanced_enabled",
    "advanced_tool_names",
]

_ADVANCED_TOOLS: set[str] = set()
_INCLUDE_ADVANCED: bool = (
    os.environ.get("PPTX_MCP_INCLUDE_ADVANCED", "").strip().lower()
    in {"1", "true", "yes"}
)


def is_advanced_enabled() -> bool:
    """Return True when ``PPTX_MCP_INCLUDE_ADVANCED`` is set to a truthy
    value at import time (``"1"``, ``"true"``, ``"yes"`` — any casing)."""
    return _INCLUDE_ADVANCED


def mcp_tool(mcp, *, advanced: bool = False):
    """Register a FastMCP tool with tier classification.

    Usage mirrors :meth:`FastMCP.tool`::

        @mcp_tool(mcp)
        def pptx_foo(...): ...

        @mcp_tool(mcp, advanced=True)
        def pptx_bar(...): ...

    When ``advanced=True`` and :envvar:`PPTX_MCP_INCLUDE_ADVANCED` is
    unset, the tool is **not** registered with the MCP server (agents
    won't see it) but the Python function remains importable for
    library-mode callers and tests. When ``advanced=False`` the tool is
    always registered.
    """

    def decorator(fn):
        if advanced:
            _ADVANCED_TOOLS.add(fn.__name__)
            if not _INCLUDE_ADVANCED:
                # Not registered with FastMCP; still importable.
                return fn
        return mcp.tool()(fn)

    return decorator


def advanced_tool_names() -> frozenset[str]:
    """Return the set of tools classified as advanced (registered or
    not). Useful for introspection and for the exhaustiveness test."""
    return frozenset(_ADVANCED_TOOLS)
