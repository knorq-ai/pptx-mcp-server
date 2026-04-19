"""
Packaging guardrail tests for the ``mcp`` optional extra (issue #36).

These tests back up the AST-level decoupling in ``tests/test_library_usage.py``
with install-time assertions, so that ``pip install pptx-mcp-server`` stays
lean for library consumers and only ``pip install 'pptx-mcp-server[mcp]'``
pulls the MCP SDK + anyio transitive deps.

Limitation
----------
In this test environment the ``mcp`` package IS importable (it is installed
for the MCP server test suite), so we cannot truly simulate a fresh venv
without ``mcp``. The AST guardrail in ``test_library_usage.py`` already proves
source-level decoupling; the subprocess-based test below shadows the ``mcp``
import to verify the helpful error path surfaces.
"""

from __future__ import annotations

import importlib
import subprocess
import sys
import tomllib
from pathlib import Path

import pytest


REPO_ROOT = Path(__file__).resolve().parent.parent
PYPROJECT = REPO_ROOT / "pyproject.toml"


def _load_pyproject() -> dict:
    return tomllib.loads(PYPROJECT.read_text())


def test_pyproject_declares_mcp_optional_extra() -> None:
    """``[project.optional-dependencies]`` must define an ``mcp`` entry."""
    data = _load_pyproject()
    extras = data.get("project", {}).get("optional-dependencies", {})
    assert "mcp" in extras, (
        "expected [project.optional-dependencies] to define an 'mcp' extra; "
        f"got keys: {sorted(extras)}"
    )
    mcp_extra = extras["mcp"]
    assert any(req.startswith("mcp") for req in mcp_extra), (
        f"'mcp' extra must pin the mcp package; got: {mcp_extra!r}"
    )


def test_pyproject_required_dependencies_exclude_mcp() -> None:
    """``mcp`` must not appear in required ``[project] dependencies``."""
    data = _load_pyproject()
    deps = data.get("project", {}).get("dependencies", [])
    offenders = [d for d in deps if d.split()[0].split(">=")[0].split("==")[0].strip() == "mcp"]
    assert not offenders, (
        "mcp must be an optional extra, not a required dependency; "
        f"found in [project] dependencies: {offenders!r}"
    )


def test_pyproject_keeps_core_dependencies_required() -> None:
    """``python-pptx`` and ``lxml`` must remain required (not moved to extras)."""
    data = _load_pyproject()
    deps = data.get("project", {}).get("dependencies", [])
    roots = {d.split(">=")[0].split("==")[0].strip() for d in deps}
    assert "python-pptx" in roots, f"python-pptx must stay required; got {deps!r}"
    assert "lxml" in roots, f"lxml must stay required; got {deps!r}"


@pytest.mark.parametrize(
    "dotted",
    [
        "pptx_mcp_server.engine.shapes",
        "pptx_mcp_server.theme",
        "pptx_mcp_server.engine.text_metrics",
    ],
)
def test_library_consumer_modules_import_without_mcp_reference(dotted: str) -> None:
    """Library-path modules must import successfully standalone.

    NOTE: this test environment already has ``mcp`` installed, so this check
    cannot prove dependency-free behaviour on its own. The AST guardrail in
    ``tests/test_library_usage.py::test_engine_has_no_mcp_imports`` proves the
    source-level split; this test is a simple smoke check that the dotted
    module paths still resolve.
    """
    mod = importlib.import_module(dotted)
    assert mod is not None


def test_server_import_without_mcp_raises_helpful_error() -> None:
    """If ``mcp`` is absent, importing ``pptx_mcp_server.server`` must raise a
    clear ImportError pointing at ``pip install 'pptx-mcp-server[mcp]'``.

    We simulate absence via a subprocess that injects a ``sitecustomize``-style
    import hook that makes any ``mcp``/``mcp.*`` import fail with
    ``ModuleNotFoundError``.
    """
    script = r"""
import sys
import importlib.abc
import importlib.machinery


class _BlockMcp(importlib.abc.MetaPathFinder):
    def find_spec(self, fullname, path=None, target=None):
        if fullname == "mcp" or fullname.startswith("mcp."):
            raise ModuleNotFoundError(f"No module named {fullname!r}")
        return None


# Purge any already-cached mcp modules, then install the blocker first.
for name in list(sys.modules):
    if name == "mcp" or name.startswith("mcp."):
        del sys.modules[name]
sys.meta_path.insert(0, _BlockMcp())

try:
    import pptx_mcp_server.server  # noqa: F401
except ImportError as e:
    msg = str(e)
    assert "mcp" in msg.lower(), f"error message missing 'mcp' hint: {msg!r}"
    assert "pptx-mcp-server[mcp]" in msg, (
        f"error message must point at the extra; got: {msg!r}"
    )
    print("OK")
else:
    raise AssertionError("expected ImportError when mcp is shadowed, got none")
"""
    result = subprocess.run(
        [sys.executable, "-c", script],
        capture_output=True,
        text=True,
        cwd=str(REPO_ROOT),
    )
    assert result.returncode == 0, (
        f"subprocess failed\nstdout={result.stdout!r}\nstderr={result.stderr!r}"
    )
    assert "OK" in result.stdout, f"unexpected stdout: {result.stdout!r}"


def test_package_main_without_mcp_raises_helpful_error() -> None:
    """The top-level ``pptx_mcp_server.main`` entry point must surface the
    same helpful error when ``mcp`` is not importable."""
    script = r"""
import sys
import importlib.abc
import importlib.machinery


class _BlockMcp(importlib.abc.MetaPathFinder):
    def find_spec(self, fullname, path=None, target=None):
        if fullname == "mcp" or fullname.startswith("mcp."):
            raise ModuleNotFoundError(f"No module named {fullname!r}")
        return None


for name in list(sys.modules):
    if name == "mcp" or name.startswith("mcp."):
        del sys.modules[name]
sys.meta_path.insert(0, _BlockMcp())

import pptx_mcp_server

try:
    pptx_mcp_server.main()
except ImportError as e:
    msg = str(e)
    assert "mcp" in msg.lower(), f"error missing 'mcp' hint: {msg!r}"
    assert "pptx-mcp-server[mcp]" in msg, (
        f"error must point at the extra; got: {msg!r}"
    )
    print("OK")
else:
    raise AssertionError("expected ImportError from main() when mcp is shadowed")
"""
    result = subprocess.run(
        [sys.executable, "-c", script],
        capture_output=True,
        text=True,
        cwd=str(REPO_ROOT),
    )
    assert result.returncode == 0, (
        f"subprocess failed\nstdout={result.stdout!r}\nstderr={result.stderr!r}"
    )
    assert "OK" in result.stdout, f"unexpected stdout: {result.stdout!r}"
