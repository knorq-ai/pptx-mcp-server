"""py.typed マーカがインストール後に同梱されていることを検証する.

PEP 561 に従い、`py.typed` マーカを含むパッケージは下流の型チェッカに
型注釈を提供するものとして扱われる。editable install・wheel install の
いずれでも、パッケージディレクトリ直下に `py.typed` が存在する必要がある。
"""

from __future__ import annotations

from pathlib import Path

import pptx_mcp_server


def test_py_typed_marker_is_installed() -> None:
    """`py.typed` がインストール済みパッケージディレクトリに存在すること."""
    package_dir = Path(pptx_mcp_server.__file__).parent
    marker = package_dir / "py.typed"
    assert marker.exists(), (
        f"py.typed marker not found at {marker}. "
        "PEP 561 requires this file to expose type annotations to downstream "
        "mypy/pyright consumers."
    )


def test_py_typed_marker_is_empty() -> None:
    """`py.typed` は PEP 561 に従い空ファイルであること."""
    package_dir = Path(pptx_mcp_server.__file__).parent
    marker = package_dir / "py.typed"
    assert marker.is_file()
    assert marker.stat().st_size == 0
