"""layout_constants 単一ソース定数を pin するテスト.

issue #42 — textbox の内側 padding は shapes.py / cards.py / validation.py の
3 箇所に magic number として散らばっていた。centralization 後もドリフトさせ
ないよう、値自体と alias の一致を明示的にアサートする。
"""

from __future__ import annotations

from pptx_mcp_server.engine import layout_constants
from pptx_mcp_server.engine import shapes as shapes_mod
from pptx_mcp_server.engine import cards as cards_mod
from pptx_mcp_server.engine.layout_constants import (
    TEXTBOX_INNER_PADDING_PER_SIDE,
    TEXTBOX_INNER_PADDING_TOTAL,
)


def test_per_side_value_pinned() -> None:
    """per-side padding は 0.05" で固定。変更は意図的に行う."""
    assert TEXTBOX_INNER_PADDING_PER_SIDE == 0.05


def test_total_value_pinned() -> None:
    """total は per-side の 2 倍 (= 0.10")."""
    assert TEXTBOX_INNER_PADDING_TOTAL == 0.10
    assert TEXTBOX_INNER_PADDING_TOTAL == 2 * TEXTBOX_INNER_PADDING_PER_SIDE


def test_shapes_alias_matches_constant() -> None:
    """shapes._AUTO_FIT_PADDING_PER_SIDE は layout_constants と同一値を指す."""
    assert shapes_mod._AUTO_FIT_PADDING_PER_SIDE == TEXTBOX_INNER_PADDING_PER_SIDE
    # 旧名 alias も同一値
    assert shapes_mod._AUTO_FIT_PADDING == TEXTBOX_INNER_PADDING_PER_SIDE


def test_cards_alias_matches_constant() -> None:
    """cards.py が layout_constants と同一の per-side 定数を使う."""
    assert cards_mod._AUTO_FIT_PADDING_PER_SIDE == TEXTBOX_INNER_PADDING_PER_SIDE


def test_exported_from_engine_package() -> None:
    """engine パッケージ経由でも同じ値を import できる."""
    from pptx_mcp_server import engine

    assert engine.TEXTBOX_INNER_PADDING_PER_SIDE == 0.05
    assert engine.TEXTBOX_INNER_PADDING_TOTAL == 0.10
    assert "TEXTBOX_INNER_PADDING_PER_SIDE" in engine.__all__
    assert "TEXTBOX_INNER_PADDING_TOTAL" in engine.__all__


def test_module_has_future_annotations() -> None:
    """layout_constants は ``from __future__ import annotations`` を使う."""
    import inspect

    src = inspect.getsource(layout_constants)
    assert "from __future__ import annotations" in src
