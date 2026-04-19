"""text_metrics モジュールのテスト.

Japanese-aware 幅推定・折り返し・高さ推定が仕様どおりに動作することを確認する。
"""

from __future__ import annotations

import pytest

from pptx_mcp_server.engine.text_metrics import (
    estimate_char_width,
    estimate_text_height,
    estimate_text_width,
    is_cjk,
    wrap_text,
)


class TestIsCjk:
    """is_cjk の判定ロジックを検証する."""

    def test_hiragana(self) -> None:
        assert is_cjk("あ") is True

    def test_katakana(self) -> None:
        assert is_cjk("ア") is True

    def test_kanji(self) -> None:
        assert is_cjk("漢") is True

    def test_fullwidth_punctuation(self) -> None:
        assert is_cjk("、") is True
        assert is_cjk("。") is True

    def test_ascii_letter(self) -> None:
        assert is_cjk("a") is False
        assert is_cjk("A") is False

    def test_ascii_digit(self) -> None:
        assert is_cjk("1") is False
        assert is_cjk("9") is False

    def test_newline_is_not_cjk(self) -> None:
        assert is_cjk("\n") is False

    def test_empty_is_not_cjk(self) -> None:
        assert is_cjk("") is False


class TestEstimateTextWidth:
    """estimate_text_width の数値検証."""

    def test_japanese_bold(self) -> None:
        # "プレミアムモルツ" = 8 文字 × 9pt × 0.0139 × 1.05 ≈ 1.050 inches
        width = estimate_text_width("プレミアムモルツ", 9, bold=True)
        expected = 8 * 9 * 0.0139 * 1.05
        assert width == pytest.approx(expected, abs=0.01)
        assert width == pytest.approx(1.050, abs=0.01)

    def test_ascii_hello_world(self) -> None:
        # "Hello World" = 11 文字 × 12pt × 0.0083 ≈ 1.096 inches
        width = estimate_text_width("Hello World", 12)
        expected = 12 * 11 * 0.0083
        assert width == pytest.approx(expected, abs=0.01)
        assert width == pytest.approx(1.096, abs=0.01)

    def test_mixed_ai_seikatsusha(self) -> None:
        # "AI生活者" = 2 ASCII + 3 CJK @ 10pt
        width = estimate_text_width("AI生活者", 10)
        expected = 2 * 10 * 0.0083 + 3 * 10 * 0.0139
        assert width == pytest.approx(expected, abs=0.001)

    def test_empty_string_has_zero_width(self) -> None:
        assert estimate_text_width("", 12) == 0.0
        assert estimate_text_width("", 12, bold=True) == 0.0

    def test_bold_multiplier(self) -> None:
        normal = estimate_text_width("Hello", 12, bold=False)
        bold = estimate_text_width("Hello", 12, bold=True)
        assert bold == pytest.approx(normal * 1.05, abs=0.0001)

    def test_newline_does_not_add_width(self) -> None:
        # 改行文字は幅に寄与しない
        w_with_nl = estimate_text_width("ab\ncd", 12)
        w_without_nl = estimate_text_width("abcd", 12)
        assert w_with_nl == pytest.approx(w_without_nl, abs=0.0001)


class TestEstimateCharWidth:
    """estimate_char_width のスポットチェック."""

    def test_ascii_char(self) -> None:
        assert estimate_char_width("a", 10) == pytest.approx(10 * 0.0083, abs=0.0001)

    def test_cjk_char(self) -> None:
        assert estimate_char_width("あ", 10) == pytest.approx(10 * 0.0139, abs=0.0001)

    def test_digit_same_as_ascii(self) -> None:
        assert estimate_char_width("0", 10) == pytest.approx(estimate_char_width("a", 10))

    def test_empty_char_zero(self) -> None:
        assert estimate_char_width("", 10) == 0.0


class TestWrapText:
    """wrap_text の分割ロジック."""

    def test_japanese_wraps_to_2_or_3_lines(self) -> None:
        # 仕様: "父の日ギフトシーンで安定露出も定番2位に留まる" を width=2.0, 10pt
        result = wrap_text("父の日ギフトシーンで安定露出も定番2位に留まる", 2.0, 10)
        assert 2 <= len(result) <= 3

    def test_empty_returns_empty_list(self) -> None:
        assert wrap_text("", 2.0, 10) == []

    def test_explicit_newline_forces_line(self) -> None:
        # 明示的 \n を含む場合は強制改行
        result = wrap_text("abc\ndef", 10.0, 12)
        assert result == ["abc", "def"]

    def test_multiple_newlines_preserve_blank_line(self) -> None:
        result = wrap_text("a\n\nb", 10.0, 12)
        assert len(result) == 3
        assert result[0] == "a"
        assert result[1] == ""
        assert result[2] == "b"

    def test_long_word_placed_on_own_line(self) -> None:
        # max_width より長い ASCII ワードも単独行として配置される (無限ループしない)
        long_word = "supercalifragilisticexpialidocious"
        result = wrap_text(long_word, 0.5, 12)
        # 少なくとも 1 行にそのまま置かれる
        assert any(long_word in line for line in result)
        assert len(result) >= 1

    def test_ascii_word_wrap_at_space(self) -> None:
        # "Hello World Foo Bar Baz" を狭めの幅で折り返す → 空白で分割される
        result = wrap_text("Hello World Foo Bar Baz", 0.6, 12)
        assert len(result) >= 2
        # 各行に空白以外の単語が含まれる
        for line in result:
            assert line != ""

    def test_single_line_fits(self) -> None:
        # 十分に広ければ 1 行で収まる
        result = wrap_text("Hello", 10.0, 12)
        assert result == ["Hello"]


class TestEstimateTextHeight:
    """estimate_text_height の算出."""

    def test_empty_string_zero_height(self) -> None:
        assert estimate_text_height("", 5.0, 12) == 0.0

    def test_single_line_height(self) -> None:
        # "Hello" は十分広い max_width では 1 行 → 12 * 0.0139 * 1.2
        height = estimate_text_height("Hello", 10.0, 12)
        expected = 1 * 12 * 0.0139 * 1.2
        assert height == pytest.approx(expected, abs=0.001)

    def test_multiline_height_scales(self) -> None:
        # 明示改行で 3 行 → 高さは 1 行の 3 倍
        single = estimate_text_height("a", 10.0, 12)
        triple = estimate_text_height("a\nb\nc", 10.0, 12)
        assert triple == pytest.approx(single * 3, abs=0.001)

    def test_line_height_factor_applied(self) -> None:
        h_default = estimate_text_height("Hello", 10.0, 12)
        h_double = estimate_text_height("Hello", 10.0, 12, line_height_factor=2.4)
        # line_height_factor 2.4 は 1.2 の 2 倍
        assert h_double == pytest.approx(h_default * 2, abs=0.001)
