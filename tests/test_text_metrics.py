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
    is_half_width_kana,
    is_zero_width,
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

    def test_cjk_ext_a(self) -> None:
        # 人名漢字など CJK Ext A (U+3400–U+4DBF)
        assert is_cjk("㐂") is True

    def test_cjk_sip_plane2(self) -> None:
        # SIP (U+20000–U+2FFFF) の 1 コードポイント
        assert is_cjk(chr(0x2000B)) is True  # 𠀋

    def test_cjk_compat_ideographs(self) -> None:
        # CJK Compatibility Ideographs (U+F900–U+FAFF)
        assert is_cjk(chr(0xF900)) is True

    def test_cjk_radicals(self) -> None:
        # CJK Radicals Supplement / Kangxi Radicals
        assert is_cjk(chr(0x2E80)) is True
        assert is_cjk(chr(0x2F00)) is True

    def test_half_width_kana_is_cjk(self) -> None:
        # 半角カタカナも CJK スクリプト扱い (幅は ASCII 相当だが判定は True)
        assert is_cjk("ｶ") is True

    def test_zwsp_is_not_cjk(self) -> None:
        assert is_cjk("\u200B") is False


class TestIsHalfWidthKana:
    """is_half_width_kana の判定."""

    def test_half_width_kana(self) -> None:
        assert is_half_width_kana("ｶ") is True
        assert is_half_width_kana("ﾀ") is True
        assert is_half_width_kana("ﾅ") is True

    def test_full_width_kana_is_not_half(self) -> None:
        assert is_half_width_kana("カ") is False
        assert is_half_width_kana("ア") is False

    def test_ascii_is_not_half_width_kana(self) -> None:
        assert is_half_width_kana("a") is False
        assert is_half_width_kana("1") is False

    def test_empty_string(self) -> None:
        assert is_half_width_kana("") is False


class TestIsZeroWidth:
    """is_zero_width の判定."""

    def test_zwsp(self) -> None:
        assert is_zero_width("\u200B") is True

    def test_zwnj_zwj(self) -> None:
        assert is_zero_width("\u200C") is True
        assert is_zero_width("\u200D") is True

    def test_word_joiner(self) -> None:
        assert is_zero_width("\u2060") is True

    def test_bom_zwnbsp(self) -> None:
        assert is_zero_width("\uFEFF") is True

    def test_variation_selector(self) -> None:
        assert is_zero_width("\uFE00") is True
        assert is_zero_width("\uFE0F") is True

    def test_combining_mark(self) -> None:
        # 結合濁点 (U+3099) は unicodedata.combining != 0 のため 0 幅扱い。
        # なお U+309B は spacing voiced sound mark で combining ではない。
        assert is_zero_width("\u3099") is True
        # 他の結合マーク例: COMBINING ACUTE ACCENT
        assert is_zero_width("\u0301") is True

    def test_ascii_is_not_zero_width(self) -> None:
        assert is_zero_width("a") is False

    def test_cjk_is_not_zero_width(self) -> None:
        assert is_zero_width("あ") is False

    def test_empty_string(self) -> None:
        assert is_zero_width("") is False


class TestEstimateTextWidth:
    """estimate_text_width の数値検証."""

    def test_japanese_bold(self) -> None:
        # "プレミアムモルツ" = 8 文字 × 9pt × 0.0139 × 1.05 ≈ 1.050 inches
        width = estimate_text_width("プレミアムモルツ", 9, bold=True)
        expected = 8 * 9 * 0.0139 * 1.05
        assert width == pytest.approx(expected, abs=0.01)
        assert width == pytest.approx(1.050, abs=0.01)

    def test_ascii_hello_world(self) -> None:
        # "Hello World" = H(wide) e(norm) l(narr) l(narr) o(norm) ' '(narr)
        #                 W(wide) o(norm) r(narr) l(narr) d(norm)
        # = 2 wide + 4 normal + 5 narrow @ 12pt
        width = estimate_text_width("Hello World", 12)
        expected = 12 * (2 * 0.01150 + 4 * 0.00765 + 5 * 0.00335)
        assert width == pytest.approx(expected, abs=0.01)

    def test_mixed_ai_seikatsusha(self) -> None:
        # "AI生活者" = A(wide) + I(narrow) + 3 CJK @ 10pt
        width = estimate_text_width("AI生活者", 10)
        expected = 10 * 0.01150 + 10 * 0.00335 + 3 * 10 * 0.0139
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
        # "a" は ASCII normal バケット。
        assert estimate_char_width("a", 10) == pytest.approx(10 * 0.00765, abs=0.0001)

    def test_cjk_char(self) -> None:
        assert estimate_char_width("あ", 10) == pytest.approx(10 * 0.0139, abs=0.0001)

    def test_digit_same_as_ascii(self) -> None:
        assert estimate_char_width("0", 10) == pytest.approx(estimate_char_width("a", 10))

    def test_empty_char_zero(self) -> None:
        assert estimate_char_width("", 10) == 0.0

    def test_half_width_kana_is_ascii_width(self) -> None:
        # 半角カタカナは CJK 全角ではなく ASCII 相当の幅を使う。
        assert estimate_char_width("ｶ", 10) == pytest.approx(
            estimate_char_width("a", 10), abs=0.0001
        )

    def test_full_width_kana_is_em_width(self) -> None:
        # 全角カタカナは 1em 幅 (0.0139"/pt)
        assert estimate_char_width("カ", 10) == pytest.approx(10 * 0.0139, abs=0.0001)

    def test_half_width_kana_half_of_full(self) -> None:
        # 半角は全角のおおよそ半分 (ASCII 0.0083 vs CJK 0.0139)
        assert estimate_char_width("ｶ", 10) < estimate_char_width("カ", 10)

    def test_zero_width_space(self) -> None:
        assert estimate_char_width("\u200B", 10) == 0.0

    def test_variation_selector(self) -> None:
        assert estimate_char_width("\uFE0F", 10) == 0.0

    def test_combining_mark_zero_width(self) -> None:
        # U+3099 は結合濁点 → 0 幅
        assert estimate_char_width("\u3099", 10) == 0.0

    def test_cjk_ext_a_is_em_width(self) -> None:
        # 人名漢字 U+3400–U+4DBF も 1em 扱い (従来は ASCII 扱いで過小評価)
        assert estimate_char_width("㐂", 10) == pytest.approx(10 * 0.0139, abs=0.0001)

    def test_sip_is_em_width(self) -> None:
        # SIP 面のコードポイントも 1em 扱い
        assert estimate_char_width(chr(0x2000B), 10) == pytest.approx(
            10 * 0.0139, abs=0.0001
        )


class TestEstimateTextWidthExtended:
    """is_cjk 拡張後の幅推定検証."""

    def test_combining_mark_adds_no_width(self) -> None:
        # "か" + 結合濁点 の幅は "か" 単独と同じ
        with_combining = estimate_text_width("か\u3099", 10)
        without = estimate_text_width("か", 10)
        assert with_combining == pytest.approx(without, abs=0.0001)

    def test_cjk_ext_a_triple(self) -> None:
        # 実プロダクションの日本人氏名ケース: "㐂㐂㐂" は 3 × CJK 幅。
        # 旧実装では ASCII 扱いで 3 × 0.0083 と過小評価されていた。
        width = estimate_text_width("㐂㐂㐂", 10)
        expected = 3 * 10 * 0.0139
        assert width == pytest.approx(expected, abs=0.001)

    def test_half_width_kana_string(self) -> None:
        # ｶﾀｶﾅ = 4 × ASCII normal 幅 (旧実装では 4 × CJK 幅で過大評価されていた)
        width = estimate_text_width("ｶﾀｶﾅ", 10)
        expected = 4 * 10 * 0.00765
        assert width == pytest.approx(expected, abs=0.001)

    def test_zwsp_does_not_add_width(self) -> None:
        # ZWSP を含む文字列は含まないものと同じ幅
        assert estimate_text_width("ab\u200Bcd", 10) == pytest.approx(
            estimate_text_width("abcd", 10), abs=0.0001
        )


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

    def test_long_word_char_broken_across_lines(self) -> None:
        # max_width を超える ASCII ワードは文字単位で強制分割される (issue #31)。
        # 通常長の単語を不必要に分割しないが、過大トークンは複数行に展開される。
        long_word = "supercalifragilisticexpialidocious"
        result = wrap_text(long_word, 0.5, 12)
        # 文字単位折り返しなので 1 行には収まらない
        assert len(result) >= 2
        # 元の文字列を復元できる (情報欠損なし)
        assert "".join(result) == long_word
        # どの行も max_width 内に収まる
        for line in result:
            assert estimate_text_width(line, 12) <= 0.5 + 1e-9

    def test_long_url_breaks_into_multiple_lines(self) -> None:
        # URL のようにスペースを含まない長い ASCII 文字列も折り返される (issue #31)。
        url = "https://example.com/very/long/path/that/wont/fit/on/one/line"
        result = wrap_text(url, 1.0, 10)
        assert len(result) >= 2
        assert "".join(result) == url

    def test_long_token_neighbors_preserved(self) -> None:
        # 長大トークンは文字分割されるが、隣接する短い単語は分割されない。
        result = wrap_text("foo bar baz_verylongwordthatexceedswidth morestuff", 1.0, 10)
        # 短い単語は原形保持されている
        joined = " ".join(result)
        assert "foo" in joined
        assert "bar" in joined
        assert "morestuff" in joined
        # 長大トークン自体は 1 行には収まっていない (複数行へ展開)
        long_tok = "baz_verylongwordthatexceedswidth"
        assert not any(long_tok in line for line in result)

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

    def test_short_words_not_split(self) -> None:
        # リグレッション: 普通の単語は decent な max_width で分割されない。
        result = wrap_text("Hello World", 5.0, 12)
        assert result == ["Hello World"]

    def test_cjk_char_break_regression(self) -> None:
        # リグレッション: 長い CJK 連続列も従来通り文字単位で折り返される。
        # CJK は _tokenize で 1 文字 1 トークン化される既存仕様により機能する。
        cjk = "あ" * 50
        result = wrap_text(cjk, 1.0, 10)
        assert len(result) >= 2
        assert "".join(result) == cjk

    def test_token_exactly_max_width_single_line(self) -> None:
        # トークン幅がちょうど max_width のときは 1 行に収まり、強制分割されない。
        token = "abcdef"  # 6 chars * 12pt * 0.0083 = 0.5976 inches
        width = estimate_text_width(token, 12)
        result = wrap_text(token, width, 12)
        assert result == [token]


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

    def test_long_ascii_token_reports_multi_line_height(self) -> None:
        # Issue #31: 長大な ASCII トークンも複数行に展開され、高さが正しく推定される。
        # 旧バグでは 1 行扱いとなり ~0.18" が返っていた。
        height = estimate_text_height("x" * 1000, 2.0, 10)
        # max_width=2.0" 内での char-break なので最低 2" 以上 (実際は数 inches)
        assert height >= 2.0
