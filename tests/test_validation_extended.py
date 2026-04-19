"""validation 拡張 (Issue #5) のテスト.

text_overflow / unreadable_text / divider_collision / inconsistent_gap を
それぞれ独立に検証し、最後に check_deck_extended の集計動作を確認する。
"""

from __future__ import annotations

import pytest
from pptx import Presentation
from pptx.util import Inches, Pt

from pptx_mcp_server.engine.validation import (
    ValidationFinding,
    check_deck_extended,
    check_divider_collision,
    check_inconsistent_gaps,
    check_text_overflow,
    check_unreadable_text,
)


# ---------------------------------------------------------------------------
# ヘルパ
# ---------------------------------------------------------------------------


def _add_textbox(
    slide,
    left_in: float,
    top_in: float,
    width_in: float,
    height_in: float,
    text: str,
    font_pt: float | None = None,
    name: str | None = None,
):
    """指定位置に text box を追加し、必要ならフォントサイズと name を設定する."""
    tb = slide.shapes.add_textbox(
        Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in)
    )
    tf = tb.text_frame
    tf.word_wrap = True
    tf.text = text
    if font_pt is not None:
        for para in tf.paragraphs:
            for run in para.runs:
                run.font.size = Pt(font_pt)
    if name is not None:
        tb.name = name
    return tb


def _add_rect(
    slide,
    left_in: float,
    top_in: float,
    width_in: float,
    height_in: float,
    name: str | None = None,
):
    """rectangle auto shape を追加する."""
    from pptx.enum.shapes import MSO_SHAPE

    sh = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in),
    )
    if name is not None:
        sh.name = name
    return sh


# ---------------------------------------------------------------------------
# text_overflow
# ---------------------------------------------------------------------------


class TestTextOverflow:
    """check_text_overflow の挙動を検証する."""

    def test_short_text_in_large_box_no_finding(self, one_slide_prs: Presentation) -> None:
        slide = one_slide_prs.slides[0]
        _add_textbox(slide, 1.0, 1.0, 6.0, 3.0, "短いテキスト", font_pt=14)
        findings = check_text_overflow(one_slide_prs)
        assert findings == []

    def test_long_japanese_in_small_box_overflows(self, one_slide_prs: Presentation) -> None:
        slide = one_slide_prs.slides[0]
        long_jp = (
            "これは非常に長い日本語のテキストであり、"
            "指定された小さなテキストボックスには到底収まらないほどの分量を持っている。"
            "さらに追加の説明文も含めて計測する。"
        )
        _add_textbox(slide, 1.0, 1.0, 2.0, 0.5, long_jp, font_pt=18, name="SmallBox")
        findings = check_text_overflow(one_slide_prs)
        assert len(findings) >= 1
        f = findings[0]
        assert f.category == "text_overflow"
        assert f.severity == "error"
        assert f.shape_name == "SmallBox"
        assert f.suggested_fix != ""


# ---------------------------------------------------------------------------
# unreadable_text
# ---------------------------------------------------------------------------


class TestUnreadableText:
    """check_unreadable_text の挙動を検証する."""

    def test_6pt_text_warns(self, one_slide_prs: Presentation) -> None:
        slide = one_slide_prs.slides[0]
        _add_textbox(slide, 1.0, 1.0, 3.0, 0.5, "tiny", font_pt=6, name="TinyBox")
        findings = check_unreadable_text(one_slide_prs)
        assert len(findings) == 1
        f = findings[0]
        assert f.category == "unreadable_text"
        assert f.severity == "warning"
        assert f.shape_name == "TinyBox"

    def test_9pt_text_no_finding(self, one_slide_prs: Presentation) -> None:
        slide = one_slide_prs.slides[0]
        _add_textbox(slide, 1.0, 1.0, 3.0, 0.5, "ok", font_pt=9, name="OkBox")
        findings = check_unreadable_text(one_slide_prs)
        assert findings == []

    def test_footer_named_shape_whitelisted(self, one_slide_prs: Presentation) -> None:
        slide = one_slide_prs.slides[0]
        _add_textbox(
            slide, 1.0, 7.0, 3.0, 0.3, "Source: example",
            font_pt=7, name="Footer_Source",
        )
        findings = check_unreadable_text(one_slide_prs)
        assert findings == []


# ---------------------------------------------------------------------------
# divider_collision
# ---------------------------------------------------------------------------


class TestDividerCollision:
    """check_divider_collision の挙動を検証する."""

    def test_title_wraps_into_divider(self, one_slide_prs: Presentation) -> None:
        slide = one_slide_prs.slides[0]
        # タイトル: width 3.0" に 14pt 長タイトルを入れ 2 行以上に折り返す想定
        long_title = (
            "これはスライドのアクションタイトルで必ず折り返される日本語テキスト"
        )
        _add_textbox(slide, 0.9, 0.45, 3.0, 0.5, long_title,
                     font_pt=14, name="Title")
        # divider: y=0.95" に薄い水平線 (高さ 0.01")
        _add_rect(slide, 0.9, 0.95, 11.5, 0.02, name="Divider")

        findings = check_divider_collision(one_slide_prs)
        assert len(findings) >= 1
        f = findings[0]
        assert f.category == "divider_collision"
        assert f.severity == "error"
        assert f.shape_name == "Title"

    def test_short_title_no_collision(self, one_slide_prs: Presentation) -> None:
        slide = one_slide_prs.slides[0]
        # 十分に収まる短いタイトル (12pt 1 行 ≈ 0.20" → top 0.65 + 0.20 = 0.85 < 0.93)
        _add_textbox(slide, 0.9, 0.65, 11.5, 0.25, "短題",
                     font_pt=12, name="Title")
        _add_rect(slide, 0.9, 0.95, 11.5, 0.02, name="Divider")
        findings = check_divider_collision(one_slide_prs)
        assert findings == []


# ---------------------------------------------------------------------------
# inconsistent_gap
# ---------------------------------------------------------------------------


class TestInconsistentGap:
    """check_inconsistent_gaps の挙動を検証する."""

    def test_uneven_row_gaps(self, one_slide_prs: Presentation) -> None:
        slide = one_slide_prs.slides[0]
        # 同じ top (=2.0") に 3 shape、幅 1.0"、x gap を [0.15, 0.30] に設定
        _add_rect(slide, 1.00, 2.0, 1.0, 1.0, name="A")
        _add_rect(slide, 2.15, 2.0, 1.0, 1.0, name="B")  # gap A→B = 0.15
        _add_rect(slide, 3.45, 2.0, 1.0, 1.0, name="C")  # gap B→C = 0.30
        # もう 1 shape で 3 つ目の gap を確保: gaps [0.15, 0.30] 差 = 0.15 > 0.05
        findings = check_inconsistent_gaps(one_slide_prs)
        assert len(findings) >= 1
        assert any(f.category == "inconsistent_gap" for f in findings)
        assert all(f.severity == "info" for f in findings)

    def test_consistent_row_gaps_no_finding(self, one_slide_prs: Presentation) -> None:
        slide = one_slide_prs.slides[0]
        # gaps 0.20, 0.20, 0.20 → 差 0
        _add_rect(slide, 1.00, 2.0, 1.0, 1.0, name="A")
        _add_rect(slide, 2.20, 2.0, 1.0, 1.0, name="B")
        _add_rect(slide, 3.40, 2.0, 1.0, 1.0, name="C")
        _add_rect(slide, 4.60, 2.0, 1.0, 1.0, name="D")
        findings = check_inconsistent_gaps(one_slide_prs)
        # 行方向は一定、列方向も揃っていないので誤検出しない
        row_findings = [f for f in findings if "row" in f.message]
        assert row_findings == []


# ---------------------------------------------------------------------------
# check_deck_extended 集計
# ---------------------------------------------------------------------------


class TestCheckDeckExtended:
    """check_deck_extended が summary を正しく集計することを確認する."""

    def test_aggregated_summary_counts(self, one_slide_prs: Presentation) -> None:
        slide = one_slide_prs.slides[0]
        # 1) text_overflow (error)
        long_jp = (
            "非常に長い日本語のテキストでありこの幅の箱に収まらない量を持つ。"
            "追加文章もあるので必ずオーバーフローする。"
        )
        _add_textbox(slide, 1.0, 1.0, 2.0, 0.5, long_jp,
                     font_pt=18, name="Overflow")
        # 2) unreadable (warning)
        _add_textbox(slide, 1.0, 5.0, 3.0, 0.3, "tiny",
                     font_pt=6, name="TinyBox")

        result = check_deck_extended(one_slide_prs)

        assert "slides" in result
        assert "summary" in result
        assert len(result["slides"]) == 1
        slide0 = result["slides"][0]
        # 後方互換キーの保持を確認
        assert "overlaps" in slide0
        assert "out_of_bounds" in slide0
        # 新キー
        assert "text_overflow" in slide0
        assert "unreadable_text" in slide0
        assert "divider_collision" in slide0
        assert "inconsistent_gaps" in slide0

        assert len(slide0["text_overflow"]) >= 1
        assert len(slide0["unreadable_text"]) >= 1

        summary = result["summary"]
        assert summary["errors"] >= 1
        assert summary["warnings"] >= 1
        assert summary["infos"] >= 0


# ---------------------------------------------------------------------------
# ValidationFinding dataclass の簡易確認
# ---------------------------------------------------------------------------


def test_validation_finding_to_dict_shape() -> None:
    f = ValidationFinding(
        severity="error",
        slide_index=0,
        shape_name="X",
        category="text_overflow",
        message="m",
        suggested_fix="fix",
    )
    d = f.to_dict()
    assert d["severity"] == "error"
    assert d["slide_index"] == 0
    assert d["shape_name"] == "X"
    assert d["category"] == "text_overflow"
    assert d["message"] == "m"
    assert d["suggested_fix"] == "fix"
