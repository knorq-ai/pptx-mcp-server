"""可変高カード行プリミティブのテスト.

``add_responsive_card_row`` の 3 つの ``height_mode`` (content / max / fill)
と、単一カード・最小高 clamp・アクセントバー描画・消費高返却の各挙動を
網羅する。
"""

from __future__ import annotations

import pytest

from pptx_mcp_server.engine.cards import (
    CardSpec,
    add_responsive_card_row,
    _content_height,
)


# EMU (914400 per inch) → inches 換算ユーティリティ
def _in(emu: int) -> float:
    return emu / 914400


def _shape_bbox(shape) -> tuple[float, float, float, float]:
    """shape の (left, top, width, height) を inches で返す."""
    return _in(shape.left), _in(shape.top), _in(shape.width), _in(shape.height)


# 各カードの先頭 shape (背景矩形) の bbox がそのカード全体の bbox になる。
# アクセントバーや textbox もそこから派生するため、背景矩形を頼りに検証する。
def _card_background_heights(slide, n_cards: int, has_accent: bool) -> list[float]:
    """各カードの背景矩形の高さを順に取り出す.

    ``add_responsive_card_row`` は各カードにつき最初に背景矩形 (AUTO_SHAPE) を
    追加する。その後、任意で accent bar (AUTO_SHAPE)、label/title/body
    (TEXT_BOX) を追加する。本テストでは shape 追加順・種別に基づき「各カード
    ごとに最初に現れる AUTO_SHAPE で、かつ label/title/body の TextBox より幅が
    広いもの」を拾う。簡易に、先頭から走査して AUTO_SHAPE が現れた時点の高さを
    採用し、同一 left の次の AUTO_SHAPE (accent bar) は無視する。
    """
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    heights: list[float] = []
    seen_lefts: set[float] = set()
    for s in slide.shapes:
        if s.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE:
            continue
        key = round(_in(s.left), 3)
        if key in seen_lefts:
            # 同じ left で 2 つ目の AUTO_SHAPE は accent bar のためスキップ
            continue
        seen_lefts.add(key)
        heights.append(_in(s.height))
        if len(heights) == n_cards:
            break
    return heights


def test_max_mode_identical_content(slide):
    """3 カードで同一の title/body を持ち height_mode='max' ならすべて同じ高さになる."""
    cards = [
        CardSpec(title="Strategy", body="Short body text."),
        CardSpec(title="Strategy", body="Short body text."),
        CardSpec(title="Strategy", body="Short body text."),
    ]
    consumed = add_responsive_card_row(
        slide, cards,
        left=0.5, top=1.0, width=12.0, max_height=4.0,
        gap=0.2, height_mode="max", min_card_height=0.5,
    )

    heights = _card_background_heights(slide, n_cards=3, has_accent=False)
    assert len(heights) == 3
    assert heights[0] == pytest.approx(heights[1], abs=1e-4)
    assert heights[1] == pytest.approx(heights[2], abs=1e-4)
    # 消費高 == 各カード高
    assert consumed == pytest.approx(heights[0], abs=1e-4)


def test_max_mode_different_content_lengths_align(slide):
    """3 カード + content 長さ差 + height_mode='max' で全カード同一高 (= 最大 content 高)."""
    short = CardSpec(title="T", body="Short.")
    medium = CardSpec(title="T", body="Medium body text that wraps.")
    long_body = "This is a very long body text that will wrap to many lines. " * 6
    long_card = CardSpec(title="T", body=long_body)

    cards = [short, medium, long_card]
    consumed = add_responsive_card_row(
        slide, cards,
        left=0.5, top=1.0, width=12.0, max_height=6.0,
        gap=0.2, height_mode="max",
    )

    heights = _card_background_heights(slide, n_cards=3, has_accent=False)
    # 全カード同一
    assert heights[0] == pytest.approx(heights[1], abs=1e-4)
    assert heights[1] == pytest.approx(heights[2], abs=1e-4)
    # 高さ == 最大 content 高
    max_content = max(_content_height(c) for c in cards)
    assert heights[0] == pytest.approx(max_content, abs=1e-3)
    assert consumed == pytest.approx(heights[0], abs=1e-4)


def test_content_mode_heights_differ(slide):
    """height_mode='content' では短い/長いカードで高さが異なる."""
    short = CardSpec(title="T", body="Short.")
    long_body = "Long body text that wraps across multiple lines. " * 8
    long_card = CardSpec(title="T", body=long_body)
    medium = CardSpec(title="T", body="Medium sample body.")

    cards = [short, medium, long_card]
    add_responsive_card_row(
        slide, cards,
        left=0.5, top=1.0, width=12.0, max_height=6.0,
        gap=0.2, height_mode="content", min_card_height=0.2,
    )

    heights = _card_background_heights(slide, n_cards=3, has_accent=False)
    # 短いカードと長いカードの高さは明確に異なる
    assert heights[0] != pytest.approx(heights[2], abs=0.1)
    assert heights[2] > heights[0]


def test_fill_mode_uses_max_height(slide):
    """height_mode='fill' で max_height=5" ならすべてのカード高が 5" になる."""
    cards = [
        CardSpec(title="A", body="Short body."),
        CardSpec(title="B", body="Short body."),
        CardSpec(title="C", body="Short body."),
    ]
    consumed = add_responsive_card_row(
        slide, cards,
        left=0.5, top=0.5, width=12.0, max_height=5.0,
        gap=0.2, height_mode="fill",
    )

    heights = _card_background_heights(slide, n_cards=3, has_accent=False)
    for h in heights:
        assert h == pytest.approx(5.0, abs=1e-4)
    assert consumed == pytest.approx(5.0, abs=1e-4)


def test_min_card_height_clamp(slide):
    """min_card_height=2 なら content が ~1" でも全カード 2" に clamp される."""
    # padding+title+body だけの軽いカード (~1" 程度)
    cards = [
        CardSpec(title="X", body="One line."),
        CardSpec(title="Y", body="One line."),
    ]
    add_responsive_card_row(
        slide, cards,
        left=0.5, top=0.5, width=10.0, max_height=6.0,
        gap=0.2, height_mode="max", min_card_height=2.0,
    )

    heights = _card_background_heights(slide, n_cards=2, has_accent=False)
    for h in heights:
        assert h == pytest.approx(2.0, abs=1e-4)


def test_single_card_fills_row_width(slide):
    """単一カード (n == 1): gap を無視し行幅すべてをそのカードが占める."""
    cards = [CardSpec(title="Only", body="Alone in the row.")]
    add_responsive_card_row(
        slide, cards,
        left=1.0, top=1.0, width=8.0, max_height=4.0,
        gap=0.5, height_mode="max",
    )

    # 単一カードの width 書き戻し
    assert cards[0].width == pytest.approx(8.0, abs=1e-4)

    # 先頭 shape = 背景矩形。width が 8.0 に一致
    first_shape = list(slide.shapes)[0]
    _, _, w, _ = _shape_bbox(first_shape)
    assert w == pytest.approx(8.0, abs=1e-3)


def test_accent_bar_drawn_with_expected_bbox(slide):
    """accent_color 非空のとき、card left と同じ位置に width ≈ 0.08" のバーが生成される."""
    cards = [CardSpec(title="T", body="Body", accent_color="FF0000")]
    add_responsive_card_row(
        slide, cards,
        left=2.0, top=1.0, width=4.0, max_height=3.0,
        gap=0.2, height_mode="max",
    )

    shapes = list(slide.shapes)
    # 背景矩形 → アクセントバー → 任意のテキストの順
    bg = shapes[0]
    accent = shapes[1]
    bg_l, bg_t, bg_w, bg_h = _shape_bbox(bg)
    ac_l, ac_t, ac_w, ac_h = _shape_bbox(accent)

    assert bg_l == pytest.approx(2.0, abs=1e-3)
    assert bg_w == pytest.approx(4.0, abs=1e-3)

    # アクセントバー: 左端一致、幅 0.08"、高さは card と一致
    assert ac_l == pytest.approx(bg_l, abs=1e-3)
    assert ac_t == pytest.approx(bg_t, abs=1e-3)
    assert ac_w == pytest.approx(0.08, abs=1e-3)
    assert ac_h == pytest.approx(bg_h, abs=1e-3)


def test_consumed_height_matches_max_for_max_and_fill(slide, one_slide_prs):
    """consumed_height == max/fill モードの共通カード高."""
    # max モード
    cards_max = [CardSpec(title="T", body="Body A"), CardSpec(title="T", body="Body B")]
    consumed_max = add_responsive_card_row(
        slide, cards_max,
        left=0.5, top=0.5, width=10.0, max_height=6.0,
        gap=0.2, height_mode="max",
    )
    heights_max = _card_background_heights(slide, n_cards=2, has_accent=False)
    assert consumed_max == pytest.approx(max(heights_max), abs=1e-4)

    # fill モード (別スライド用意)
    layout = one_slide_prs.slide_layouts[6]
    one_slide_prs.slides.add_slide(layout)
    slide2 = one_slide_prs.slides[1]
    cards_fill = [CardSpec(title="T", body="A"), CardSpec(title="T", body="B")]
    consumed_fill = add_responsive_card_row(
        slide2, cards_fill,
        left=0.5, top=0.5, width=10.0, max_height=4.5,
        gap=0.2, height_mode="fill",
    )
    heights_fill = _card_background_heights(slide2, n_cards=2, has_accent=False)
    assert consumed_fill == pytest.approx(4.5, abs=1e-4)
    for h in heights_fill:
        assert h == pytest.approx(4.5, abs=1e-4)
