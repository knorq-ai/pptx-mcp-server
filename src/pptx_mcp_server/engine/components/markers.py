"""Page marker / slide footer block components (v0.6.0, closes #135).

2 つの軽量ブロックコンポーネントを提供する:

* ``add_page_marker`` - スライド右上に 2 行 (section / page) の小さな
  テキストマーカーを描画する。IR や四半期レポート系の定型フォーマットで
  各ページ上部に "FINANCIAL SUMMARY" / "P.05 ／ FY Q3" のように表示する
  慣習的な位置専用の薄いヘルパ。
* ``add_slide_footer`` - スライド下部に左右 2 テキストのフッタを描画する。
  右側テキストが空文字の場合は右テキストボックスを作らない。

どちらも位置・サイズは固定のオフセット定数 (モジュール冒頭で定義) を
使う。「慣習的な位置」専用であり、任意位置指定は呼び出し側が
``add_auto_fit_textbox`` を直接使えば足りる。

どちらの関数も背後では ``add_auto_fit_textbox`` のみを使い、theme 解決
(``"text_secondary"`` など) はそちらに委ねる (#131 の二重解決を避ける)。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from ._util import resolve_component_color as _resolve_color

# ``engine.shapes`` imports ``components.container`` at module load time —
# if we import ``add_auto_fit_textbox`` at module top here, ``shapes.py``
# and ``components/__init__.py`` deadlock during partial init (shapes
# triggers components init → markers import → shapes re-entry). Defer the
# import to call time; it is cheap after first load.


# ---------------------------------------------------------------------------
# Fixed offsets (inches). 本モジュールは "慣習的な位置" 専用であり、可変化
# しない。呼び出し側が別座標を使いたい場合は add_auto_fit_textbox を直接
# 使うべき、というのが設計上の境界 (issue #135)。
# ---------------------------------------------------------------------------

# スライド左右端からの水平マージン。
_MARGIN: float = 0.5

# Page marker (top-right)
_PAGE_MARKER_TOP: float = 0.35           # 1 行目の top
_PAGE_MARKER_TEXTBOX_W: float = 2.5      # 各行のテキストボックス幅
_PAGE_MARKER_LINE_H: float = 0.25        # 各行のテキストボックス高さ

# Slide footer (bottom)
_FOOTER_HEIGHT: float = 0.3              # フッタテキストボックス高さ
_FOOTER_BOTTOM_OFFSET: float = 0.4       # スライド下端からの top オフセット
_FOOTER_LEFT_W: float = 4.0              # 左テキストボックス幅
_FOOTER_RIGHT_W: float = 6.0             # 右テキストボックス幅


@dataclass
class PageMarkerSpec:
    """右上ページマーカーの内容定義.

    Attributes:
        section: 1 行目。セクション名 (e.g. ``"FINANCIAL SUMMARY"``)。
            呼び出し側で uppercase にすることを想定するが強制はしない。
        page: 2 行目。ページ表記 (e.g. ``"P.05 ／ FY Q3"``)。
        font_size_pt: 両行に適用する font size (pt)。
        color: 文字色トークン or hex (``#`` なし)。``theme`` 指定時は
            ``resolve_theme_color`` 経由で解決される。既定は
            ``"text_secondary"``。
    """

    section: str
    page: str
    font_size_pt: float = 10
    color: str = "text_secondary"


@dataclass
class SlideFooterSpec:
    """スライド下部フッタの内容定義.

    Attributes:
        left_text: 左側フッタテキスト。
        right_text: 右側フッタテキスト。空文字列のときは右テキストボックス
            を描画しない (``add_slide_footer`` の ``right_shape`` は
            ``None``)。
        font_size_pt: 両側に適用する font size (pt)。
        color: 文字色トークン or hex (``#`` なし)。``theme`` 指定時は
            ``resolve_theme_color`` 経由で解決される。既定は
            ``"text_secondary"``。
    """

    left_text: str = "IR Presentation · FY Q3"
    right_text: str = ""
    font_size_pt: float = 10
    color: str = "text_secondary"


def add_page_marker(
    slide,
    spec: PageMarkerSpec,
    *,
    slide_width: float,
    slide_height: float,  # noqa: ARG001 - API consistency; not used for top placement
    theme: Optional[str] = None,
) -> dict:
    """スライド右上に 2 行のページマーカーを描画する.

    配置は固定で、右端が ``slide_width - _MARGIN`` (= 0.5" マージン)、
    1 行目の top が ``_PAGE_MARKER_TOP`` (= 0.35")、2 行目はその直下
    (``_PAGE_MARKER_LINE_H`` = 0.25" オフセット)。どちらの行も右寄せ、
    折り返しなし (``wrap=False``) で描画する。

    Args:
        slide: python-pptx の Slide オブジェクト。
        spec: :class:`PageMarkerSpec` — 内容とスタイル。
        slide_width: スライド幅 (inches)。右端座標の計算に使う。
        slide_height: スライド高さ (inches)。API 一貫性のため受け取るが
            top-right マーカーでは未使用。
        theme: テーマ名 (e.g. ``"ir"``)。``spec.color`` がトークンのとき
            解決に使う。

    Returns:
        ``{"section_shape": ..., "page_shape": ..., "bounds": {...}}``。
        ``bounds`` は 2 行全体の外接矩形 (inches)。
    """
    from ..shapes import add_auto_fit_textbox  # lazy: avoid circular import

    left = slide_width - _MARGIN - _PAGE_MARKER_TEXTBOX_W
    top1 = _PAGE_MARKER_TOP
    top2 = _PAGE_MARKER_TOP + _PAGE_MARKER_LINE_H
    color = _resolve_color(spec.color, theme)

    section_shape, _ = add_auto_fit_textbox(
        slide,
        spec.section,
        left=left,
        top=top1,
        width=_PAGE_MARKER_TEXTBOX_W,
        height=_PAGE_MARKER_LINE_H,
        font_size_pt=spec.font_size_pt,
        color_hex=color,
        align="right",
        wrap=False,
        theme=theme,
    )
    page_shape, _ = add_auto_fit_textbox(
        slide,
        spec.page,
        left=left,
        top=top2,
        width=_PAGE_MARKER_TEXTBOX_W,
        height=_PAGE_MARKER_LINE_H,
        font_size_pt=spec.font_size_pt,
        color_hex=color,
        align="right",
        wrap=False,
        theme=theme,
    )

    return {
        "section_shape": section_shape,
        "page_shape": page_shape,
        "bounds": {
            "left": left,
            "top": top1,
            "width": _PAGE_MARKER_TEXTBOX_W,
            "height": _PAGE_MARKER_LINE_H * 2,
        },
    }


def add_slide_footer(
    slide,
    spec: SlideFooterSpec,
    *,
    slide_width: float,
    slide_height: float,
    theme: Optional[str] = None,
) -> dict:
    """スライド下部に左右 2 テキストのフッタを描画する.

    配置は固定で、top は ``slide_height - _FOOTER_BOTTOM_OFFSET``、
    左側は ``left=_MARGIN`` の左寄せテキスト、右側は
    ``left=slide_width - _MARGIN - _FOOTER_RIGHT_W`` の右寄せテキスト。
    右側は ``spec.right_text`` が空文字のとき描画をスキップする (戻り値の
    ``right_shape`` が ``None``)。

    Args:
        slide: python-pptx の Slide オブジェクト。
        spec: :class:`SlideFooterSpec` — 内容とスタイル。
        slide_width: スライド幅 (inches)。
        slide_height: スライド高さ (inches)。top の計算に使う。
        theme: テーマ名。``spec.color`` がトークンのとき解決に使う。

    Returns:
        ``{"left_shape": ..., "right_shape": ... or None, "bounds": {...}}``。
        ``bounds`` は左右両端を含む外接矩形 (inches、右側未描画でも幅は
        ``slide_width - 2*_MARGIN``)。
    """
    from ..shapes import add_auto_fit_textbox  # lazy: avoid circular import

    top = slide_height - _FOOTER_BOTTOM_OFFSET
    color = _resolve_color(spec.color, theme)

    left_shape, _ = add_auto_fit_textbox(
        slide,
        spec.left_text,
        left=_MARGIN,
        top=top,
        width=_FOOTER_LEFT_W,
        height=_FOOTER_HEIGHT,
        font_size_pt=spec.font_size_pt,
        color_hex=color,
        align="left",
        wrap=False,
        theme=theme,
    )

    right_shape = None
    if spec.right_text:
        right_shape, _ = add_auto_fit_textbox(
            slide,
            spec.right_text,
            left=slide_width - _MARGIN - _FOOTER_RIGHT_W,
            top=top,
            width=_FOOTER_RIGHT_W,
            height=_FOOTER_HEIGHT,
            font_size_pt=spec.font_size_pt,
            color_hex=color,
            align="right",
            wrap=False,
            theme=theme,
        )

    return {
        "left_shape": left_shape,
        "right_shape": right_shape,
        "bounds": {
            "left": _MARGIN,
            "top": top,
            "width": slide_width - 2 * _MARGIN,
            "height": _FOOTER_HEIGHT,
        },
    }
