"""可変高カード行プリミティブ.

横並びの N 枚カードをレイアウトするヘルパを提供する。各カードは content 量に
応じて自動で高さを決定し、``height_mode`` により「個別 content 高」「全カード
最大高に揃える」「行高いっぱいに埋める」を切り替える。カードの描画要素は
背景矩形、任意の左アクセントバー、上から label → title → body の順のテキスト
ブロックで構成される。

# 設計メモ
- 各カードの幅 ``width`` は ``(row_width - gap*(n-1)) / n`` で一意に決まる。
  計算結果は呼び出し元で再利用できるよう ``CardSpec.width`` に書き戻す
  (入力 dataclass を破壊的に変更する)。
- 高さ推定は ``estimate_text_height`` を使う。label は単一行想定のため
  ``label_size_pt * 0.0139`` をそのまま採用する。
- 各最終高は ``[min_card_height, max_height]`` で clamp する。
- 単一カード (n == 1) のとき ``gap`` は無視する。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

from .shapes import _add_shape, add_auto_fit_textbox
from .text_metrics import estimate_text_height

CardHeightMode = Literal["content", "max", "fill"]


# 1em inches 換算係数 (text_metrics の _CJK_WIDTH_PER_PT に相当)
_EM_PER_PT: float = 0.0139

# ブロック間の内部余白 (inches)
_INTRA_GAP: float = 0.05

# 左アクセントバーの幅 (inches)
_ACCENT_BAR_W: float = 0.08


@dataclass
class CardSpec:
    """1 枚のカードの内容・スタイル定義.

    Attributes:
        title: タイトル文字列。空文字列でブロックを省略する。
        body: 本文文字列。空文字列でブロックを省略する。
        label: title の上に小さく表示する任意のラベル文字列。単一行想定。
        accent_color: 左アクセントバーの色 (6 桁 hex、``#`` なし)。空文字列で
            アクセントバーを描画しない。
        fill_color: カード背景色 (6 桁 hex)。
        title_size_pt: タイトル font size (pt)。
        body_size_pt: 本文 font size (pt)。
        title_color: タイトル文字色 (6 桁 hex)。
        body_color: 本文文字色 (6 桁 hex)。
        label_size_pt: ラベル font size (pt)。
        label_color: ラベル文字色 (6 桁 hex)。
        padding: カード内側 padding (inches、上下左右共通)。
        width: 行レイアウトが計算したカード幅 (inches)。呼び出し側で事前に
            設定する必要はない。``add_responsive_card_row`` が書き込む。
    """

    title: str = ""
    body: str = ""
    label: str = ""
    accent_color: str = ""
    fill_color: str = "F5F7FA"
    title_size_pt: float = 14
    body_size_pt: float = 10
    title_color: str = "051C2C"
    body_color: str = "333333"
    label_size_pt: float = 9
    label_color: str = "666666"
    padding: float = 0.2
    width: float = 0.0


def _estimate_block_heights(card: CardSpec) -> tuple[float, float, float, int]:
    """カードの label / title / body それぞれの推定高 (inches) と非空ブロック数を返す.

    label は単一行想定のため ``label_size_pt * 0.0139`` をそのまま使う。
    title / body は ``estimate_text_height`` により wrap 後の総高を算出する。
    """
    inner_w = max(card.width - 2 * card.padding, 0.01)

    label_h = card.label_size_pt * _EM_PER_PT if card.label else 0.0
    title_h = (
        estimate_text_height(card.title, inner_w, card.title_size_pt)
        if card.title
        else 0.0
    )
    body_h = (
        estimate_text_height(card.body, inner_w, card.body_size_pt)
        if card.body
        else 0.0
    )

    n_blocks = sum(1 for h in (label_h, title_h, body_h) if h > 0)
    return label_h, title_h, body_h, n_blocks


def _content_height(card: CardSpec) -> float:
    """カードの content 高 (padding 込み、intra_gap 含む) を返す."""
    label_h, title_h, body_h, n_blocks = _estimate_block_heights(card)
    if n_blocks == 0:
        return card.padding * 2
    return (
        card.padding * 2
        + label_h
        + title_h
        + body_h
        + _INTRA_GAP * max(n_blocks - 1, 0)
    )


def _render_card(
    slide,
    card: CardSpec,
    left: float,
    top: float,
    w: float,
    h: float,
    *,
    mode: CardHeightMode,
) -> None:
    """1 枚のカードを実際に描画する内部ヘルパ.

    描画順:
        1. 背景矩形 (``fill_color``)
        2. 任意の左アクセントバー (``accent_color`` が非空のとき)
        3. label → title → body のテキスト 3 ブロック

    ``mode="fill"`` かつ content が h より短い場合、label は先頭固定のまま
    title + body のグループを垂直方向に中央揃えで配置する。
    """
    # 1) 背景矩形
    _add_shape(
        slide,
        shape_type="rectangle",
        left=left,
        top=top,
        width=w,
        height=h,
        fill_color=card.fill_color,
        no_line=True,
    )

    # 2) 任意の左アクセントバー
    if card.accent_color:
        _add_shape(
            slide,
            shape_type="rectangle",
            left=left,
            top=top,
            width=_ACCENT_BAR_W,
            height=h,
            fill_color=card.accent_color,
            no_line=True,
        )

    # 3) テキストブロック
    inner_left = left + card.padding
    inner_w = max(w - 2 * card.padding, 0.01)

    label_h, title_h, body_h, n_blocks = _estimate_block_heights(card)

    # label は常に上端に固定で配置する。
    cursor = top + card.padding
    if card.label:
        add_auto_fit_textbox(
            slide,
            card.label,
            left=inner_left,
            top=cursor,
            width=inner_w,
            height=max(label_h, 0.1),
            font_size_pt=card.label_size_pt,
            min_size_pt=max(card.label_size_pt - 2, 6),
            color_hex=card.label_color,
            align="left",
            vertical_anchor="top",
            truncate_with_ellipsis=True,
        )
        cursor += label_h + _INTRA_GAP

    # title + body の配置。fill モードで余白が出る場合のみ中央寄せにする。
    tb_blocks = [blk for blk in ((card.title, title_h), (card.body, body_h)) if blk[1] > 0]
    if not tb_blocks:
        return

    # title + body を合わせた所要高 (intra_gap 込み)
    tb_total = sum(h for _, h in tb_blocks) + _INTRA_GAP * max(len(tb_blocks) - 1, 0)

    # fill モードで全 content < h の場合のみ中央寄せ。それ以外は cursor をそのまま使う。
    label_base = top + card.padding + (label_h + _INTRA_GAP if card.label else 0.0)
    available_from_label = (top + h - card.padding) - label_base
    if mode == "fill" and tb_total < available_from_label:
        # title/body グループを label 直下〜底辺までの領域で中央に寄せる
        cursor = label_base + (available_from_label - tb_total) / 2

    # 底辺を超えないよう残り領域も考慮して各ブロックの height を割り当てる。
    # body は残余を埋める方針だが、content / max モードでは推定値どおりに積む。
    if card.title:
        # title は推定分だけ割り当てる
        t_h = title_h
        add_auto_fit_textbox(
            slide,
            card.title,
            left=inner_left,
            top=cursor,
            width=inner_w,
            height=max(t_h, 0.1),
            font_size_pt=card.title_size_pt,
            min_size_pt=max(card.title_size_pt - 4, 7),
            bold=True,
            color_hex=card.title_color,
            align="left",
            vertical_anchor="top",
            truncate_with_ellipsis=True,
        )
        cursor += t_h + (_INTRA_GAP if card.body else 0.0)

    if card.body:
        # body は残りの垂直領域いっぱい使う。ただし content/max モードで短い
        # content の場合は推定 body_h をそのまま使えば十分。fill モードの
        # 中央寄せ時は上で cursor を調整済みのため、body_h を使う。
        if mode == "fill":
            b_h = body_h
        else:
            # body は残余 (padding 下辺まで) を埋める
            b_h = max((top + h - card.padding) - cursor, body_h)
        add_auto_fit_textbox(
            slide,
            card.body,
            left=inner_left,
            top=cursor,
            width=inner_w,
            height=max(b_h, 0.1),
            font_size_pt=card.body_size_pt,
            min_size_pt=max(card.body_size_pt - 3, 6),
            color_hex=card.body_color,
            align="left",
            vertical_anchor="top",
            truncate_with_ellipsis=True,
        )


def add_responsive_card_row(
    slide,
    cards: list[CardSpec],
    *,
    left: float,
    top: float,
    width: float,
    max_height: float,
    gap: float = 0.2,
    height_mode: CardHeightMode = "max",
    min_card_height: float = 1.0,
) -> float:
    """N 枚のカードを横並びで配置する.

    各カードの幅は ``(width - gap * (n - 1)) / n`` で均等分割する
    (n == 1 のとき ``gap`` は無視される)。計算した幅は入力 ``cards`` の
    各 ``CardSpec.width`` に書き戻される (dataclass を破壊的に変更する点に注意)。

    Args:
        slide: 対象スライドオブジェクト。
        cards: ``CardSpec`` のリスト。空リストなら何も描画せず 0 を返す。
        left, top: 行の左上座標 (inches)。
        width: 行全体の幅 (inches)。
        max_height: 各カードが取り得る最大高 (inches)。``height_mode`` により
            上限として、または実際の採用高として機能する。
        gap: カード間の水平ギャップ (inches)。
        height_mode:
            - ``"content"``: 各カードが自身の content 高をそのまま使う (高さは
              カードごとに異なり得る)。``max_height`` は上限として機能する。
            - ``"max"``: 各カードの content 高を測定し、その最大値を全カードに
              適用する (底辺が揃う)。
            - ``"fill"``: ``max_height`` いっぱいを使う。content が短いカードは
              title/body グループを中央寄せする (label はトップ固定)。
        min_card_height: 全カードの下限高 (inches)。

    Returns:
        行が実際に消費した縦サイズ (inches)。height_mode が ``"content"`` の
        場合は個別カード高の最大値、それ以外は共通カード高と一致する。
    """
    n = len(cards)
    if n == 0:
        return 0.0

    # 各カードの幅を計算して CardSpec.width に書き戻す
    effective_gap = gap if n > 1 else 0.0
    card_w = (width - effective_gap * (n - 1)) / n
    for card in cards:
        card.width = card_w

    # 各カードの content 高を事前測定する
    content_heights = [_content_height(card) for card in cards]

    # height_mode に応じて各カードの最終高を決定する
    final_heights: list[float] = []
    if height_mode == "content":
        for ch in content_heights:
            h = max(min(ch, max_height), min_card_height)
            final_heights.append(h)
    elif height_mode == "max":
        max_ch = max(content_heights)
        common_h = max(min(max_ch, max_height), min_card_height)
        final_heights = [common_h] * n
    elif height_mode == "fill":
        common_h = max(max_height, min_card_height)
        final_heights = [common_h] * n
    else:  # 型的にはここには来ないが、防御的分岐
        final_heights = [min(max(ch, min_card_height), max_height) for ch in content_heights]

    # 各カードを描画
    x = left
    for card, h in zip(cards, final_heights):
        _render_card(slide, card, x, top, card.width, h, mode=height_mode)
        x += card.width + effective_gap

    return max(final_heights)
