"""データテーブルを textbox グリッドとして描画するプリミティブ.

# 設計メモ

IR / コンサルの資料で多用される「数値が並ぶ密なテーブル」は、python-pptx の
``Table`` オブジェクトでは表現しづらい。既定の枠線、セル幅・alignment の
制御しにくさが理由である。そこで本モジュールは ``Table`` オブジェクトを使わ
ず、各セルを個別の **textbox** として並べる grid layout を提供する。

主な機能:
    - ヘッダー行 / データ行を 1 セル = 1 textbox で描画する。
    - alt-row (縞模様) の背景シェーディング。
    - 特定 1 行のハイライトシェーディング (alt-row より優先)。
    - 行間・ヘッダー直下のヘアライン罫線 (< 0.02" の細長い矩形として描画)。
    - 列幅は呼び出し側で合計幅と整合させる。合計が ``width`` と一致しない
      場合は比例スケーリングで揃える (呼び出し側の意図を殺さないよう、
      ``EngineError`` は出さず黙って調整する)。

設計判断:
    - python-pptx ``Table`` は使わない (枠線・alignment の制御が不足)。
    - 各セルは ``add_auto_fit_textbox`` を ``wrap=False`` で呼び出し、セル幅
      からはみ出す場合は末尾を省略記号で切り詰める。
    - 背景シェーディングは textbox より **先に** 追加することで、確実に
      textbox の背後 (z-order 下) に置かれる。
    - 列幅合計不一致は ``(width / sum(column.width))`` のスケーリング係数で
      各列幅に掛けて強制一致させる (#110)。

Issue: #110
"""

from __future__ import annotations

from dataclasses import dataclass, fields
from typing import Any, Literal

from .pptx_io import EngineError, ErrorCode
from .shapes import _add_shape, add_auto_fit_textbox


# ヘアライン罫線の既定厚 (inches)。< 0.02" は PowerPoint / LibreOffice の
# 描画エンジンで 1 px 相当にクランプされることが多く、意図した細線となる。
_HAIRLINE_THICKNESS: float = 0.01


@dataclass
class TableColumnSpec:
    """テーブルの 1 列の仕様.

    Attributes:
        header: ヘッダーセルに表示する文字列。
        align: データセルの水平方向 alignment。"left" | "center" | "right"。
        width: 列幅 (inches)。合計が ``add_data_table`` の ``width`` と一致
            しない場合は比例スケーリングで揃えられる。
        font_size_pt: データセルの開始 font size (pt)。auto-fit で縮小され得る。
        header_bold: ヘッダー文字を太字にするか。
        header_font_size_pt: ヘッダーの開始 font size (pt)。
        value_color: データセルの文字色 (6 桁 hex、``#`` なし)。
        header_color: ヘッダーの文字色 (6 桁 hex)。
    """

    header: str
    align: Literal["left", "center", "right"] = "left"
    width: float = 1.5
    font_size_pt: float = 10
    header_bold: bool = True
    header_font_size_pt: float = 10
    value_color: str = "333333"
    header_color: str = "6B7280"


def _scale_column_widths(
    columns: list[TableColumnSpec],
    total_width: float,
) -> list[float]:
    """列幅の合計を ``total_width`` に強制一致させる.

    呼び出し側の列幅比を保ったまま、合計が ``total_width`` となるよう
    スケーリング係数を掛ける。いずれかの列幅が非正の場合は
    ``INVALID_PARAMETER`` を投げる。
    """
    for i, c in enumerate(columns):
        if c.width <= 0:
            raise EngineError(
                ErrorCode.INVALID_PARAMETER,
                f"columns[{i}].width must be > 0, got {c.width}",
            )
    raw_sum = sum(c.width for c in columns)
    if raw_sum <= 0:
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            "columns widths sum must be > 0",
        )
    factor = total_width / raw_sum
    return [c.width * factor for c in columns]


def _stringify(value: Any) -> str:
    """セル値を表示用文字列に変換する.

    None は空文字列、その他は ``str()`` に委譲する。float の表記整形は呼び
    出し側の責務とする (本モジュールでは何も format しない)。
    """
    if value is None:
        return ""
    return str(value)


def add_data_table(
    slide,
    rows: list[list[Any]],
    columns: list[TableColumnSpec],
    *,
    left: float,
    top: float,
    width: float,
    row_height: float = 0.35,
    header_height: float = 0.4,
    alt_row_color: str | None = None,
    highlight_row_index: int | None = None,
    highlight_color: str = "F0F0F0",
    rule_color: str | None = "E0E0E0",
    rule_thickness: float = _HAIRLINE_THICKNESS,
    header_rule: bool = True,
    font_name: str = "Arial",
) -> dict:
    """行 × 列の grid として ``rows`` を描画する.

    各セルは個別の textbox として配置される (python-pptx の ``Table``
    オブジェクトは使わない)。alt-row シェーディング、ハイライト行、ヘア
    ライン罫線を組み合わせて IR / コンサル資料向けの密なテーブルを
    表現する。

    Args:
        slide: 描画対象の python-pptx スライドオブジェクト。
        rows: データ行のリスト。各行は ``len(columns)`` 個の値を持つ必要
            がある。値は ``str`` / ``int`` / ``float`` / ``None`` 等、``str()``
            で文字列化できる任意の型。
        columns: 列仕様のリスト。header, align, width 等を定義する。
        left: テーブル左端 x 座標 (inches)。
        top: テーブル上端 y 座標 (inches)。
        width: テーブル全幅 (inches)。列幅の合計はこの値に比例スケーリング
            される (呼び出し側の列幅比は保持される)。
        row_height: データ行 1 行の高さ (inches)。
        header_height: ヘッダー行の高さ (inches)。
        alt_row_color: 偶数 index (0-indexed で 1, 3, 5...) 行の背景色。None
            でシェーディング無効。
        highlight_row_index: ハイライト 1 行の index (0-indexed、データ行
            基準)。None で無効。alt-row と重なる場合はハイライトが優先。
        highlight_color: ハイライト行の背景色 (6 桁 hex)。
        rule_color: 行間 / ヘッダー下の罫線色。None で罫線無効。
        rule_thickness: 罫線の太さ (inches)。既定 0.01" はヘアライン相当。
        header_rule: ヘッダーとデータ行の間に罫線を引くか。
        font_name: セルのフォント名。

    Returns:
        dict: ``{"consumed_height": float, "header_y_bottom": float,
        "shape_count": int}``。``consumed_height`` はテーブル全体の占有高、
        ``header_y_bottom`` はヘッダー行の下端 y 座標、``shape_count`` は
        追加された shape 数。

    Raises:
        EngineError: 各行の長さが ``len(columns)`` と一致しない、列幅が
            非正、``highlight_row_index`` が範囲外などの場合
            (``INVALID_PARAMETER``)。
    """
    # ── 入力検証 ─────────────────────────────────────────
    if not columns:
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            "columns must be a non-empty list",
        )
    if width <= 0 or row_height <= 0 or header_height <= 0:
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            "width, row_height, header_height must all be > 0",
        )
    if rule_thickness < 0:
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            f"rule_thickness must be >= 0; got {rule_thickness:.3f}",
        )

    n_cols = len(columns)
    for ri, row in enumerate(rows):
        if len(row) != n_cols:
            raise EngineError(
                ErrorCode.INVALID_PARAMETER,
                (
                    f"rows[{ri}] has {len(row)} values, but columns has "
                    f"{n_cols} entries"
                ),
            )

    if highlight_row_index is not None:
        if highlight_row_index < 0 or highlight_row_index >= len(rows):
            raise EngineError(
                ErrorCode.INVALID_PARAMETER,
                (
                    f"highlight_row_index {highlight_row_index} out of range "
                    f"(rows has {len(rows)} entries)"
                ),
            )

    # ── 列幅のスケーリング ──────────────────────────────
    scaled_widths = _scale_column_widths(columns, width)

    shape_count_before = len(slide.shapes)

    # ── 1) 背景シェーディング (alt-row / highlight) ────
    # テキストセルより **先** に追加することで z-order で背後に置く。
    data_top = top + header_height
    for ri in range(len(rows)):
        row_top = data_top + ri * row_height
        fill: str | None = None
        # alt-row: 1, 3, 5, ... 行目 (0-indexed で奇数)
        if alt_row_color and ri % 2 == 1:
            fill = alt_row_color
        # highlight: 指定行は alt-row より優先
        if highlight_row_index is not None and ri == highlight_row_index:
            fill = highlight_color
        if fill is not None:
            _add_shape(
                slide,
                "rectangle",
                left,
                row_top,
                width,
                row_height,
                fill_color=fill,
                no_line=True,
            )

    # ── 2) ヘッダー罫線 / 行間罫線 ──────────────────────
    if rule_color is not None:
        # ヘッダー直下
        if header_rule:
            _add_shape(
                slide,
                "rectangle",
                left,
                top + header_height - rule_thickness / 2,
                width,
                rule_thickness,
                fill_color=rule_color,
                no_line=True,
            )
        # 各データ行の下端 (最終行を含む)
        for ri in range(len(rows)):
            rule_y = data_top + (ri + 1) * row_height - rule_thickness / 2
            _add_shape(
                slide,
                "rectangle",
                left,
                rule_y,
                width,
                rule_thickness,
                fill_color=rule_color,
                no_line=True,
            )

    # ── 3) ヘッダーセル ─────────────────────────────────
    col_x: list[float] = []
    x_cursor = left
    for w in scaled_widths:
        col_x.append(x_cursor)
        x_cursor += w

    for ci, col in enumerate(columns):
        add_auto_fit_textbox(
            slide,
            col.header,
            col_x[ci],
            top,
            scaled_widths[ci],
            header_height,
            font_name=font_name,
            font_size_pt=col.header_font_size_pt,
            min_size_pt=max(6.0, col.header_font_size_pt - 3.0),
            bold=col.header_bold,
            color_hex=col.header_color,
            align=col.align,
            vertical_anchor="middle",
            wrap=False,
            truncate_with_ellipsis=True,
        )

    # ── 4) データセル ───────────────────────────────────
    for ri, row in enumerate(rows):
        row_top = data_top + ri * row_height
        for ci, value in enumerate(row):
            col = columns[ci]
            add_auto_fit_textbox(
                slide,
                _stringify(value),
                col_x[ci],
                row_top,
                scaled_widths[ci],
                row_height,
                font_name=font_name,
                font_size_pt=col.font_size_pt,
                min_size_pt=max(6.0, col.font_size_pt - 3.0),
                bold=False,
                color_hex=col.value_color,
                align=col.align,
                vertical_anchor="middle",
                wrap=False,
                truncate_with_ellipsis=True,
            )

    # ── 返り値 ─────────────────────────────────────────
    consumed_height = header_height + row_height * len(rows)
    # 行が空でも ``header_rule`` が有効ならその厚みだけ余計に占有する
    # (レイアウト呼び出し側が次要素の y を決めるとき、罫線までを含めた
    # 下端が欲しいケースが多いため算入する)。
    if not rows and header_rule and rule_color is not None:
        consumed_height += rule_thickness

    return {
        "consumed_height": consumed_height,
        "header_y_bottom": top + header_height,
        "shape_count": len(slide.shapes) - shape_count_before,
    }


# 公開キー (server.py の strict validation で使う)
TABLE_COLUMN_SPEC_KEYS: frozenset[str] = frozenset(f.name for f in fields(TableColumnSpec))
