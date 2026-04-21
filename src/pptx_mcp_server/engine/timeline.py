"""マイルストーン・タイムラインプリミティブ.

IR デッキの成長ストーリースライドで頻出する「フェーズ帯 + 垂直ルール +
年次マイルストーン注記」の構図を、既存チャート領域の上に重ねて描画する
プリミティブである。呼び出し元は別途チャートを配置し、その ``chart_top``
および ``chart_bottom`` を本関数に渡すことで、フェーズ境界線とマイルストー
ン・ルールを同じ垂直レンジに整列させる。

# 設計メモ
- フェーズ帯は ``add_flex_container`` の ``direction="row"`` + N 個の均等
  ``grow=1`` 項目で構成する。帯内の各項目は「index_label (小) + label
  (太字) + year_range (小)」の 3 ブロックを縦積みで描画する。
- フェーズ境界ルール (``N-1`` 本) はフェーズ帯の下端から ``chart_bottom``
  まで垂直に伸びる極細矩形である。``width = 0.01"`` で描く (dashed line
  はサポート対象外; 要件は「視覚的に罫線として機能する」こと)。
- マイルストーン・ルールは、各マイルストーンの ``x_pos`` (0.0-1.0) を
  ``left + x_pos * width`` に変換し、フェーズ帯の下端から ``chart_bottom
  - 1.5"`` (primary) まで垂直に伸ばす。ルール直上にテキストボックスを
  置き、``year`` (太字・大) と ``label`` (細・小) を上下に並べる。
- マイルストーン帯の座標系はチャート領域 (``left`` ~ ``left + width``) を
  ``[0, 1]`` に射影したものである。チャート・データ点そのものの x 座標は
  caller が別途計測して渡す前提 (MVP では shape ツリーから chart data を
  取得できないため)。

# 本関数のスコープ外
- ダッシュライン描画 (視覚的にはソリッド極細罫線で代替する)。
- x_pos のフェーズ境界との衝突解消 (caller 責任。本関数はマイルストーン
  ルールを境界上に重ねても警告しない)。
- チャート・データ点の自動抽出 (x_pos は caller が 0.0-1.0 で渡す)。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from .pptx_io import EngineError, ErrorCode
from .shapes import _add_shape, add_auto_fit_textbox
from ..theme import resolve_theme_color


# マイルストーン帯のスタイル → 文字色・太字・ルール色のマッピング。
# primary はナビー基調で IR 資料の主要イベント (IPO 等) を強調する想定。
# secondary はグレー基調で補助イベント (プロダクトローンチ等) 向け。
_STYLE_PRIMARY_COLOR: str = "051C2C"   # navy
_STYLE_SECONDARY_COLOR: str = "666666"  # mid gray

# フェーズ帯内部レイアウト。index_label・label・year_range の 3 ブロック
# を縦積みにする。各ブロックの高さ比率はフェーズ帯全体 (``phase_band_height``)
# を基準に算出する。
_PHASE_INDEX_HEIGHT_RATIO: float = 0.25
_PHASE_LABEL_HEIGHT_RATIO: float = 0.45
_PHASE_YEAR_HEIGHT_RATIO: float = 0.30

# 垂直ルールの幅 (inches)。PowerPoint では 0.01" 程度が hairline 相当。
_RULE_WIDTH: float = 0.01

# マイルストーン・ルールの下端余白 (primary スタイル)。チャート底ではなく
# 少し内側で止めることで、caller が配置したチャート・データ点と被せ過ぎ
# ないようにする。
_PRIMARY_RULE_BOTTOM_MARGIN: float = 1.5
# secondary スタイルはやや短め (caller のチャート枠下側を大きく空ける)。
_SECONDARY_RULE_BOTTOM_MARGIN: float = 2.0

# マイルストーンラベル矩形の高さ (inches)。
_MILESTONE_LABEL_HEIGHT: float = 0.7
# マイルストーンラベルはルールの真上に中央揃えで配置する。幅はこの定数で
# 左右対称に広げる (画面内に収まらない極端な x_pos には caller が責任を持つ)。
_MILESTONE_LABEL_HALF_WIDTH: float = 0.9


@dataclass
class TimelinePhase:
    """タイムラインの 1 フェーズ.

    Attributes:
        label: フェーズ名 (例: ``"事業基盤の確立"``)。
        index_label: フェーズ番号などの短い ID 文字列 (例: ``"01"``)。
        year_range: 期間表記 (例: ``"2012 — 2016"``)。
    """

    label: str
    index_label: str
    year_range: str


@dataclass
class TimelineMilestone:
    """タイムライン上の 1 マイルストーン.

    Attributes:
        x_pos: チャート領域の幅を ``[0.0, 1.0]`` に正規化した水平位置。
        year: 年次表記 (例: ``"2016"``)。
        label: イベント説明 (改行可。例: ``"IPO\\n東証マザーズ上場"``)。
        style: ``"primary"`` (ナビー基調) または ``"secondary"`` (グレー基調)。
    """

    x_pos: float
    year: str
    label: str
    style: str = "primary"


_STYLE_TOKEN = {
    "primary": "primary",
    "secondary": "text_secondary",
}


def _style_color(style: str, theme_name: Optional[str] = None) -> str:
    """Resolve style → hex. When ``theme_name`` registered, pulls from theme
    (``primary`` → theme.primary, ``secondary`` → theme.text_secondary).
    Falls back to hardcoded defaults otherwise (#125).
    """
    if style not in _STYLE_TOKEN:
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            f"Unknown milestone style {style!r}; expected 'primary' or 'secondary'.",
        )
    if theme_name:
        resolved = resolve_theme_color(_STYLE_TOKEN[style], theme_name)
        if resolved:
            return resolved
    return _STYLE_PRIMARY_COLOR if style == "primary" else _STYLE_SECONDARY_COLOR


def _style_rule_bottom_margin(style: str) -> float:
    if style == "primary":
        return _PRIMARY_RULE_BOTTOM_MARGIN
    return _SECONDARY_RULE_BOTTOM_MARGIN


def _validate_inputs(
    phases: list[TimelinePhase],
    milestones: list[TimelineMilestone],
    *,
    left: float,
    top: float,
    width: float,
    phase_band_height: float,
    chart_top: float,
    chart_bottom: float,
) -> None:
    """入力パラメタをまとめて検証する.

    caller のロジックバグをサイレントにスライド外配置へ落とさないよう、
    エントリポイントで失敗シグナルを返す。

    Raises:
        EngineError: 下記のいずれかに該当するとき (``INVALID_PARAMETER``)。
            - ``chart_top >= chart_bottom``
            - ``width``・``phase_band_height`` が負
            - ``milestone.x_pos`` が ``[0.0, 1.0]`` を外れる
            - ``milestone.style`` が ``'primary'``・``'secondary'`` 以外
    """
    if width < 0:
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            f"width={width:.2f} must be >= 0.",
        )
    if phase_band_height < 0:
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            f"phase_band_height={phase_band_height:.2f} must be >= 0.",
        )
    if chart_top >= chart_bottom:
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            (
                f"chart_top={chart_top:.2f} must be < "
                f"chart_bottom={chart_bottom:.2f}."
            ),
        )
    for i, m in enumerate(milestones):
        if not (0.0 <= m.x_pos <= 1.0):
            raise EngineError(
                ErrorCode.INVALID_PARAMETER,
                (
                    f"milestones[{i}]: x_pos={m.x_pos:.3f} must be in "
                    "[0.0, 1.0] (relative to chart area)."
                ),
            )
        if m.style not in ("primary", "secondary"):
            raise EngineError(
                ErrorCode.INVALID_PARAMETER,
                (
                    f"milestones[{i}]: style={m.style!r} must be "
                    "'primary' or 'secondary'."
                ),
            )


def _render_phase_band(
    slide,
    phases: list[TimelinePhase],
    *,
    left: float,
    top: float,
    width: float,
    phase_band_height: float,
) -> tuple[list[dict], list[float]]:
    """フェーズ帯を描画し、各フェーズの右端 x 座標 (境界候補) を返す.

    戻り値:
        - phase_shapes: 各フェーズについて ``{"shape_index", "actual_font_size"}``
          の 3 ブロック (index・label・year_range) を追加順に記録。
        - right_edges: len(phases) 個の x 座標。最後の要素はフェーズ帯の
          右端で、境界ルールの描画対象ではない (``right_edges[:-1]`` が
          inter-phase 境界に対応する)。
    """
    phase_shapes: list[dict] = []
    right_edges: list[float] = []

    n = len(phases)
    if n == 0 or phase_band_height <= 0 or width <= 0:
        return phase_shapes, right_edges

    # 均等幅: gap は設けず、隣接フェーズは境界ルールで分離する想定。
    cell_w = width / n

    index_h = phase_band_height * _PHASE_INDEX_HEIGHT_RATIO
    label_h = phase_band_height * _PHASE_LABEL_HEIGHT_RATIO
    year_h = phase_band_height * _PHASE_YEAR_HEIGHT_RATIO

    for i, phase in enumerate(phases):
        cell_left = left + i * cell_w
        cursor_y = top

        # index_label (小)
        idx_before = len(slide.shapes)
        _shape, actual = add_auto_fit_textbox(
            slide,
            phase.index_label,
            cell_left,
            cursor_y,
            cell_w,
            index_h,
            font_size_pt=10,
            min_size_pt=7,
            bold=True,
            color_hex=_STYLE_PRIMARY_COLOR,
            align="left",
            vertical_anchor="top",
        )
        phase_shapes.append({
            "kind": "phase_index",
            "phase_index": i,
            "shape_index": idx_before,
            "actual_font_size": actual,
        })
        cursor_y += index_h

        # label (太字・中)
        idx_before = len(slide.shapes)
        _shape, actual = add_auto_fit_textbox(
            slide,
            phase.label,
            cell_left,
            cursor_y,
            cell_w,
            label_h,
            font_size_pt=13,
            min_size_pt=9,
            bold=True,
            color_hex=_STYLE_PRIMARY_COLOR,
            align="left",
            vertical_anchor="top",
        )
        phase_shapes.append({
            "kind": "phase_label",
            "phase_index": i,
            "shape_index": idx_before,
            "actual_font_size": actual,
        })
        cursor_y += label_h

        # year_range (小・グレー)
        idx_before = len(slide.shapes)
        _shape, actual = add_auto_fit_textbox(
            slide,
            phase.year_range,
            cell_left,
            cursor_y,
            cell_w,
            year_h,
            font_size_pt=9,
            min_size_pt=7,
            bold=False,
            color_hex=_STYLE_SECONDARY_COLOR,
            align="left",
            vertical_anchor="top",
        )
        phase_shapes.append({
            "kind": "phase_year",
            "phase_index": i,
            "shape_index": idx_before,
            "actual_font_size": actual,
        })

        right_edges.append(cell_left + cell_w)

    return phase_shapes, right_edges


def _render_phase_rules(
    slide,
    right_edges: list[float],
    *,
    top: float,
    bottom: float,
    color: str,
) -> list[dict]:
    """フェーズ境界の垂直ルールを描画する.

    ``right_edges`` は ``_render_phase_band`` が返す各フェーズの右端 x 座標。
    最後の要素はフェーズ帯の右端 (= チャート領域の右端) であり、境界では
    ないのでここでは無視する (``[:-1]``)。
    """
    rule_shapes: list[dict] = []
    if bottom <= top:
        return rule_shapes
    height = bottom - top
    # 最後の要素を除いた N-1 本を境界として描画する
    for boundary_index, x in enumerate(right_edges[:-1]):
        idx = _add_shape(
            slide,
            "rectangle",
            x - _RULE_WIDTH / 2,
            top,
            _RULE_WIDTH,
            height,
            fill_color=color,
            no_line=True,
        )
        rule_shapes.append({
            "kind": "phase_rule",
            "boundary_index": boundary_index,
            "shape_index": idx,
        })
    return rule_shapes


def _render_milestones(
    slide,
    milestones: list[TimelineMilestone],
    *,
    left: float,
    width: float,
    rule_top: float,
    chart_bottom: float,
    rule_color: str,
    font_size_pt: float,
    year_font_size_pt: float,
    theme: Optional[str] = None,
) -> tuple[list[dict], list[dict]]:
    """マイルストーン・ルールと注記テキストを描画する.

    各マイルストーンにつき:
      1. 垂直ルール (``rule_top`` から ``chart_bottom - style_margin``)
      2. year + label を 1 つのテキストボックスにまとめた注記 (ルール直上)

    Returns:
        ``(rule_shapes, label_shapes)``。それぞれ len = len(milestones)。
    """
    rule_shapes: list[dict] = []
    label_shapes: list[dict] = []

    for i, m in enumerate(milestones):
        x = left + m.x_pos * width
        color = _style_color(m.style, theme_name=theme)
        rule_bottom_margin = _style_rule_bottom_margin(m.style)
        rule_bottom = chart_bottom - rule_bottom_margin

        # rule_top > rule_bottom (= chart_bottom が十分低くない) ケースでは
        # ルールを描画しない (サイレントに逆向き矩形を描くより安全)。
        if rule_bottom > rule_top:
            idx = _add_shape(
                slide,
                "rectangle",
                x - _RULE_WIDTH / 2,
                rule_top,
                _RULE_WIDTH,
                rule_bottom - rule_top,
                fill_color=rule_color,
                no_line=True,
            )
            rule_shapes.append({
                "kind": "milestone_rule",
                "milestone_index": i,
                "shape_index": idx,
            })

        # ラベル矩形はルールの直上に配置する。ルール自体が描けなくても
        # ラベルは残す (テキストとして最低限の情報を提供する)。
        label_left = x - _MILESTONE_LABEL_HALF_WIDTH
        label_top = max(rule_top - _MILESTONE_LABEL_HEIGHT, 0.0)

        # 1 つの textbox に year + label を詰める。改行で区切る。
        combined = f"{m.year}\n{m.label}" if m.label else m.year
        idx_before = len(slide.shapes)
        _shape, actual = add_auto_fit_textbox(
            slide,
            combined,
            label_left,
            label_top,
            _MILESTONE_LABEL_HALF_WIDTH * 2,
            _MILESTONE_LABEL_HEIGHT,
            font_size_pt=max(font_size_pt, year_font_size_pt),
            min_size_pt=min(font_size_pt, year_font_size_pt) - 1,
            bold=(m.style == "primary"),
            color_hex=color,
            align="center",
            vertical_anchor="bottom",
        )
        label_shapes.append({
            "kind": "milestone_label",
            "milestone_index": i,
            "shape_index": idx_before,
            "actual_font_size": actual,
        })

    return rule_shapes, label_shapes


def add_milestone_timeline(
    slide,
    phases: list[TimelinePhase],
    milestones: list[TimelineMilestone],
    *,
    left: float,
    top: float,
    width: float,
    phase_band_height: float = 0.9,
    chart_top: float,
    chart_bottom: float,
    phase_rule_color: str = "E0E0E0",
    milestone_rule_color: str = "C0C0C0",
    milestone_font_size_pt: float = 9,
    milestone_year_font_size_pt: float = 11,
    theme: Optional[str] = None,
) -> dict:
    """フェーズ帯 + 垂直ルール + マイルストーン注記を描画する.

    呼び出し元は本関数に先立ってチャート等を ``chart_top`` から
    ``chart_bottom`` の垂直レンジに配置する想定である。本関数はチャート
    自体は描かず、その上に重ねる「注記レイヤ」を担当する。

    Args:
        slide: 対象スライド。
        phases: フェーズの順序付きリスト。空リスト可 (フェーズ帯をスキップ)。
        milestones: マイルストーンの順序付きリスト。空リスト可。
        left: タイムライン領域の左端 (inches)。
        top: タイムライン領域 (フェーズ帯) の上端 (inches)。
        width: タイムライン領域の幅 (inches)。チャート領域の幅と一致させる。
        phase_band_height: フェーズ帯の高さ (inches)。
        chart_top: チャート領域の上端 (inches)。マイルストーン・ルールの
            上端 (= ``top + phase_band_height``) と等しいか低い位置を想定する。
        chart_bottom: チャート領域の下端 (inches)。境界ルール・マイルス
            トーン・ルールの下限の基準となる。
        phase_rule_color: フェーズ境界ルールの色。theme トークン
            (例: ``"rule_subtle"``) または hex を受け付ける。
        milestone_rule_color: マイルストーン・ルールの色。theme トークン
            または hex。
        milestone_font_size_pt: マイルストーン label (年を除く) の font size。
        milestone_year_font_size_pt: マイルストーン year の font size。
        theme: テーマ名 (``"mckinsey"``, ``"deloitte"``, ``"neutral"``,
            ``"ir"``)。色引数に theme token が渡されたときの解決に使う。
            ``None`` もしくは未登録なら hex passthrough のみ。

    Returns:
        ``{"phase_shapes": [...], "milestone_shapes": [...],
           "rule_shapes": [...], "consumed_height": phase_band_height}``
        の dict。``phase_shapes`` は各フェーズ 3 要素 (index/label/year) を
        含み、``rule_shapes`` はフェーズ境界ルールとマイルストーン・ルールを
        混在させた全ルールを追加順に含む。``milestone_shapes`` は各マイルス
        トーンの注記テキストを含む。

    Raises:
        EngineError: 入力が無効なとき (``INVALID_PARAMETER``)。
    """
    _validate_inputs(
        phases,
        milestones,
        left=left,
        top=top,
        width=width,
        phase_band_height=phase_band_height,
        chart_top=chart_top,
        chart_bottom=chart_bottom,
    )

    # theme token → 6-hex 解決 (#125)。
    phase_rule_color = resolve_theme_color(phase_rule_color, theme)
    milestone_rule_color = resolve_theme_color(milestone_rule_color, theme)

    phase_shapes, right_edges = _render_phase_band(
        slide,
        phases,
        left=left,
        top=top,
        width=width,
        phase_band_height=phase_band_height,
    )

    # ルールは caller-declared chart area の上端 (chart_top) から下に引く。
    # 旧実装は `top + phase_band_height` を使って chart_top 引数を無視していた (#119)。
    rule_top = chart_top

    phase_rule_shapes = _render_phase_rules(
        slide,
        right_edges,
        top=rule_top,
        bottom=chart_bottom,
        color=phase_rule_color,
    )

    milestone_rule_shapes, milestone_label_shapes = _render_milestones(
        slide,
        milestones,
        left=left,
        width=width,
        rule_top=rule_top,
        chart_bottom=chart_bottom,
        rule_color=milestone_rule_color,
        font_size_pt=milestone_font_size_pt,
        year_font_size_pt=milestone_year_font_size_pt,
        theme=theme,
    )

    # rule_shapes は「フェーズ境界 → マイルストーン」の順でまとめる。
    rule_shapes = phase_rule_shapes + milestone_rule_shapes

    return {
        "phase_shapes": phase_shapes,
        "milestone_shapes": milestone_label_shapes,
        "rule_shapes": rule_shapes,
        "consumed_height": phase_band_height,
    }
