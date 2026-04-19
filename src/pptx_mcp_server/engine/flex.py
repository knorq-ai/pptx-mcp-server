"""CSS Flexbox 相当のコンテナプリミティブ.

PPTX レイアウトにおいて、行方向 (``row``) または列方向 (``column``) に
子要素をサイジング・配置する。CSS Flexbox の ``fixed``・``grow``・
``content`` 3 種のサイジングモデルをサポートする MVP 実装である。

# 設計メモ
- 子要素の描画は呼び出し元が指定する ``render(x, y, w, h)`` コールバック
  に委譲する。flex 自身は形状を生成せず、配置計算のみを担う。
- ``wrap`` ・ ``justify`` ・ ``align-self`` 等の CSS 機能はスコープ外。
- cross 軸の実寸 (intrinsic cross size) は測定しないため、非 stretch
  でも cross サイズは usable_cross を割り当てる (MVP の制約)。

# アルゴリズム概略 (row の場合)
1. usable_main = width - 2*padding - gap*(n-1)
2. fixed_total と content_total をまず確定させる。
3. 残余 ``remain = max(0, usable_main - fixed_total - content_total)`` を
   grow 項目の ``grow`` 比率で按分する。
4. 各 grow 項目を ``[min_size, max_size]`` にクランプし、クランプにより
   余った分を未クランプ項目で再配分する (CSS Flex 方式)。
5. main 軸方向に padding から積み上げて配置し、cross 軸は padding から
   usable_cross を割り当てる。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable, Literal

from .pptx_io import EngineError, ErrorCode, open_pptx, save_pptx, _get_slide
from .shapes import _add_shape, add_auto_fit_textbox

SizingMode = Literal["fixed", "grow", "content"]
Direction = Literal["row", "column"]
Align = Literal["start", "center", "end", "stretch"]


@dataclass
class FlexItem:
    """Flex コンテナの子要素.

    Attributes:
        sizing: サイジングモード。``"fixed"``・``"grow"``・``"content"``。
        render: ``render(x, y, w, h)`` 形式の描画コールバック。実際に
            shape を生成する責務は呼び出し元にある。
        size: ``sizing="fixed"`` のときの main 軸サイズ (inches)。
        grow: ``sizing="grow"`` のときの成長係数。0 以下は 0 扱い。
        content_size: ``sizing="content"`` のときの main 軸サイズ (inches)。
            呼び出し元が事前に計測した値を渡す。
        min_size: クランプ下限 (inches)。主に grow 項目で有効。
        max_size: クランプ上限 (inches)。主に grow 項目で有効。
    """

    sizing: SizingMode
    render: Callable[[float, float, float, float], None]
    size: float = 0.0
    grow: float = 1.0
    content_size: float = 0.0
    min_size: float = 0.0
    max_size: float = float("inf")


def _item_base_main_size(item: FlexItem) -> float:
    """grow でない項目の main 軸ベースサイズを返す."""
    if item.sizing == "fixed":
        return max(item.size, 0.0)
    if item.sizing == "content":
        return max(item.content_size, 0.0)
    # grow 項目はこの関数の対象外
    return 0.0


def _distribute_grow(
    items: list[FlexItem],
    indices: list[int],
    remain: float,
) -> dict[int, float]:
    """grow 項目にサイズを按分する (CSS Flex のクランプ再配分を実装).

    ``indices`` は ``items`` 内で ``sizing=="grow"`` である要素の index。
    各要素の ``max(grow, 0)`` を係数として ``remain`` を按分する。クランプ
    が発生した項目は pool から除外し、残余を未クランプ項目で再配分する。
    この手続きをクランプが発生しなくなるまで繰り返す。

    Returns:
        index -> 決定サイズ の dict。
    """
    result: dict[int, float] = {}
    pool: list[int] = []
    for i in indices:
        g = max(items[i].grow, 0.0)
        if g <= 0.0:
            # 成長しない grow 項目は min_size を割り当てる (0 起点の clamp)
            clamped = max(items[i].min_size, 0.0)
            result[i] = clamped
            remain -= clamped
        else:
            pool.append(i)

    # pool が空 = クランプ以外で処理済み。
    while pool:
        total_grow = sum(max(items[i].grow, 0.0) for i in pool)
        if total_grow <= 0.0:
            for i in pool:
                result[i] = max(items[i].min_size, 0.0)
            break

        # 分配対象の残余が負なら 0 としてクランプ (縮めない)
        dist = max(remain, 0.0)
        newly_clamped: list[int] = []
        for i in pool:
            g = max(items[i].grow, 0.0)
            share = dist * (g / total_grow)
            if share < items[i].min_size:
                result[i] = items[i].min_size
                newly_clamped.append(i)
            elif share > items[i].max_size:
                result[i] = items[i].max_size
                newly_clamped.append(i)

        if not newly_clamped:
            # 残った pool には純粋な按分値を割り当てる
            for i in pool:
                g = max(items[i].grow, 0.0)
                result[i] = dist * (g / total_grow)
            break

        # クランプで確定した分を remain から差し引き、pool から除外
        for i in newly_clamped:
            remain -= result[i]
        pool = [i for i in pool if i not in newly_clamped]

    return result


def add_flex_container(
    slide,
    items: list[FlexItem],
    *,
    left: float,
    top: float,
    width: float,
    height: float,
    direction: Direction = "row",
    gap: float = 0.15,
    padding: float = 0.0,
    align: Align = "stretch",
) -> list[tuple[float, float, float, float]]:
    """子要素のサイズ・位置を計算し、各 ``item.render(x, y, w, h)`` を呼ぶ.

    Args:
        slide: 対象スライド (本関数はスライド自体を参照せず、render 側に委譲する)。
        items: 子要素のリスト。空リストでも可。
        left, top, width, height: コンテナの位置・寸法 (inches)。
        direction: 主軸方向。``"row"`` または ``"column"``。
        gap: 子要素間のスペース (inches)。
        padding: コンテナ内側の余白 (inches、全 4 辺に適用)。
        align: cross 軸方向の整列。``"stretch"`` 以外は現状、cross サイズ
            自体は stretch と同等に扱う (intrinsic cross 未測定のため MVP)。

    Returns:
        各 item に割り当てられた ``(left, top, width, height)`` のタプル
        のリスト。``items`` と同順。
    """
    if not items:
        return []

    n = len(items)

    if direction == "row":
        usable_main = width - 2 * padding - gap * max(n - 1, 0)
        usable_cross = height - 2 * padding
    else:
        usable_main = height - 2 * padding - gap * max(n - 1, 0)
        usable_cross = width - 2 * padding

    # fixed + content の合計サイズ
    fixed_total = sum(
        max(it.size, 0.0) for it in items if it.sizing == "fixed"
    )
    content_total = sum(
        max(it.content_size, 0.0) for it in items if it.sizing == "content"
    )
    remain = max(0.0, usable_main - fixed_total - content_total)

    grow_indices = [i for i, it in enumerate(items) if it.sizing == "grow"]
    grow_sizes = _distribute_grow(items, grow_indices, remain)

    # 各 item の main 軸サイズを決定
    main_sizes: list[float] = []
    for i, it in enumerate(items):
        if it.sizing == "grow":
            main_sizes.append(grow_sizes.get(i, 0.0))
        else:
            main_sizes.append(_item_base_main_size(it))

    # cross 軸サイズは MVP では常に usable_cross (負にならないようクランプ)
    cross_size = max(usable_cross, 0.0)

    # 配置
    allocations: list[tuple[float, float, float, float]] = []
    if direction == "row":
        cursor = left + padding
        cross_origin = top + padding
        for i, it in enumerate(items):
            w = main_sizes[i]
            x = cursor
            y = cross_origin
            h = cross_size
            allocations.append((x, y, w, h))
            cursor += w + gap
    else:
        cursor = top + padding
        cross_origin = left + padding
        for i, it in enumerate(items):
            h = main_sizes[i]
            x = cross_origin
            y = cursor
            w = cross_size
            allocations.append((x, y, w, h))
            cursor += h + gap

    # 描画コールバックを呼ぶ
    for it, (x, y, w, h) in zip(items, allocations):
        it.render(x, y, w, h)

    return allocations


# --- 宣言的 (declarative) 項目からの flex 生成 ---------------------------

_SUPPORTED_ITEM_TYPES = ("text", "rectangle")


def _build_declarative_item(
    slide,
    spec: dict[str, Any],
    created_shape_indices: list[dict[str, Any]],
) -> FlexItem:
    """MCP 層から渡される dict 仕様から ``FlexItem`` を生成する.

    サポートする ``type``:

    - ``"text"``: ``add_auto_fit_textbox`` を使って描画する。
    - ``"rectangle"``: ``_add_shape`` で塗りつぶし矩形を描画する。

    描画結果 (shape_index) は ``created_shape_indices`` に追記される。
    """
    sizing: SizingMode = spec.get("sizing", "grow")
    if sizing not in ("fixed", "grow", "content"):
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            f"Unknown sizing mode '{sizing}'. Expected fixed|grow|content.",
        )

    item_type = spec.get("type")
    if item_type not in _SUPPORTED_ITEM_TYPES:
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            f"Unknown flex item type '{item_type}'. "
            f"Supported: {', '.join(_SUPPORTED_ITEM_TYPES)}.",
        )

    size = float(spec.get("size", 0.0))
    grow = float(spec.get("grow", 1.0))
    content_size = float(spec.get("content_size", 0.0))
    min_size = float(spec.get("min_size", 0.0))
    max_size_raw = spec.get("max_size", None)
    max_size = float("inf") if max_size_raw is None else float(max_size_raw)

    if item_type == "text":
        text = spec.get("text", "")
        font_name = spec.get("font_name", "Arial")
        font_size_pt = float(spec.get("font_size_pt", 11))
        min_size_pt = float(spec.get("min_size_pt", 7))
        bold = bool(spec.get("bold", False))
        color_hex = spec.get("color_hex", "333333")
        align = spec.get("align", "left")
        vertical_anchor = spec.get("vertical_anchor", "top")
        truncate = bool(spec.get("truncate_with_ellipsis", True))

        def render_text(x: float, y: float, w: float, h: float) -> None:
            shape, actual = add_auto_fit_textbox(
                slide, text, x, y, w, h,
                font_name=font_name,
                font_size_pt=font_size_pt,
                min_size_pt=min_size_pt,
                bold=bold,
                color_hex=color_hex,
                align=align,
                vertical_anchor=vertical_anchor,
                truncate_with_ellipsis=truncate,
            )
            idx = -1
            for i, s in enumerate(slide.shapes):
                if s is shape:
                    idx = i
                    break
            created_shape_indices.append({
                "type": "text",
                "shape_index": idx,
                "actual_font_size": actual,
            })

        render = render_text

    else:  # rectangle
        fill_color = spec.get("fill_color")
        line_color = spec.get("line_color")
        line_width = spec.get("line_width")
        no_line = bool(spec.get("no_line", False))

        def render_rect(x: float, y: float, w: float, h: float) -> None:
            idx = _add_shape(
                slide, "rectangle", x, y, w, h,
                fill_color=fill_color,
                line_color=line_color,
                line_width=line_width,
                no_line=no_line,
            )
            created_shape_indices.append({
                "type": "rectangle",
                "shape_index": idx,
            })

        render = render_rect

    return FlexItem(
        sizing=sizing,
        render=render,
        size=size,
        grow=grow,
        content_size=content_size,
        min_size=min_size,
        max_size=max_size,
    )


def add_flex_container_file(
    file_path: str,
    slide_index: int,
    items: list[dict[str, Any]],
    *,
    left: float,
    top: float,
    width: float,
    height: float,
    direction: Direction = "row",
    gap: float = 0.15,
    padding: float = 0.0,
    align: Align = "stretch",
) -> dict[str, Any]:
    """File-based wrapper: 宣言的仕様を受け取って flex コンテナを配置する.

    Args:
        file_path: PPTX ファイルパス。
        slide_index: 対象スライドの index。
        items: ``{"sizing": ..., "type": ..., ...}`` 形式の dict のリスト。
        left, top, width, height: コンテナ寸法 (inches)。
        direction, gap, padding, align: flex 設定。

    Returns:
        ``{"allocations": [...], "shapes": [...], "slide_index": int}``
        を含む dict。``allocations`` は ``(left, top, width, height)`` 4-tuple
        のリスト、``shapes`` は生成された各子要素の shape 識別情報。
    """
    prs = open_pptx(file_path)
    slide = _get_slide(prs, slide_index)

    created_shape_indices: list[dict[str, Any]] = []
    flex_items: list[FlexItem] = [
        _build_declarative_item(slide, s, created_shape_indices)
        for s in items
    ]

    allocations = add_flex_container(
        slide, flex_items,
        left=left, top=top, width=width, height=height,
        direction=direction, gap=gap, padding=padding, align=align,
    )

    save_pptx(prs, file_path)
    return {
        "slide_index": slide_index,
        "allocations": [list(a) for a in allocations],
        "shapes": created_shape_indices,
    }
