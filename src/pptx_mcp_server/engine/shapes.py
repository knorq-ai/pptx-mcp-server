"""
Shape operations -- add textbox, add shape, edit text, add paragraph, delete, list.
"""

from __future__ import annotations

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

from .pptx_io import (
    EngineError, ErrorCode,
    open_pptx, save_pptx, _get_slide, _get_shape, _parse_color,
)
from .text_metrics import estimate_text_height, estimate_text_width, wrap_text

from ..theme import Theme, resolve_color

# Mapping of alignment strings to enums
_ALIGN_MAP = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
    "justify": PP_ALIGN.JUSTIFY,
}

_ANCHOR_MAP = {
    "top": MSO_ANCHOR.TOP,
    "middle": MSO_ANCHOR.MIDDLE,
    "bottom": MSO_ANCHOR.BOTTOM,
}

# Common auto shapes
_SHAPE_MAP = {
    "rectangle": MSO_SHAPE.RECTANGLE,
    "rounded_rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
    "oval": MSO_SHAPE.OVAL,
    "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
    "diamond": MSO_SHAPE.DIAMOND,
    "pentagon": MSO_SHAPE.PENTAGON,
    "hexagon": MSO_SHAPE.HEXAGON,
    "chevron": MSO_SHAPE.CHEVRON,
    "arrow_right": MSO_SHAPE.RIGHT_ARROW,
    "arrow_left": MSO_SHAPE.LEFT_ARROW,
    "arrow_up": MSO_SHAPE.UP_ARROW,
    "arrow_down": MSO_SHAPE.DOWN_ARROW,
    "callout": MSO_SHAPE.RECTANGULAR_CALLOUT,
    "star_5": MSO_SHAPE.STAR_5_POINT,
}


# ── Internal helpers ────────────────────────────────────────────


def _apply_font(
    paragraph,
    font_name=None,
    font_size=None,
    font_color=None,
    bold=None,
    italic=None,
    underline=None,
    theme=None,
):
    """Apply font formatting to a paragraph's font."""
    font = paragraph.font
    if font_name is not None:
        font.name = font_name
    if font_size is not None:
        font.size = Pt(font_size)
    if font_color is not None:
        color_hex = resolve_color(theme, font_color) if theme else font_color
        font.color.rgb = _parse_color(color_hex)
    if bold is not None:
        font.bold = bold
    if italic is not None:
        font.italic = italic
    if underline is not None:
        font.underline = underline


def _apply_paragraph(
    paragraph,
    alignment=None,
    line_spacing=None,
):
    """Apply paragraph-level formatting."""
    if alignment and alignment in _ALIGN_MAP:
        paragraph.alignment = _ALIGN_MAP[alignment]
    if line_spacing is not None:
        paragraph.line_spacing = Pt(line_spacing)


# ── In-memory primitives ────────────────────────────────────────


def _add_textbox(
    slide,
    left,
    top,
    width,
    height,
    text="",
    font_name=None,
    font_size=None,
    font_color=None,
    bold=None,
    italic=None,
    alignment=None,
    vertical_anchor=None,
    word_wrap=True,
    line_spacing=None,
    underline=None,
    theme=None,
):
    """In-memory: add textbox to an existing slide object. Returns shape_index."""
    if theme:
        font_name = font_name or theme.fonts.get("body")
        if font_color:
            font_color = resolve_color(theme, font_color)

    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height),
    )
    tf = txBox.text_frame
    tf.word_wrap = word_wrap

    if vertical_anchor and vertical_anchor in _ANCHOR_MAP:
        from pptx.oxml.ns import qn
        bodyPr = tf._txBody.find(qn("a:bodyPr"))
        anchor_val = {"top": "t", "middle": "ctr", "bottom": "b"}[vertical_anchor]
        bodyPr.set("anchor", anchor_val)

    p = tf.paragraphs[0]
    p.text = text

    _apply_font(p, font_name, font_size, font_color, bold, italic, underline, theme=theme)
    _apply_paragraph(p, alignment, line_spacing)

    return len(list(slide.shapes)) - 1


def _add_image(slide, image_path, left, top, width=None, height=None):
    """In-memory: add an image to a slide. Returns shape_index.

    If only width is given, height is auto-calculated to maintain aspect ratio (and vice versa).
    If both are given, image is stretched to fit.
    If neither is given, image is placed at its native size.
    """
    import os
    if not os.path.exists(image_path):
        raise EngineError(ErrorCode.FILE_NOT_FOUND, f"Image not found: {image_path}")

    kwargs = {"left": Inches(left), "top": Inches(top)}
    if width is not None:
        kwargs["width"] = Inches(width)
    if height is not None:
        kwargs["height"] = Inches(height)

    slide.shapes.add_picture(image_path, **kwargs)
    return len(list(slide.shapes)) - 1


def _add_shape(
    slide,
    shape_type,
    left,
    top,
    width,
    height,
    fill_color=None,
    line_color=None,
    line_width=None,
    no_line=False,
    text=None,
    font_name=None,
    font_size=None,
    font_color=None,
    bold=None,
    alignment=None,
    theme=None,
):
    """In-memory: add an auto shape to a slide. Returns shape_index."""
    if theme:
        font_name = font_name or theme.fonts.get("body")
        if fill_color:
            fill_color = resolve_color(theme, fill_color)
        if line_color:
            line_color = resolve_color(theme, line_color)
        if font_color:
            font_color = resolve_color(theme, font_color)

    shape_enum = _SHAPE_MAP.get(shape_type.lower())
    if shape_enum is None:
        available = ", ".join(sorted(_SHAPE_MAP.keys()))
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            f"Unknown shape type '{shape_type}'. Available: {available}",
        )

    shape = slide.shapes.add_shape(
        shape_enum, Inches(left), Inches(top), Inches(width), Inches(height),
    )

    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = _parse_color(fill_color)

    if no_line:
        shape.line.fill.background()
    elif line_color:
        shape.line.color.rgb = _parse_color(line_color)
    if line_width is not None:
        shape.line.width = Pt(line_width)

    if text is not None:
        tf = shape.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        _apply_font(p, font_name, font_size, font_color, bold, None, None, theme=None)
        _apply_paragraph(p, alignment, None)

    return len(list(slide.shapes)) - 1


def _edit_text(
    slide,
    shape_index,
    text=None,
    paragraph_index=0,
    font_name=None,
    font_size=None,
    font_color=None,
    bold=None,
    italic=None,
    underline=None,
    alignment=None,
    line_spacing=None,
    theme=None,
):
    """In-memory: edit text content and formatting in a shape's paragraph."""
    shape = _get_shape(slide, shape_index)

    if not shape.has_text_frame:
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            f"Shape [{shape_index}] does not have a text frame",
        )

    tf = shape.text_frame
    if paragraph_index < 0 or paragraph_index >= len(tf.paragraphs):
        raise EngineError(
            ErrorCode.INDEX_OUT_OF_RANGE,
            f"Paragraph index {paragraph_index} out of range (0-{len(tf.paragraphs) - 1})",
        )

    if theme:
        font_name = font_name or None  # don't override with theme default for edit
        if font_color:
            font_color = resolve_color(theme, font_color)

    p = tf.paragraphs[paragraph_index]
    if text is not None:
        p.text = text
    _apply_font(p, font_name, font_size, font_color, bold, italic, underline, theme=None)
    _apply_paragraph(p, alignment, line_spacing)


def _add_paragraph(
    slide,
    shape_index,
    text,
    font_name=None,
    font_size=None,
    font_color=None,
    bold=None,
    italic=None,
    underline=None,
    alignment=None,
    line_spacing=None,
    theme=None,
):
    """In-memory: append a new paragraph to an existing shape's text frame.
    Returns paragraph index."""
    shape = _get_shape(slide, shape_index)

    if not shape.has_text_frame:
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            f"Shape [{shape_index}] does not have a text frame",
        )

    if theme:
        if font_color:
            font_color = resolve_color(theme, font_color)

    tf = shape.text_frame
    p = tf.add_paragraph()
    p.text = text
    _apply_font(p, font_name, font_size, font_color, bold, italic, underline, theme=None)
    _apply_paragraph(p, alignment, line_spacing)

    return len(tf.paragraphs) - 1


def _delete_shape(slide, shape_index, theme=None):
    """In-memory: delete a shape from a slide by index."""
    shape = _get_shape(slide, shape_index)
    sp = shape._element
    sp.getparent().remove(sp)


def _list_shapes(slide, slide_index, theme=None):
    """In-memory: list all shapes on a slide."""
    lines = [f"Slide [{slide_index}] -- {len(slide.shapes)} shapes"]
    for si, shape in enumerate(slide.shapes):
        stype = "TABLE" if shape.has_table else ("TEXTBOX" if shape.has_text_frame else str(shape.shape_type))
        pos = f"({shape.left / 914400:.2f}\", {shape.top / 914400:.2f}\")"
        size = f"{shape.width / 914400:.2f}\" x {shape.height / 914400:.2f}\""
        text_preview = ""
        if shape.has_text_frame:
            txt = shape.text_frame.text[:50]
            if txt:
                text_preview = f" \"{txt}\""
        lines.append(f"  [{si}] {stype} @ {pos} {size}{text_preview}")
    return "\n".join(lines)


# ── File-based public wrappers ──────────────────────────────────


def add_textbox(
    file_path,
    slide_index,
    left,
    top,
    width,
    height,
    text="",
    font_name=None,
    font_size=None,
    font_color=None,
    bold=None,
    italic=None,
    alignment=None,
    vertical_anchor=None,
    word_wrap=True,
    line_spacing=None,
    underline=None,
):
    """File-based wrapper: add a text box to a slide."""
    prs = open_pptx(file_path)
    slide = _get_slide(prs, slide_index)
    idx = _add_textbox(
        slide, left, top, width, height, text,
        font_name, font_size, font_color, bold, italic,
        alignment, vertical_anchor, word_wrap, line_spacing, underline,
    )
    save_pptx(prs, file_path)
    return f"Added textbox [{idx}] on slide [{slide_index}]"


def add_image(file_path, slide_index, image_path, left, top, width=None, height=None):
    """File-based wrapper: add an image to a slide."""
    prs = open_pptx(file_path)
    slide = _get_slide(prs, slide_index)
    idx = _add_image(slide, image_path, left, top, width, height)
    save_pptx(prs, file_path)
    return f"Added image [{idx}] on slide [{slide_index}]"


def add_shape(
    file_path,
    slide_index,
    shape_type,
    left,
    top,
    width,
    height,
    fill_color=None,
    line_color=None,
    line_width=None,
    no_line=False,
    text=None,
    font_name=None,
    font_size=None,
    font_color=None,
    bold=None,
    alignment=None,
):
    """File-based wrapper: add an auto shape to a slide."""
    prs = open_pptx(file_path)
    slide = _get_slide(prs, slide_index)
    idx = _add_shape(
        slide, shape_type, left, top, width, height,
        fill_color, line_color, line_width, no_line,
        text, font_name, font_size, font_color, bold, alignment,
    )
    save_pptx(prs, file_path)
    return f"Added {shape_type} [{idx}] on slide [{slide_index}]"


def edit_text(
    file_path,
    slide_index,
    shape_index,
    text=None,
    paragraph_index=0,
    font_name=None,
    font_size=None,
    font_color=None,
    bold=None,
    italic=None,
    underline=None,
    alignment=None,
    line_spacing=None,
):
    """File-based wrapper: edit text content and formatting."""
    prs = open_pptx(file_path)
    slide = _get_slide(prs, slide_index)
    _edit_text(
        slide, shape_index, text, paragraph_index,
        font_name, font_size, font_color, bold, italic, underline,
        alignment, line_spacing,
    )
    save_pptx(prs, file_path)
    return f"Edited shape [{shape_index}] paragraph [{paragraph_index}] on slide [{slide_index}]"


def add_paragraph(
    file_path,
    slide_index,
    shape_index,
    text,
    font_name=None,
    font_size=None,
    font_color=None,
    bold=None,
    italic=None,
    underline=None,
    alignment=None,
    line_spacing=None,
):
    """File-based wrapper: append a paragraph to a shape's text frame."""
    prs = open_pptx(file_path)
    slide = _get_slide(prs, slide_index)
    p_idx = _add_paragraph(
        slide, shape_index, text,
        font_name, font_size, font_color, bold, italic, underline,
        alignment, line_spacing,
    )
    save_pptx(prs, file_path)
    return f"Added paragraph [{p_idx}] to shape [{shape_index}] on slide [{slide_index}]"


def delete_shape(file_path, slide_index, shape_index):
    """File-based wrapper: delete a shape from a slide."""
    prs = open_pptx(file_path)
    slide = _get_slide(prs, slide_index)
    _delete_shape(slide, shape_index)
    save_pptx(prs, file_path)
    return f"Deleted shape [{shape_index}] from slide [{slide_index}]"


def list_shapes(file_path, slide_index):
    """File-based wrapper: list all shapes on a slide."""
    prs = open_pptx(file_path)
    slide = _get_slide(prs, slide_index)
    return _list_shapes(slide, slide_index)


# ── Auto-fit textbox ────────────────────────────────────────────

# 内側 padding (inches、左右各側)。左右に同量の padding を見込むため、
# usable width は width - 2 * _AUTO_FIT_PADDING_PER_SIDE となる。
# cards.py など他モジュールからも参照できるよう module-level の定数として公開する。
_AUTO_FIT_PADDING_PER_SIDE: float = 0.05

# 後方互換のための別名 (旧名)。新規コードは _AUTO_FIT_PADDING_PER_SIDE を使うこと。
_AUTO_FIT_PADDING: float = _AUTO_FIT_PADDING_PER_SIDE

# font size 縮小ステップ (pt)
_AUTO_FIT_STEP_PT: float = 0.5

# 高さ推定に使う行高倍率
_AUTO_FIT_LINE_HEIGHT: float = 1.2

# 省略記号
_ELLIPSIS: str = "\u2026"


def _fit_font_size(
    text: str,
    usable_width: float,
    height: float,
    font_name: str,
    font_size_pt: float,
    min_size_pt: float,
) -> tuple[float, bool]:
    """指定 box に収まる font size を二分探索ではなく 0.5pt ステップで決定する.

    Returns:
        (size, fits): ``size`` は採用する font size。``fits`` は最終的にその
        size で text が box 内に収まるかどうか (min に達してもなおオーバー
        フローする場合は False)。
    """
    size = float(font_size_pt)
    min_size = float(min_size_pt)
    while size > min_size:
        height_est = estimate_text_height(
            text, usable_width, size, font_name,
            line_height_factor=_AUTO_FIT_LINE_HEIGHT,
        )
        if height_est <= height:
            return size, True
        size = round(size - _AUTO_FIT_STEP_PT, 2)
    # min_size でチェック
    height_est = estimate_text_height(
        text, usable_width, min_size, font_name,
        line_height_factor=_AUTO_FIT_LINE_HEIGHT,
    )
    return min_size, height_est <= height


def _truncate_to_fit(
    text: str,
    usable_width: float,
    height: float,
    font_name: str,
    size_pt: float,
) -> str:
    """``size_pt`` の size で ``(usable_width, height)`` に収まるよう末尾を
    省略記号で切り詰める.

    方針:
    - まず wrap_text で折り返し結果を得る。
    - 高さが許す行数 ``max_lines`` を算出する。
    - 最終行は末尾から 1 文字ずつ削りながら ``line + 省略記号`` の幅が
      ``usable_width`` 以下になるまで縮める。
    - ``max_lines == 0`` の場合は空文字列を返す。
    """
    line_h = size_pt * 0.0139 * _AUTO_FIT_LINE_HEIGHT
    if line_h <= 0:
        return text
    max_lines = int(height // line_h)
    if max_lines <= 0:
        return ""

    lines = wrap_text(text, usable_width, size_pt, font_name)
    if not lines:
        return ""

    if len(lines) <= max_lines:
        # 高さには収まるが、幅の再確認は wrap で保証済み。そのまま返す。
        return "\n".join(lines)

    retained = lines[:max_lines]
    last = retained[-1]
    # last + ellipsis が usable_width を超える間、末尾から 1 文字ずつ削る。
    while last and estimate_text_width(last + _ELLIPSIS, size_pt, font_name) > usable_width:
        last = last[:-1]
    retained[-1] = (last + _ELLIPSIS) if last else _ELLIPSIS
    return "\n".join(retained)


def add_auto_fit_textbox(
    slide,
    text: str,
    left: float,
    top: float,
    width: float,
    height: float,
    *,
    font_name: str = "Arial",
    font_size_pt: float = 11,
    min_size_pt: float = 7,
    bold: bool = False,
    color_hex: str = "333333",
    align: str = "left",
    vertical_anchor: str = "top",
    truncate_with_ellipsis: bool = True,
) -> tuple[object, float]:
    """指定 box に収まる最大 font size でテキストを描画する.

    動作:
        (a) 内側 padding 0.05" を左右に見込んで usable width を計算する。
        (b) ``font_size_pt`` から開始し、``estimate_text_height(wrapped)`` が
            ``height`` 以下になるまで 0.5pt 刻みで縮小する。
        (c) ``min_size_pt`` に達しても収まらない場合、
            ``truncate_with_ellipsis=True`` なら末尾を省略記号で切り詰めて
            描画する。False なら full text をそのまま描画しオーバーフロー
            を許容する。
        (d) 決定した font size で ``_add_textbox`` 経由で textframe を生成し、
            ``vertical_anchor`` に応じて ``MSO_ANCHOR`` を設定する。

    Args:
        slide: 対象スライドオブジェクト。
        text: 描画するテキスト。
        left, top, width, height: 位置と寸法 (inches)。
        font_name: フォント名。デフォルトは ``"Arial"``。
        font_size_pt: 開始 font size (pt)。
        min_size_pt: 縮小の下限 (pt)。
        bold: 太字にするかどうか。
        color_hex: 文字色 (hex、``#`` なし)。
        align: 水平方向の揃え ``"left" | "center" | "right"``。
        vertical_anchor: 垂直方向の揃え ``"top" | "middle" | "bottom"``。
        truncate_with_ellipsis: min size でも収まらない場合に末尾を
            省略記号で切り詰めるかどうか。

    Returns:
        ``(shape, actual_font_size)`` のタプル。テスト・デバッグ用途。
    """
    usable_width = max(width - 2 * _AUTO_FIT_PADDING, 0.01)

    actual_size, fits = _fit_font_size(
        text, usable_width, height, font_name, font_size_pt, min_size_pt,
    )

    rendered_text = text
    if not fits and truncate_with_ellipsis:
        rendered_text = _truncate_to_fit(
            text, usable_width, height, font_name, actual_size,
        )

    idx = _add_textbox(
        slide,
        left, top, width, height,
        text=rendered_text,
        font_name=font_name,
        font_size=actual_size,
        font_color=color_hex,
        bold=bold if bold else None,
        italic=None,
        alignment=align,
        vertical_anchor=vertical_anchor,
        word_wrap=True,
        line_spacing=None,
        underline=None,
    )

    shape = slide.shapes[idx]
    tf = shape.text_frame
    tf.word_wrap = True
    # auto_size は手動でサイズ決定済みなので明示的に None にする。
    tf.auto_size = None
    if vertical_anchor in _ANCHOR_MAP:
        tf.vertical_anchor = _ANCHOR_MAP[vertical_anchor]

    return shape, actual_size


def add_auto_fit_textbox_file(
    file_path: str,
    slide_index: int,
    text: str,
    left: float,
    top: float,
    width: float,
    height: float,
    font_name: str = "Arial",
    font_size_pt: float = 11,
    min_size_pt: float = 7,
    bold: bool = False,
    color_hex: str = "333333",
    align: str = "left",
    vertical_anchor: str = "top",
    truncate_with_ellipsis: bool = True,
) -> dict:
    """File-based wrapper: 指定 box に収まるよう自動縮小する textbox を追加する.

    Returns:
        ``{"shape_index": int, "slide_index": int, "actual_font_size": float}``
        を含む dict。MCP ツール経由で呼び出される想定。
    """
    prs = open_pptx(file_path)
    slide = _get_slide(prs, slide_index)
    shape, actual_size = add_auto_fit_textbox(
        slide, text, left, top, width, height,
        font_name=font_name,
        font_size_pt=font_size_pt,
        min_size_pt=min_size_pt,
        bold=bold,
        color_hex=color_hex,
        align=align,
        vertical_anchor=vertical_anchor,
        truncate_with_ellipsis=truncate_with_ellipsis,
    )
    # shape_index を決定 (slide 内で shape を線形検索)
    shape_index = -1
    for i, s in enumerate(slide.shapes):
        if s is shape:
            shape_index = i
            break
    save_pptx(prs, file_path)
    return {
        "slide_index": slide_index,
        "shape_index": shape_index,
        "shape_name": shape.name,
        "actual_font_size": actual_size,
    }
