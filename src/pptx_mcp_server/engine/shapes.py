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
