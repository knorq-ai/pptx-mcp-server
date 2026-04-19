"""
Formatting operations -- shape formatting, slide dimensions.

Note: format_text has been merged into edit_text in shapes.py.
"""

from __future__ import annotations

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor

from .pptx_io import (
    EngineError, ErrorCode,
    open_pptx, save_pptx, _get_slide, _get_shape, _parse_color,
)

from ..theme import Theme, resolve_color


# ── In-memory primitives ────────────────────────────────────────


def _format_shape(
    slide,
    shape_index,
    left=None,
    top=None,
    width=None,
    height=None,
    fill_color=None,
    no_fill=False,
    line_color=None,
    line_width=None,
    no_line=False,
    rotation=None,
    theme=None,
):
    """In-memory: reposition, resize, or restyle a shape."""
    shape = _get_shape(slide, shape_index)

    if theme:
        if fill_color:
            fill_color = resolve_color(theme, fill_color)
        if line_color:
            line_color = resolve_color(theme, line_color)

    if left is not None:
        shape.left = Inches(left)
    if top is not None:
        shape.top = Inches(top)
    if width is not None:
        shape.width = Inches(width)
    if height is not None:
        shape.height = Inches(height)
    if rotation is not None:
        shape.rotation = rotation

    if no_fill:
        shape.fill.background()
    elif fill_color is not None:
        shape.fill.solid()
        shape.fill.fore_color.rgb = _parse_color(fill_color)

    if no_line:
        shape.line.fill.background()
    elif line_color is not None:
        shape.line.color.rgb = _parse_color(line_color)
    if line_width is not None:
        shape.line.width = Pt(line_width)


def _set_slide_dimensions(prs, width, height, theme=None):
    """In-memory: set presentation slide dimensions."""
    prs.slide_width = Inches(width)
    prs.slide_height = Inches(height)


# ── File-based public wrappers ──────────────────────────────────


def format_shape(
    file_path,
    slide_index,
    shape_index,
    left=None,
    top=None,
    width=None,
    height=None,
    fill_color=None,
    no_fill=False,
    line_color=None,
    line_width=None,
    no_line=False,
    rotation=None,
):
    """File-based wrapper: reposition, resize, or restyle a shape."""
    prs = open_pptx(file_path)
    slide = _get_slide(prs, slide_index)
    _format_shape(
        slide, shape_index,
        left, top, width, height,
        fill_color, no_fill, line_color, line_width, no_line, rotation,
    )
    save_pptx(prs, file_path)
    return f"Formatted shape [{shape_index}] on slide [{slide_index}]"


def set_slide_dimensions(file_path, width, height):
    """File-based wrapper: set slide dimensions."""
    prs = open_pptx(file_path)
    _set_slide_dimensions(prs, width, height)
    save_pptx(prs, file_path)
    return f"Set slide dimensions to {width}\" x {height}\""
