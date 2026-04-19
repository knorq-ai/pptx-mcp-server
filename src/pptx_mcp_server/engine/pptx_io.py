"""
PPTX file I/O — open, save, create, error types.
"""

from __future__ import annotations

import os
from enum import Enum
from pptx import Presentation
from pptx.util import Inches, Pt, Emu


class ErrorCode(str, Enum):
    FILE_NOT_FOUND = "FILE_NOT_FOUND"
    INVALID_PPTX = "INVALID_PPTX"
    SLIDE_NOT_FOUND = "SLIDE_NOT_FOUND"
    SHAPE_NOT_FOUND = "SHAPE_NOT_FOUND"
    INDEX_OUT_OF_RANGE = "INDEX_OUT_OF_RANGE"
    INVALID_PARAMETER = "INVALID_PARAMETER"
    TABLE_ERROR = "TABLE_ERROR"
    CHART_ERROR = "CHART_ERROR"


class EngineError(Exception):
    def __init__(self, code: ErrorCode, message: str):
        super().__init__(message)
        self.code = code


def open_pptx(file_path: str) -> Presentation:
    """Open an existing PPTX file."""
    if not os.path.exists(file_path):
        raise EngineError(ErrorCode.FILE_NOT_FOUND, f"File not found: {file_path}")
    try:
        return Presentation(file_path)
    except Exception as e:
        raise EngineError(ErrorCode.INVALID_PPTX, f"Not a valid PPTX file: {file_path} ({e})")


def save_pptx(prs: Presentation, file_path: str) -> None:
    """Save a presentation to disk."""
    prs.save(file_path)


def create_presentation(
    file_path: str,
    width_inches: float = 13.333,
    height_inches: float = 7.5,
) -> str:
    """Create a new blank PPTX file with specified dimensions."""
    prs = Presentation()
    prs.slide_width = Inches(width_inches)
    prs.slide_height = Inches(height_inches)
    save_pptx(prs, file_path)
    return f"Created presentation: {file_path} ({width_inches}\" x {height_inches}\")"


def _get_slide(prs: Presentation, slide_index: int):
    """Get slide by 0-based index with bounds checking."""
    if slide_index < 0 or slide_index >= len(prs.slides):
        raise EngineError(
            ErrorCode.SLIDE_NOT_FOUND,
            f"Slide index {slide_index} out of range (0-{len(prs.slides) - 1})",
        )
    return prs.slides[slide_index]


def _get_shape(slide, shape_index: int):
    """Get shape by 0-based index with bounds checking."""
    shapes = list(slide.shapes)
    if shape_index < 0 or shape_index >= len(shapes):
        raise EngineError(
            ErrorCode.SHAPE_NOT_FOUND,
            f"Shape index {shape_index} out of range (0-{len(shapes) - 1})",
        )
    return shapes[shape_index]


def _parse_color(color_str: str):
    """Parse hex color string like '#FF0000' or 'FF0000' to RGBColor."""
    from pptx.dml.color import RGBColor
    color_str = color_str.lstrip("#")
    if len(color_str) != 6:
        raise EngineError(ErrorCode.INVALID_PARAMETER, f"Invalid color: #{color_str}")
    try:
        return RGBColor(
            int(color_str[0:2], 16),
            int(color_str[2:4], 16),
            int(color_str[4:6], 16),
        )
    except ValueError:
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            f"Invalid hex color: '#{color_str}'. Use 6-digit hex like 'FF0000'.",
        )
