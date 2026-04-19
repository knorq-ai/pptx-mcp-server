"""
Layout validation — detect overlapping shapes, out-of-bounds elements, etc.

Run after build_slide / build_deck to catch layout issues before delivery.
"""

from __future__ import annotations

from typing import List, Tuple

from pptx import Presentation
from pptx.util import Emu


def _get_shape_bounds(shape) -> Tuple[float, float, float, float]:
    """Return (left, top, right, bottom) in inches for a shape."""
    emu_to_in = 1 / 914400
    left = shape.left * emu_to_in
    top = shape.top * emu_to_in
    right = left + shape.width * emu_to_in
    bottom = top + shape.height * emu_to_in
    return (left, top, right, bottom)


def _boxes_overlap(a: Tuple, b: Tuple, margin: float = 0.0) -> bool:
    """Check if two (left, top, right, bottom) boxes overlap.

    margin: minimum gap required between shapes (inches).
    Negative margin means shapes must overlap by that amount to trigger.
    """
    a_left, a_top, a_right, a_bottom = a
    b_left, b_top, b_right, b_bottom = b

    # No overlap if separated on any axis
    if a_right + margin <= b_left:
        return False
    if b_right + margin <= a_left:
        return False
    if a_bottom + margin <= b_top:
        return False
    if b_bottom + margin <= a_top:
        return False
    return True


def _is_small_decorative(shape) -> bool:
    """Check if a shape is small/decorative (divider lines, accent bars)."""
    emu_to_in = 1 / 914400
    w = shape.width * emu_to_in
    h = shape.height * emu_to_in
    # Thin lines (dividers, accent bars)
    if h < 0.08 or w < 0.08:
        return True
    # Very small shapes (icon accent bars in KPI cards)
    if w * h < 0.05:
        return True
    return False


def _is_container(bounds_a: Tuple, bounds_b: Tuple) -> bool:
    """Check if one shape fully contains the other (parent-child relationship).

    Text placed inside a background rectangle is intentional, not an overlap.
    """
    a_left, a_top, a_right, a_bottom = bounds_a
    b_left, b_top, b_right, b_bottom = bounds_b

    # A contains B
    if a_left <= b_left and a_top <= b_top and a_right >= b_right and a_bottom >= b_bottom:
        return True
    # B contains A
    if b_left <= a_left and b_top <= a_top and b_right >= a_right and b_bottom >= a_bottom:
        return True
    return False


def check_slide_overlaps(
    slide,
    margin: float = 0.03,
    slide_width: float = 13.333,
    slide_height: float = 7.5,
) -> List[str]:
    """Check a slide for overlapping shapes and insufficient gaps.

    margin: minimum gap required between shapes (inches).
            0.03 = shapes touching or within 0.03" trigger a warning.

    Returns list of warning strings. Empty = no issues.
    """
    warnings = []
    shapes = list(slide.shapes)

    # Filter out small decorative shapes (divider lines, accent bars)
    content_shapes = [s for s in shapes if not _is_small_decorative(s)]

    # Check pairwise overlaps
    for i in range(len(content_shapes)):
        for j in range(i + 1, len(content_shapes)):
            si = content_shapes[i]
            sj = content_shapes[j]

            bi = _get_shape_bounds(si)
            bj = _get_shape_bounds(sj)

            if _boxes_overlap(bi, bj, margin):
                # Skip if one contains the other (text inside background shape)
                if _is_container(bi, bj):
                    continue

                # Compute actual overlap area
                overlap_x = max(0, min(bi[2], bj[2]) - max(bi[0], bj[0]))
                overlap_y = max(0, min(bi[3], bj[3]) - max(bi[1], bj[1]))
                overlap_area = overlap_x * overlap_y

                # Skip small labels/badges placed on charts (< 1.0 sq in shape)
                area_i = (bi[2] - bi[0]) * (bi[3] - bi[1])
                area_j = (bj[2] - bj[0]) * (bj[3] - bj[1])
                if min(area_i, area_j) < 1.0:
                    continue

                name_i = si.name or f"Shape {i}"
                name_j = sj.name or f"Shape {j}"

                if overlap_area >= 0.15:
                    warnings.append(
                        f"OVERLAP ({overlap_area:.1f} sq in): "
                        f"'{name_i}' at ({bi[0]:.1f},{bi[1]:.1f},{bi[2]:.1f},{bi[3]:.1f}) "
                        f"× '{name_j}' at ({bj[0]:.1f},{bj[1]:.1f},{bj[2]:.1f},{bj[3]:.1f})"
                    )
                else:
                    # Compute minimum gap between edges
                    h_gap = max(0, max(bi[0], bj[0]) - min(bi[2], bj[2]))
                    v_gap = max(0, max(bi[1], bj[1]) - min(bi[3], bj[3]))
                    min_gap = max(h_gap, v_gap)
                    # Both shapes are significant (> 1 sq in each)
                    if area_i > 1.0 and area_j > 1.0:
                        warnings.append(
                            f"TOO CLOSE (gap {min_gap:.2f}\"): "
                            f"'{name_i}' at ({bi[0]:.1f},{bi[1]:.1f},{bi[2]:.1f},{bi[3]:.1f}) "
                            f"× '{name_j}' at ({bj[0]:.1f},{bj[1]:.1f},{bj[2]:.1f},{bj[3]:.1f})"
                        )

    # Check out-of-bounds
    for s in shapes:
        if _is_small_decorative(s):
            continue
        b = _get_shape_bounds(s)
        if b[2] > slide_width + 0.1:
            warnings.append(
                f"OUT OF BOUNDS (right): '{s.name}' extends to {b[2]:.1f}\" "
                f"(slide width={slide_width}\")"
            )
        if b[3] > slide_height + 0.1:
            warnings.append(
                f"OUT OF BOUNDS (bottom): '{s.name}' extends to {b[3]:.1f}\" "
                f"(slide height={slide_height}\")"
            )

    return warnings


def check_deck_overlaps(pptx_path: str) -> str:
    """Check all slides in a deck for layout issues.

    Returns a report string. Empty = all clean.
    """
    prs = Presentation(pptx_path)
    slide_w = prs.slide_width / 914400
    slide_h = prs.slide_height / 914400

    all_warnings = []
    for i, slide in enumerate(prs.slides):
        warnings = check_slide_overlaps(
            slide, margin=-0.05,
            slide_width=slide_w, slide_height=slide_h,
        )
        for w in warnings:
            all_warnings.append(f"Slide {i + 1}: {w}")

    if not all_warnings:
        return "All slides clean — no overlaps or out-of-bounds detected."

    return f"Found {len(all_warnings)} layout issues:\n" + "\n".join(all_warnings)
