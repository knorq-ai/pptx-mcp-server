#!/usr/bin/env python3
"""
pptx-mcp-server -- MCP server for PPTX presentation editing.
"""

import json

from mcp.server.fastmcp import FastMCP

from .engine import (
    EngineError,
    create_presentation,
    get_presentation_info,
    read_slide,
    add_slide,
    move_slide,
    delete_slide,
    duplicate_slide,
    set_slide_background,
    add_textbox,
    add_auto_fit_textbox_file,
    add_shape,
    add_image,
    edit_text,
    add_paragraph,
    delete_shape,
    list_shapes,
    add_table,
    edit_table_cell,
    edit_table_cells,
    format_table,
    format_shape,
    set_slide_dimensions,
    add_content_slide,
    add_section_divider,
    add_kpi_row,
    add_bullet_block,
    build_slide,
    build_deck,
    render_slide,
    add_chart,
    add_icon,
    list_icons_formatted,
    add_connector,
    add_callout,
    check_deck_overlaps,
    check_deck_extended,
)

INSTRUCTIONS = """
# pptx-editor — Professional Presentation Builder

## IMPORTANT: Before Building Any Deck
**Always ask the user which color palette to use before creating slides.** Present these options:
1. **McKinsey** (default) — Dark navy `#051C2C` + Bright blue `#2251FF`. Professional, authoritative.
2. **Deloitte** — Navy `#002776` + Green `#81BC00`. Corporate, multi-color palette with blues and greens.
3. **Neutral** — Dark gray `#333333` + Soft blue `#4A90D9`. Clean, universally safe.
4. **Custom** — Ask the user for their brand primary color (hex) and accent color (hex). You can pass these as overrides.

Once chosen, pass `"theme": "mckinsey"` (or `"deloitte"` / `"neutral"`) in every slide spec.

## Quick Start
1. Ask user for color palette preference (see above)
2. `pptx_create` — create a new 16:9 PPTX file
3. `pptx_build_deck` — build an ENTIRE deck from a JSON spec (most efficient)
4. `pptx_render_slide` — render to PNG for visual verification

## Recommended Workflow
- **Always ask for color palette first** before building any new deck
- Use `pptx_build_deck` for new decks (1 call = all slides, 1 file I/O)
- Use `pptx_build_slide` to add individual slides
- Use `pptx_render_slide` to verify visually (returns PNG path — read with Read tool)
- Use `pptx_check_layout` after building to detect overlaps before delivery
- Use primitive tools (`pptx_edit_text`, `pptx_format_shape`) for fine-grained edits

## McKinsey-Style Layout Rules
- **Slide dimensions**: 16:9 widescreen (13.333" x 7.5")
- **Margins**: 0.9" left/right, content width = 11.533"
- **Action title**: 22pt Arial Bold, max ~30 chars for single line (auto-shrinks if longer)
- **Title zone**: 0.45" to 0.95" (bottom-anchored, divider line at 0.95")
- **Body zone**: 1.15" to 6.5" — USE THE FULL 5.35" HEIGHT. Distribute content evenly.
- **Footer zone**: 6.65" (source line + page number)
- **Font**: Arial for everything (sans-serif works best with Japanese)

## Color Palette (McKinsey-inspired)
- Primary text: `051C2C` (dark navy)
- Accent/highlight: `2251FF` (bright blue)
- Secondary text: `666666`
- Footnote: `A2AAAD`
- Background alt: `F5F5F5`
- Positive: `2E7D32` (green)
- Negative: `C62828` (red)
- Table border: `D0D0D0`

## Table Formatting (McKinsey-style, automatic)
- Dark navy header (`051C2C`) with white text
- No vertical borders, thin horizontal borders only
- Alternating row shading (`F5F5F5`)
- Numbers right-aligned, text left-aligned
- Use `pptx_add_table` or `"type": "table"` in build_slide spec

## Data Density Guidelines (Consulting Quality)
- **Fill every slide**: Body zone (1.15" to 6.5") should be 90%+ utilized. No large blank areas.
- **Use small fonts for detail**: Tables at 9-10pt, bullet text at 10-12pt. Only titles at 22pt.
- **Pack information**: A chart+sidebar slide should have 4+ text bullets AND a table in the sidebar. A table slide should have 8+ rows.
- **Multi-zone layouts**: Combine chart (left 60%) + sidebar (right 40%) with title + bullets + table in the sidebar.
- **Bullet blocks should describe, not list**: Each bullet should be a full sentence with specifics, not a 3-word fragment.
- **For CAGR/growth annotations**: Use a floating `rounded_rectangle` shape badge, NOT a callout with arrow.

## Common Pitfalls to Avoid
1. **Text behind shapes**: Don't put text on a shape if you'll overlay other shapes on top. Use `pptx_add_textbox` for labels over background shapes.
2. **Sparse slides**: Don't leave the bottom half empty. If you have 4 bullets, use the full body height — line spacing auto-distributes.
3. **Long titles**: Keep action titles under 30 chars. The tool auto-shrinks and warns if too long.
4. **Too many tool calls**: Use `pptx_build_deck` (1 call for whole deck) instead of individual `pptx_add_*` calls.
5. **Arrows to nowhere**: Callout arrows must point at a specific data point. For general annotations (CAGR, labels), use a shaped textbox instead.

## Slide Spec Format (for build_slide / build_deck)
```json
{
  "layout": "content",        // "content" | "section_divider" | "blank"
  "title": "Action title",    // for content/section_divider
  "subtitle": "...",           // section_divider only
  "background": "051C2C",     // optional hex color
  "source": "Source: ...",     // optional footnote
  "page_number": 1,           // optional
  "elements": [
    {"type": "textbox", "left": 0.9, "top": 1.2, "width": 11.5, "height": 0.3,
     "text": "...", "font_size": 14, "font_color": "2251FF", "bold": true,
     "alignment": "left", "vertical_anchor": "top"},
    {"type": "shape", "shape_type": "rectangle", ...},
    {"type": "table", "rows": [["H1","H2"],["v1","v2"]], "left": 0.9, "top": 3.0,
     "width": 11.5, "col_widths": [0.5, 0.5]},
    {"type": "kpi_row", "kpis": [{"value":"100","label":"Metric"}], "y": 1.2},
    {"type": "bullet_block", "items": ["Point 1","Point 2"],
     "left": 0.9, "top": 2.0, "width": 11.5, "height": 2.0},
    {"type": "image", "image_path": "/path/img.png", "left": 1.0, "top": 1.0, "width": 3.0},
    {"type": "chart", "chart_type": "stacked_column",
     "left": 0.9, "top": 1.15, "width": 7.0, "height": 5.0,
     "categories": ["2020","2021","2022"],
     "series": [{"name":"Revenue","values":[10,20,30],"color":"2251FF"}],
     "data_labels_show": true, "legend_position": "bottom"}
  ]
}
```

## Chart Element (for build_slide / pptx_add_chart)
chart_type: bar, stacked_bar, stacked_bar_100, column, stacked_column, stacked_column_100, line, line_markers, pie, area, area_stacked, doughnut, radar.

Key fields (all flat — no nesting):
- `categories`: ["A","B","C"] — category labels
- `series`: [{"name":"S1","values":[1,2,3],"color":"2251FF"}] — data series (color optional, auto-assigned from theme)
- `data_labels_show`: true/false — show values on chart
- `data_labels_position`: center, outside_end, inside_end, inside_base, above, below
- `data_labels_number_format`: "#,##0", "0.0%", etc.
- `legend_position`: bottom, top, right, left, null (hidden)
- `axis_value_title`: "Revenue (M)" — Y-axis label
- `axis_value_min/max/major_unit`: scale control
- `axis_value_gridlines`: true/false
- `gap_width`: 150 (bar gap), `overlap`: 100 (stacked)

## Icon Element (for build_slide / pptx_add_icon)
640 built-in vector icons. Use `pptx_list_icons` to browse, or use directly:
```json
{"type": "icon", "icon_id": "briefcase", "left": 2.0, "top": 3.0, "width": 0.8,
 "color": "2251FF", "outline_color": "051C2C"}
```
Common icons: briefcase, chart, person, globe, airplane, laptop, phone, car, building, arrow,
calendar, clock, document, email, gear, handshake, key, lock, money, star, target, trophy.

## Card Grid Element (for build_slide) — auto-balanced layout
Use `card_grid` for 2×2 frameworks, feature grids, strategy cards, etc. Cards auto-size to fill the body zone.
```json
{"type": "card_grid", "cards": [
  {"title": "Strategy 1", "bullets": ["Point A", "Point B"], "icon_id": "target", "icon_color": "2251FF"},
  {"title": "Strategy 2", "body": "Description text", "icon_id": "globe"}
]}
```
Each card: title (required), body or bullets, optional icon_id + icon_color.
Grid auto-computes: 1-2 cards → 1 row, 3-4 → 2×2, 5-6 → 2×3, 7-9 → 3×3.

## Connector Element (for build_slide / pptx_add_connector)
```json
{"type": "connector", "begin_x": 2.0, "begin_y": 3.5, "end_x": 6.0, "end_y": 5.0,
 "color": "accent", "end_arrow": "triangle", "dash_style": "dash"}
```
connector_type: straight (default), elbow, curve.
Arrows: none, triangle, stealth, diamond, oval, open.
dash_style: solid, dash, dot, dash_dot, long_dash.

## Callout Element (for build_slide / pptx_add_callout)
Annotation textbox + arrow pointing to target. Auto-places label if label_x/label_y omitted.
```json
{"type": "callout", "text": "+15% YoY", "target_x": 5.5, "target_y": 3.0,
 "font_color": "negative", "bg_color": "F5F5F5", "arrow_end": "stealth"}
```

## Themes
Set per-slide via `"theme": "mckinsey"` in spec. Available: mckinsey (default), deloitte, neutral.

## Rendering (Optional)
`pptx_render_slide` requires LibreOffice. Install:
- macOS: `brew install --cask libreoffice`
- Ubuntu/Debian: `sudo apt install libreoffice`
- Windows: Download from https://www.libreoffice.org/download/
"""

mcp = FastMCP("pptx-editor", instructions=INSTRUCTIONS)


def _auto_render(file_path: str, slide_index: int) -> str:
    """Try to render a slide preview. Returns preview info or empty string if unavailable."""
    try:
        png = render_slide(file_path, slide_index=slide_index, dpi=100)
        # render_slide may return multiple lines; take the last one (target slide)
        lines = png.strip().split("\n")
        return f"\n📸 Preview: {lines[-1]} (open with Read tool to verify visually)"
    except Exception:
        return ""


def _err(e: Exception) -> str:
    if isinstance(e, EngineError):
        return f"[{e.code.value}] {e}"
    return f"[INTERNAL_ERROR] {e}"


# --- Presentation ------------------------------------------------

@mcp.tool()
def pptx_create(
    file_path: str,
    width_inches: float = 13.333,
    height_inches: float = 7.5,
) -> str:
    """Create a new blank PPTX file. Default is 16:9 widescreen."""
    try:
        return create_presentation(file_path, width_inches, height_inches)
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_get_info(file_path: str) -> str:
    """Get presentation overview: slide count, dimensions, shape summaries."""
    try:
        return get_presentation_info(file_path)
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_read_slide(file_path: str, slide_index: int) -> str:
    """Read detailed content of a slide -- all shapes, text, tables."""
    try:
        return read_slide(file_path, slide_index)
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_list_shapes(file_path: str, slide_index: int) -> str:
    """List all shapes on a slide with indices, types, positions, text preview."""
    try:
        return list_shapes(file_path, slide_index)
    except Exception as e:
        return _err(e)


# --- Slides -------------------------------------------------------

@mcp.tool()
def pptx_add_slide(file_path: str, layout_index: int = 6) -> str:
    """Add a new slide. Layout 6 = Blank (most common)."""
    try:
        return add_slide(file_path, layout_index)
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_move_slide(file_path: str, from_index: int, to_index: int) -> str:
    """Move a slide from one position to another. 0-based indices."""
    try:
        return move_slide(file_path, from_index, to_index)
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_delete_slide(file_path: str, slide_index: int) -> str:
    """Delete a slide by 0-based index."""
    try:
        return delete_slide(file_path, slide_index)
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_duplicate_slide(file_path: str, slide_index: int) -> str:
    """Duplicate a slide (appended at end)."""
    try:
        return duplicate_slide(file_path, slide_index)
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_set_slide_background(file_path: str, slide_index: int, color: str) -> str:
    """Set solid background color for a slide. Color as hex e.g. '051C2C' (without #)."""
    try:
        return set_slide_background(file_path, slide_index, color)
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_set_dimensions(file_path: str, width: float, height: float) -> str:
    """Set presentation slide dimensions in inches (applies to all slides)."""
    try:
        return set_slide_dimensions(file_path, width, height)
    except Exception as e:
        return _err(e)


# --- Textboxes ----------------------------------------------------

@mcp.tool()
def pptx_add_textbox(
    file_path: str,
    slide_index: int,
    left: float,
    top: float,
    width: float,
    height: float,
    text: str = "",
    font_name: str = None,
    font_size: float = None,
    font_color: str = None,
    bold: bool = None,
    italic: bool = None,
    alignment: str = None,
    vertical_anchor: str = None,
    word_wrap: bool = True,
    line_spacing: float = None,
    underline: bool = None,
) -> str:
    """Add a text box to a slide. Position and size in inches. Alignment: left/center/right. Vertical anchor: top/middle/bottom."""
    try:
        return add_textbox(
            file_path, slide_index, left, top, width, height, text,
            font_name, font_size, font_color, bold, italic,
            alignment, vertical_anchor, word_wrap, line_spacing, underline,
        )
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_add_auto_fit_textbox(
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
) -> str:
    """Add a textbox that auto-shrinks font size to fit a fixed box. Starts from font_size_pt and steps down 0.5pt until text fits height, or reaches min_size_pt. If still overflowing at min and truncate_with_ellipsis=True, trailing chars are replaced with an ellipsis. Returns a JSON object with shape_index, shape_name, and actual_font_size."""
    try:
        result = add_auto_fit_textbox_file(
            file_path, slide_index, text, left, top, width, height,
            font_name=font_name,
            font_size_pt=font_size_pt,
            min_size_pt=min_size_pt,
            bold=bold,
            color_hex=color_hex,
            align=align,
            vertical_anchor=vertical_anchor,
            truncate_with_ellipsis=truncate_with_ellipsis,
        )
        return json.dumps(result)
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_edit_text(
    file_path: str,
    slide_index: int,
    shape_index: int,
    text: str = None,
    paragraph_index: int = 0,
    font_name: str = None,
    font_size: float = None,
    font_color: str = None,
    bold: bool = None,
    italic: bool = None,
    underline: bool = None,
    alignment: str = None,
    line_spacing: float = None,
) -> str:
    """Edit text content and formatting in an existing shape's paragraph. Supports all formatting: font, color, bold, italic, underline, alignment, line spacing."""
    try:
        return edit_text(
            file_path, slide_index, shape_index, text, paragraph_index,
            font_name, font_size, font_color, bold, italic, underline,
            alignment, line_spacing,
        )
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_add_paragraph(
    file_path: str,
    slide_index: int,
    shape_index: int,
    text: str,
    font_name: str = None,
    font_size: float = None,
    font_color: str = None,
    bold: bool = None,
    italic: bool = None,
    underline: bool = None,
    alignment: str = None,
    line_spacing: float = None,
) -> str:
    """Append a new paragraph to an existing shape's text frame. Useful for multi-line text."""
    try:
        return add_paragraph(
            file_path, slide_index, shape_index, text,
            font_name, font_size, font_color, bold, italic, underline,
            alignment, line_spacing,
        )
    except Exception as e:
        return _err(e)


# --- Shapes -------------------------------------------------------

@mcp.tool()
def pptx_add_shape(
    file_path: str,
    slide_index: int,
    shape_type: str,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_color: str = None,
    line_color: str = None,
    line_width: float = None,
    no_line: bool = False,
    text: str = None,
    font_name: str = None,
    font_size: float = None,
    font_color: str = None,
    bold: bool = None,
    alignment: str = None,
) -> str:
    """Add an auto shape. Types: rectangle, rounded_rectangle, oval, triangle, diamond, chevron, arrow_right, arrow_left, arrow_up, arrow_down, callout, star_5, hexagon, pentagon. Position/size in inches. Colors as hex. WARNING: text inside shapes renders BEHIND any shapes placed on top. For labels over background shapes, use pptx_add_textbox instead."""
    try:
        return add_shape(
            file_path, slide_index, shape_type, left, top, width, height,
            fill_color, line_color, line_width, no_line,
            text, font_name, font_size, font_color, bold, alignment,
        )
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_add_image(
    file_path: str,
    slide_index: int,
    image_path: str,
    left: float,
    top: float,
    width: float = None,
    height: float = None,
) -> str:
    """Add an image (PNG, JPG, SVG) to a slide. Position in inches. If only width or height is given, aspect ratio is preserved. If both given, image stretches to fit."""
    try:
        return add_image(file_path, slide_index, image_path, left, top, width, height)
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_delete_shape(file_path: str, slide_index: int, shape_index: int) -> str:
    """Delete a shape from a slide by its 0-based index."""
    try:
        return delete_shape(file_path, slide_index, shape_index)
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_format_shape(
    file_path: str,
    slide_index: int,
    shape_index: int,
    left: float = None,
    top: float = None,
    width: float = None,
    height: float = None,
    fill_color: str = None,
    no_fill: bool = False,
    line_color: str = None,
    line_width: float = None,
    no_line: bool = False,
    rotation: float = None,
) -> str:
    """Reposition, resize, or restyle an existing shape. Dimensions in inches."""
    try:
        return format_shape(
            file_path, slide_index, shape_index,
            left, top, width, height,
            fill_color, no_fill, line_color, line_width, no_line, rotation,
        )
    except Exception as e:
        return _err(e)


# --- Tables -------------------------------------------------------

@mcp.tool()
def pptx_add_table(
    file_path: str,
    slide_index: int,
    rows_json: str,
    left: float,
    top: float,
    width: float,
    col_widths_json: str = "",
    row_height: float = 0.36,
    font_size: float = 12,
    header_bg: str = "051C2C",
    header_fg: str = "FFFFFF",
    alt_row_bg: str = "F5F5F5",
    border_color: str = "D0D0D0",
    no_vertical_borders: bool = True,
) -> str:
    """Add a professionally formatted table. rows_json: JSON 2D array e.g. '[["A","B"],["1","2"]]'. First row = header. col_widths_json: JSON array of fractions e.g. '[0.5, 0.5]'."""
    try:
        rows = json.loads(rows_json)
        col_widths = json.loads(col_widths_json) if col_widths_json else None
        return add_table(
            file_path, slide_index, rows, left, top, width,
            col_widths, row_height, font_size,
            header_bg, header_fg, alt_row_bg, border_color,
            0.5, no_vertical_borders,
        )
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_edit_table_cell(
    file_path: str,
    slide_index: int,
    shape_index: int,
    row: int,
    col: int,
    text: str = None,
    font_size: float = None,
    font_color: str = None,
    bold: bool = None,
    bg_color: str = None,
    alignment: str = None,
) -> str:
    """Edit a single table cell's text and formatting."""
    try:
        return edit_table_cell(
            file_path, slide_index, shape_index, row, col,
            text, font_size, font_color, bold, bg_color, alignment,
        )
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_edit_table_cells(
    file_path: str,
    slide_index: int,
    shape_index: int,
    edits_json: str,
) -> str:
    """Batch edit multiple table cells. edits_json: JSON array of objects e.g. '[{"row":0,"col":1,"text":"new"}]'. Each: {row, col, text?, font_size?, font_color?, bold?, bg_color?}."""
    try:
        edits = json.loads(edits_json)
        return edit_table_cells(file_path, slide_index, shape_index, edits)
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_format_table(
    file_path: str,
    slide_index: int,
    shape_index: int,
    font_name: str = None,
    font_size: float = None,
    header_bg: str = None,
    header_fg: str = None,
    alt_row_bg: str = None,
) -> str:
    """Apply bulk formatting to an entire table (font, header colors, alternating rows)."""
    try:
        return format_table(
            file_path, slide_index, shape_index,
            font_name, font_size, header_bg, header_fg, alt_row_bg,
        )
    except Exception as e:
        return _err(e)


# --- Composites ---------------------------------------------------

@mcp.tool()
def pptx_add_content_slide(
    file_path: str,
    title: str,
    source: str = None,
    page_number: int = None,
) -> str:
    """Add a content slide with action title (auto-shrink to fit), divider line, optional source footnote and page number. McKinsey-style layout. Auto-renders a preview PNG. LAYOUT GUIDE: Body area is 1.15\" to 6.65\" (5.5\" usable height). Distribute content evenly across this range — avoid clustering content in the top half with empty bottom space."""
    try:
        result = add_content_slide(file_path, title, source, page_number)
        # Extract slide index from result like "Added content slide [0]..."
        idx = int(result.split("[")[1].split("]")[0])
        result += _auto_render(file_path, idx)
        return result
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_add_section_divider(
    file_path: str,
    title: str,
    subtitle: str = "",
) -> str:
    """Add a section divider slide with dark background, centered title, and accent stripes. Auto-renders a preview PNG."""
    try:
        result = add_section_divider(file_path, title, subtitle)
        idx = int(result.split("[")[1].split("]")[0])
        result += _auto_render(file_path, idx)
        return result
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_add_kpi_row(
    file_path: str,
    slide_index: int,
    kpis_json: str,
    y: float,
) -> str:
    """Add a row of KPI callout boxes. kpis_json: JSON array e.g. '[{"value":"107.8M","label":"Revenue"}]'. y = vertical position in inches. Auto-renders a preview PNG."""
    try:
        result = add_kpi_row(file_path, slide_index, kpis_json, y)
        result += _auto_render(file_path, slide_index)
        return result
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_add_bullet_block(
    file_path: str,
    slide_index: int,
    items_json: str,
    left: float,
    top: float,
    width: float,
    height: float,
) -> str:
    """Add a bulleted text block with multiple items. items_json: JSON array of strings e.g. '["Item 1","Item 2"]'. Auto-renders a preview PNG."""
    try:
        result = add_bullet_block(file_path, slide_index, items_json, left, top, width, height)
        result += _auto_render(file_path, slide_index)
        return result
    except Exception as e:
        return _err(e)


# --- Batch Build --------------------------------------------------

@mcp.tool()
def pptx_build_slide(
    file_path: str,
    spec_json: str,
) -> str:
    """Build an entire slide in ONE call from a JSON spec. Single file open/save. Much faster than individual tool calls.

    spec_json format:
    {
        "layout": "content" | "section_divider" | "blank",
        "title": "Action title",
        "background": "051C2C",  (optional)
        "source": "Source: ...", (optional)
        "page_number": 1,       (optional)
        "elements": [
            {"type": "textbox", "left": 0.9, "top": 1.2, "width": 11.5, "height": 0.3,
             "text": "...", "font_size": 14, "font_color": "2251FF", "bold": true},
            {"type": "shape", "shape_type": "rectangle", "left": 0.9, "top": 2.0,
             "width": 5.0, "height": 1.0, "fill_color": "F5F5F5", "no_line": true},
            {"type": "table", "rows": [["H1","H2"],["v1","v2"]], "left": 0.9,
             "top": 3.0, "width": 11.5, "col_widths": [0.5, 0.5]},
            {"type": "kpi_row", "kpis": [{"value":"100","label":"Metric"}], "y": 1.2},
            {"type": "bullet_block", "items": ["Point 1","Point 2"],
             "left": 0.9, "top": 2.0, "width": 11.5, "height": 2.0},
            {"type": "image", "image_path": "/path/img.png", "left": 1.0, "top": 1.0, "width": 3.0}
        ]
    }
    LAYOUT GUIDE: Body area is 1.15" to 6.65" (5.5" usable). Distribute elements evenly."""
    try:
        result = build_slide(file_path, spec_json)
        # Auto-render
        import json
        spec = json.loads(spec_json) if isinstance(spec_json, str) else spec_json
        idx_str = result.split("[")[1].split("]")[0]
        result += _auto_render(file_path, int(idx_str))
        return result
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_build_deck(
    file_path: str,
    slides_json: str,
) -> str:
    """Build an ENTIRE DECK in ONE call from a JSON array of slide specs. Single file open/save for all slides. Use this for generating complete presentations efficiently.

    slides_json: JSON array where each element is a slide spec (same format as pptx_build_slide).
    Example: '[{"layout":"content","title":"Slide 1","elements":[...]},{"layout":"section_divider","title":"Section"}]'"""
    try:
        result = build_deck(file_path, slides_json)
        result += _auto_render(file_path, -1)
        return result
    except Exception as e:
        return _err(e)


# --- Connectors & Callouts ----------------------------------------

@mcp.tool()
def pptx_add_connector(
    file_path: str,
    slide_index: int,
    begin_x: float,
    begin_y: float,
    end_x: float,
    end_y: float,
    connector_type: str = "straight",
    color: str = None,
    width: float = None,
    dash_style: str = None,
    begin_arrow: str = "none",
    end_arrow: str = "triangle",
    arrow_size: str = "medium",
) -> str:
    """Add a connector line between two points. Position in inches.
    connector_type: straight/elbow/curve.
    Arrows: none/triangle/stealth/diamond/oval/open.
    dash_style: solid/dash/dot/dash_dot/long_dash."""
    try:
        result = add_connector(
            file_path, slide_index, begin_x, begin_y, end_x, end_y,
            connector_type, color, width, dash_style,
            begin_arrow, end_arrow, arrow_size,
        )
        result += _auto_render(file_path, slide_index)
        return result
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_add_callout(
    file_path: str,
    slide_index: int,
    text: str,
    target_x: float,
    target_y: float,
    label_x: float = None,
    label_y: float = None,
    label_width: float = 2.0,
    label_height: float = 0.4,
    connector_type: str = "straight",
    font_size: float = 10,
    font_color: str = None,
    font_bold: bool = True,
    line_color: str = None,
    line_width: float = 1.0,
    arrow_end: str = "triangle",
    bg_color: str = None,
    border_color: str = None,
) -> str:
    """Add a callout annotation: textbox + connector arrow pointing to target.
    Auto-places label if label_x/label_y omitted. Position in inches."""
    try:
        result = add_callout(
            file_path, slide_index, text, target_x, target_y,
            label_x, label_y, label_width, label_height,
            connector_type, font_size, font_color, font_bold,
            line_color, line_width, arrow_end, bg_color, border_color,
        )
        result += _auto_render(file_path, slide_index)
        return result
    except Exception as e:
        return _err(e)


# --- Icons --------------------------------------------------------

@mcp.tool()
def pptx_list_icons(
    category: str = "",
    search: str = "",
) -> str:
    """List available icons from the built-in icon library (640 vector icons).
    Filter by category and/or keyword search.

    Categories: business, people, technology, transport, medical, education, nature, objects, general.
    Example: pptx_list_icons(category="business") or pptx_list_icons(search="chart")

    Common icons: briefcase, chart, person, globe, airplane, laptop, phone, car, building, arrow,
    calendar, clock, document, email, gear, handshake, key, lock, money, star, target, trophy, user."""
    try:
        return list_icons_formatted(
            category=category or None,
            search=search or None,
        )
    except Exception as e:
        return _err(e)


@mcp.tool()
def pptx_add_icon(
    file_path: str,
    slide_index: int,
    icon_id: str,
    left: float,
    top: float,
    width: float = None,
    height: float = None,
    color: str = None,
    outline_color: str = None,
) -> str:
    """Add a vector icon from the built-in library to a slide.
    Position in inches. If only width or height given, aspect ratio is preserved.
    Colors as hex (e.g. '2251FF') or theme token ('accent', 'primary').
    Use pptx_list_icons to browse available icons."""
    try:
        result = add_icon(file_path, slide_index, icon_id, left, top, width, height, color, outline_color)
        result += _auto_render(file_path, slide_index)
        return result
    except Exception as e:
        return _err(e)


# --- Charts -------------------------------------------------------

@mcp.tool()
def pptx_add_chart(
    file_path: str,
    slide_index: int,
    chart_json: str,
) -> str:
    """Add a professional chart to a slide. chart_json is a JSON object with:

    Required: chart_type (column/stacked_column/bar/stacked_bar/line/pie/area/doughnut/radar),
    categories (array of labels), series (array of {name, values, color?}).

    Optional: title, legend_position (bottom/top/right/left/null), data_labels_show (bool),
    data_labels_position, data_labels_number_format, axis_value_title, axis_value_min/max,
    axis_value_gridlines, gap_width, overlap, theme (mckinsey/deloitte/neutral).

    Example: '{"chart_type":"stacked_column","categories":["Q1","Q2"],"series":[{"name":"Rev","values":[10,20],"color":"2251FF"}],"data_labels_show":true}'"""
    try:
        spec = json.loads(chart_json) if isinstance(chart_json, str) else chart_json
        result = add_chart(file_path, slide_index, spec)
        result += _auto_render(file_path, slide_index)
        return result
    except json.JSONDecodeError as e:
        return f"[INVALID_PARAMETER] Invalid JSON in chart_json: {e}"
    except Exception as e:
        return _err(e)


# --- Rendering ----------------------------------------------------

@mcp.tool()
def pptx_render_slide(
    file_path: str,
    slide_index: int = -1,
    dpi: int = 150,
) -> str:
    """Render PPTX slide(s) to PNG image(s) for visual verification. Returns path(s) to PNG files that can be viewed with the Read tool. slide_index: 0-based (-1 = all slides). dpi: 150 for review, 300 for print."""
    try:
        return render_slide(file_path, slide_index, dpi=dpi)
    except Exception as e:
        return _err(e)


# --- Entry Point --------------------------------------------------

@mcp.tool()
def pptx_check_layout(
    file_path: str,
    min_readable_pt: float = 8.0,
    overflow_tolerance_pct: float = 5.0,
) -> str:
    """Validate slide layouts: overlaps, out-of-bounds, text overflow,
    unreadable font, title/divider collision, inconsistent gaps.

    Returns a JSON string with per-slide findings and a summary block.
    Run after building a deck to catch layout issues before delivery."""
    try:
        from pptx import Presentation
        prs = Presentation(file_path)
        result = check_deck_extended(
            prs,
            min_readable_pt=min_readable_pt,
            overflow_tolerance_pct=overflow_tolerance_pct,
        )
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return _err(e)


def main():
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
