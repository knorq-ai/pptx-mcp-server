from .pptx_io import open_pptx, save_pptx, create_presentation, EngineError, ErrorCode
from .slides import (
    get_presentation_info,
    read_slide,
    add_slide,
    move_slide,
    delete_slide,
    duplicate_slide,
    set_slide_background,
    _get_presentation_info,
    _read_slide,
    _add_slide,
    _move_slide,
    _delete_slide,
    _duplicate_slide,
    _set_slide_background,
)
from .shapes import (
    add_textbox,
    add_shape,
    add_image,
    edit_text,
    add_paragraph,
    delete_shape,
    list_shapes,
    _add_textbox,
    _add_shape,
    _add_image,
    _edit_text,
    _add_paragraph,
    _delete_shape,
    _list_shapes,
)
from .tables import (
    add_table,
    edit_table_cell,
    edit_table_cells,
    format_table,
    _add_table,
    _edit_table_cell,
    _edit_table_cells,
    _format_table,
)
from .formatting import (
    format_shape,
    set_slide_dimensions,
    _format_shape,
    _set_slide_dimensions,
)
from .rendering import (
    render_slide,
    render_slide_to_path,
)
from .charts import (
    add_chart,
    _add_chart,
)
from .icons import (
    add_icon,
    list_icons_formatted,
    _add_icon,
)
from .validation import (
    check_deck_overlaps,
    check_slide_overlaps,
)
from .connectors import (
    add_connector,
    add_callout,
    _add_connector,
    _add_callout,
)
from .composites import (
    add_content_slide,
    add_section_divider,
    add_kpi_row,
    add_bullet_block,
    build_slide,
    build_deck,
    _add_content_slide,
    _add_section_divider,
    _add_kpi_row,
    _add_bullet_block,
    _build_slide,
)

__all__ = [
    # I/O
    "open_pptx", "save_pptx", "create_presentation",
    "EngineError", "ErrorCode",
    # Slides
    "get_presentation_info", "read_slide", "add_slide", "move_slide",
    "delete_slide", "duplicate_slide", "set_slide_background",
    # Shapes
    "add_textbox", "add_shape", "add_image", "edit_text", "add_paragraph",
    "delete_shape", "list_shapes",
    # Tables
    "add_table", "edit_table_cell", "edit_table_cells", "format_table",
    # Formatting
    "format_shape", "set_slide_dimensions",
    # Charts
    "add_chart",
    # Icons
    "add_icon", "list_icons_formatted",
    # Connectors
    "add_connector", "add_callout",
    # Validation
    "check_deck_overlaps",
    # Composites
    "add_content_slide", "add_section_divider", "add_kpi_row", "add_bullet_block",
    "build_slide", "build_deck",
    # Rendering
    "render_slide", "render_slide_to_path",
]
