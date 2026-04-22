from .layout_constants import (
    TEXTBOX_INNER_PADDING_PER_SIDE,
    TEXTBOX_INNER_PADDING_TOTAL,
)
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
    add_auto_fit_textbox,
    add_auto_fit_textbox_file,
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
    check_deck_extended,
    check_text_overflow,
    check_unreadable_text,
    check_divider_collision,
    check_inconsistent_gaps,
    check_containment,
    ValidationFinding,
)
from .components.container import (
    ContainerBounds,
    begin_container,
)
from .components.markers import (
    PageMarkerSpec,
    SlideFooterSpec,
    add_page_marker,
    add_slide_footer,
)
from .components.section_header import (
    SectionHeaderSpec,
    add_section_header,
)
from .connectors import (
    add_connector,
    add_callout,
    _add_connector,
    _add_callout,
)
from .flex import (
    FlexItem,
    add_flex_container,
    add_flex_container_file,
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
from .cards import (
    CardSpec,
    CardHeightMode,
    CardPlacement,
    add_responsive_card_row,
)
from .tables_grid import (
    TableColumnSpec,
    add_data_table,
)
from .timeline import (
    TimelinePhase,
    TimelineMilestone,
    add_milestone_timeline,
)

__all__ = [
    # Layout constants
    "TEXTBOX_INNER_PADDING_PER_SIDE", "TEXTBOX_INNER_PADDING_TOTAL",
    # I/O
    "open_pptx", "save_pptx", "create_presentation",
    "EngineError", "ErrorCode",
    # Slides
    "get_presentation_info", "read_slide", "add_slide", "move_slide",
    "delete_slide", "duplicate_slide", "set_slide_background",
    # Shapes
    "add_textbox", "add_shape", "add_image", "edit_text", "add_paragraph",
    "delete_shape", "list_shapes",
    "add_auto_fit_textbox", "add_auto_fit_textbox_file",
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
    "check_deck_overlaps", "check_slide_overlaps", "check_deck_extended",
    "check_text_overflow", "check_unreadable_text",
    "check_divider_collision", "check_inconsistent_gaps",
    "check_containment",
    "ValidationFinding",
    # Components
    "ContainerBounds", "begin_container",
    "PageMarkerSpec", "SlideFooterSpec",
    "add_page_marker", "add_slide_footer",
    "SectionHeaderSpec", "add_section_header",
    # Composites
    "add_content_slide", "add_section_divider", "add_kpi_row", "add_bullet_block",
    "build_slide", "build_deck",
    # Flex
    "FlexItem", "add_flex_container", "add_flex_container_file",
    # Cards
    "CardSpec", "CardHeightMode", "CardPlacement", "add_responsive_card_row",
    # Data tables (textbox grid)
    "TableColumnSpec", "add_data_table",
    # Timeline
    "TimelinePhase", "TimelineMilestone", "add_milestone_timeline",
    # Rendering
    "render_slide", "render_slide_to_path",
]
