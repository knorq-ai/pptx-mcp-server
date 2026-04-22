"""Component-layer primitives for pptx-mcp-server (v0.6.0+).

This sub-package hosts higher-level "block" components (cards, containers,
etc.) that compose atomic primitives from the engine into reusable layout
units. The atomic primitive API (``engine.shapes`` etc.) is intentionally
kept unchanged — components live here so v0.5.x users are not affected.
"""

from .container import (
    ContainerBounds,
    begin_container,
    clear_slide_containers,
)
from .markers import (
    PageMarkerSpec,
    SlideFooterSpec,
    add_page_marker,
    add_slide_footer,
)
from .section_header import (
    SectionHeaderSpec,
    add_section_header,
)
from .numbered_list import (
    NumberedItem,
    add_numbered_list,
)
from .kpi_row import (
    KPISpec,
    # Aliased to avoid collision with legacy `engine.composites.add_kpi_row`.
    add_kpi_row as add_kpi_row_block,
)

__all__ = [
    "ContainerBounds",
    "begin_container",
    "clear_slide_containers",
    "PageMarkerSpec",
    "SlideFooterSpec",
    "add_page_marker",
    "add_slide_footer",
    "SectionHeaderSpec",
    "add_section_header",
    "NumberedItem",
    "add_numbered_list",
    "KPISpec",
    "add_kpi_row_block",
]
