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
from .metric_card import (
    MetricEntry,
    MetricCardSpec,
    add_metric_card,
    add_metric_card_row,
)

__all__ = [
    "ContainerBounds",
    "begin_container",
    "clear_slide_containers",
    "MetricEntry",
    "MetricCardSpec",
    "add_metric_card",
    "add_metric_card_row",
]
