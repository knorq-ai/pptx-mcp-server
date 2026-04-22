"""Container primitive — declare a bounded rectangle for child shapes.

``begin_container`` is a context manager that declares a logical bounding box
on a slide. Shapes added by atomic primitives (``_add_shape``, ``_add_textbox``,
``_add_image``, ``add_auto_fit_textbox``) inside the ``with`` block are auto-
registered against the innermost active container via a thread-local stack.
The registry is validator-time metadata only — it is **not** serialized into
the PPTX file (see Issue #130).

Two data structures back this feature:

1. ``_CONTAINER_STACK`` — a ``threading.local`` list used during ``with``
   blocks to route newly created shapes to the innermost open container.
2. ``_SLIDE_REGISTRY`` — a module-level ``dict`` keyed by ``id(slide)`` that
   persists container declarations beyond the ``with`` block so the
   validator (``check_containment``) can inspect them later.

The validator key is ``id(slide)`` which is stable for the lifetime of the
Python slide object (that is all we need; containment is checked before the
slide is garbage collected). For shape identity we store the shape's XML
``name`` plus its bbox at tag time — python-pptx rebuilds shape wrappers on
each access, so ``id(shape)`` is unreliable, but the name survives round-
trips within a single in-memory session.
"""

from __future__ import annotations

import threading
from contextlib import contextmanager
from dataclasses import dataclass, field
from typing import Dict, Iterator, List, Optional, Tuple


@dataclass
class ShapeRef:
    """Lightweight handle to a shape tagged inside a container.

    Holds the shape's ``name`` (as assigned by python-pptx on creation) plus
    its bbox at tag time. ``check_containment`` re-resolves the live shape
    by walking ``slide.shapes`` and matching on ``name`` — falling back to
    the cached bbox if the name cannot be found.
    """

    name: str
    left: float
    top: float
    width: float
    height: float


@dataclass
class ContainerBounds:
    """Declared bounding rectangle for a logical component on a slide.

    Attributes:
        name: Semantic label (e.g. ``"metric_card_0"``). Used in validator
            findings so the author can quickly locate the offending
            container in their code.
        left, top, width, height: Outer bbox in inches.
        padding: Inset (inches) subtracted equally from each side when
            computing the effective bounds children must fit inside.
            Default 0 (bounds == outer bbox).
        children: Shapes registered while this container was the innermost
            active entry on the thread-local stack. Populated by
            ``_register_with_active_container``.
    """

    name: str
    left: float
    top: float
    width: float
    height: float
    padding: float = 0.0
    children: List[ShapeRef] = field(default_factory=list)

    def inner_bounds(self) -> Tuple[float, float, float, float]:
        """Return ``(left, top, right, bottom)`` after applying ``padding``."""
        pad = self.padding
        return (
            self.left + pad,
            self.top + pad,
            self.left + self.width - pad,
            self.top + self.height - pad,
        )


# Thread-local stack of (slide_id, ContainerBounds) tuples. Innermost container
# is at the end of the list. Only the innermost matching-slide entry receives
# the new shape on registration.
_thread_local = threading.local()


def _get_stack() -> List[Tuple[int, ContainerBounds]]:
    """Return this thread's container stack, lazily initializing it."""
    stack = getattr(_thread_local, "stack", None)
    if stack is None:
        stack = []
        _thread_local.stack = stack
    return stack


# Persistent per-slide registry: ``id(slide) -> list[ContainerBounds]``.
# Entries live until ``clear_container_registry`` is called (tests) or the
# slide object is replaced. This is validator-time metadata — never
# serialized into the PPTX XML.
_SLIDE_REGISTRY: Dict[int, List[ContainerBounds]] = {}


def _get_slide_containers(slide) -> List[ContainerBounds]:
    """Return the list of declared containers for ``slide`` (may be empty)."""
    return _SLIDE_REGISTRY.get(id(slide), [])


def iter_slide_containers(slide) -> Iterator[ContainerBounds]:
    """Iterate declared containers for ``slide`` (validator consumption)."""
    for bounds in _get_slide_containers(slide):
        yield bounds


def clear_container_registry() -> None:
    """Drop all registered containers. Primarily for test isolation."""
    _SLIDE_REGISTRY.clear()
    # Also clear the thread-local stack for safety; tests that interleave
    # with un-matched begin_container/exit flows should not leak state.
    stack = getattr(_thread_local, "stack", None)
    if stack is not None:
        stack.clear()


def clear_slide_containers(slide) -> None:
    """Remove any container entries registered against this slide.

    Called at the end of a validation pass so long-running processes
    don't leak memory, and so ``id(slide)`` reuse can't bleed stale
    bounds into a freshly-loaded slide.
    """
    _SLIDE_REGISTRY.pop(id(slide), None)


@contextmanager
def begin_container(
    slide,
    *,
    name: str,
    left: float,
    top: float,
    width: float,
    height: float,
    padding: float = 0.0,
):
    """Declare a bounded container on ``slide``.

    Children added while this context is active are tagged against this
    container via a thread-local stack and — after the block exits — can be
    inspected by ``check_containment`` to flag any child whose bbox escapes
    the declared bounds.

    Nested containers are supported: a ``begin_container`` call inside
    another ``begin_container`` pushes onto the same thread-local stack, and
    newly added shapes are registered against the innermost entry only.

    Args:
        slide: The python-pptx slide object children will be added to.
        name: Semantic label for validator messages.
        left, top, width, height: Outer bbox in inches.
        padding: Inner inset in inches (default 0). ``check_containment``
            enforces ``child.bbox ⊆ (outer shrunk by padding)``.

    Yields:
        The ``ContainerBounds`` instance. Callers may use it to read back
        the bounds they declared (e.g. for relative-positioning math).
    """
    bounds = ContainerBounds(
        name=name,
        left=float(left),
        top=float(top),
        width=float(width),
        height=float(height),
        padding=float(padding),
    )

    # Persistent registry for the validator.
    _SLIDE_REGISTRY.setdefault(id(slide), []).append(bounds)

    # Thread-local stack for auto-tagging new shapes.
    stack = _get_stack()
    stack.append((id(slide), bounds))
    try:
        yield bounds
    finally:
        # Pop our entry. Find from the top in case an inner block forgot to
        # pop (defensive; normally stack top IS our entry).
        for i in range(len(stack) - 1, -1, -1):
            if stack[i][1] is bounds:
                del stack[i]
                break


def _register_with_active_container(
    slide,
    shape,
    left: float,
    top: float,
    width: float,
    height: float,
) -> None:
    """If a matching container is active on this thread, register ``shape``.

    Called by atomic primitives (``_add_shape``, ``_add_textbox``,
    ``_add_image``, ``add_auto_fit_textbox``) after the shape has been
    created. Silently no-ops if:

    - No container is active on this thread.
    - The innermost container's slide does not match (e.g. the caller
      switched slides inside the ``with`` block).
    - ``shape`` has no ``.name`` attribute (e.g. a stubbed slide used in
      unit tests that pass a mock — don't crash).

    Only the innermost matching container receives the ShapeRef. Outer
    containers implicitly "see" the shape too via the nesting relationship
    (inner bounds are expected to lie inside outer bounds); this matches
    how CSS / DOM containment is typically reasoned about and keeps the
    registration cheap.
    """
    stack = _get_stack()
    if not stack:
        return
    slide_id = id(slide)
    # Innermost matching slide only.
    for i in range(len(stack) - 1, -1, -1):
        entry_slide_id, bounds = stack[i]
        if entry_slide_id == slide_id:
            shape_name = _safe_shape_name(shape)
            bounds.children.append(
                ShapeRef(
                    name=shape_name,
                    left=float(left),
                    top=float(top),
                    width=float(width),
                    height=float(height),
                )
            )
            return


def _safe_shape_name(shape) -> str:
    """Best-effort extraction of ``shape.name`` for registry lookups."""
    try:
        n = shape.name
    except Exception:  # pragma: no cover - defensive
        n = None
    if not n:
        return ""
    return str(n)
