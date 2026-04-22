"""Shared helpers for v0.6.0+ block components."""

from __future__ import annotations

# Canonical fallback palette used when no ``theme`` is supplied.
# Mirrors the MCKINSEY theme for the 4 tokens that block components use
# as defaults. Keep in sync with engine.theme.MCKINSEY.
_FALLBACK_TOKENS: dict[str, str] = {
    "primary": "051C2C",
    "text_secondary": "666666",
    "rule_subtle": "E0E0E0",
    "highlight_row": "F8F9F5",
}


def resolve_component_color(token_or_hex: str, theme: str | None) -> str:
    """Resolve a component color token to a raw 6-hex string.

    Behavior:
    - Empty / None input → return as-is (empty means "disable fill" etc.)
    - If ``theme`` is a non-empty string → delegate to ``resolve_theme_color``
      (centralised engine/theme.py helper; same behavior as v0.5.0
      atomic primitives).
    - Else (theme is None): look up in ``_FALLBACK_TOKENS``; if not a known
      token, strip any leading ``#`` and return the input unchanged (treat
      as raw hex; downstream ``_parse_color`` will validate).

    This replaces 5 per-component ``_resolve_color`` helpers that had crept
    into subtle behavior drift. All block components must use this helper.
    """
    if not token_or_hex:
        return token_or_hex
    if theme:
        from ...theme import resolve_theme_color
        return resolve_theme_color(token_or_hex, theme)
    token = _FALLBACK_TOKENS.get(token_or_hex)
    if token is not None:
        return token
    return token_or_hex.lstrip("#")
