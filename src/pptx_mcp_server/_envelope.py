"""Structured response envelope helpers for the MCP server.

本モジュールは `server.py` から切り出したヘルパである。公開 API
ではなく、モジュール名に先頭アンダースコアを付けている。

Tool 戻り値はすべて JSON 文字列で、次の包に統一する (issue #88):

- 成功: ``{"ok": true, "result": <legacy value>}``
- 失敗: ``{"ok": false, "error": {"code": ..., "message": ...,
        "parameter"?: ..., "hint"?: ..., "issue"?: ...}}``

``code`` は :class:`pptx_mcp_server.engine.EngineError` の ``ErrorCode``
値と一致する (``INVALID_PARAMETER`` 等)。それ以外の例外は
``INTERNAL_ERROR`` に畳む。
"""

from __future__ import annotations

import json
from typing import Any, Dict, Optional

from .engine import EngineError


def _success(result: Any) -> str:
    """Wrap a successful tool result in ``{"ok": true, "result": ...}``.

    ``result`` は legacy tool の戻り値 (通常は human-readable string) を
    そのまま格納する。JSON で表現できない object は呼び出し側で事前に
    serialize すること。
    """
    return json.dumps({"ok": True, "result": result}, ensure_ascii=False)


def _error(
    code: str,
    message: str,
    *,
    parameter: Optional[str] = None,
    hint: Optional[str] = None,
    issue: Optional[int] = None,
) -> str:
    """Build a structured error payload and return JSON-string.

    Shape::

        {"ok": false, "error": {"code": <str>, "message": <str>,
         "parameter": <optional str>, "hint": <optional str>,
         "issue": <optional int>}}
    """
    err: Dict[str, Any] = {"code": code, "message": message}
    if parameter is not None:
        err["parameter"] = parameter
    if hint is not None:
        err["hint"] = hint
    if issue is not None:
        err["issue"] = issue
    return json.dumps({"ok": False, "error": err}, ensure_ascii=False)


def _err(e: Exception) -> str:
    """Translate an exception into a structured error JSON string.

    ``EngineError`` は ``code`` enum をそのまま error.code として流用する。
    それ以外は ``INTERNAL_ERROR`` に fall back する。
    """
    if isinstance(e, EngineError):
        return _error(e.code.value, str(e))
    return _error("INTERNAL_ERROR", f"{type(e).__name__}: {e}")


def _success_with_render(primary: Any, render_info: Dict[str, Any]) -> str:
    """Compose a success payload plus the auto-render outcome.

    - Render disabled → plain ``{"ok": true, "result": <primary>}``.
    - Render succeeded → result wraps ``{"value": primary, "preview_path": ...}``.
    - Render failed/timed out → result wraps ``{"value": primary,
      "render_warning": {...}}``.
    """
    if not render_info.get("rendered") and render_info.get("reason") == "disabled":
        return _success(primary)
    if render_info.get("rendered"):
        return _success(
            {"value": primary, "preview_path": render_info.get("preview_path")}
        )
    # Failed / timeout — primary still succeeded.
    return _success({"value": primary, "render_warning": render_info})
