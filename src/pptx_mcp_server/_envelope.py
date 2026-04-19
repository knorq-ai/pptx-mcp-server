"""Structured response envelope helpers for the MCP server.

本モジュールは `server.py` から切り出したヘルパである。公開 API
ではなく、モジュール名に先頭アンダースコアを付けている。

Tool 戻り値はすべて JSON 文字列で、次の包に統一する (issue #88, #98):

- 成功: ``{"ok": true, "result": <dict>}``  — ``result`` は **常に dict**
- 失敗: ``{"ok": false, "error": {"code": ..., "message": ...,
        "parameter"?: ..., "hint"?: ..., "issue"?: ...}}``

v0.3.0 (#98): ``result`` は常に dict である。legacy で string を返していた
tool は ``{"message": "..."}`` に wrap して ``_success`` に渡す。auto-render
や追加情報は同じ dict に key を足す (``preview_path`` / ``render_warning``
/ ``slides`` 等)。

``code`` は :class:`pptx_mcp_server.engine.EngineError` の ``ErrorCode``
値と一致する (``INVALID_PARAMETER`` 等)。それ以外の例外は
``INTERNAL_ERROR`` に畳む。
"""

from __future__ import annotations

import json
from typing import Any, Dict, Optional

from .engine import EngineError


def _success(result: Dict[str, Any]) -> str:
    """Wrap a successful tool result in ``{"ok": true, "result": <dict>}``.

    v0.3.0 (#98): ``result`` は **常に dict**。legacy string を返していた
    tool は ``{"message": "..."}`` の形で wrap して呼び出す。auto-render
    や追加情報は同じ dict に key を足す (``preview_path`` / ``render_warning``
    / ``slides`` 等)。
    """
    if not isinstance(result, dict):
        # Defensive guard: 呼び出し漏れ検出用。本 branch に落ちるのは bug。
        raise TypeError(
            f"_success() expects dict; got {type(result).__name__}. "
            "v0.3.0 以降 result は常に dict を渡す (#98)."
        )
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


def _success_with_render(message: str, render_info: Dict[str, Any]) -> str:
    """Compose a success payload plus the auto-render outcome (v0.3.0 shape).

    v0.3.0 (#98): ``result`` は常に dict で, ``message`` キーに legacy
    human-readable string を入れる。auto-render の結果は同じ dict に
    ``preview_path`` か ``render_warning`` を **追加** する。

    - Render disabled → ``{"ok": true, "result": {"message": <msg>}}``
    - Render succeeded → ``{"message": <msg>, "preview_path": ...}``
    - Render failed/timed out → ``{"message": <msg>, "render_warning": {...}}``
    """
    payload: Dict[str, Any] = {"message": message}
    if not render_info.get("rendered") and render_info.get("reason") == "disabled":
        return _success(payload)
    if render_info.get("rendered"):
        payload["preview_path"] = render_info.get("preview_path")
        return _success(payload)
    # Failed / timeout — primary still succeeded.
    payload["render_warning"] = render_info
    return _success(payload)
