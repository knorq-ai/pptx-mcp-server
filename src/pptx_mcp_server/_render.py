"""Auto-render gating and timeout plumbing for the MCP server.

本モジュールは `server.py` から切り出したヘルパである。公開 API
ではなく、モジュール名に先頭アンダースコアを付けている。

Auto-render は issue #86 で opt-in 化した。既定 OFF で、環境変数
``PPTX_MCP_AUTO_RENDER`` が truthy のときのみ LibreOffice を fork して
PNG プレビューを生成する。timeout (秒) は ``PPTX_MCP_RENDER_TIMEOUT``
で制御し、既定 10 秒である。タイムアウトや失敗で primary tool を落と
さず、``render_warning`` フィールドとして呼び出し元にそのまま返す。

`render_slide` 呼び出しは callable として注入する。これは ``server``
モジュール側で ``render_slide`` を名前空間に持ちたい (テストで
``monkeypatch.setattr(server_mod, "render_slide", ...)`` される) ため。
"""

from __future__ import annotations

import concurrent.futures
import os
from typing import Any, Callable, Dict

_DEFAULT_RENDER_TIMEOUT_S = 10.0


def _auto_render_enabled() -> bool:
    """``PPTX_MCP_AUTO_RENDER`` が truthy なら auto-render を実行する."""
    v = os.environ.get("PPTX_MCP_AUTO_RENDER", "").strip().lower()
    return v in {"1", "true", "yes", "on"}


def _auto_render_timeout() -> float:
    """``PPTX_MCP_RENDER_TIMEOUT`` (秒) を float で返す. 既定 10 秒."""
    raw = os.environ.get("PPTX_MCP_RENDER_TIMEOUT", "").strip()
    if not raw:
        return _DEFAULT_RENDER_TIMEOUT_S
    try:
        v = float(raw)
        if v <= 0:
            return _DEFAULT_RENDER_TIMEOUT_S
        return v
    except ValueError:
        return _DEFAULT_RENDER_TIMEOUT_S


def _run_auto_render(
    file_path: str,
    slide_index: int,
    *,
    render_fn: Callable[..., str],
) -> Dict[str, Any]:
    """Render a slide preview if enabled; else return a neutral "skipped" payload.

    Always returns a dict — never raises, never fails the caller. Shape::

        {"rendered": false, "reason": "disabled"}                       # off
        {"rendered": true, "preview_path": "/.../slide-01.png"}         # ok
        {"rendered": false, "reason": "timeout", "timeout_s": 10.0}     # slow
        {"rendered": false, "reason": "failed", "error": "<msg>"}       # crash

    Opt-in via ``PPTX_MCP_AUTO_RENDER=1``. Timeout via
    ``PPTX_MCP_RENDER_TIMEOUT`` (default 10s). The caller should only invoke
    this AFTER the primary action has succeeded.

    ``render_fn`` は ``engine.render_slide`` 互換の callable。呼び出し元
    (server.py) は自身の名前空間にある ``render_slide`` を渡す — テストが
    ``monkeypatch.setattr(server_mod, "render_slide", ...)`` でパッチする
    ためである。
    """
    if not _auto_render_enabled():
        return {"rendered": False, "reason": "disabled"}

    timeout = _auto_render_timeout()

    def _do_render() -> str:
        return render_fn(file_path, slide_index=slide_index, dpi=100)

    # ThreadPoolExecutor で走らせ future.result(timeout) で上限を掛ける。
    # `with` block を使うと __exit__ で shutdown(wait=True) が呼ばれ、
    # 裏の slow スレッドが終わるまでブロックするため timeout が機能しない。
    # 代わりに明示的に shutdown(wait=False) を呼ぶ。
    # timeout 到達時はスレッドが裏で生きたままだが、subprocess 側にも
    # 120 秒 / 60 秒の独自 timeout があるため無限ハングはしない。
    ex = concurrent.futures.ThreadPoolExecutor(max_workers=1)
    try:
        future = ex.submit(_do_render)
        try:
            out = future.result(timeout=timeout)
        except concurrent.futures.TimeoutError:
            return {
                "rendered": False,
                "reason": "timeout",
                "timeout_s": timeout,
            }
        except Exception as e:  # renderer itself raised
            return {"rendered": False, "reason": "failed", "error": f"{type(e).__name__}: {e}"}
    finally:
        ex.shutdown(wait=False)

    try:
        # render_fn may return multiple lines; take the last one (target slide)
        lines = out.strip().split("\n")
        return {"rendered": True, "preview_path": lines[-1]}
    except Exception as e:
        return {"rendered": False, "reason": "failed", "error": f"{type(e).__name__}: {e}"}
