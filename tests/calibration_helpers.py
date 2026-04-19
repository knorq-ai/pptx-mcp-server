"""キャリブレーション用ヘルパ (後方互換の thin wrapper).

実装は ``pptx_mcp_server.engine.font_metrics`` に昇格済み (Issue #91)。
既存 test/スクリプトが ``from tests.calibration_helpers import
advance_width_inches`` で参照していたため、互換のため re-export する。

新規コードは ``pptx_mcp_server.engine.font_metrics`` から直接 import
すること。
"""

from __future__ import annotations

from pptx_mcp_server.engine.font_metrics import (  # noqa: F401
    advance_width_inches,
    text_width_inches,
)
