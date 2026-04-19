"""幾何パラメータの fail-fast 検証ヘルパ.

# 背景 (#85)

エンジンのプリミティブ (``add_auto_fit_textbox`` / ``add_flex_container`` /
``add_responsive_card_row``) はこれまで ``width <= 0``・``padding < 0``・
``min_size_pt > font_size_pt`` などの不正な幾何指定をサイレントに受け入れ、
「技術的には成功したが結果が誤っている」出力を返していた。エージェント
消費者にとっては最悪の失敗モードであり、ツールは成功を返すのに出力は
壊れている、という状況を生む。

本モジュールは entry point で呼ぶ最小限のバリデータを提供する。
メッセージにはパラメータ名と不正値を必ず含めることで、LLM エージェントが
再試行時に修正点を特定できる形式を担保する。

NaN / inf も拒否する (``math.isfinite``)。
"""

from __future__ import annotations

import math

from .pptx_io import EngineError, ErrorCode


def _reject_nonfinite(caller: str, name: str, value: float) -> None:
    """NaN / inf を ``INVALID_PARAMETER`` として拒否する."""
    if not math.isfinite(value):
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            f"{caller}: {name}={value!r} must be finite (no NaN/inf).",
        )


def _require_positive(caller: str, name: str, value: float) -> None:
    """``value > 0`` を要求する. ``0`` ちょうどは拒否する."""
    _reject_nonfinite(caller, name, value)
    if value <= 0:
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            f"{caller}: {name}={value:.4f} must be > 0.",
        )


def _require_non_negative(caller: str, name: str, value: float) -> None:
    """``value >= 0`` を要求する. ``0`` ちょうどは許容する."""
    _reject_nonfinite(caller, name, value)
    if value < 0:
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            f"{caller}: {name}={value:.4f} must be >= 0.",
        )


def validate_auto_fit_geometry(
    *,
    caller: str = "add_auto_fit_textbox",
    width: float,
    height: float,
    font_size_pt: float,
    min_size_pt: float,
) -> None:
    """``add_auto_fit_textbox`` の幾何/フォント引数を検証する.

    Raises:
        EngineError: いずれかの引数が不正なとき (``INVALID_PARAMETER``)。
    """
    _require_positive(caller, "width", width)
    _require_positive(caller, "height", height)
    _require_positive(caller, "font_size_pt", font_size_pt)
    _require_positive(caller, "min_size_pt", min_size_pt)
    if min_size_pt > font_size_pt:
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            (
                f"{caller}: min_size_pt={min_size_pt:.2f} must be <= "
                f"font_size_pt={font_size_pt:.2f} "
                "(auto-fit starts at font_size_pt and steps down toward min_size_pt)."
            ),
        )


def validate_flex_geometry(
    *,
    caller: str = "add_flex_container",
    width: float,
    height: float,
    padding: float,
    gap: float,
) -> None:
    """``add_flex_container`` の幾何引数を検証する.

    Raises:
        EngineError: いずれかの引数が不正なとき (``INVALID_PARAMETER``)。
    """
    _require_positive(caller, "width", width)
    _require_positive(caller, "height", height)
    _require_non_negative(caller, "padding", padding)
    _require_non_negative(caller, "gap", gap)


def validate_card_row_geometry(
    *,
    caller: str = "add_responsive_card_row",
    width: float,
    max_height: float,
    gap: float,
    min_card_height: float,
) -> None:
    """``add_responsive_card_row`` の幾何引数を検証する.

    Raises:
        EngineError: いずれかの引数が不正なとき (``INVALID_PARAMETER``)。
    """
    _require_positive(caller, "width", width)
    _require_positive(caller, "max_height", max_height)
    _require_non_negative(caller, "gap", gap)
    _require_non_negative(caller, "min_card_height", min_card_height)
    if max_height < min_card_height:
        raise EngineError(
            ErrorCode.INVALID_PARAMETER,
            (
                f"{caller}: max_height={max_height:.2f} < "
                f"min_card_height={min_card_height:.2f}; "
                "constraints are contradictory."
            ),
        )
