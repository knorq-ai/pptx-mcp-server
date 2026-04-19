"""text_metrics ヒューリスティックのキャリブレーションテスト.

fontTools で TTF の実 advance 幅を独立に測定し、ヒューリスティック推定
との乖離が許容帯に収まるかを検証する。

対応フォントが検出できないときは `pytest.skip` で正直にスキップする
(barebones 環境でも既存スイートが通ることを保証する)。実行するには:

    # macOS: Arial + Hiragino がシステム同梱のため自動検出される。
    # Ubuntu: apt install fonts-liberation fonts-noto-cjk

許容帯は「桁違いのミスキャリブレーション」を検出する水準に設定している。
ヒューリスティック自体が ±10-15% の近似モデルであり、ピクセル精度の
テストではない点に注意。実測で決めた許容帯は PR 本文の測定表を参照。
"""

from __future__ import annotations

import os

import pytest

# fontTools が無ければモジュール単位でスキップ。
pytest.importorskip("fontTools.ttLib")

from pptx_mcp_server.engine.text_metrics import (
    estimate_char_width,
    estimate_text_width,
)

from tests.calibration_helpers import advance_width_inches


# Arial 互換フォントの候補 (Liberation Sans は metric-compatible).
_ARIAL_CANDIDATES = [
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    "/Library/Fonts/Arial.ttf",
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "/Windows/Fonts/arial.ttf",
    "C:\\Windows\\Fonts\\arial.ttf",
]

# CJK フォントの候補.
_CJK_CANDIDATES = [
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJKjp-Regular.otf",
    "/System/Library/Fonts/Hiragino Sans GB.ttc",
    "/Library/Fonts/Yu Gothic Medium.ttc",
    "C:\\Windows\\Fonts\\YuGothM.ttc",
]


def _first_existing(paths: list[str]) -> str | None:
    for p in paths:
        if os.path.exists(p):
            return p
    return None


def _bold_path_for(regular_path: str) -> str | None:
    """Regular フォントと同じディレクトリから Bold バリアントを探す."""
    candidates = [
        regular_path.replace("-Regular.ttf", "-Bold.ttf"),
        regular_path.replace("LiberationSans-Regular", "LiberationSans-Bold"),
        regular_path.replace("Arial.ttf", "Arial Bold.ttf"),
        regular_path.replace("arial.ttf", "arialbd.ttf"),
    ]
    for c in candidates:
        if c != regular_path and os.path.exists(c):
            return c
    return None


@pytest.fixture(scope="module")
def arial_font() -> str:
    path = _first_existing(_ARIAL_CANDIDATES)
    if path is None:
        pytest.skip(
            "No Arial/Liberation Sans font found. "
            "Install fonts-liberation (Ubuntu) or use macOS/Windows."
        )
    return path


@pytest.fixture(scope="module")
def cjk_font() -> str:
    path = _first_existing(_CJK_CANDIDATES)
    if path is None:
        pytest.skip(
            "No CJK font found. "
            "Install fonts-noto-cjk (Ubuntu) or use macOS/Windows."
        )
    return path


# --- ASCII キャリブレーション ---

# 3-tier バケットモデル (#69) 導入後は per-char テスト対象に narrow / wide の
# sentinel (``i``, ``l``, ``M``, ``W``) も含める。単一定数モデル時代に必要だった
# 「4× 境界のみ」の workaround は不要になった。
@pytest.mark.parametrize(
    ("char", "size_pt"),
    [
        ("a", 10),
        ("a", 12),
        ("a", 14),
        ("H", 10),
        ("H", 18),
        ("0", 10),
        ("9", 12),
        ("e", 12),
        ("o", 12),
        ("i", 12),
        ("l", 12),
        ("M", 12),
        ("W", 12),
    ],
)
def test_ascii_per_char_calibration(arial_font: str, char: str, size_pt: float) -> None:
    """ASCII 文字について ±20% 以内で推定できることを確認する.

    #69 の 3-tier バケットモデル導入により narrow (``i``/``l``) と
    wide (``M``/``W``) も同一の許容帯で扱えるようになった。
    """
    measured = advance_width_inches(arial_font, char, size_pt)
    estimated = estimate_char_width(char, size_pt)
    rel_err = abs(estimated - measured) / measured
    assert rel_err <= 0.20, (
        f"char={char!r} pt={size_pt}: estimated={estimated:.4f} "
        f"vs measured={measured:.4f} ({rel_err * 100:.1f}% off)"
    )


def test_ascii_mixed_string_calibration(arial_font: str) -> None:
    """代表的な混在文字列に対して ±10% 以内で推定できることを確認する.

    #69 の 3-tier バケットモデル導入前は +21% 過大評価だったが、
    narrow/normal/wide を別定数化したことで系統的バイアスが解消された。
    """
    sample = "Hello World 12345 McKinsey"
    measured = sum(advance_width_inches(arial_font, ch, 12) for ch in sample)
    estimated = estimate_text_width(sample, 12)
    rel_err = abs(estimated - measured) / measured
    assert rel_err <= 0.10, (
        f"estimated={estimated:.3f} vs measured={measured:.3f} "
        f"({rel_err * 100:.1f}% off)"
    )


def test_ascii_bold_multiplier_calibration(arial_font: str) -> None:
    """Bold 倍率がおおむね測定値の比率と一致することを確認する.

    Arial Regular/Bold 両面の advance 比 (文字列 ``"Hello"`` 12pt で
    測定) に対し、ヒューリスティックの Bold 倍率 1.05 が ±5 パーセンテ
    ージポイント以内で追従していることを確認する。
    """
    bold_path = _bold_path_for(arial_font)
    if bold_path is None:
        pytest.skip(f"No Bold variant co-located with {arial_font}")
    sample = "Hello"
    measured_regular = sum(advance_width_inches(arial_font, ch, 12) for ch in sample)
    measured_bold = sum(advance_width_inches(bold_path, ch, 12) for ch in sample)
    estimated_regular = estimate_text_width(sample, 12, bold=False)
    estimated_bold = estimate_text_width(sample, 12, bold=True)
    measured_ratio = measured_bold / measured_regular
    estimated_ratio = estimated_bold / estimated_regular
    assert abs(estimated_ratio - measured_ratio) <= 0.05, (
        f"estimated_ratio={estimated_ratio:.4f} vs measured_ratio={measured_ratio:.4f}"
    )


# --- CJK キャリブレーション ---


@pytest.mark.parametrize(
    ("char", "size_pt"),
    [
        ("あ", 10),
        ("漢", 10),
        ("ア", 12),
        ("、", 10),
        ("日", 12),
    ],
)
def test_cjk_per_char_calibration(cjk_font: str, char: str, size_pt: float) -> None:
    """CJK 全角文字について ±20% 以内で推定できることを確認する."""
    measured = advance_width_inches(cjk_font, char, size_pt)
    estimated = estimate_char_width(char, size_pt)
    rel_err = abs(estimated - measured) / measured
    assert rel_err <= 0.20, (
        f"char={char!r} pt={size_pt}: estimated={estimated:.4f} "
        f"vs measured={measured:.4f} ({rel_err * 100:.1f}% off)"
    )


def test_cjk_mixed_string_calibration(cjk_font: str) -> None:
    """代表的な日本語混在文字列に対して ±15% 以内で推定できることを確認する."""
    sample = "父の日ギフトシーンで安定"
    measured = sum(advance_width_inches(cjk_font, ch, 10) for ch in sample)
    estimated = estimate_text_width(sample, 10)
    rel_err = abs(estimated - measured) / measured
    assert rel_err <= 0.15, (
        f"estimated={estimated:.3f} vs measured={measured:.3f} "
        f"({rel_err * 100:.1f}% off)"
    )


def test_half_width_kana_calibration(cjk_font: str) -> None:
    """半角カタカナは ASCII 相当幅モデルで推定する前提を実測で検証する.

    フォントに半角カナ glyph が無い場合は gracefully にスキップする。
    """
    try:
        measured = advance_width_inches(cjk_font, "ｶ", 10)
    except KeyError:
        pytest.skip("Font does not include half-width kana glyphs")
    estimated = estimate_char_width("ｶ", 10)
    rel_err = abs(estimated - measured) / measured
    # 半角カナの advance 幅はフォント差が大きいため、粗い許容帯を設定する。
    assert rel_err <= 0.40, (
        f"estimated={estimated:.4f} vs measured={measured:.4f} "
        f"({rel_err * 100:.1f}% off)"
    )


# --- ゼロ幅文字のサニティチェック (フォント非依存) ---


@pytest.mark.parametrize("zw", ["\u200b", "\u200d", "\ufe0f"])
def test_zero_width_has_zero_estimate(zw: str) -> None:
    """ゼロ幅文字は必ず 0.0" と推定される (フォント非依存)."""
    assert estimate_char_width(zw, 10) == 0.0
