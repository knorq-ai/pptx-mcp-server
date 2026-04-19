"""Real font advance-width measurement via fontTools.

Used by ``check_text_overflow`` の real-font path (#91) として、
``text_metrics`` ヒューリスティックとは独立した真実源を提供する。
TTF/TTC の ``hmtx`` / ``cmap`` を直接読むため、ヒューリスティック
の echo chamber を破壊できる。

fontTools は optional: ``pip install pptx-mcp-server[validation]``
を要求する。未インストール環境では ``_load_font`` が ImportError を
送出するので、呼び出し側が friendly message に変換すること。

本モジュールは ``tests/calibration_helpers.py`` から昇格したものである。
テスト側はこのモジュールを re-export する薄いラッパに変更済み。
"""

from __future__ import annotations

import os
from functools import lru_cache
from typing import Any, Dict, List


# ---------------------------------------------------------------------------
# fontTools ロード (optional)
# ---------------------------------------------------------------------------


@lru_cache(maxsize=16)
def _load_font(path: str) -> Any:
    """TTF/TTC を遅延読込しキャッシュする.

    TTC の場合は先頭フェイスを返す。日本語システムフォントの TTC は
    通常 Regular ウェイトから順に格納されているため、基準としては
    先頭で十分である。

    fontTools が未インストールの場合は ImportError を送出する
    (呼び出し側で ``pip install ...[validation]`` メッセージに
    変換することを想定)。
    """
    from fontTools.ttLib import TTCollection, TTFont  # type: ignore

    if path.endswith(".ttc"):
        coll = TTCollection(path)
        return coll.fonts[0]
    return TTFont(path, recalcTimestamp=False)


def advance_width_inches(font_path: str, char: str, size_pt: float) -> float:
    """``char`` の ``size_pt`` における advance 幅を inches で返す.

    複数コードポイントが渡された場合は各コードポイントの advance を
    加算する (合字・シェイピングは考慮しない点に注意。ヒューリスティック
    と対比可能な粒度で測るのが目的である)。

    cmap に glyph が存在しないとき ``KeyError`` を送出する。
    """
    font = _load_font(font_path)
    cmap = font.getBestCmap()
    hmtx = font["hmtx"]
    units_per_em = font["head"].unitsPerEm
    total_units = 0
    for ch in char:
        glyph_name = cmap.get(ord(ch))
        if glyph_name is None:
            raise KeyError(
                f"Character {ch!r} (U+{ord(ch):04X}) not in font {font_path}"
            )
        advance, _ = hmtx[glyph_name]
        total_units += advance
    # 1 pt = 1/72 inch. advance_in_em = units / units_per_em.
    return total_units / units_per_em * size_pt / 72.0


def text_width_inches(font_path: str, text: str, size_pt: float) -> float:
    """``text`` 全体の advance 幅合計を inches で返す.

    cmap に無いコードポイントは幅 0 として黙って無視する
    (例: 制御文字や異体字セレクタ)。ヒューリスティックとの対比
    用途では "ほぼすべて" の文字で幅を得られれば十分なためである。
    ``\n`` も幅 0 として無視する (行分割は呼び出し側の責務)。
    """
    font = _load_font(font_path)
    cmap = font.getBestCmap()
    hmtx = font["hmtx"]
    units_per_em = font["head"].unitsPerEm
    total_units = 0
    for ch in text:
        if ch == "\n":
            continue
        glyph_name = cmap.get(ord(ch))
        if glyph_name is None:
            continue
        advance, _ = hmtx[glyph_name]
        total_units += advance
    return total_units / units_per_em * size_pt / 72.0


# ---------------------------------------------------------------------------
# システムフォント探索
# ---------------------------------------------------------------------------


# フォント名 → 候補パスのマップ。Linux CI (Liberation + Noto CJK)、
# macOS (Arial + Hiragino + Yu Gothic)、Windows (Arial + Yu Gothic +
# Meiryo) のいずれでも最低 1 本当たれば十分なように網羅的に列挙する。
_FONT_CANDIDATES: Dict[str, List[str]] = {
    "Arial": [
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
        "/Library/Fonts/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Windows/Fonts/arial.ttf",
        "C:\\Windows\\Fonts\\arial.ttf",
    ],
    "Liberation Sans": [
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    ],
    "Yu Gothic": [
        "/Library/Fonts/Yu Gothic Medium.ttc",
        "/System/Library/Fonts/YuGothicMedium.ttc",
        "C:\\Windows\\Fonts\\YuGothM.ttc",
        "/Windows/Fonts/YuGothM.ttc",
    ],
    "Meiryo": [
        "C:\\Windows\\Fonts\\meiryo.ttc",
        "/Windows/Fonts/meiryo.ttc",
    ],
    "Hiragino Sans": [
        "/System/Library/Fonts/Hiragino Sans GB.ttc",
        "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
        "/Library/Fonts/Hiragino Sans GB.ttc",
    ],
    "Noto Sans CJK": [
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJKjp-Regular.otf",
    ],
    "Noto Sans": [
        "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
        "/usr/share/fonts/opentype/noto/NotoSans-Regular.ttf",
    ],
}


def discover_system_fonts() -> Dict[str, str]:
    """一般的な場所から font name → path のマップを構築して返す.

    実在するパスのみを含む。何も見つからなければ空 dict を返す。
    Consumers は ``check_text_overflow(..., font_paths=discover_system_fonts())``
    の形で zero-config な real validation を実行できる。
    """
    resolved: Dict[str, str] = {}
    for name, candidates in _FONT_CANDIDATES.items():
        for p in candidates:
            if os.path.exists(p):
                resolved[name] = p
                break
    return resolved
