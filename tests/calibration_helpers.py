"""キャリブレーション用ヘルパ: TTF/OTF から実 advance 幅を読み出す.

`text_metrics` ヒューリスティックとは独立したモジュールである。
fontTools の hmtx テーブルを直接参照するため、決定論的かつ
システム依存のラスタライズ誤差を受けない。

fontTools は test extra でのみ要求されるため、このモジュールは
`tests/` 配下に置き、本体パッケージには取り込まない。
"""

from __future__ import annotations

from functools import lru_cache
from typing import Any


@lru_cache(maxsize=8)
def _load_font(path: str) -> Any:
    """TTF/TTC を遅延読込しキャッシュする.

    TTC の場合は先頭フェイスを返す。日本語システムフォントの TTC は
    通常 Regular ウェイトから順に格納されているため、キャリブレーション
    基準としては先頭で十分である。
    """
    from fontTools.ttLib import TTCollection, TTFont

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
