"""テキスト描画メトリクスの簡易推定モジュール.

自動レイアウト用のテキスト幅・高さ推定を提供する。
実フォントメトリクスファイルに依存せず、Arial/Helvetica を想定した平均字幅
モデルで近似する。CJK は 1em 全角、ASCII は 0.5em 相当として扱う。

# 精度 (Accuracy)

本モジュールはヒューリスティックであり、実測値との誤差帯は以下のとおりである。

- Arial Latin: ±10%
- CJK (Yu Gothic / Meiryo など日本語システムフォント): ±15%
- Italic / Condensed / 非 Arial 系: それ以上に悪化し得る

``font`` 引数は現状「助言ラベル」に過ぎず、幅定数は Arial 向けに較正されて
いる点に注意すること。将来のフォント別較正テーブル導入時に意味を持たせる
余地を残すため引数シグネチャに残してある。

# 既知の制限
- カーニングやヒンティングは考慮しない。
- フォントファミリ差は現時点で無視する (引数は将来拡張のため保持)。
- プロポーショナルフォントであっても全 ASCII 文字を同一の平均幅として扱う。
- イタリック・装飾体の幅差分は無視する。Bold のみ 1.05 倍の補正を行う。
- 合字・絵文字の幅は保証しない。結合マーク・ゼロ幅文字は 0 幅で扱う。
"""

from __future__ import annotations

import unicodedata

# Calibrated empirically for Arial on Windows PowerPoint at 100% zoom.
# ASCII: 0.0083"/pt ≈ average advance width of lowercase (upper+digits widen by ~20%).
# CJK:   0.0139"/pt = 1 em, matches Yu Gothic/Meiryo full-width advance.
# フォント別テーブルを導入する際はここを差し替え可能な構造に拡張する想定。

# ASCII 平均文字幅 (size_pt あたりの inches 係数)
_ASCII_WIDTH_PER_PT: float = 0.0083

# CJK 全角 1em 幅 (size_pt あたりの inches 係数)
_CJK_WIDTH_PER_PT: float = 0.0139

# Bold 補正倍率
_BOLD_MULTIPLIER: float = 1.05


# CJK スクリプトに属する Unicode 範囲 (全角相当の幅を持つもの).
# 半角カタカナ U+FF61–U+FF9F は CJK スクリプトではあるが幅は ASCII 相当のため
# `_HALF_WIDTH_KANA_RANGE` として別途管理し、この表から除外する。
_CJK_RANGES: tuple[tuple[int, int], ...] = (
    (0x2E80, 0x2EFF),  # CJK Radicals Supplement
    (0x2F00, 0x2FDF),  # Kangxi Radicals
    (0x3000, 0x303F),  # CJK Symbols and Punctuation
    (0x3040, 0x309F),  # Hiragana
    (0x30A0, 0x30FF),  # Katakana (full-width)
    (0x3400, 0x4DBF),  # CJK Unified Ideographs Extension A
    (0x4E00, 0x9FFF),  # CJK Unified Ideographs
    (0xF900, 0xFAFF),  # CJK Compatibility Ideographs
    (0xFF01, 0xFF60),  # Fullwidth ASCII + symbols (pre half-width kana)
    (0xFFE0, 0xFFEF),  # Fullwidth symbols (post half-width kana)
    (0x20000, 0x2FFFF),  # SIP: CJK Ext B/C/D/E/F + Compat Ideographs Supplement
)

# 半角カタカナの範囲 (CJK スクリプトだが幅は ASCII 相当).
_HALF_WIDTH_KANA_RANGE: tuple[int, int] = (0xFF61, 0xFF9F)


def is_cjk(char: str) -> bool:
    """CJK スクリプトに属する文字かを判定する.

    判定対象の Unicode 範囲 (全角相当 + 半角カタカナ):

    - CJK Radicals Supplement: U+2E80 – U+2EFF
    - Kangxi Radicals: U+2F00 – U+2FDF
    - CJK Symbols and Punctuation: U+3000 – U+303F
    - Hiragana: U+3040 – U+309F
    - Katakana (full-width): U+30A0 – U+30FF
    - CJK Unified Ideographs Extension A: U+3400 – U+4DBF
    - CJK Unified Ideographs: U+4E00 – U+9FFF
    - CJK Compatibility Ideographs: U+F900 – U+FAFF
    - Halfwidth and Fullwidth Forms: U+FF01 – U+FF60, U+FF61 – U+FF9F, U+FFE0 – U+FFEF
    - Supplementary Ideographic Plane (Ext B/C/D/E/F): U+20000 – U+2FFFF

    空文字列および ``\\n`` / ``\\r`` / ``\\t`` は False を返す。
    半角カタカナは CJK スクリプトの一部であるため True を返すが、字幅は
    ASCII 相当である点に注意すること (:func:`estimate_char_width` で別処理)。
    """
    if not char:
        return False
    # 明示的に ASCII 改行類は False
    if char in ("\n", "\r", "\t"):
        return False
    code = ord(char[0])
    # 半角カタカナは CJK スクリプト扱いとする (字幅は別管理)。
    if _HALF_WIDTH_KANA_RANGE[0] <= code <= _HALF_WIDTH_KANA_RANGE[1]:
        return True
    for start, end in _CJK_RANGES:
        if start <= code <= end:
            return True
    return False


def is_half_width_kana(char: str) -> bool:
    """半角カタカナ (U+FF61 – U+FF9F) かを判定する.

    半角カタカナは CJK スクリプトに属するものの、実描画では ASCII 相当の
    ~0.5em 幅となる。``estimate_char_width`` 側で CJK 全角幅ではなく
    ASCII 幅を割り当てるためにこの述語を利用する。空文字列は False を返す。
    """
    if not char:
        return False
    code = ord(char[0])
    return _HALF_WIDTH_KANA_RANGE[0] <= code <= _HALF_WIDTH_KANA_RANGE[1]


def is_zero_width(char: str) -> bool:
    """描画上の前進幅を持たない文字かを判定する.

    以下の文字を 0 幅として扱う:

    - U+200B (ZERO WIDTH SPACE)
    - U+200C (ZERO WIDTH NON-JOINER)
    - U+200D (ZERO WIDTH JOINER)
    - U+2060 (WORD JOINER)
    - U+FEFF (ZERO WIDTH NO-BREAK SPACE / BOM)
    - U+FE00 – U+FE0F (Variation Selectors)
    - ``unicodedata.combining(ch) != 0`` を満たす結合マーク全般

    空文字列は False を返す。
    """
    if not char:
        return False
    ch = char[0]
    code = ord(ch)
    if code in (0x200B, 0x200C, 0x200D, 0x2060, 0xFEFF):
        return True
    if 0xFE00 <= code <= 0xFE0F:
        return True
    if unicodedata.combining(ch) != 0:
        return True
    return False


def estimate_char_width(char: str, size_pt: float, font: str = "Arial") -> float:
    """単一文字の推定描画幅 (inches) を返す.

    判定優先順位:

    1. ゼロ幅文字 (``is_zero_width``) は 0.0 を返す。
    2. 半角カタカナ (``is_half_width_kana``) は ASCII 相当の ``size_pt * 0.0083``。
    3. CJK 全角 (``is_cjk``) は ``size_pt * 0.0139`` (1em)。
    4. それ以外の ASCII 相当は ``size_pt * 0.0083`` (平均幅)。

    ``font`` は将来拡張のため引数に残すが、現状は Arial/Helvetica を前提
    として無視される (助言ラベル)。

    Returns:
        Estimated width in inches. Accuracy: ±10% for Arial Latin,
        ±15% for CJK with Japanese system fonts (Yu Gothic/Meiryo),
        worse for italic/condensed/non-Arial.
    """
    if not char:
        return 0.0
    if is_zero_width(char):
        return 0.0
    if is_half_width_kana(char):
        return size_pt * _ASCII_WIDTH_PER_PT
    if is_cjk(char):
        return size_pt * _CJK_WIDTH_PER_PT
    return size_pt * _ASCII_WIDTH_PER_PT


def estimate_text_width(
    text: str,
    size_pt: float,
    font: str = "Arial",
    bold: bool = False,
) -> float:
    """テキスト全体の推定描画幅 (inches) を返す.

    各文字について :func:`estimate_char_width` の合計を取る。``bold=True``
    のとき全体に 1.05 倍の補正を掛ける。改行文字 (``\\n``) の幅は 0 として
    扱う (改行後の新しい行の開始とみなす想定)。

    Returns:
        Estimated width in inches. Accuracy: ±10% for Arial Latin,
        ±15% for CJK with Japanese system fonts (Yu Gothic/Meiryo),
        worse for italic/condensed/non-Arial. The ``font`` argument is
        currently an advisory label — width constants are calibrated for
        Arial only.
    """
    if not text:
        return 0.0
    total = 0.0
    for ch in text:
        if ch == "\n":
            continue
        total += estimate_char_width(ch, size_pt, font)
    if bold:
        total *= _BOLD_MULTIPLIER
    return total


def wrap_text(
    text: str,
    max_width_inches: float,
    size_pt: float,
    font: str = "Arial",
) -> list[str]:
    """``max_width_inches`` に収まるようにテキストを改行分割する.

    ルール:
    - 日本語/CJK は文字単位で折り返す。
    - ASCII は word 単位で折り返す (スペース区切り)。
    - 混在時は文字単位で現在行幅を追跡し、ASCII 部では直近の空白位置で折り
      返せる場合はそこで折り返す。
    - 単一の ASCII ワードが ``max_width_inches`` を超える場合は、そのワード
      のみ文字単位で強制分割する (URL/ハッシュタグ/長大 ASCII 対策)。
      通常長の単語については従来通り途中では切らない。
    - 入力内の ``\\n`` は強制改行として扱う。

    Returns:
        List of wrapped lines. Line breaks are determined by the same
        width heuristic as :func:`estimate_text_width`; accuracy ±10% for
        Arial Latin, ±15% for CJK with Japanese system fonts
        (Yu Gothic/Meiryo), worse for italic/condensed/non-Arial. Border
        cases may produce one extra or one fewer line than PowerPoint's
        own layout.
    """
    if not text:
        return []

    # 明示改行で分割し、各セグメントを独立に折り返す。
    segments = text.split("\n")
    lines: list[str] = []
    for seg in segments:
        if seg == "":
            lines.append("")
            continue
        lines.extend(_wrap_segment(seg, max_width_inches, size_pt, font))
    return lines


def _wrap_segment(
    text: str,
    max_width_inches: float,
    size_pt: float,
    font: str,
) -> list[str]:
    """改行を含まない 1 セグメントを折り返す内部関数.

    処理方針:
    - テキストを「トークン」に分割する。CJK 文字とスペースはそれぞれ単独トークン。
      連続する ASCII (非 CJK・非空白) 文字は 1 ワードとして 1 トークン化する。
    - トークン単位で現在行に詰め、max_width を超える直前で改行する。
    - ASCII ワード単体で max_width を超える場合は、そのトークンのみ文字単位で
      強制分割する (URL / 長い識別子 / ハッシュタグ等が 1 行扱いされて高さが
      過小評価される問題への対策)。他のトークンに対する語中分割は行わない。
    """
    tokens: list[str] = _tokenize(text)

    lines: list[str] = []
    current: str = ""
    current_width: float = 0.0

    def flush() -> None:
        nonlocal current, current_width
        lines.append(current.rstrip(" "))
        current = ""
        current_width = 0.0

    for tok in tokens:
        tok_width = _measure(tok, size_pt, font)

        # ASCII ワード単体が max_width を超える場合は文字単位で分割する。
        # CJK 1 文字・スペース・max_width 内に収まる通常ワードはこの経路に入らない。
        if (
            len(tok) > 1
            and not is_cjk(tok[0])
            and tok != " "
            and tok_width > max_width_inches
        ):
            for ch in tok:
                ch_width = _measure(ch, size_pt, font)
                if current == "":
                    current = ch
                    current_width = ch_width
                elif current_width + ch_width <= max_width_inches:
                    current += ch
                    current_width += ch_width
                else:
                    flush()
                    current = ch
                    current_width = ch_width
            continue

        if current == "":
            # 行頭: 空白トークンは無視する (行頭空白は詰める)
            if tok == " ":
                continue
            current = tok
            current_width = tok_width
            continue

        if current_width + tok_width <= max_width_inches:
            current += tok
            current_width += tok_width
        else:
            # 行を確定して新しい行を始める
            # 末尾の空白は trim
            flush()
            if tok == " ":
                continue
            current = tok
            current_width = tok_width

    if current != "":
        lines.append(current.rstrip(" "))
    return lines


def _tokenize(text: str) -> list[str]:
    """wrap 用のトークン分割.

    - CJK 文字: 1 文字 1 トークン
    - スペース: 1 トークン
    - その他 (ASCII ワード等): 連続する非 CJK・非空白をまとめて 1 トークン
    """
    tokens: list[str] = []
    buf: str = ""
    for ch in text:
        if ch == " ":
            if buf:
                tokens.append(buf)
                buf = ""
            tokens.append(" ")
        elif is_cjk(ch):
            if buf:
                tokens.append(buf)
                buf = ""
            tokens.append(ch)
        else:
            buf += ch
    if buf:
        tokens.append(buf)
    return tokens


def _measure(text: str, size_pt: float, font: str) -> float:
    """内部用: bold 無し・改行無しの幅を返す."""
    total = 0.0
    for ch in text:
        total += estimate_char_width(ch, size_pt, font)
    return total


def estimate_text_height(
    text: str,
    max_width_inches: float,
    size_pt: float,
    font: str = "Arial",
    line_height_factor: float = 1.2,
) -> float:
    """wrap した結果の総高さ (inches) を返す.

    ``len(wrap_text(...)) * size_pt * 0.0139 * line_height_factor`` で算出する。
    空文字列は 0 を返す。``size_pt * 0.0139`` は 1em の inches 換算値である。

    Returns:
        Estimated total height in inches. Accuracy inherits from
        :func:`wrap_text` — ±10% for Arial Latin, ±15% for CJK with
        Japanese system fonts (Yu Gothic/Meiryo), worse for
        italic/condensed/non-Arial. The ``line_height_factor`` default of
        1.2 matches PowerPoint's "single spacing"; adjust explicitly for
        tighter/looser leading.
    """
    if not text:
        return 0.0
    lines = wrap_text(text, max_width_inches, size_pt, font)
    if not lines:
        return 0.0
    return len(lines) * size_pt * _CJK_WIDTH_PER_PT * line_height_factor
