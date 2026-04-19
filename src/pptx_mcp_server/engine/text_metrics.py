"""テキスト描画メトリクスの簡易推定モジュール.

自動レイアウト用のテキスト幅・高さ推定を提供する。
実フォントメトリクスファイルに依存せず、Arial/Helvetica を想定した平均字幅
モデルで近似する。CJK は 1em 全角、ASCII は 0.5em 相当として扱う。

# 既知の制限
- カーニングやヒンティングは考慮しない。
- フォントファミリ差は現時点で無視する (引数は将来拡張のため保持)。
- プロポーショナルフォントであっても全 ASCII 文字を同一の平均幅として扱う。
- イタリック・装飾体の幅差分は無視する。Bold のみ 1.05 倍の補正を行う。
- 合字・結合文字・絵文字の幅は保証しない。
"""

from __future__ import annotations

# ASCII 平均文字幅 (size_pt あたりの inches 係数)
_ASCII_WIDTH_PER_PT: float = 0.0083

# CJK 全角 1em 幅 (size_pt あたりの inches 係数)
_CJK_WIDTH_PER_PT: float = 0.0139

# Bold 補正倍率
_BOLD_MULTIPLIER: float = 1.05


def is_cjk(char: str) -> bool:
    """CJK Unified / Hiragana / Katakana / 全角句読点・記号を判定する.

    判定対象の Unicode 範囲:
    - CJK Unified Ideographs: U+4E00 – U+9FFF
    - Hiragana: U+3040 – U+309F
    - Katakana: U+30A0 – U+30FF
    - CJK Symbols and Punctuation: U+3000 – U+303F
    - Halfwidth and Fullwidth Forms: U+FF00 – U+FFEF

    改行文字 (``\\n`` など) は False を返す。
    """
    if not char:
        return False
    # 明示的に ASCII 改行類は False
    if char in ("\n", "\r", "\t"):
        return False
    code = ord(char[0])
    if 0x4E00 <= code <= 0x9FFF:
        return True
    if 0x3040 <= code <= 0x309F:
        return True
    if 0x30A0 <= code <= 0x30FF:
        return True
    if 0x3000 <= code <= 0x303F:
        return True
    if 0xFF00 <= code <= 0xFFEF:
        return True
    return False


def estimate_char_width(char: str, size_pt: float, font: str = "Arial") -> float:
    """単一文字の推定描画幅 (inches) を返す.

    CJK 全角は ``size_pt * 0.0139`` (1em)、ASCII 相当は ``size_pt * 0.0083``
    (平均幅) として近似する。``font`` は将来拡張のため引数に残すが、現状は
    Arial/Helvetica を前提として無視される。
    """
    if not char:
        return 0.0
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
      を単独行として配置する (無限ループ回避)。
    - 入力内の ``\\n`` は強制改行として扱う。
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
    - ASCII ワードが単独で max_width を超える場合でも、そのワードをそのまま
      1 行として配置し、途中で切らない (仕様)。
    """
    tokens: list[str] = _tokenize(text)

    lines: list[str] = []
    current: str = ""
    current_width: float = 0.0

    for tok in tokens:
        tok_width = _measure(tok, size_pt, font)
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
            lines.append(current.rstrip(" "))
            if tok == " ":
                current = ""
                current_width = 0.0
            else:
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
    """
    if not text:
        return 0.0
    lines = wrap_text(text, max_width_inches, size_pt, font)
    if not lines:
        return 0.0
    return len(lines) * size_pt * _CJK_WIDTH_PER_PT * line_height_factor
