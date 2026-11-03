"""書き込む値の字の確かめ（どの形式にも共通。api.fill が、形式ごとの書き込みの前に呼ぶ）。

- 制御文字（C0：U+0000〜U+001F のうち、タブ・改行のほか。U+007F も）：Excel・Word・PDF のどれでも、ファイルが壊れるか、
  見えない字として残る。書かない（どの欄の、何文字目の字かを示す）
- 対になっていないサロゲート（U+D800〜U+DFFF の1つだけ）：UTF-8 にも UTF-16 にもできない字。書かない
- 字でない符号位置（U+FFFE・U+FFFF）：XML に置けない。書かない
- 改行は「\\n」にそろえる（「\\r\\n」と、1つだけの「\\r」は「\\n」にする）。各形式の書き手が、その形式の改行にする

値の中の字だけでなく、リスト（空欄ごとの値）・{"old", "new"} などの入れ子の字も、欄の名前（鍵）も確かめる。
"""

from __future__ import annotations

import re

ALLOWED_C0 = {"\t", "\n", "\r"}
_BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\ud800-\udfff￾￿]")


class InputError(ValueError):
    pass


def _bad(ch: str) -> str | None:
    o = ord(ch)
    if o in (0x0B, 0x0C, 0x1C, 0x1D, 0x1E):
        return f"改行の働きをする制御文字 U+{o:04X}（改行は \\n で書いてください）"
    if (o < 0x20 and ch not in ALLOWED_C0) or o == 0x7F:
        return f"制御文字 U+{o:04X}"
    if 0xD800 <= o <= 0xDFFF:
        return f"対になっていないサロゲート U+{o:04X}"
    if o in (0xFFFE, 0xFFFF):
        return f"字でない符号位置 U+{o:04X}"
    return None


def check_text(s: str, where: str) -> str:
    """字を確かめ、改行をそろえた字を返す（書けない字があれば InputError）。"""
    m = _BAD.search(s)
    if m:
        i = m.start()
        why = _bad(s[i])
        if why:
            around = "".join(c if _bad(c) is None else "�" for c in s[max(0, i - 8):i + 8])
            raise InputError(f"{where}: {i + 1} 文字目に、書けない字（{why}）があります（前後：「{around}」）。"
                             "コピーしたときに紛れ込んだ字のことが多いので、取り除いてから書いてください")
    if "\r" in s:
        s = s.replace("\r\n", "\n").replace("\r", "\n")
    return s


def normalize_values(values: dict) -> dict:
    """values の字をすべて確かめ、改行をそろえた新しい dict を返す（元の dict は変えない）。"""
    if not isinstance(values, dict):
        raise InputError("values は {欄の id: 値} の形で指定してください")

    def walk(v, where: str):
        if isinstance(v, str):
            return check_text(v, where)
        if isinstance(v, bool) or v is None:
            return v
        if isinstance(v, float):
            import math

            if not math.isfinite(v):
                raise InputError(f"{where}: 値 {v!r} は数ではありません（無限大・NaN は書けません）")
            return v
        if isinstance(v, (list, tuple)):
            return [walk(x, f"{where}［{n}つ目］") for n, x in enumerate(v, 1)]
        if isinstance(v, dict):
            return {check_text(k, where) if isinstance(k, str) else k: walk(x, f"{where}・{k}") for k, x in v.items()}
        return v

    out = {}
    for k, v in values.items():
        key = check_text(k, "欄の id") if isinstance(k, str) else k
        out[key] = walk(v, str(key))
    return out


# ====================================================================== チェックの欄の値（どの形式も、この 1 つで読む）
# 受ける値は 2 種類だけ：
# 1. 真偽値：JSON の true / false（Python の True / False）と、その綴りの字「true」「false」（大文字・小文字は問わない）
# 2. その書類・その形式が宣言した状態の値（on・off）：PDF のチェックの「入」の状態の名前と Off（PDF 32000-1 12.7.4.2.3）、
#    XFA の <items> の値、ActiveX の Value、Word などの箱の字（☐・☑ など。GLYPH_ON・GLYPH_OFF）
# 「はい」「あり」「有」「○」のような言葉は受けない（言葉の意味で決めない）。本人の答えを true / false に直すのはエージェント。
CHECKBOX_HINT = ("チェックの欄には true / false（JSON の真偽値）か、この欄が宣言した状態の値で書いてください。"
                 "「はい」「あり」などの言葉は受けません（本人の答えを true / false に直して書いてください）")


def _glyph_states() -> tuple[frozenset, frozenset]:
    from . import chars

    return chars.CHECKED_BOXES | chars.MAYBE_CHECKED, chars.EMPTY_BOXES


GLYPH_ON, GLYPH_OFF = _glyph_states()  # 箱の字で表すチェック（Word・一太郎・テキスト・PowerPoint など）の、状態の字


def checkbox_state(value, on=(), off=()) -> bool | None:
    """チェックの欄に書く値 → True（入）・False（切）。受けない値は None（言葉・数・空の字・None も None）。
    on・off：その欄が宣言した状態の値（字のとおりに比べる）。"""
    if isinstance(value, bool):
        return value
    if not isinstance(value, str):
        return None
    if value in on:
        return True
    if value in off:
        return False
    t = value.strip()
    if t in on:
        return True
    if t in off:
        return False
    low = t.lower()
    if low in ("true", "false"):
        return low == "true"
    return None


def checkbox_value(value, on=(), off=(), where: str = "") -> bool:
    """checkbox_state と同じ。受けない値なら InputError（宣言した状態の値があれば、それを添える）。"""
    got = checkbox_state(value, on, off)
    if got is None:
        declared = [str(x) for x in (*on, *off) if str(x)]
        extra = f"（この欄の状態の値：{'・'.join(dict.fromkeys(declared))}）" if declared else ""
        raise InputError(f"{where + ': ' if where else ''}{CHECKBOX_HINT}{extra}（「{value}」）")
    return got
