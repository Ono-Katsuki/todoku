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
