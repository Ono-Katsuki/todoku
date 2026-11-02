"""RTF の字句（Rich Text Format Specification 1.9.1 の「RTF Syntax」）：グループ「{」「}」・制御語「\\語 数」・制御記号「\\x」・
「\\'hh」・「\\binN」のバイナリ・文字。どれも、ファイルの中のバイトの位置 [s, e) を持つ。

制御語の区切りの空白は、制御語の一部（e に含む）。文字の中の CR・LF は、RTF では意味を持たない（字ではない）。
"""

from __future__ import annotations

import re
from dataclasses import dataclass

_WORD = re.compile(rb"\\([a-zA-Z]{1,32})(-?\d{1,10})? ?")


@dataclass(slots=True)
class Tok:
    kind: str  # open / close / word / sym / hex / bin / text / nl
    s: int
    e: int
    word: str = ""  # 制御語・制御記号の名前
    param: int | None = None
    data: bytes = b""  # 文字（text）・\'hh の1バイト・バイナリ


def tokens(blob: bytes):
    n = len(blob)
    i = 0
    while i < n:
        b = blob[i]
        if b == 0x7B:
            yield Tok("open", i, i + 1)
            i += 1
        elif b == 0x7D:
            yield Tok("close", i, i + 1)
            i += 1
        elif b == 0x5C:
            if i + 1 >= n:
                yield Tok("text", i, i + 1, data=b"\\")
                i += 1
                continue
            c = blob[i + 1]
            if (0x41 <= c <= 0x5A) or (0x61 <= c <= 0x7A):
                m = _WORD.match(blob, i)
                word = m[1].decode("ascii")
                param = int(m[2]) if m[2] is not None else None
                if word == "bin" and param is not None and param > 0:
                    end = m.end()
                    yield Tok("bin", i, end + param, word, param, blob[end:end + param])
                    i = end + param
                    continue
                yield Tok("word", i, m.end(), word, param)
                i = m.end()
            elif c == 0x27:  # \'hh
                h = blob[i + 2:i + 4]
                try:
                    v = int(h, 16)
                except ValueError:
                    yield Tok("sym", i, i + 2, "'")
                    i += 2
                    continue
                yield Tok("hex", i, i + 4, "'", v, bytes([v]))
                i += 4
            elif c in (0x0D, 0x0A):  # \ と改行は \par と同じ
                yield Tok("word", i, i + 2, "par")
                i += 2
            else:
                yield Tok("sym", i, i + 2, chr(c))
                i += 2
        elif b in (0x0D, 0x0A):
            j = i
            while j < n and blob[j] in (0x0D, 0x0A):
                j += 1
            yield Tok("nl", i, j)
            i = j
        else:
            j = i
            while j < n and blob[j] not in (0x7B, 0x7D, 0x5C, 0x0D, 0x0A):
                j += 1
            yield Tok("text", i, j, data=blob[i:j])
            i = j


def control_word_ends_at(blob: bytes, pos: int) -> bool:
    """pos の直前が、区切りのない制御語（「\\b」「\\f12」など）で終わっているか（そこに字を足すと、制御語とつながってしまう）。"""
    j = pos
    while j > 0 and (0x30 <= blob[j - 1] <= 0x39):
        j -= 1
    if j > 0 and blob[j - 1] == 0x2D:  # 負の数
        j -= 1
    k = j
    while k > 0 and ((0x41 <= blob[k - 1] <= 0x5A) or (0x61 <= blob[k - 1] <= 0x7A)):
        k -= 1
    if k == j or k == 0 or blob[k - 1] != 0x5C:
        return False
    # 「\\\\b」（逆斜線そのもの＋字）なら、制御語ではない
    back = 0
    m = k - 1
    while m >= 0 and blob[m] == 0x5C:
        back += 1
        m -= 1
    return back % 2 == 1
