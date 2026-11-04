"""Word の EQ フィールド（結果の字を持たず、命令の中の字を組んで見せるもの）を、字にする。

ふりがな（ルビ）は .doc・.rtf では `EQ \\* jc2 \\* "Font:…" \\* hps10 \\o\\ad(\\s\\up 9(かんじ),漢字)` の形で保存され、結果の字がない
（ECMA-376 第1部 17.16.5.20 の Overlay と Superscript の組み合わせ）。読み飛ばすと、見えている字（漢字）が窓から消える。
読めない形は、命令のまま返す（黙って落とさない）。
"""

from __future__ import annotations

import re

_OVERLAY = re.compile(r"\\o\s*(?:\\a[lcdr]\s*)*\(", re.I)
_UP = re.compile(r"\s*\\s\s*\\up\s*-?\d+\s*\(", re.I)


def _close_paren(s: str, i: int) -> int:
    """s[i] の次から数えて、開いている括弧を閉じる位置。なければ -1。\\ の次の字は括弧に数えない。"""
    depth = 1
    while i < len(s):
        c = s[i]
        if c == "\\":
            i += 2
            continue
        if c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
            if depth == 0:
                return i
        i += 1
    return -1


def _plain(s: str) -> str:
    return re.sub(r"\\([\\(),;])", r"\1", s).strip()


def eq_text(instr: str) -> tuple[str, str | None] | None:
    """(見える字, ふりがな)。EQ でなければ None。ふりがなの形でなければ (命令のままの説明, None)。"""
    s = instr.replace("\x13", "").replace("\x14", "").replace("\x15", "").strip()
    if not re.match(r"eq\b", s, re.I):
        return None
    m = _OVERLAY.search(s)
    if m:
        end = _close_paren(s, m.end())
        up = _UP.match(s, m.end())
        if end > 0 and up:
            e2 = _close_paren(s, up.end())
            if 0 < e2 < end:
                rest = s[e2 + 1:end].lstrip()
                if rest[:1] in (",", ";", "，", "、") and _plain(rest[1:]):
                    return _plain(rest[1:]), _plain(s[up.end():e2])
    return f"〔数式のフィールド（読めない形）：{' '.join(s.split())}〕", None
