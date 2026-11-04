"""書類の中のリンクの先（URL）。見える字だけでは、先が分からない（申請先のページ・様式の入手先など）。

書類の形ごとの読み方はそれぞれの部品にあり、ここは Word の HYPERLINK フィールドの命令の読み方（ECMA-376 第1部 17.16.5.25）と、窓の節の書き方だけ。
"""

from __future__ import annotations

import re

_FIELD = re.compile(r"""\s*HYPERLINK\s+(?:"([^"]*)"|(\S+))?""", re.I)
_ANCHOR = re.compile(r"""\\l\s+(?:"([^"]*)"|(\S+))""")


def field_target(instr: str) -> str | None:
    """HYPERLINK フィールドの先。ほかのフィールドなら None。"""
    instr = instr.replace("\x13", "").replace("\x14", "").replace("\x15", "").replace("\x01", "")
    m = _FIELD.match(instr)
    if not m or not re.match(r"\s*HYPERLINK\b", instr, re.I):
        return None
    t = m[1] if m[1] is not None else (m[2] or "")
    if t.startswith("\\"):  # 先がなく \\l だけ
        t = ""
    t = t.replace("\\\\", "\\")
    a = _ANCHOR.search(instr)
    if a:
        t += "#" + (a[1] if a[1] is not None else a[2])
    return t or None


def markdown(links: list) -> list[str]:
    """窓の節。links：(見える字, 先) の並び。"""
    if not links:
        return []
    out = ["## リンク（見える字とその先。ファイルに書かれた先のまま。開いて確かめてはいません）", ""]
    out += [f"- {(t or '').strip() or '（字なし）'} → {u}" for t, u in links]
    out.append("")
    return out
