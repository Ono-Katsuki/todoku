"""Word（.docx・.doc）の字の書式の事実：取り消し線・字の色。表計算（xlsx/fontfacts.py）と同じ印で、窓に事実としてだけ出す。

字の色の「ふつう」は、文書でいちばん字数の多い色（自動の色は黒として数える）。ふつうの色の字には印を付けない。段落の字がみな同じなら
段落の前に〔取り消し線〕〔字の色 #RRGGBB〕、一部だけなら段落の後ろに〔うち字の色 #RRGGBB：「字」〕。字そのものは変えない。"""

from __future__ import annotations

import collections

from ..xlsx.fontfacts import AUTO, _rgb


def normal_color(doc) -> str:
    """文書のふつうの字の色（文書で 1 回だけ数える）。"""
    got = doc.__dict__.get("_normal_font_color")
    if got is not None:
        return got
    from .document import iter_paras

    n: collections.Counter = collections.Counter()
    n[AUTO] += 0
    for _k, _p, blocks in doc.all_blocks():
        for p in iter_paras(blocks):
            for s in p.segs:
                if s.look is not None and not s.hidden and s.kind in ("t", "sym"):
                    n[_rgb(s.look[1])] += sum(1 for ch in s.text if not ch.isspace())
    got = doc.__dict__["_normal_font_color"] = n.most_common(1)[0][0]
    return got


def marks(p, normal: str) -> tuple[str, str]:
    """(段落の前に置く印, 段落の後ろに置く印)。"""
    shown: list[list] = []
    for s in p.segs:
        if s.look is None or s.hidden or s.kind not in ("t", "sym") or not s.text.strip():
            continue
        strike, color = bool(s.look[0]), _rgb(s.look[1])
        if shown and shown[-1][1] == strike and shown[-1][2] == color:
            shown[-1][0] += s.text
        else:
            shown.append([s.text, strike, color])
    if not shown:
        return "", ""
    whole_strike = all(x[1] for x in shown)
    colors = {x[2] for x in shown}
    whole_color = next(iter(colors)) if len(colors) == 1 else normal
    parts = []
    for t, st, c in shown:
        what = []
        if st and not whole_strike:
            what.append("取り消し線")
        if c != whole_color:
            what.append(f"字の色 #{c}")
        if what:
            parts.append(f"〔うち{'・'.join(what)}：「{t.replace(chr(10), ' ').strip()}」〕")
    head = []
    if whole_strike:
        head.append("取り消し線")
    if whole_color != normal:
        head.append(f"字の色 #{whole_color}")
    return (f"〔{'・'.join(head)}〕" if head else ""), "".join(parts)
