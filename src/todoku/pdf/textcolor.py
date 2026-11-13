"""PDF の字の色の事実。ほかの形式（docx/fontfacts.py・xlsx/fontfacts.py）と同じ印で、窓に事実としてだけ出す。

字の色は、字を描いた色：塗りの色（輪郭だけを描く字は線の色）を RGB に直したもの（pdf/content.py。CMYK・Separation などは、
ICC の色の変換をしない単純な換算）。模様（Pattern）で描いた字は「字の色（模様）」。見えない字（Tr 3・7）の色は見ない。
「ふつう」は、ページで字数のいちばん多い色（PDF はページの範囲で読むので、ページごとに決める。同じ字数なら黒）。ふつうの色の字には
印を付けない。行（マスの中の 1 行）の字がみな同じ色なら、行の前に〔字の色 #RRGGBB〕、一部だけなら、その字のすぐ後ろに
〔うち字の色 #RRGGBB：「字」〕。字そのものは変えない。色の意味（記入例・注意・消した字）は決めない。"""

from __future__ import annotations

import bisect
import collections

from ..docx.fontfacts import _color_name, most_common_color

BLACK = "000000"
PATTERN = "模様"  # 模様（Pattern）で描いた字：1 つの色として言えない

_hex: dict = {}


def hexcolor(rgb) -> str:
    """(r, g, b)（0〜1）→ "RRGGBB"。None（模様）は PATTERN。"""
    if rgb is None:
        return PATTERN
    got = _hex.get(rgb)
    if got is None:
        got = "%02X%02X%02X" % tuple(max(0, min(255, round(float(v) * 255))) for v in (tuple(rgb) + (0.0, 0.0, 0.0))[:3])
        if len(_hex) < 65536:
            _hex[rgb] = got
    return got


def name(color: str) -> str:
    return _color_name(color)


def _counts(c) -> bool:
    return not c.invisible and not c.text.isspace()


def _marked(c) -> bool:
    """色の印を付ける字：見える字と、下の色と同じ色で塗った字（画面には出ないが、色はある：印で、行のどの字かが分かる）。"""
    return not c.text.isspace() and (not c.invisible or c.faint == "same")


def normal(chars) -> str:
    """ページのふつうの字の色：見える字（空白でない字）の字数のいちばん多い色。"""
    n: collections.Counter = collections.Counter()
    for c in chars:
        if _counts(c):
            n[c.color] += len(c.text)
    return most_common_color(n)


def per_text(ln) -> list:
    """行の字（ln.text）の 1 字ごとの色：空白・どの字か決まらない所・見えない字は None。字の位置（ln.pos）から決める。"""
    spans = sorted(((c.y0, c.y1, c) if ln.vertical else (c.x0, c.x1, c) for c in ln.chars), key=lambda s: (s[0], s[1]))
    starts = [s[0] for s in spans]
    widest = max((s[1] - s[0] for s in spans), default=0.0)
    out = []
    for k, ch in enumerate(ln.text):
        if ch.isspace() or k >= len(ln.pos):
            out.append(None)
            continue
        mid = (ln.pos[k][0] + ln.pos[k][1]) / 2
        j = bisect.bisect_right(starts, mid + 0.01) - 1
        found = fallback = None
        while j >= 0 and spans[j][0] >= mid - widest - 0.01:
            lo, hi, c = spans[j]
            if hi >= mid - 0.01:
                if ch in c.text:
                    found = c
                    break
                fallback = fallback or c
            j -= 1
        c = found or fallback
        out.append(c.color if c is not None and _marked(c) else None)
    return out


def line_marks(ln, normal_color: str) -> tuple[str, list[tuple[int, int, str]]]:
    """(行の前に置く印, [(始め, 終わり, 色)]：ふつうと違う色の、行の一部の字)。"""
    colors = {c.color for c in ln.chars if _marked(c)}
    if not colors or (len(colors) == 1 and normal_color in colors):
        return "", []
    if len(colors) == 1:
        return f"〔{name(next(iter(colors)))}〕", []
    runs: list[list] = []
    for k, col in enumerate(per_text(ln)):
        if col is None:
            continue
        if runs and runs[-1][2] == col:
            runs[-1][1] = k + 1
        else:
            runs.append([k, k + 1, col])
    return "", [(a, b, col) for a, b, col in runs if col != normal_color]
