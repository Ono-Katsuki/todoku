"""字の幅の見積もり（Unicode の東アジアの幅、UAX #11 East Asian Width）。

- W（全角）・F（全角の互換の字）：2 つ分
- A（曖昧：○・①・※・ギリシャ文字・キリル文字など）：東アジアの文脈（同じ字の並びに W・F の字がある、または東アジアの
  書体）では 2 つ分、ほかは 1 つ分（UAX #11 第5節の勧め）
- 結合の記号（Mn・Me）・書式の字（Cf）：0
- ほか（Na・H・N）：1 つ分
1 つ分の字は、多くの書体で字ごとに幅が違う（プロポーショナル）。その字の位置は見積もりで、確かめが要る。
"""

from __future__ import annotations

import unicodedata


def east_asian_context(text: str) -> bool:
    """同じ字の並びに、全角（W・F）の字があるか（曖昧な幅の字を全角として数える文脈か）。"""
    return any(unicodedata.east_asian_width(ch) in ("W", "F") for ch in text)


def cells(ch: str, ambiguous_wide: bool = False) -> int:
    """1 字の幅（半角の字を 1 とする）。"""
    if unicodedata.category(ch) in ("Mn", "Me", "Cf"):
        return 0
    eaw = unicodedata.east_asian_width(ch)
    if eaw in ("W", "F"):
        return 2
    if eaw == "A":
        return 2 if ambiguous_wide else 1
    return 1


def text_cells(s: str, ambiguous_wide: bool | None = None) -> int:
    """字の並びの幅。ambiguous_wide が None なら、並びそのものの文脈で決める。"""
    wide = east_asian_context(s) if ambiguous_wide is None else ambiguous_wide
    return sum(cells(ch, wide) for ch in s)


def fixed_width(ch: str, ambiguous_wide: bool = False) -> bool:
    """幅が字の大きさで決まる字（全角の字。東アジアの文脈の曖昧な幅の字）か。ほかの字は書体で幅が変わる。"""
    return cells(ch, ambiguous_wide) == 2
