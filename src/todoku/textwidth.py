"""字の幅の見積もり（Unicode の東アジアの幅、UAX #11 East Asian Width）。

- W（全角）・F（全角の互換の字）：2 つ分
- A（曖昧：○・①・※・ギリシャ文字・キリル文字など）：東アジアの文脈（同じ字の並びに W・F の字がある、または東アジアの
  書体）では 2 つ分、ほかは 1 つ分（UAX #11 第5節の勧め）
- 結合の記号（Mn・Me）・書式の字（Cf）：0
- ほか（Na・H・N）：1 つ分
1 つ分の字は、多くの書体で字ごとに幅が違う（プロポーショナル）。その字の位置は見積もりで、確かめが要る。
曖昧な幅の字の数え方は、書き込みでエージェントが決められる（ambiguous_width："wide"・"narrow"。align_options）。
文字の中の空きにそろえて書くときの前後の空白は gap_padding（Word・Excel の書き込みと証明が同じ決まりで使う）。
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


ALIGNS = ("left", "center", "right")
AMBIGUOUS = {"wide": True, "narrow": False}  # 曖昧な幅（A）の字の数え方を、エージェントが決めるとき（ambiguous_width）


def align_options(key: str, value) -> tuple[dict, object]:
    """文字の中の空き（Word の段落・Excel のセルの字の中の空欄）に書く値の、置き方の指定：
    {"value": 値, "align": "left"|"center"|"right", "ambiguous_width": "wide"|"narrow"}。
    返す：(指定, 値)。指定がなければ ({}, 値)。形が違えば ValueError（呼ぶ側で、その形式の誤りの型に直す）。"""
    if not (isinstance(value, dict) and ("align" in value or "ambiguous_width" in value)):
        return {}, value
    if set(value) - {"value", "align", "ambiguous_width"} or "value" not in value:
        raise ValueError(f'{key}: 空きの中にそろえて書くときは {{"value": 値, "align": "left"|"center"|"right"}} の形で'
                         '指定してください（曖昧な幅の字の数え方は "ambiguous_width": "wide"|"narrow"）')
    opts = {k: value[k] for k in ("align", "ambiguous_width") if k in value}
    if opts.get("align", "left") not in ALIGNS:
        raise ValueError(f'{key}: align は "left"（既定）・"center"・"right" のどれかです')
    if "ambiguous_width" in opts and opts["ambiguous_width"] not in AMBIGUOUS:
        raise ValueError(f'{key}: ambiguous_width は "wide"（全角として数える）か "narrow"（半角として数える）です')
    return opts, value["value"]


def gap_padding(gap: str, value: str, align: str = "left", context: str = "",
                ambiguous_wide: bool | None = None) -> tuple[str, str]:
    """空き gap の幅を保って value を置くときに、value の前と後ろに足す半角の空白（字の数で数える：半角の字を 1 とした幅
    text_cells の差）。曖昧な幅（A）の字は ambiguous_wide のとおり。None なら、context と value の並びの文脈で決める
    （east_asian_context：同じ並びに全角の字があれば全角）。value が空きより広ければ、何も足さない。
    align：left は後ろに、right は前に、center は前後に半分ずつ（余りの 1 は後ろ）。"""
    wide = east_asian_context(context + value) if ambiguous_wide is None else ambiguous_wide
    room = text_cells(gap, wide) - text_cells(value, wide)
    if room <= 0:
        return "", ""
    lead = {"left": 0, "right": room, "center": room // 2}[align]
    return " " * lead, " " * (room - lead)
