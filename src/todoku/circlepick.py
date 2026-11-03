"""○で囲む字の選び方のうち、PDF・Word・Excel で同じ規則（字の形だけで決め、意味は見ない）。

囲む字が near の字そのものの中にあるとき（circle「昭」near「大・昭・平・令」・circle「都」near「都県市」）：
- near の字が選択肢の並び（中に区切りの字がある）で、囲む字がその中で区切りにはさまれた一つの項目なら（「大・昭・平・令」の
  「昭」・「S・H」の「S」）、その所を囲む（エージェントは、選択肢の並びを目印に書いた）
- そうでなければ（「都県市」の「都」：字の並びの中の字）、near の中の所は選ばない。ほかの候補が near と同じ行にあれば
  それを選び、別の行にしかなければ、黙って別の行へ移らずに、候補（near の中の所も、occurrence の番号つき）を添えて断る
"""

from __future__ import annotations

import unicodedata


def sep(ch: str) -> bool:
    """字の区切り：字（文字・数字・結合の字。Unicode の一般カテゴリ L*・N*・M*）でないもの。"""
    return ch.isspace() or unicodedata.category(ch)[0] not in ("L", "N", "M")


def option_item(text: str, a: int, b: int, lo: int, hi: int) -> bool:
    """text[lo:hi]（near の字の所）の中の text[a:b]（囲む字）が、選択肢の並びの一つの項目か：near の字の中に、字と字の
    あいだの区切り（「・」「／」「、」・空白など）があり、囲む字の前後が near の端か区切り。"""
    if not (lo <= a < b <= hi) or (a == lo and b == hi):
        return False
    inner = text[lo:hi].strip()
    if not any(sep(ch) for ch in inner[1:-1]):
        return False
    return (a == lo or sep(text[a - 1])) and (b == hi or sep(text[b]))


def inside_message(target: str, near: str, listing: str) -> str:
    return (f"「{target}」は near の「{near}」の字の中にもあり（区切られた選択肢の項目ではありません）、ほかの「{target}」は "
            f"near と別の行にあります。どれを囲むか、黙って別の行には移らないので、occurrence で1つに決めてください"
            f"（near の中の字も選べます）：{listing}")


def items_message(target: str, near: str, n: int, listing: str) -> str:
    return (f"「{target}」が、near の「{near}」（選択肢の並び）と同じ字の所の中に {n} か所あります。occurrence で1つに"
            f"決めてください：{listing}")
