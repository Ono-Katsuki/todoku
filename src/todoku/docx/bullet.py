"""箇条書きの印の □（段落番号の定義 w:lvlText の字）を、1つの段落だけチェックした印にする（ECMA-376 第1部 17.9）。

段落番号の定義は、同じ定義を使う段落すべてに効くので、定義そのものは変えない。代わりに、
- numbering.xml に、同じ抽象の定義（w:abstractNum）を指す新しい w:num を1つ足す。その段落の段の定義（w:lvl）だけを
  w:lvlOverride で上書きし、上書きの w:lvl は元の段の定義のバイト列のまま、w:lvlText の字だけを変える
  （□ → ■、☐ → ☑、Wingdings の □ → Wingdings の ☑。フォント・大きさ・字下げは元のまま）
- その段落の w:numPr の w:numId を、新しい w:num の番号にする（段落のスタイルで箇条書きにしている段落は、段落の w:pPr に
  w:numPr を足す）
書き込み（fill）と証明（proof）は、同じこの決め方を使う。証明は、numbering.xml の差が「足した w:num」だけで、足した w:num の
段の定義が元の定義と lvlText の字のほかは同じであることを、バイト列で確かめる。
"""

from __future__ import annotations

import re

from .. import symfont
from .xmlpos import Node, parse

PART = "word/numbering.xml"
# チェックした印：Unicode の字は同じ形の塗った字（window.box_mark と同じ決まり）、記号の書体の印は、その書体の印のついた箱の番号
# （symfont.box_codes。Word と同じく、私用の領域 U+F0xx の字で書く）


class BulletError(ValueError):
    pass


def _root(xml: bytes) -> Node:
    root = parse(xml)
    if not root.kids:
        raise BulletError("段落番号の定義（numbering.xml）を読めません")
    return root.kids[0]


def level_of(xml: bytes, num_id: str, ilvl: int) -> tuple[str, Node, Node | None]:
    """(抽象の定義の番号, 段の定義 w:lvl, 元の w:num)。w:num の上書きに w:lvl があれば、そちら。"""
    top = _root(xml)
    num = next((n for n in top.findall("num") if n.get("numId") == num_id), None)
    if num is None:
        raise BulletError(f"段落番号の定義 {num_id} がありません")
    an = num.find("abstractNumId")
    aid = an.get("val") if an is not None else None
    for lo in num.findall("lvlOverride"):
        if int(lo.get("ilvl", 0) or 0) == ilvl and lo.find("lvl") is not None:
            return aid, lo.find("lvl"), num
    absn = next((a for a in top.findall("abstractNum") if a.get("abstractNumId") == aid), None)
    if absn is None:
        raise BulletError(f"抽象の段落番号の定義 {aid} がありません")
    if absn.find("numStyleLink") is not None:
        raise BulletError("段落番号の定義が、番号のスタイル（numStyleLink）を指しています（この形の箇条書きには、まだ書けません）")
    lvl = next((x for x in absn.findall("lvl") if int(x.get("ilvl", 0) or 0) == ilvl), None)
    if lvl is None:
        raise BulletError(f"段落番号の定義に、段 {ilvl} がありません")
    return aid, lvl, num


def checked_text(lvl_text: str, fonts, symbols=()) -> str:
    """箇条書きの印の字（lvlText）の、チェックした印。fonts は印の字の書体 (ascii, hAnsi)。決められなければ BulletError。"""
    from .document import BULLET_BOXES, sym_text
    from .window import box_mark

    shown = sym_text(lvl_text, fonts, symbols)
    if shown not in BULLET_BOXES:
        raise BulletError(f"箇条書きの印「{lvl_text}」は、チェックの □ ではありません")
    if shown != lvl_text:  # 記号の書体の印：その書体の、印のついた箱の番号
        o = ord(lvl_text)
        font = (fonts[0] if o >= 0xF000 or o < 0x80 else fonts[1]) or fonts[1] or fonts[0]
        codes = symfont.box_codes(font)
        if codes is None:
            raise BulletError(f"記号の書体（{font or '不明'}）の箇条書きの印で、チェックした印の字が決められません")
        return chr(0xF000 + codes[0])
    return box_mark(lvl_text, True)


def _font(lvl: Node) -> tuple:
    from .document import _fonts_of

    return _fonts_of(lvl.find("rPr"))


def _esc_attr(t: str) -> bytes:
    return t.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;").encode("utf-8")


def checked_lvl(xml: bytes, lvl: Node, symbols=()) -> bytes:
    """段の定義のバイト列の、lvlText の字だけをチェックした印にしたもの。symbols：文書が記号の書体と宣言した書体の名前。"""
    lt = lvl.find("lvlText")
    if lt is None:
        raise BulletError("段落番号の定義に、印の字（lvlText）がありません")
    new = checked_text(lt.get("val", "") or "", _font(lvl), symbols)
    tag = xml[lt.s:lt.se]
    tag2, n = re.subn(rb'(\b(?:\w+:)?val=")[^"]*(")', lambda m: m[1] + _esc_attr(new) + m[2], tag, count=1)
    if n != 1:
        raise BulletError("段落番号の印の字（lvlText の val）を書き換えられません")
    return xml[lvl.s:lt.s] + tag2 + xml[lt.se:lvl.e]


def new_num(xml: bytes, pre: str, num_id: str, ilvl: int, new_id: str, symbols=()) -> bytes:
    """足す w:num：元の w:num と同じ抽象の定義を指し、段 ilvl だけを、印をチェックした段の定義で上書きする。
    元の w:num のほかの段の上書き（startOverride など）は、そのまま写す。"""
    aid, lvl, num = level_of(xml, num_id, ilvl)
    keep = b""
    if num is not None:
        keep = b"".join(xml[lo.s:lo.e] for lo in num.findall("lvlOverride") if int(lo.get("ilvl", 0) or 0) != ilvl)
    over = f'<{pre}:lvlOverride {pre}:ilvl="{ilvl}">'.encode() + checked_lvl(xml, lvl, symbols) + f"</{pre}:lvlOverride>".encode()
    return (f'<{pre}:num {pre}:numId="{new_id}"><{pre}:abstractNumId {pre}:val="{aid}"/>'.encode() + keep + over
            + f"</{pre}:num>".encode())


def insert_at(xml: bytes) -> int:
    """足す w:num を置く所：いちばん後ろの w:num の後ろ（なければ、最後の w:abstractNum の後ろ）。CT_Numbering の並び
    （numPicBullet・abstractNum・num・numIdMacAtCleanup）を守る。"""
    top = _root(xml)
    nums = top.findall("num") or top.findall("abstractNum")
    if not nums:
        raise BulletError("段落番号の定義（numbering.xml）に、w:num も w:abstractNum もありません")
    return nums[-1].e


def next_ids(xml: bytes, n: int) -> list[str]:
    top = _root(xml)
    used = [int(x.get("numId")) for x in top.findall("num") if (x.get("numId") or "").isdigit()]
    base = max(used + [0]) + 1
    return [str(base + k) for k in range(n)]
