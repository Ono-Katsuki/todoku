"""一太郎の文書に書き込む。

1. 字数の同じ上書き：空白・伏せ字（○○）・□ の上に、その字数までの値を上書きする（残りは元の空白のまま。伏せ字の残りは
   同じ幅の空白にする）。書き換えるのは、字の単位の 2 バイトずつだけ（字数・位置・書式と、ほかのバイトは変えない）
2. 字数を変える書き込み（edit.py）：空のマス・空の段落に字を足す、空白・伏せ字より長い値に替える、字を目印にした
   書き込みで old と new の字数が違う。本文（DocumentText）だけで、文書の構造の意味を確かめられたときだけ書く
   （確かめられないときは、その理由を添えて断る）。字数を変えると、行・ページの割り付けの控えを取り除く。本文が、
   1 つのセグメントで確かめた大きさを超えるときは、「QLSTV.01」の一覧と片に分けて書く（edit.py の 9.）。
   空のマス・段落の、前の記録と段落の終わりの印の書式が違う所（書式の切り替わる所）に足す字は、終わりの印の書式を
   すべて言い直した区切りにして書く（edit.py の 2.）
圧縮して保存した文書は、中の文書に書いてから -lh5- で縮め直す。
"""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass, field

from ..docx.fill import FillError, Patch, _as_bool, _is_text_value
from ..docx.window import BOX, CHECKED, Spots, WriteError, box_mark, resolve_write
from . import edit as E
from .book import JtdFile, para_units, tokens

SPACES = " 　"
LIMIT = ("一太郎の文書は、字数を変える書き込みを、文書の構造の意味を確かめられた所（本文）だけにします")


@dataclass
class JtdFillResult:
    blob: bytes
    patches: list[Patch] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    edits: dict = field(default_factory=dict)  # (文字のセグメントの番号, 単位の番号) → 書く字（字数の同じ上書き）
    ops: list = field(default_factory=list)  # 字数を変える書き込み（edit.Op）
    notes: list[str] = field(default_factory=list)  # 直した構造の説明


def _width(s: str) -> int:
    return sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in s)


def _check_chars(key: str, text: str) -> None:
    for ch in text:
        o = ord(ch)
        if o < 0x20 or o > 0xFFFF or 0xD800 <= o <= 0xDFFF or o in (0xFFFE, 0xFFFF):
            raise FillError(f"{key}: 「{ch!r}」は一太郎の文書に書けない字です（改行・制御の字・ U+FFFF より後ろの字は書けません）")


def _pad(orig: str) -> str:
    """値の後ろに残る所：空白は元のまま、伏せ字などは同じ幅の空白にする。"""
    return "".join(ch if ch in SPACES else ("　" if _width(ch) == 2 else " ") for ch in orig)


def capacity(doc: JtdFile, it: dict) -> int:
    """その欄に、字数を変えずに上書きできる字数（0 は、上書きできる字がない所）。"""
    p = it.get("para")
    if p is None:
        return 0
    us = para_units(doc, p)
    kind = it.get("kind")
    if kind == "cell":
        t = p.text
        if t and not t.strip(SPACES) and all(u is not None for u in us[:len(t)]):
            return len(t)
        return 0
    a, b = it.get("span", (0, 0))
    if b <= a or b > len(us) or any(u is None for u in us[a:b]):
        return 0
    if kind == "checkbox":
        return 1 if p.text[a:b] in BOX | CHECKED else 0  # 空の箱（□・☐）と、チェック済みの印らしい字（☑・■ など）
    return b - a


def span_op(doc: JtdFile, p, a: int, b: int, value: str, key: str) -> E.Op:
    """段落の字の範囲 [a, b) を value にする書き込み → 本文の単位の書き換え（字数を変える）。決められなければ FillError。"""
    us = para_units(doc, p)
    t = p.text
    if b > a:
        part = us[a:b]
        if any(u is None for u in part) or len({u[0] for u in part}) != 1 \
                or any(y[1] != x[1] + 1 for x, y in zip(part, part[1:])):
            raise FillError(f"{key}: 書き換える字が、本文の続いた字ではありません（添え書き・記録をまたいでいます）")
        return E.Op(part[0][0], part[0][1], part[-1][1] + 1, value, key, "prev")
    if a > 0:
        u = us[a - 1] if a - 1 < len(us) else None
        if u is None:
            raise FillError(f"{key}: 字を足す所の前が、本文の字ではありません")
        return E.Op(u[0], u[1] + 1, u[1] + 1, value, key, "prev")
    if t:
        u = us[0]
        if u is None:
            raise FillError(f"{key}: 字を足す所の後ろが、本文の字ではありません")
        return E.Op(u[0], u[1], u[1], value, key, "next")
    return empty_op(doc, p, value, key)


def empty_op(doc: JtdFile, p, value: str, key: str) -> E.Op:
    """字のない段落（空のマスなど）に字を足す：段落の終わりの印の前に入れる。段落の始まりから終わりの印までに、
    段落の書式の記録のほかのもの（図・組など）があれば断る。"""
    end = getattr(p, "jend", None)
    start = getattr(p, "jstart", None)
    if end is None or start is None or p.segs or p.boxes:
        raise FillError(f"{key}: 字を足す所（段落の終わり）を決められません")
    st = doc.texts[end.k]
    toks = {t[1]: t for t in tokens(st.units)}
    i = start.s
    while i < end.s:
        t = toks.get(i)
        if t is None or t[0] != "rec" or t[2][0] != 0x10:
            raise FillError(f"{key}: 段落の中に、段落の書式のほかの記録（図・枠など）があるので、字を足す所を決められません")
        i += t[3]
    term = toks.get(end.s)
    if term is None or not (term[0] == "ctl" and term[2] in (0x0A, 0x0E) or term[0] == "rec" and term[2][0] == 0x30):
        raise FillError(f"{key}: 段落の終わりの印（段落の終わり・行の終わり・次のマス）が見つかりません")
    return E.Op(end.k, end.s, end.s, value, key, "next")


def growable(doc: JtdFile, it: dict) -> str | None:
    """この欄に、字数を変える書き込み（字を足す・長い値に替える）ができるか。できなければ理由。"""
    p = it.get("para")
    if p is None:
        return "欄の段落がありません"
    try:
        if it.get("kind") == "cell":
            t = p.text
            op = span_op(doc, p, 0, len(t), "あ" * (len(t) + 1), it.get("id", "")) if t else empty_op(doc, p, "あ", it.get("id", ""))
        else:
            a, b = it.get("span", (0, 0))
            op = span_op(doc, p, a, b, "あ" * (b - a + 1), it.get("id", ""))
    except FillError as e:
        return str(e).split(": ", 1)[-1]
    return E.check_op(doc, op)


def plan(doc: JtdFile, values: dict, *, force: bool = False):
    """書き込み → (字数の同じ上書き {(k, i): 字}, 字数を変える書き込み [Op], 書いた所の一覧, 知らせること)。
    書き込みと証明で同じものを使う。"""
    sp = Spots(doc)
    edits: dict[tuple[int, int], str] = {}
    ops: list[E.Op] = []
    patches: list[Patch] = []
    warnings: list[str] = []
    problems: list[str] = []
    used: dict = {}
    for key, value in values.items():
        try:
            anchor = None
            edits_here: dict = {}
            if key.startswith("@"):
                if not isinstance(value, dict):
                    raise FillError(f'{key}: 字を目印にした書き込みは {{"old": …, "new": …}} の形で指定してください')
                try:
                    it = resolve_write(sp, value, used)
                except WriteError as e:
                    raise FillError(f"{key}: {e}") from None
                anchor = dict(value)
                fid, text = it["id"], it["value"]
                _check_chars(key, text)
                p = it["para"]
                a, b = it["span"]
                us = para_units(doc, p)
                if len(text) == b - a:
                    if any(u is None for u in us[a:b]):
                        raise FillError(f"{key}: 書き換える字が、本文の字ではありません")
                    edits_here = {us[a + j]: ch for j, ch in enumerate(text) if ch != p.text[a + j]}
                else:
                    op = span_op(doc, p, a, b, text, key)
                    why = E.check_op(doc, op)
                    if why:
                        raise FillError(f"{key}: old（{len(str(value.get('old', '')))} 字）と new（{len(str(value.get('new', '')))} 字）の"
                                        f"字数が違うので、字数を変える書き込みになりますが、ここではできません：{why}。"
                                        "同じ字数にすれば（足りない分は空白「　」で埋める）、上書きできます")
                    ops.append(op)
                patches.append(Patch(fid, "input", text, "jtd", anchor=anchor, key=key))
            else:
                it = sp.items.get(key)
                if it is None:
                    raise FillError(f"{key}: そのような欄はありません")
                fid = key
                p = it["para"]
                us = para_units(doc, p)
                kind = it["kind"]
                if kind == "checkbox":
                    on = _as_bool(value)
                    a, b = it["span"]
                    if p.text[a:b] not in BOX | CHECKED or us[a] is None:
                        raise FillError(f"{key}: チェックの印（□）を書き換えられません")
                    if box_mark(p.text[a], on) != p.text[a]:
                        edits_here = {us[a]: box_mark(p.text[a], on)}
                    patches.append(Patch(fid, "control", on, "jtd", key=key))
                else:
                    text = _is_text_value(value)
                    _check_chars(key, text)
                    room = capacity(doc, it)
                    a, b = (0, room) if kind == "cell" else it["span"]
                    if text and len(text) <= room:
                        orig = p.text[a:b]
                        new = text + _pad(orig[len(text):])
                        if _width(text) > _width(orig[:len(text)]):
                            warnings.append(f"{fid}: 「{text}」は、上書きした元の字より幅が広いので、その行の後ろの字が少し右にずれます")
                        edits_here = {us[a + j]: ch for j, ch in enumerate(new) if ch != orig[j]}
                    elif text:
                        # 字数を変える：空のマスに足す・空白や伏せ字を長い値に替える
                        if kind == "cell":
                            op = span_op(doc, p, 0, len(p.text), text, key) if p.text else empty_op(doc, p, text, key)
                        else:
                            op = span_op(doc, p, a, b, text, key)
                        why = E.check_op(doc, op)
                        if why:
                            msg = (f"「{text}」は {len(text)} 字で、ここに上書きできるのは {room} 字までです" if room else
                                   "ここには上書きできる字がありません（空のマス・段落）")
                            raise FillError(f"{key}: {msg}。字数を変える書き込みも、ここではできません：{why}")
                        ops.append(op)
                    patches.append(Patch(fid, "input", text, "jtd", anchor=anchor, key=key))
            for u, ch in edits_here.items():
                if u in edits and edits[u] != ch:
                    raise FillError(f"{key}: ほかの書き込みと同じ字に重なりました")
                edits[u] = ch
        except FillError as e:
            problems.append(str(e))
    if problems:
        raise FillError("\n".join(problems))
    del force
    return edits, ops, patches, warnings


def apply(doc: JtdFile, edits: dict, ops: list | None = None) -> tuple[bytes, list[str]]:
    """書き込んだファイルと、直した構造の説明。字数の同じ上書きだけなら、ストリームの中の字の 2 バイトずつを書き換える
    （圧縮した文書は、中の文書に書いて縮め直す）。字数を変える書き込みがあれば、edit.rewrite で構造ごと直す。"""
    ops = list(ops or [])
    if not edits and not ops:
        return doc.blob, []
    if ops:
        rw = E.rewrite(doc, ops, edits)
        return rw.blob, rw.notes
    native = E.overwrite(doc.cfb.blob, doc, edits)
    if doc.compressed:
        return E.recompress(doc.blob, native), ["中の文書の字を上書きし、-lh5- で縮め直しました"]
    return native, []


def fill(doc: JtdFile, values: dict, *, force: bool = False) -> JtdFillResult:
    edits, ops, patches, warnings = plan(doc, values, force=force)
    blob, notes = apply(doc, edits, ops)
    if ops:
        # 行・ページの割り付けの控え（LineMark・PageMark・PaperMark）は、字数が変わると合わなくなるので外した。
        # 控えのない一太郎の文書は公開されているが、todoku が書いた文書を一太郎で開いて確かめたことはまだない
        warnings.append("字数を変えて書いたので、行とページの割り付けの控えを外しました（一太郎が開くときに割り付け直します）。"
                        "このように書いた文書を、一太郎で開いて確かめたことはまだありません。出す前に、一太郎で一度開いて見てもらってください")
        if any("分けて持つ" in n or "片を分け" in n for n in notes):
            # 大きな本文の分け方は、集めた文書の 1 件（2 つの片）で確かめた形に合わせたもの
            warnings.append("本文が大きくなったので、一太郎が大きな本文を持つ形（「QLSTV.01」の一覧と、いくつかの片）に分けて書きました。"
                            "この形は、集めた文書の 1 件で確かめたものです。一太郎で開いて、本文の終わりまで字と書式が元のとおりかを見てください")
    return JtdFillResult(blob, patches, warnings, edits, ops, notes)
