"""一太郎の文書に書き込む：元の字を、同じ字数の字で上書きする（字数・位置・書式と、ほかのバイトは変えない）。

一太郎の形式には公開された仕様がなく、字数を変えたときにずらすべき位置の表（文字の書式の区切り・行やページの印・
位置の表など）を、すべては確かめられていない。そのため、確かめられる範囲だけを書く：
- 空白・伏せ字（○○）・□ の上に、その字数までの値を上書きする（残りは元の空白のまま。伏せ字の残りは同じ幅の空白にする）
- 字を目印にした書き込み（{"old", "new"}）は、old と new が同じ字数のときだけ
- 字のない所（空のマス・段落の終わりの後ろ）に字を足すこと、字数を変えることは、壊れた文書を作らないために断る
文字の単位はストリームの中の決まった位置にあるので、書き換えるのはその 2 バイトずつだけ。複合ファイルのほかのバイトは変えない。
"""

from __future__ import annotations

import struct
import unicodedata
from dataclasses import dataclass, field

from ..cfb import CFB
from ..docx.fill import FillError, Patch, _as_bool, _is_text_value
from ..docx.window import Spots, WriteError, resolve_write
from .book import JtdFile, para_units

SPACES = " 　"
LIMIT = ("一太郎の文書は、字数を変えずに上書きします（一太郎の形式は公開された仕様がなく、字数を変えると、"
         "ずらすべき位置の表を確かめきれないため）")


@dataclass
class JtdFillResult:
    blob: bytes
    patches: list[Patch] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    edits: dict = field(default_factory=dict)  # (文字のセグメントの番号, 単位の番号) → 書く字


def _width(s: str) -> int:
    return sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in s)


def _check_chars(key: str, text: str) -> None:
    for ch in text:
        o = ord(ch)
        if o < 0x20 or o > 0xFFFF or 0xD800 <= o <= 0xDFFF or o in (0xFFFE, 0xFFFF):
            raise FillError(f"{key}: 「{ch!r}」は一太郎の文書に上書きできない字です（改行・制御の字・ U+FFFF より後ろの字は書けません）")


def _pad(orig: str) -> str:
    """値の後ろに残る所：空白は元のまま、伏せ字などは同じ幅の空白にする。"""
    return "".join(ch if ch in SPACES else ("　" if _width(ch) == 2 else " ") for ch in orig)


def capacity(doc: JtdFile, it: dict) -> int:
    """その欄に上書きできる字数（0 は書けない所）。"""
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
        return 1 if p.text[a:b] in ("□", "☐") else 0
    return b - a


def plan(doc: JtdFile, values: dict, *, force: bool = False):
    """書き込み → (単位の書き換え {(k, i): 字}, 書いた所の一覧, 知らせること)。書き込みと証明で同じものを使う。"""
    sp = Spots(doc)
    edits: dict[tuple[int, int], str] = {}
    patches: list[Patch] = []
    warnings: list[str] = []
    problems: list[str] = []
    used: dict = {}
    for key, value in values.items():
        try:
            anchor = None
            if key.startswith("@"):
                if not isinstance(value, dict):
                    raise FillError(f'{key}: 字を目印にした書き込みは {{"old": …, "new": …}} の形で指定してください')
                old, new = str(value.get("old", "")), str(value.get("new", ""))
                if len(old) != len(new):
                    raise FillError(f"{key}: {LIMIT}。old（{len(old)} 字）と new（{len(new)} 字）を同じ字数にしてください"
                                    "（足りない分は空白「　」で埋める）")
                try:
                    it = resolve_write(sp, value, used)
                except WriteError as e:
                    raise FillError(f"{key}: {e}") from None
                anchor = dict(value)
                fid, text = it["id"], it["value"]
            else:
                it = sp.items.get(key)
                if it is None:
                    raise FillError(f"{key}: そのような欄はありません")
                fid = key
                text = None
            p = it["para"]
            us = para_units(doc, p)
            kind = it["kind"]
            if kind == "checkbox":
                on = _as_bool(value)
                a, b = it["span"]
                if p.text[a:b] not in ("□", "☐") or us[a] is None:
                    raise FillError(f"{key}: チェックの印（□）を書き換えられません")
                if on:
                    edits_here = {us[a]: "■" if p.text[a] == "□" else "☑"}
                else:
                    edits_here = {}
                patches.append(Patch(fid, "control", on, "jtd", key=key))
            else:
                if text is None:
                    text = _is_text_value(value)
                _check_chars(key, text)
                if kind == "cell":
                    room = capacity(doc, it)
                    a, b = 0, room
                else:
                    a, b = it["span"]
                    room = capacity(doc, it)
                if room == 0:
                    raise FillError(f"{key}: ここには字がないので書けません。{LIMIT}。字のない所（空のマス・段落の終わり）に"
                                    "字を足すことはできません。空白や伏せ字のある所を、字を目印にして上書きしてください")
                if len(text) > room:
                    raise FillError(f"{key}: 「{text}」は {len(text)} 字です。{LIMIT}。ここに書けるのは {room} 字までです")
                if anchor is not None and len(text) != b - a:
                    raise FillError(f"{key}: {LIMIT}。old と new を同じ字数にしてください")
                orig = p.text[a:b]
                new = text + _pad(orig[len(text):])
                if _width(text) > _width(orig[:len(text)]):
                    warnings.append(f"{fid}: 「{text}」は、上書きした元の字より幅が広いので、その行の後ろの字が少し右にずれます")
                edits_here = {us[a + j]: ch for j, ch in enumerate(new) if ch != orig[j]}
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
    return edits, patches, warnings


def apply(doc: JtdFile, edits: dict) -> bytes:
    """単位の書き換えを、ストリームの中の位置に書き込んだファイルを返す。ストリームの大きさは変わらない。"""
    if not edits:
        return doc.blob
    if doc.compressed:
        raise FillError("圧縮して保存された一太郎の文書には、まだ書けません（書いたあとに LHA で縮め直す処理と、その証明が要るため）。"
                        "一太郎で「圧縮しない」で保存し直した文書なら書けます")
    by_stream: dict[str, dict[int, str]] = {}
    for (k, i), ch in edits.items():
        st = doc.texts[k]
        by_stream.setdefault(st.stream, {})[st.unit_offset(i)] = ch
    blob = doc.blob
    for name, at in by_stream.items():
        cfb = CFB(blob)
        data = bytearray(cfb.read(name))
        for off, ch in at.items():
            struct.pack_into(">H", data, off, ord(ch))
        blob = cfb.replace(name, bytes(data))
    return blob


def fill(doc: JtdFile, values: dict, *, force: bool = False) -> JtdFillResult:
    edits, patches, warnings = plan(doc, values, force=force)
    return JtdFillResult(apply(doc, edits), patches, warnings, edits)
