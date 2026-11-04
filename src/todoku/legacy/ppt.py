"""古い PowerPoint（.ppt、MS-PPT）：スライドとノートの字・表の字を読む（読むだけ）。

字の取り出しは embeds/ppt.py（CurrentUserAtom → UserEditAtom → PersistDirectoryAtom → スライド）。表は、.ppt では
図形のまとまりなので、マスごとの字として順に出す。絵（Pictures ストリーム）は字として読まず、一覧にして view_image で見せる（imaging/docimages.py）。
"""

from __future__ import annotations

import struct

from .. import readonlydoc as rd
from ..cfb import CFB

KIND_JA = "古い PowerPoint（.ppt。読むだけ）"
WRITE = "PowerPoint 97〜2003 の形（.ppt）には書きません。本人か発行元に .pptx で保存し直してもらえば書けます"
ENCRYPTED_TOKEN = 0xF3D1C4DF  # CurrentUserAtom の headerToken：暗号化された .ppt（MS-PPT 2.3.2）


class PptError(ValueError):
    pass


def _storage(blob: bytes) -> tuple[CFB, dict]:
    c = CFB(blob)
    names = {n.lower(): n for n in c.streams()}
    if "powerpoint document" not in names:
        raise PptError("PowerPoint Document のストリームがありません")
    if "current user" in names:
        cu = c.read(names["current user"])
        if len(cu) >= 16 and struct.unpack_from("<I", cu, 12)[0] == ENCRYPTED_TOKEN:
            raise PptError("パスワードで保護された .ppt です。この形の暗号は、まだ解けません。パスワードを解いて保存し直したものか、PDF を発行元に頼んでください")
    return c, names


def pictures(c: CFB, names: dict) -> int:
    """Pictures ストリームの絵の数（OfficeArtBStoreContainerFileBlock のレコードを数える）。"""
    if "pictures" not in names:
        return 0
    d = c.read(names["pictures"])
    pos = n = 0
    while pos + 8 <= len(d):
        ln = struct.unpack_from("<I", d, pos + 4)[0]
        pos += 8 + ln
        n += 1
    return n


def build(blob: bytes, name: str) -> dict:
    from ..embeds.ppt import ppt_slides

    c, names = _storage(blob)
    slides = ppt_slides(c)
    n = pictures(c, names)
    unread = [f"絵 {n} 個の中身は字として読めません（view_image で見られます）"] if n else []
    from ..embeds.ppt import PP95_HEADING, UNORDERED_HEADING

    if slides and slides[0][0] in (UNORDERED_HEADING, PP95_HEADING):
        if not slides[0][1]:  # 字がないのではなく、作りが読めない：「字はありません」と見せない
            raise PptError("PowerPoint 97〜2003 の作り（編集の記録）が読めず、ほかに字を取り出せる所もありません。中身を読めていません。"
                           "PowerPoint か LibreOffice で .pptx に保存し直してもらえば読めます")
        unread.append("スライドの作りが読めないので、スライドの区切り・順・ノートと本文の別を読めていません（字はファイルにある順に出しています）")
    return rd.make("ppt", KIND_JA, name, blob, slides, "元のファイルが正本。スライドとノートの字を、図形の順に読みました（字の位置と飾りは読みません）。" + WRITE, unread,
                   links=hyperlinks(c, names))


def hyperlinks(c: CFB, names: dict) -> list[tuple[str, str]]:
    """文書のリンクの一覧（[MS-PPT] 2.10.16 ExHyperlinkContainer。文書の ExObjList の中）：(名前, 先)。名前は FriendlyNameAtom（instance 0）、
    先は TargetAtom（1）と LocationAtom（3。「#」でつなぐ）。いまの編集の persist の表が指す文書だけを読む。どのスライドの字のリンクかまでは読まない。"""
    from ..embeds.ppt import RT_DOCUMENT, _children, _header, _persist

    RT_EXOBJLIST, RT_HYPERLINK, RT_CSTRING = 0x0409, 0x0FD7, 0x0FBA
    d = c.read(names["powerpoint document"])
    try:
        persist, doc_ref = _persist(c, names, d)
        pos = persist[doc_ref]
        _v, _i, t, ln = _header(d, pos)
        if t != RT_DOCUMENT:
            return []
        out = []
        for p, _ver, _inst, t2, l2 in _children(d, pos + 8, pos + 8 + ln):
            if t2 != RT_EXOBJLIST:
                continue
            for q, _v3, _i3, t3, l3 in _children(d, p + 8, p + 8 + l2):
                if t3 != RT_HYPERLINK:
                    continue
                got = {i4: d[r + 8:r + 8 + l4].decode("utf-16-le", "replace").rstrip("\0")
                       for r, _v4, i4, t4, l4 in _children(d, q + 8, q + 8 + l3) if t4 == RT_CSTRING}
                target = got.get(1, "") + (f"#{got[3]}" if got.get(3) else "")
                if target:
                    out.append((got.get(0, ""), target))
        return out
    except (struct.error, KeyError, IndexError, ValueError):
        return []


def raw_texts(c: CFB, names: dict) -> list[str]:
    """窓とは別に：いまの編集の persist の表が指す入れ物（スライド・ノート・文書）の中の、字の atom をすべて拾う。
    前の編集の古いスライドと、マスターの見本の字（スライドマスター・配布資料マスター・ノートマスター）は除く。"""
    from ..embeds.ppt import (RT_DOCUMENT, RT_NOTES, RT_SLIDE, _all_atoms, _atom_text, _children, _header, _persist, text_atoms)

    d = c.read(names["powerpoint document"])
    try:
        persist, doc_ref = _persist(c, names, d)
        pos = persist[doc_ref]
        _v, _i, t, ln = _header(d, pos)
        if t != RT_DOCUMENT:
            raise ValueError
        notes_master = None
        for p, _ver, _inst, t2, _ln2 in _children(d, pos + 8, pos + 8 + ln):
            if t2 == 0x03E9:  # DocumentAtom：slideSize 8・notesSize 8・serverZoom 8 の後ろに notesMasterPersistIdRef
                notes_master = struct.unpack_from("<I", d, p + 8 + 24)[0]
        out = []
        for pid, off in sorted(persist.items()):
            if off + 8 > len(d) or pid == notes_master:
                continue
            _v, _i, t, ln = _header(d, off)
            if t in (RT_DOCUMENT, RT_SLIDE, RT_NOTES):
                out += [_atom_text(d, p, tt, ll) for p, tt, ll in _all_atoms(d, off + 8, off + 8 + ln, skip_masters=True)]
        return out
    except (struct.error, KeyError, IndexError, ValueError):
        return text_atoms(c)


def grade(blob: bytes, name: str):
    from ..grade import _locked

    try:
        win = build(blob, name)
    except PptError as e:
        return _locked("読めない", str(e), "読めないので、書けません")
    return rd.grade(win, "B", ".ppt は字を読む", "スライドとノートの字を、図形の順に読みます。図形の位置・表の形は読まないので、並びの意味はエージェントが判断します",
                    ".ppt には書かない", WRITE)


def audit(blob: bytes, name: str):
    from ..xlsx.audit import Audit, Check

    try:
        win = build(blob, name)
        c, names = _storage(blob)
    except PptError as e:
        au = Audit(name)
        au.checks.append(Check("読む", False, str(e)))
        return au
    return rd.audit(win, name, raw_texts(c, names), "スライド・ノートの字の記録（TextCharsAtom・TextBytesAtom。マスターの見本の字を除く）", WRITE)
