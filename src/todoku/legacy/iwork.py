"""Apple の Pages・Numbers・Keynote（.pages・.numbers・.key）：字を読む（読むだけ）。

Apple は形式を公開していない。ここに書いたのは、公開された解析（Sean Patrick O'Brien「iWork 2013 File Format」・keynote-parser）で知られた形：
- 入れ物は ZIP。Index/*.iwa か、中の Index.zip。古い形（iWork '09 まで）は index.xml / index.apxl の XML
- .iwa：[0x00, 長さ 3 バイト] ＋ Snappy（長さの前置きつきの生の形。CRC なし）の塊の並び。解くと ArchiveInfo（varint の長さつき）と中身の並び
- ArchiveInfo：field 2（MessageInfo）の field 1 が型、field 3 が中身の長さ
- 字：TSWP.StorageArchive（型 2001）の field 3（字の並び）、field 1 が種類（0 本文）。Numbers のマスの字：TST.TableDataList（型 6005）の
  ListEntry（field 3）の field 3

非公開の形を、知られた範囲で読むので、段階は C。字の順は本文のほかは決まらず、Numbers の表の形（どのマスの字か）は読まない。推測で補わない。
"""

from __future__ import annotations

import re

from .. import readonlydoc as rd
from ..zipraw import RawZip

KIND_JA = "Apple の iWork（.pages・.numbers・.key。読むだけ）"
WRITE = "iWork の形には書きません（形式が公開されていません）。Pages・Numbers で Word・Excel・PDF に書き出してもらってください"
STORAGE, DATALIST = 2001, 6005
MAX_OUT = 256 * 1024 * 1024


class IworkError(ValueError):
    pass


def is_iwork(names: set) -> bool:
    if any(n.startswith("Index/") and n.endswith(".iwa") for n in names) or "Index.zip" in names:
        return True
    return ("index.apxl" in names or "index.xml" in names or "index.xml.gz" in names) and any(
        n.startswith(("QuickLook/", "buildVersionHistory", "Metadata/")) for n in names)


def unsnappy(d: bytes) -> bytes:
    """Snappy の生の形（google/snappy の format_description）：varint の元の長さ・字そのまま／前の字の写しのタグ。"""
    i = 0
    n = 0
    shift = 0
    while True:
        b = d[i]
        i += 1
        n |= (b & 0x7F) << shift
        shift += 7
        if not b & 0x80:
            break
        if shift > 35:
            raise IworkError("Snappy の長さが壊れています")
    if n > MAX_OUT:
        raise IworkError("展開すると大きすぎます")
    out = bytearray()
    ln = len(d)
    while i < ln:
        t = d[i]
        i += 1
        k = t & 3
        if k == 0:
            m = t >> 2
            if m >= 60:
                c = m - 59
                m = int.from_bytes(d[i:i + c], "little")
                i += c
            m += 1
            out += d[i:i + m]
            i += m
            continue
        if k == 1:
            m = 4 + ((t >> 2) & 7)
            off = ((t >> 5) << 8) | d[i]
            i += 1
        elif k == 2:
            m = (t >> 2) + 1
            off = d[i] | d[i + 1] << 8
            i += 2
        else:
            m = (t >> 2) + 1
            off = int.from_bytes(d[i:i + 4], "little")
            i += 4
        if off == 0 or off > len(out):
            raise IworkError("Snappy の参照が壊れています")
        s = len(out) - off
        if off >= m:
            out += out[s:s + m]
        else:
            for j in range(m):
                out.append(out[s + j])
        if len(out) > n:
            raise IworkError("Snappy の長さが合いません")
    if len(out) != n:
        raise IworkError("Snappy の長さが合いません")
    return bytes(out)


def iwa_plain(d: bytes) -> bytes:
    i = 0
    out = bytearray()
    while i < len(d):
        if d[i] != 0 or i + 4 > len(d):
            raise IworkError(".iwa の塊の印が合いません")
        ln = int.from_bytes(d[i + 1:i + 4], "little")
        i += 4
        out += unsnappy(d[i:i + ln])
        i += ln
    return bytes(out)


def _varint(d: bytes, i: int) -> tuple[int, int]:
    v = s = 0
    while True:
        b = d[i]
        i += 1
        v |= (b & 0x7F) << s
        s += 7
        if not b & 0x80:
            return v, i
        if s > 70:
            raise IworkError("varint が壊れています")


def fields(d: bytes):
    """protobuf の (field, 値)。値は int か bytes。"""
    i = 0
    n = len(d)
    while i < n:
        key, i = _varint(d, i)
        f, w = key >> 3, key & 7
        if w == 0:
            v, i = _varint(d, i)
        elif w == 1:
            v = d[i:i + 8]
            i += 8
        elif w == 2:
            ln, i = _varint(d, i)
            v = d[i:i + ln]
            i += ln
        elif w == 5:
            v = d[i:i + 4]
            i += 4
        else:
            raise IworkError(f"protobuf の種類 {w} は読めません")
        yield f, v


def messages(plain: bytes):
    """(識別子, 型, 中身)。"""
    i = 0
    while i < len(plain):
        ln, i = _varint(plain, i)
        info = plain[i:i + ln]
        i += ln
        ident = 0
        infos = []
        for f, v in fields(info):
            if f == 1 and isinstance(v, int):
                ident = v
            elif f == 2 and isinstance(v, bytes):
                t = size = 0
                for ff, vv in fields(v):
                    if ff == 1 and isinstance(vv, int):
                        t = vv
                    elif ff == 3 and isinstance(vv, int):
                        size = vv
                infos.append((t, size))
        for t, size in infos:
            yield ident, t, plain[i:i + size]
            i += size


def _iwa_parts(z: RawZip, names: set):
    for n in sorted(names):
        if n.startswith("Index/") and n.endswith(".iwa"):
            yield n, z.read(n)
    if "Index.zip" in names:
        inner = RawZip(z.read("Index.zip"))
        for n in sorted(inner.names()):
            if n.endswith(".iwa"):
                yield n, inner.read(n)


def _strings(data: list[tuple[int, int, bytes]]) -> tuple[list[str], list[str]]:
    body, others = [], []
    for ident, t, payload in sorted(data, key=lambda x: x[0]):
        if t == STORAGE:
            kind = 3
            texts = []
            for f, v in fields(payload):
                if f == 1 and isinstance(v, int):
                    kind = v
                elif f == 3 and isinstance(v, bytes):
                    texts.append(v.decode("utf-8", "replace"))
            lines = [x for tx in texts for x in re.split(r"[\n\r  ]", tx) if x.strip()]
            (body if kind == 0 else others).extend(lines)
        elif t == DATALIST:
            for f, v in fields(payload):
                if f == 3 and isinstance(v, bytes):
                    for ff, vv in fields(v):
                        if ff == 3 and isinstance(vv, bytes):
                            s = vv.decode("utf-8", "replace")
                            if s.strip():
                                others.append(s)
    return body, others


def build(blob: bytes, name: str) -> dict:
    z = RawZip(blob)
    names = set(z.names())
    data = []
    unread = []
    n_parts = 0
    for part, raw in _iwa_parts(z, names):
        n_parts += 1
        try:
            for ident, t, payload in messages(iwa_plain(raw)):
                if t in (STORAGE, DATALIST):
                    data.append((ident, t, payload))
        except (IworkError, IndexError) as e:
            unread.append(f"部品 {part} を読めませんでした（{e}）")
    if n_parts:
        body, others = _strings(data)
        secs = []
        if body:
            secs.append(("本文", body))
        if others:
            secs.append(("そのほかの字（図形・表のマスなど。順は書類の順とは限りません）", others))
    else:
        from .. import markup

        part = next((n for n in ("index.apxl", "index.xml") if n in names), None)
        if part is None:
            raise IworkError("字の部品がありません（index.xml.gz の古い形は読みません）。Word・Excel・PDF に書き出したものを頼んでください")
        secs = [("字（XML のタグの間の字）", markup.build(z.read(part), "index.xml")["lines"])]
    unread.append("Apple は形式を公開していないので、知られた形の範囲で字だけを読みました。表の形（どのマスの字か）・字の位置・絵は読みません")
    return rd.make("iwork", KIND_JA, name, blob, secs, "元のファイルが正本。" + WRITE, unread, parts=n_parts)


def grade(blob: bytes, name: str):
    from ..grade import _locked

    try:
        win = build(blob, name)
    except (IworkError, ValueError) as e:
        return _locked("読めない", str(e), "読めないので、書けません")
    if not win["sections"]:
        return _locked("字を読めない iWork", "字の部品から字を取り出せませんでした。Pages・Numbers で Word・Excel・PDF に書き出してもらってください", "読めないので、書けません")
    return rd.grade(win, "C", "非公開の形の字", "形式が公開されていないので、知られた形で字だけを取り出しました。エージェントが読んで、足りない所を本人に確かめます",
                    "iWork には書かない", WRITE)


def audit(blob: bytes, name: str):
    """窓とは別に：TSWP の字を、同じ部品を解き直して型 2001 だけ拾い、窓に出ているかを照らす（表の字は窓が拾う）。"""
    from ..xlsx.audit import Audit, Check

    try:
        win = build(blob, name)
    except (IworkError, ValueError) as e:
        au = Audit(name)
        au.checks.append(Check("読む", False, str(e)))
        return au
    z = RawZip(blob)
    raw = []
    for _p, d in _iwa_parts(z, set(z.names())):
        try:
            for _id, t, payload in messages(iwa_plain(d)):
                if t == STORAGE:
                    raw += [v.decode("utf-8", "replace") for f, v in fields(payload) if f == 3 and isinstance(v, bytes)]
        except (IworkError, IndexError):
            continue
    return rd.audit(win, name, raw, "文章の部品（TSWP.StorageArchive）", WRITE)
