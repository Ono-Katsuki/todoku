"""PowerPoint のスライドの字を読む（読むだけ）。

- .pptx（PresentationML の部品）：ppt/presentation.xml の p:sldIdLst の順にスライドの部品をたどり、図形の並び（p:spTree）の
  順に、図形の字（p:txBody の a:p）・表のマス（a:tbl）・グラフ（chart.read_chart）・図の代替テキストを読む。ノート（notesSlide）も読む
- 古い形（.ppt、MS-PPT）：Current User ストリームの CurrentUserAtom（2.3.2）→ 最新の UserEditAtom（2.3.3）→ 前の編集を順に
  たどって persist の表（PersistDirectoryAtom 2.3.4。新しい編集の値が優先）を作り、DocumentContainer の
  SlideListWithTextContainer（2.4.14.3。recInstance 0 スライド・2 ノート）の SlidePersistAtom の順にスライドを読む。
  スライドの字は、スライドの絵（RT_Drawing 0x040C の中の OfficeArtClientTextbox 0xF00D）の TextCharsAtom（0x0FA0、UTF-16）・
  TextBytesAtom（0x0FA8、1 バイトの字。上のバイトは 0）と、プレースホルダーの字（OutlineTextRefAtom 0x0F9E の番号で、
  SlideListWithText の中の TextHeaderAtom の後ろの字を指す）。どこからも指されない SlideListWithText の字も、落とさず出す
"""

from __future__ import annotations

import re
import struct
import xml.etree.ElementTree as ET

from .chart import _kid, _kids, _ln, _rich, chart_dict, read_chart, rel_types, table_rows

RT_DOCUMENT, RT_SLIDE, RT_NOTES, RT_DRAWING = 0x03E8, 0x03EE, 0x03F0, 0x040C
RT_OUTLINEREF, RT_TEXTHEADER, RT_TEXTCHARS, RT_TEXTBYTES = 0x0F9E, 0x0F9F, 0x0FA0, 0x0FA8
RT_SLWT, RT_SLIDEPERSIST, RT_USEREDIT, RT_PERSISTDIR = 0x0FF0, 0x03F3, 0x0FF5, 0x1772
CLIENT_TEXTBOX = 0xF00D
RT_MAINMASTER, RT_HANDOUT = 0x03F8, 0x0FC9


# ====================================================================== .pptx
def pptx_slides(package) -> list[tuple[str, list[str]]]:
    """[(「スライド1」, 字の行)]。package は .has・.read を持つもの（RawZip）。"""
    pres = "ppt/presentation.xml"
    for _rid, (typ, target) in rel_types(package, "").items():
        if typ == "officeDocument":
            pres = target
    slides: list[str] = []
    if package.has(pres):
        rels = rel_types(package, pres)
        root = ET.fromstring(package.read(pres))
        for sid in root.iter():
            if _ln(sid.tag) == "sldId":
                rid = next((v for k, v in sid.attrib.items() if _ln(k) == "id" and k.startswith("{")), None)
                if rid in rels:
                    slides.append(rels[rid][1])
    if not slides:
        names = [n for n in package.names() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)] if hasattr(package, "names") else []
        slides = sorted(names, key=lambda n: int(re.search(r"(\d+)\.xml$", n)[1]))
    out = []
    n_chart = [0]
    for k, part in enumerate(slides, 1):
        if not package.has(part):
            continue
        lines = _slide_lines(package, part, n_chart)
        for typ, target in rel_types(package, part).values():
            if typ == "notesSlide" and package.has(target):
                notes = [x for x in _slide_lines(package, target, n_chart, notes=True) if x.strip()]
                if notes:
                    lines.append("ノート：" + " ／ ".join(notes))
        out.append((f"スライド{k}", lines))
    return out


def _slide_lines(package, part: str, n_chart: list[int], notes: bool = False) -> list[str]:
    root = ET.fromstring(package.read(part))
    rels = rel_types(package, part)
    tree = next((e for e in root.iter() if _ln(e.tag) == "spTree"), None)
    out: list[str] = []

    def walk(el):
        for sh in el:
            name = _ln(sh.tag)
            if name == "grpSp":
                walk(sh)
            elif name == "AlternateContent":
                ch = _kid(sh, "Choice")
                walk(ch if ch is not None else sh)
            elif name in ("sp", "cxnSp"):
                if notes and any(_ln(p.tag) == "ph" and p.get("type") in ("sldNum", "sldImg") for p in sh.iter()):
                    continue
                body = _kid(sh, "txBody")
                t = _rich(body) if body is not None else ""
                if t.strip():
                    out.append(t.replace("\n", " ／ "))
            elif name == "pic":
                nv = next((e for e in sh.iter() if _ln(e.tag) == "cNvPr"), None)
                alt = (nv.get("descr") or nv.get("title") or "") if nv is not None else ""
                if alt:
                    out.append(f"〔図: {alt}〕")
            elif name == "graphicFrame":
                for g in sh.iter():
                    if _ln(g.tag) == "tbl":
                        for tr in _kids(g, "tr"):
                            cells = [_rich(_kid(tc, "txBody")).replace("\n", " ") for tc in _kids(tr, "tc")]
                            if any(c.strip() for c in cells):
                                out.append(" ｜ ".join(cells))
                    elif _ln(g.tag) == "chart":
                        rid = next((v for k, v in g.attrib.items() if _ln(k) == "id"), "")
                        target = rels.get(rid, ("", ""))[1]
                        if target and package.has(target):
                            n_chart[0] += 1
                            out += chart_lines(chart_dict(read_chart(package.read(target), target, package), n_chart[0]))
                    elif _ln(g.tag) == "relIds":
                        from .smartart import smartart_lines

                        dm = next((v for k, v in g.attrib.items() if _ln(k) == "dm"), "")
                        target = rels.get(dm, ("", ""))[1]
                        if target and package.has(target):
                            out += ["SmartArt：" + " ／ ".join(x.strip() for x in smartart_lines(package.read(target)))]

    if tree is not None:
        walk(tree)
    return out


# ====================================================================== .ppt（バイナリ）
def _header(d: bytes, pos: int) -> tuple[int, int, int, int]:
    vi, t, ln = struct.unpack_from("<HHI", d, pos)
    return vi & 0x0F, vi >> 4, t, ln


def _children(d: bytes, start: int, end: int):
    """コンテナの中のレコード：(位置, 版, instance, 種類, 長さ)。"""
    pos = start
    while pos + 8 <= end:
        ver, inst, t, ln = _header(d, pos)
        yield pos, ver, inst, t, ln
        if ln > end - pos - 8:
            return
        pos += 8 + ln


def _atom_text(d: bytes, pos: int, t: int, ln: int) -> str:
    body = d[pos + 8:pos + 8 + ln]
    s = body.decode("utf-16-le", errors="replace") if t == RT_TEXTCHARS else body.decode("latin-1")
    return s.replace("\r", "\n").replace("\x0b", "\n").strip("\n")


def ppt_slides(storage) -> list[tuple[str, list[str]]]:
    """古い形の PowerPoint（.ppt の入れ物：PowerPoint Document と Current User のストリーム）のスライドの字。"""
    names = {n.lower(): n for n in storage.streams()}
    d = storage.read(names["powerpoint document"])
    try:
        persist, doc_ref = _persist(storage, names, d)
        return _slides_by_persist(d, persist, doc_ref)
    except (struct.error, KeyError, IndexError, ValueError):
        # 編集の記録をたどれない：ストリームの中の字の atom を、順にすべて出す（落とさない）
        texts = [_atom_text(d, p, t, ln) for p, t, ln in _all_atoms(d, 0, len(d))]
        return [("スライド（順は不明）", [t for t in texts if t.strip()])]


def _persist(storage, names: dict, d: bytes) -> tuple[dict[int, int], int]:
    cu = storage.read(names["current user"])
    offset = struct.unpack_from("<I", cu, 16)[0]  # rh 8・size 4・headerToken 4 の後ろ
    persist: dict[int, int] = {}
    doc_ref = None
    seen = set()
    while offset and offset not in seen and offset + 8 <= len(d):
        seen.add(offset)
        _v, _i, t, _ln = _header(d, offset)
        if t != RT_USEREDIT:
            raise ValueError("UserEditAtom がありません")
        last, pdir, dref = struct.unpack_from("<III", d, offset + 16)
        doc_ref = dref if doc_ref is None else doc_ref
        _v, _i, t, ln = _header(d, pdir)
        if t != RT_PERSISTDIR:
            raise ValueError("PersistDirectoryAtom がありません")
        p = pdir + 8
        while p < pdir + 8 + ln:
            v = struct.unpack_from("<I", d, p)[0]
            pid, n = v & 0xFFFFF, v >> 20
            for k in range(n):
                persist.setdefault(pid + k, struct.unpack_from("<I", d, p + 4 + 4 * k)[0])  # 新しい編集が優先
            p += 4 + 4 * n
        offset = last
    if doc_ref is None:
        raise ValueError("編集の記録がありません")
    return persist, doc_ref


def _slides_by_persist(d: bytes, persist: dict[int, int], doc_ref: int) -> list[tuple[str, list[str]]]:
    doc = persist[doc_ref]
    _v, _i, t, ln = _header(d, doc)
    if t != RT_DOCUMENT:
        raise ValueError("DocumentContainer がありません")
    lists: dict[int, list[tuple[int, list[str]]]] = {}
    for pos, _ver, inst, t, ln2 in _children(d, doc + 8, doc + 8 + ln):
        if t != RT_SLWT:
            continue
        items: list[tuple[int, list[str]]] = []
        for p2, _v2, _i2, t2, ln3 in _children(d, pos + 8, pos + 8 + ln2):
            if t2 == RT_SLIDEPERSIST:
                items.append((struct.unpack_from("<I", d, p2 + 8)[0], []))
            elif t2 == RT_TEXTHEADER and items:
                items[-1][1].append("")  # 字のない TextHeaderAtom も番号を数える（OutlineTextRefAtom の番号）
            elif t2 in (RT_TEXTCHARS, RT_TEXTBYTES) and items and items[-1][1]:
                items[-1][1][-1] += _atom_text(d, p2, t2, ln3)
        lists[inst] = items
    out = []
    for k, (pid, slwt) in enumerate(lists.get(0, []), 1):
        out.append((f"スライド{k}", _container_lines(d, persist.get(pid), slwt)))
    for k, (pid, slwt) in enumerate(lists.get(2, []), 1):
        lines = [x for x in _container_lines(d, persist.get(pid), slwt) if x.strip()]
        if lines:
            out.append((f"ノート{k}", lines))
    return out


def _container_lines(d: bytes, pos: int | None, slwt: list[str]) -> list[str]:
    out: list[str] = []
    used = set()
    if pos is not None and pos + 8 <= len(d):
        _v, _i, t, ln = _header(d, pos)
        for p, t2, ln2 in _all_records(d, pos + 8, pos + 8 + ln):
            if t2 != CLIENT_TEXTBOX:
                continue
            for p3, _v3, _i3, t3, ln3 in _children(d, p + 8, p + 8 + ln2):
                if t3 in (RT_TEXTCHARS, RT_TEXTBYTES):
                    txt = _atom_text(d, p3, t3, ln3)
                    if txt.strip():
                        out.append(txt.replace("\n", " ／ "))
                elif t3 == RT_OUTLINEREF:
                    i = struct.unpack_from("<I", d, p3 + 8)[0]
                    if i < len(slwt) and i not in used:
                        used.add(i)
                        if slwt[i].strip():
                            out.append(slwt[i].replace("\n", " ／ "))
    for i, txt in enumerate(slwt):
        if i not in used and txt.strip():
            out.append(txt.replace("\n", " ／ "))
    return out


def _all_records(d: bytes, start: int, end: int):
    """コンテナ（版 0xF）の中を、文書の順に深くたどった、すべてのレコード：(位置, 種類, 長さ)。"""
    out = []

    def walk(s_, e_, depth):
        for p, ver, _inst, t, ln in _children(d, s_, e_):
            out.append((p, t, ln))
            if ver == 0x0F and t != CLIENT_TEXTBOX and depth < 50:
                walk(p + 8, min(p + 8 + ln, e_), depth + 1)

    walk(start, end, 0)
    return out


def _all_atoms(d: bytes, start: int, end: int, skip_masters: bool = False):
    """ストリームを頭から（コンテナの中も）たどった、字の atom：(位置, 種類, 長さ)。skip_masters なら、スライドマスター
    （RT_MainMaster 0x03F8）と配布資料のマスター（RT_Handout 0x0FC9）の中は除く（「マスター タイトルの書式設定」などの見本の字）。"""
    out = []

    def walk(s, e, depth):
        for p, ver, _inst, t, ln in _children(d, s, e):
            if t in (RT_TEXTCHARS, RT_TEXTBYTES):
                out.append((p, t, ln))
            elif ver == 0x0F and depth < 50 and not (skip_masters and t in (RT_MAINMASTER, RT_HANDOUT)):
                walk(p + 8, min(p + 8 + ln, e), depth + 1)

    walk(start, end, 0)
    return out


def text_atoms(storage) -> list[str]:
    """（検査用）ストリームを頭からたどった、字の atom のすべての字。"""
    names = {n.lower(): n for n in storage.streams()}
    d = storage.read(names["powerpoint document"])
    return [_atom_text(d, p, t, ln) for p, t, ln in _all_atoms(d, 0, len(d), skip_masters=True)]


def pptx_texts(package) -> list[str]:
    """（検査用）スライドとノートの部品の a:t の字を、XML から直接拾う。"""
    out = []
    names = package.names() if hasattr(package, "names") else []
    for n in names:
        if re.fullmatch(r"ppt/(slides/slide|notesSlides/notesSlide)\d+\.xml", n):
            xml = package.read(n)
            for m in re.finditer(rb"<a:t>([^<]*)</a:t>", xml):
                out.append(_unxml(m[1].decode("utf-8", "replace")))
    return out


def _unxml(s: str) -> str:
    return s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&")



def chart_lines(d: dict) -> list[str]:
    """グラフを、行の字にする（スライドの中のグラフなど、表を置けない所に出すとき）。"""
    head = f"グラフ{d['n']}：{d.get('title') or '（題名なし）'}（{'・'.join(d.get('kinds') or []) or 'グラフ'}）"
    out = [head] + [f"{a['axis']}の題名：{a['title']}" for a in d.get("axes") or []] + list(d.get("texts") or [])
    for sub, rows in table_rows(d):
        out += [(f"系列「{sub}」：" if sub else "") + " ｜ ".join(r) for r in rows]
    return out
