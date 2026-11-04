"""Excel のブック（.xlsx・.xls）のシートごとに、グラフ・SmartArt・埋め込みを集めて窓の形にする（読むだけ）。

- .xlsx：シートの図（drawing の部品）の graphicFrame の a:graphicData → c:chart・cx:chart（グラフの部品）・dgm:relIds（SmartArt）。
  グラフのシート（xl/chartsheets）は、その図のグラフ。埋め込みは、シートの oleObjects/oleObject（r:id → xl/embeddings の部品。
  mc:AlternateContent の Choice に置き場所 objectPr/anchor がある）
- .xls：グラフは biffchart。埋め込みは、OBJ（ot＝8 の絵）のうち FtPioGrbit（MS-XLS 2.5.151）の fCtl（ActiveX）・fDde（リンク）・
  fCamera・fPrstm（Ctls ストリーム）が 0 のもので、FtPictFmla（2.5.150）の lPosInCtlStm から入れ物の名前「MBD」＋8桁の16進。
  入れ物は複合ファイルのいちばん上にある。RC4 CryptoAPI で暗号化したブックでは、暗号を解いた encryption ストリーム
  （MS-OFFCRYPTO 2.3.5.4）の中にある（bincrypto.summary_entries）
"""

from __future__ import annotations

import posixpath
import re
import struct
import xml.etree.ElementTree as ET

from ..formula.parser import addr
from .chart import (Chart, _attr, _kid, _ln, chart_dict, chart_markdown, read_chart, rel_types,
                    workbook_resolver)
from .ole import describe, object_markdown


def add_to_window(wb, doc: dict) -> None:
    """窓（xlsx.window.build の結果）の各シートに、グラフ（charts）と埋め込み（objects）を足す。番号はブックの中で通し。"""
    n_chart, n_obj = 0, 0
    for sh, d in zip(wb.sheets, doc["sheets"]):
        try:
            charts, objs = sheet_contents(wb, sh)
        except Exception as e:  # noqa: BLE001
            charts, objs = [Chart(part=sh.path, error=f"{type(e).__name__}: {e}")], []
        if charts:
            items = []
            for ch in charts:
                n_chart += 1
                items.append(chart_dict(ch, n_chart))
            d["charts"] = items
        if objs:
            for o in objs:
                n_obj += 1
                o["n"] = n_obj
            d["objects"] = objs


def sheet_markdown(sh: dict) -> list[str]:
    """シートの窓の、グラフと埋め込みの節。"""
    out: list[str] = []
    for c in sh.get("charts") or []:
        out += chart_markdown(c, "###")
    for o in sh.get("objects") or []:
        if o.get("kind") == "smartart":
            out += [f"### SmartArt{o['n']}（{o.get('where') or '位置不明'}）", ""] + [f"- {x}" for x in o["lines"]] + [""]
        else:
            out += object_markdown(o, "###")
    return out


def sheet_contents(wb, sh) -> tuple[list[Chart], list[dict]]:
    kind = getattr(wb, "kind", "xlsx")
    if kind == "xls":
        return _xls(wb, sh)
    if kind in ("ods", "xlsb", "xls2", "wk1", "dbf", "mdb"):  # .xlsb の図（BrtDrawing の先）は読まず、シートの unread に書く
        # OpenDocument の表計算：OOXML の部品（drawing・oleObject）はない。シートの図形・グラフ（draw:frame）は、
        # odf/sheet.py がシートの shapes として読んで窓に出す
        return [], []
    return _xlsx(wb, sh)


# ====================================================================== .xlsx
def _xlsx(wb, sh) -> tuple[list[Chart], list[dict]]:
    z = wb.zip
    charts: list[Chart] = []
    objs: list[dict] = []
    if not sh.path or not z.has(sh.path):
        return charts, objs
    rels = rel_types(z, sh.path)
    for typ, target in rels.values():
        if typ == "drawing" and z.has(target):
            c, o = _drawing(wb, target)
            charts += c
            objs += o
    if sh.state.startswith("non-"):
        return charts, objs
    root = ET.fromstring(z.read(sh.path))
    seen = set()
    for ole in root.iter():
        if _ln(ole.tag) != "oleObject":
            continue
        rid = _attr(ole, "id")
        sid = ole.get("shapeId", "")
        if not rid or sid in seen or rid not in rels:
            continue
        typ, target = rels[rid]
        if typ not in ("oleObject", "package") or not z.has(target):
            continue
        seen.add(sid)
        pr = _kid(ole, "objectPr")
        where = ""
        anc = _kid(pr, "anchor") if pr is not None else None
        if anc is not None:
            where = _span(_kid(anc, "from"), _kid(anc, "to"))
        d = describe(z.read(target), target, ole.get("progId", ""), (pr.get("altText") or "") if pr is not None else "")
        d.update({"storage": target, "prog": ole.get("progId", ""), "where": where})
        objs.append(d)
    return charts, objs


def _cell(el) -> tuple[int, int] | None:
    if el is None:
        return None
    try:
        c = int(next(x.text for x in el if _ln(x.tag) == "col"))
        r = int(next(x.text for x in el if _ln(x.tag) == "row"))
    except (StopIteration, TypeError, ValueError):
        return None
    return r + 1, c + 1


def _span(frm, to) -> str:
    a, b = _cell(frm), _cell(to)
    if a is None:
        return ""
    return f"{addr(*a)}〜{addr(*b)}" if b else f"{addr(*a)} 付近"


def _drawing(wb, part: str) -> tuple[list[Chart], list[dict]]:
    z = wb.zip
    root = ET.fromstring(z.read(part))
    rels = rel_types(z, part)
    charts, objs = [], []
    for anchor in root.iter():
        if _ln(anchor.tag) not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
            continue
        where = _span(_kid(anchor, "from"), _kid(anchor, "to"))
        for g in _graphic_data(anchor):
            frame = g["frame"]
            nv = next((e for e in frame.iter() if _ln(e.tag) == "cNvPr"), None)
            name = nv.get("name", "") if nv is not None else ""
            alt = (nv.get("descr") or nv.get("title") or "") if nv is not None else ""
            target = rels.get(g["rid"], ("", ""))[1]
            if not target or not z.has(target):
                continue
            if g["kind"] == "chart":
                try:
                    ch = read_chart(z.read(target), target, z, workbook_resolver(wb), getattr(wb, "date1904", False))
                except Exception as e:  # noqa: BLE001
                    ch = Chart(part=target, error=str(e))
                ch.where, ch.name, ch.alt = where, name, alt
                charts.append(ch)
            else:
                from .smartart import smartart_lines

                objs.append({"kind": "smartart", "part": target, "where": where, "name": name, "alt": alt,
                             "lines": smartart_lines(z.read(target)), "label": "SmartArt"})
    return charts, objs


def _graphic_data(anchor):
    """図の中の graphicData（AlternateContent は Choice だけ）：{"kind": chart/smartart, "rid", "frame"}。"""
    out = []

    def walk(el, frame):
        for k in el:
            ln = _ln(k.tag)
            if ln == "AlternateContent":
                ch = _kid(k, "Choice")
                if ch is not None:
                    walk(ch, frame)
                continue
            if ln == "graphicFrame":
                frame = k
            if ln == "graphicData":
                for x in k:
                    if _ln(x.tag) == "chart":
                        out.append({"kind": "chart", "rid": _attr(x, "id"), "frame": frame if frame is not None else anchor})
                    elif _ln(x.tag) == "relIds":
                        out.append({"kind": "smartart", "rid": _attr(x, "dm"), "frame": frame if frame is not None else anchor})
                continue
            walk(k, frame)

    walk(anchor, None)
    return out


# ====================================================================== .xls
def _xls(wb, sh) -> tuple[list[Chart], list[dict]]:
    from .biffchart import sheet_charts

    charts = sheet_charts(wb, sh)
    objs = [] if sh.state.startswith("non-") else _xls_objects(wb, sh)
    return charts, objs


def pict_info(b: bytes) -> dict | None:
    """OBJ の中身から、絵（ot＝8）の FtPioGrbit と FtPictFmla：{"flags", "pos"（lPosInCtlStm）, "cls"（strClass）}。"""
    pos = 0
    out: dict = {}
    while pos + 4 <= len(b):
        ft, cb = struct.unpack_from("<HH", b, pos)
        body = pos + 4
        if ft == 0x15:
            if struct.unpack_from("<H", b, body)[0] != 0x08:
                return None
        elif ft == 0x08:
            out["flags"] = struct.unpack_from("<H", b, body)[0]
        elif ft == 0x09:
            cbf = struct.unpack_from("<H", b, body)[0]
            q = body + 2
            cce = struct.unpack_from("<H", b, q)[0] & 0x7FFF
            rgce = b[q + 6:q + 6 + cce]
            if rgce[:1] == b"\x02":  # PtgTbl：埋め込み（PictFmlaEmbedInfo と lPosInCtlStm がある）
                e = q + 6 + cce
                if e + 3 <= len(b) and b[e] == 0x03:
                    cch = b[e + 1]  # strClass の字の数（XLUnicodeStringNoCch：fHighByte の 1 バイトと字）
                    if cch and e + 3 < len(b):
                        high = b[e + 3] & 1
                        raw = b[e + 4:e + 4 + cch * (2 if high else 1)]
                        out["cls"] = raw.decode("utf-16-le" if high else "latin-1", errors="replace").rstrip("\x00")
                if body + 2 + cbf + 4 <= len(b):
                    out["pos"] = struct.unpack_from("<I", b, body + 2 + cbf)[0]
            break
        elif ft == 0x00:
            break
        pos = body + cb
    return out


def _xls_objects(wb, sh) -> list[dict]:
    from ..xls.book import _drawing_anchors
    from .biffchart import BOF, CONTINUE, EOF, substream

    try:
        lb = int(str(sh.path).lstrip("#"))
    except ValueError:
        return []
    recs = substream(wb.data, lb)
    draw, objs = [], []
    depth = 0
    for i, (_pos, t, b) in enumerate(recs):
        if t == BOF:
            depth += 1
            continue
        if t == EOF:
            depth -= 1
            continue
        if depth != 1:
            continue
        if t == 0x00EC:
            draw.append(b)
            j = i + 1
            while j < len(recs) and recs[j][1] == CONTINUE and recs[j - 1][1] in (0x00EC, CONTINUE):
                draw.append(recs[j][2])
                j += 1
        elif t == 0x005D:
            data = b
            j = i + 1
            while j < len(recs) and recs[j][1] == CONTINUE:
                data += recs[j][2]
                j += 1
            objs.append(pict_info(data))
    anchors = _drawing_anchors(b"".join(draw))
    out = []
    for k, info in enumerate(objs):
        if not info or "pos" not in info:
            continue
        flags = info.get("flags", 0)
        if flags & (0x02 | 0x10 | 0x20 | 0x80):  # fDde・fCtl（ActiveX）・fPrstm（Ctls）・fCamera
            continue
        name = f"MBD{info['pos']:08X}"
        a = anchors[k] if k < len(anchors) else {}
        where = f"{addr(a['r1'], a['c1'])}〜{addr(a['r2'], a['c2'])}" if a.get("r1") else ""
        st = storage_of(wb, name)
        if st is None:
            d = {"kind": "other", "label": "埋め込み", "class": info.get("cls", ""), "markdown": [],
                 "error": f"入れ物 {name} が見つかりません"}
        else:
            d = describe(st, name, info.get("cls", ""), a.get("alt", ""))
        d.update({"storage": name, "prog": info.get("cls", ""), "where": where})
        if a.get("name"):
            d["name"] = a["name"]
        out.append(d)
    return out


def storage_of(wb, name: str):
    """ブックの入れ物（MBD…）：複合ファイルのいちばん上、なければ暗号を解いた encryption ストリームの中。"""
    from ..cfb import CFB, SubStorage

    cfb = wb.cfb
    if any(p.startswith(name + "/") for p in cfb.streams()):
        return SubStorage(cfb, name)
    enc = next((p for p in cfb.streams() if p.lower() == "encryption"), None)
    if enc is None:
        return None
    from ..bincrypto import summary_entries

    for ename, is_stream, data in summary_entries(cfb.read(enc)):
        if ename == name and not is_stream and data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            return CFB(data)
    return None


del posixpath, re
