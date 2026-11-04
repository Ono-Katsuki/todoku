"""書類の中の画像（Word・Excel・PowerPoint・ODF・HWPX・旧 Office・RTF・PDF）を一覧にして、エージェントに渡す。

- list_images：窓に出す一覧（id・位置・大きさ・説明・部品の名前・形式）。書類を描かなくても、ファイルの作りだけで決まる
- get_image：id の画像の中身（view_image）。PNG・JPEG はそのまま（大きすぎるものは縮める）、ほかは PNG に描いて渡す
  （GIF・BMP・TIFF・WebP・JPEG 2000 は imaging.formats、EMF・WMF は imaging.metaraster。描けない記録は not_drawn で伝える）

id：窓の絵の id（図1・シート!図1・p1.i1）。部品の名前そのもの（word/media/image1.png）でも指せる。本文から指されていない
画像の部品（ヘッダーの絵など）は、部品の名前を id にする。旧形式（.doc・.xls）の絵は blip1 の形。
"""

from __future__ import annotations

import base64
import posixpath
import re
import struct
import time
import xml.etree.ElementTree as ET
import zlib

from ..limits import inflate

MAX_BYTES = 4_000_000  # エージェントに渡す画像のバイトの上限（API の画像の上限より少し下）
MAX_SIDE = 4000
_IMG_EXT = {"png", "jpg", "jpeg", "jpe", "gif", "bmp", "dib", "tif", "tiff", "emf", "wmf", "webp", "jp2", "j2k", "wdp", "svg", "pct", "pict", "ico", "emz", "wmz"}


class ImageNotFound(ValueError):
    pass


def _ext(name: str) -> str:
    return name.rsplit(".", 1)[-1].lower() if "." in name else ""


def _zip(blob: bytes):
    from ..zipraw import RawZip

    return RawZip(blob)


# ====================================================================== 一覧
def list_images(k: str, blob: bytes, win: dict) -> list[dict]:
    """窓の JSON に足す一覧：[{id, part, format, where, alt, decorative, size_mm}]。"""
    try:
        if k in ("docx", "pptx"):
            return _ooxml_pictures(blob, win, k)
        if k == "xlsx":
            return _xlsx_pictures(blob)
        if k == "xlsb":
            return _xlsb_pictures(blob)
        if k == "ppt":
            return [{kk: v for kk, v in b.items() if kk != "data"} for b in _legacy_blips(blob, k)]
        if k in ("odt", "ods"):
            return _odf_pictures(blob)
        if k == "hwpx":
            return _hwpx_pictures(blob)
        if k == "hwp":
            return _hwp5_pictures(blob)
        if k in ("doc", "xls"):
            return [{kk: v for kk, v in b.items() if kk != "data"} for b in _legacy_blips(blob, k)]
        if k == "rtf":
            return [{kk: v for kk, v in b.items() if kk != "data"} for b in _rtf_pictures(blob)]
    except Exception:  # noqa: BLE001  一覧を作れないだけ（窓は出す）
        return []
    return []


def _rest_of_parts(z, names: list[str], used: set, prefixes: tuple) -> list[dict]:
    """どこからも指されていない画像の部品も、部品の名前を id にして並べる（ヘッダー・フッター・図形の塗りなど）。"""
    out = []
    for n in names:
        if n in used or _ext(n) not in _IMG_EXT or not n.startswith(prefixes):
            continue
        out.append({"id": n, "part": n, "format": _ext(n), "where": "本文の絵の一覧にない部品（ヘッダー・フッター・図形の塗りなど）", "alt": None, "decorative": False})
    return out


def _ooxml_pictures(blob: bytes, win: dict, k: str) -> list[dict]:
    z = _zip(blob)
    names = list(z.names())
    out = []
    used = set()
    for p in win.get("pictures") or []:
        media = p.get("media") or ""
        if media:
            used.add(media)
        out.append({"id": p["id"], "part": media, "format": _ext(media), "where": p.get("where", ""), "alt": p.get("alt"), "decorative": p.get("decorative", False),
                    **({"size_mm": p["size_mm"]} if p.get("size_mm") else {})})
    return out + _rest_of_parts(z, names, used, ("word/", "ppt/"))


_XDR = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _rels(z, part: str) -> dict:
    d, f = posixpath.split(part)
    rp = posixpath.join(d, "_rels", f + ".rels")
    try:
        root = ET.fromstring(z.read(rp))
    except Exception:  # noqa: BLE001
        return {}
    out = {}
    for r in root.iter(_REL + "Relationship"):
        t = r.get("Target") or ""
        if r.get("TargetMode") == "External":
            continue
        out[r.get("Id")] = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(d, t))
    return out


def _xlsx_pictures(blob: bytes) -> list[dict]:
    from ..formula.parser import addr

    z = _zip(blob)
    names = list(z.names())
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    wr = _rels(z, "xl/workbook.xml")
    out = []
    used = set()
    for sh in wb.iter(_MAIN + "sheet"):
        name = sh.get("name") or ""
        part = wr.get(sh.get(_R + "id"))
        if not part or part not in set(names):
            continue
        sr = _rels(z, part)
        try:
            root = ET.fromstring(z.read(part))
        except Exception:  # noqa: BLE001
            continue
        n = 0
        for d in root.iter(_MAIN + "drawing"):
            dp = sr.get(d.get(_R + "id"))
            if not dp:
                continue
            droot = ET.fromstring(z.read(dp))
            dr = _rels(z, dp)
            for anchor in droot:
                frm = anchor.find(_XDR + "from")
                pos = None
                if frm is not None:
                    pos = (int(frm.find(_XDR + "row").text) + 1, int(frm.find(_XDR + "col").text) + 1)
                ext = anchor.find(_XDR + "ext")
                for pic in anchor.iter(_XDR + "pic"):
                    n += 1
                    nv = pic.find(".//" + _XDR + "cNvPr")
                    pname = (nv.get("name") or "") if nv is not None else ""
                    alt = ((nv.get("descr") or nv.get("title") or "") if nv is not None else "").strip() or None
                    deco = nv is not None and any(e.tag.endswith("}decorative") and e.get("val") in ("1", "true") for e in nv.iter())
                    blip = pic.find(".//" + _A + "blip")
                    media = dr.get(blip.get(_R + "embed")) if blip is not None else ""
                    if media:
                        used.add(media)
                    size = None
                    x = pic.find(".//" + _A + "ext")
                    for e in (x, ext):
                        if e is not None and e.get("cx") and e.get("cy"):
                            size = [round(int(e.get("cx")) / 36000, 1), round(int(e.get("cy")) / 36000, 1)]
                            break
                    item = {"id": f"{name}!{pname or f'図{n}'}", "part": media or "", "format": _ext(media or ""), "alt": alt, "decorative": deco,
                            "where": f"シート「{name}」" + (f"の {addr(*pos)} 付近" if pos else "")}
                    if size:
                        item["size_mm"] = size
                    out.append(item)
    return out + _rest_of_parts(z, names, used, ("xl/",))


def _xlsb_pictures(blob: bytes) -> list[dict]:
    """.xlsb：シートは BIFF12 だが、図形の部品は xlsx と同じ DrawingML。シートの関係から drawing の部品をたどる。"""
    from ..formula.parser import addr
    from ..xlsb.book import XlsbWorkbook

    wb = XlsbWorkbook(blob, "")
    z = wb.zip
    names = list(z.names())
    out: list[dict] = []
    used: set = set()
    for sh in wb.sheets:
        if sh.state.startswith("non-worksheet"):
            continue
        n = 0
        for typ, dp in wb._rels(sh.path).values():
            if typ != "drawing" or not z.has(dp):
                continue
            droot = ET.fromstring(z.read(dp))
            dr = _rels(z, dp)
            for anchor in droot:
                frm = anchor.find(_XDR + "from")
                pos = (int(frm.find(_XDR + "row").text) + 1, int(frm.find(_XDR + "col").text) + 1) if frm is not None else None
                for pic in anchor.iter(_XDR + "pic"):
                    n += 1
                    nv = pic.find(".//" + _XDR + "cNvPr")
                    pname = (nv.get("name") or "") if nv is not None else ""
                    alt = ((nv.get("descr") or nv.get("title") or "") if nv is not None else "").strip() or None
                    blip = pic.find(".//" + _A + "blip")
                    media = dr.get(blip.get(_R + "embed")) if blip is not None else ""
                    if media:
                        used.add(media)
                    out.append({"id": f"{sh.name}!{pname or f'図{n}'}", "part": media or "", "format": _ext(media or ""), "alt": alt, "decorative": False,
                                "where": f"シート「{sh.name}」" + (f"の {addr(*pos)} 付近" if pos else "")})
    return out + _rest_of_parts(z, names, used, ("xl/",))


_ODF_DRAW = "{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}"
_ODF_XLINK = "{http://www.w3.org/1999/xlink}"
_ODF_SVG = "{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}"


def _odf_length_mm(v: str | None) -> float | None:
    m = re.fullmatch(r"\s*([\d.]+)\s*(mm|cm|in|pt|pc|px)?\s*", v or "")
    if not m:
        return None
    k = {"mm": 1.0, "cm": 10.0, "in": 25.4, "pt": 25.4 / 72, "pc": 25.4 / 6, "px": 25.4 / 96, None: 1.0}[m[2]]
    return round(float(m[1]) * k, 1)


def _odf_pictures(blob: bytes) -> list[dict]:
    z = _zip(blob)
    names = list(z.names())
    out = []
    used = set()
    root = ET.fromstring(z.read("content.xml"))
    n = 0
    for fr in root.iter(_ODF_DRAW + "frame"):
        img = fr.find(_ODF_DRAW + "image")
        if img is None:
            continue
        n += 1
        href = (img.get(_ODF_XLINK + "href") or "").lstrip("./")
        title = fr.find(_ODF_SVG + "title")
        desc = fr.find(_ODF_SVG + "desc")
        alt = " ".join(x for x in (("".join(title.itertext()) if title is not None else "").strip(), ("".join(desc.itertext()) if desc is not None else "").strip()) if x) or None
        if href:
            used.add(href)
        w, h = _odf_length_mm(fr.get(_ODF_SVG + "width")), _odf_length_mm(fr.get(_ODF_SVG + "height"))
        item = {"id": f"図{n}", "part": href, "format": _ext(href), "where": f"文書の {n} 番目の絵（{fr.get(_ODF_DRAW + 'name') or '名前なし'}）", "alt": alt,
                "decorative": (fr.get("{urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0}decorative") or "").lower() == "true"}
        if w and h:
            item["size_mm"] = [w, h]
        out.append(item)
    return out + _rest_of_parts(z, names, used, ("Pictures/", "media/"))


def _hwpx_manifest(z) -> dict:
    """Contents/content.hpf の id → 部品。"""
    out = {}
    try:
        root = ET.fromstring(z.read("Contents/content.hpf"))
    except Exception:  # noqa: BLE001
        return out
    for it in root.iter():
        if it.tag.endswith("}item") and it.get("id") and it.get("href"):
            out[it.get("id")] = it.get("href")
    return out


def _hwpx_pictures(blob: bytes) -> list[dict]:
    z = _zip(blob)
    names = list(z.names())
    man = _hwpx_manifest(z)
    out = []
    used = set()
    n = 0
    for part in sorted(x for x in names if re.fullmatch(r"Contents/section\d+\.xml", x)):
        root = ET.fromstring(z.read(part))
        for pic in root.iter():
            if not pic.tag.endswith("}pic"):
                continue
            n += 1
            img = next((e for e in pic.iter() if e.tag.endswith("}img")), None)
            ref = img.get("binaryItemIDRef") if img is not None else None
            media = man.get(ref or "", "") if ref else ""
            if media:
                used.add(media)
            sc = next((e for e in pic if e.tag.endswith("}shapeComment")), None)
            alt = ("".join(sc.itertext()).strip() or None) if sc is not None else None
            sz = next((e for e in pic if e.tag.endswith("}sz")), None)
            item = {"id": f"図{n}", "part": media, "format": _ext(media), "where": f"{posixpath.basename(part)} の {n} 番目の絵", "alt": alt, "decorative": False}
            if sz is not None and sz.get("width") and sz.get("height"):
                item["size_mm"] = [round(int(sz.get("width")) / 7200 * 25.4, 1), round(int(sz.get("height")) / 7200 * 25.4, 1)]
            out.append(item)
    return out + _rest_of_parts(z, names, used, ("BinData/",))


def _hwp5_pictures(blob: bytes) -> list[dict]:
    from ..cfb import CFB

    c = CFB(blob)
    out = []
    for s in c.streams():
        if s.lower().startswith("bindata/") or s.lower().startswith("bindata\\"):
            out.append({"id": s, "part": s, "format": _ext(s), "where": "HWP の BinData のストリーム", "alt": None, "decorative": False})
    return out


# ---------------------------------------------------------------------- 旧形式：OfficeArt の BLIP（MS-ODRAW 2.2.23〜）
_BLIP_FORMAT = {0xF01A: "emf", 0xF01B: "wmf", 0xF01C: "pict", 0xF01D: "jpeg", 0xF01E: "png", 0xF01F: "dib", 0xF029: "tiff", 0xF02A: "jpeg"}
_TWO_UID = {0x3D5, 0x217, 0x543, 0x46B, 0x6E3, 0x6E1, 0x7A9, 0x6E5}


def _blips(buf: bytes) -> list[tuple[str, bytes]]:
    """OfficeArt の記録の見出しをたどって BLIP を拾う（入れ物（ver=0xF）の中に入る。形式の印で中身を確かめる）。"""
    out: list = []

    def walk(pos: int, end: int, depth: int) -> None:
        while pos + 8 <= end:
            vi, typ, ln = struct.unpack_from("<HHI", buf, pos)
            ver = vi & 15
            inst = vi >> 4
            body = pos + 8
            if body + ln > end and typ not in _BLIP_FORMAT:
                return
            fmt = _BLIP_FORMAT.get(typ)
            if fmt:
                out_data = _blip_payload(buf[body:body + ln], fmt, inst)
                if out_data:
                    out.append((fmt, out_data))
            elif typ == 0xF007 and ln >= 36 and depth < 8:
                # OfficeArtFBSE（MS-ODRAW 2.2.32）：36 バイトの見出し（33 バイト目が名前の長さ）の後に BLIP の記録
                walk(body + 36 + buf[body + 33], body + ln, depth + 1)
            elif ver == 15 and depth < 8 and ln >= 8:
                walk(body, body + ln, depth + 1)
            if ln == 0 and ver != 15 and typ < 0xF000:
                return
            pos = body + ln

    walk(0, len(buf), 0)
    return out


def _blip_payload(d: bytes, fmt: str, inst: int) -> bytes | None:
    uid = 32 if inst in _TWO_UID else 16
    if fmt in ("emf", "wmf", "pict"):
        at = uid
        if len(d) < at + 34:
            return None
        comp = d[at + 32]
        raw = d[at + 34:]
        if comp == 0:
            raw = inflate(raw, what="OfficeArt の絵", tolerant=True)
        if fmt == "wmf" and raw[:4] != b"\xd7\xcd\xc6\x9a":
            # MS-ODRAW の WMF は置ける見出しを持たない：metafileHeader の rcBounds から足す
            l_, t_, r_, b_ = struct.unpack_from("<4i", d, at + 4)
            # rcBounds が 16 ビットに入らないとき（EMU の値など）は、1 ポイント＝12700 の単位として直す
            big = max(abs(l_), abs(t_), abs(r_), abs(b_)) > 32767
            k = 12700 if big else 1
            raw = struct.pack("<IH4hHIH", 0x9AC6CDD7, 0, l_ // k, t_ // k, r_ // k, b_ // k, 72 if big else 1440, 0, 0) + raw
        return raw
    at = uid + 1
    if fmt == "dib":
        return d[at:]
    return d[at:]


def _legacy_blips(blob: bytes, k: str) -> list[dict]:
    from ..cfb import CFB

    c = CFB(blob)
    streams = {s.lower(): s for s in c.streams()}
    found: list = []
    if k == "doc":
        for key in ("worddocument", "data", "0table", "1table"):
            if key in streams:
                found += _blips_in_word_stream(c.read(streams[key]))
    elif k == "ppt":
        if "pictures" in streams:  # Pictures ストリーム（MS-PPT 2.1.2）：BLIP の記録の並び
            found += _blips(c.read(streams["pictures"]))
    else:
        name = streams.get("workbook") or streams.get("book")
        if name:
            buf = _biff_drawing_groups(c.read(name))
            found += _blips(buf)
    out = []
    seen = set()
    for fmt, data in found:
        key = zlib.crc32(data), len(data)
        if key in seen:
            continue
        seen.add(key)
        n = len(out) + 1
        out.append({"id": f"blip{n}", "part": f"OfficeArt の絵 {n}", "format": fmt, "where": "文書の中の絵（この形式では、位置を部品から決めません）", "alt": None, "decorative": False,
                    "data": data})
    return out


def _blips_in_word_stream(buf: bytes) -> list:
    """Word の Data・表のストリーム：先頭から入れ物の記録の並びではないので、OfficeArt の BLIP の入れ物の型の並びを探して、見出しで確かめる。"""
    out = []
    for typ, fmt in _BLIP_FORMAT.items():
        sig = struct.pack("<H", typ)
        i = buf.find(sig)
        while i >= 2:
            start = i - 2
            vi = struct.unpack_from("<H", buf, start)[0]
            if vi & 15 == 0 and start + 8 <= len(buf):
                ln = struct.unpack_from("<I", buf, start + 4)[0]
                inst = vi >> 4
                if ln and start + 8 + ln <= len(buf) and _inst_ok(fmt, inst):
                    data = _blip_payload(buf[start + 8:start + 8 + ln], fmt, inst)
                    if data and _plausible(fmt, data):
                        out.append((start, fmt, data))
                        i = buf.find(sig, start + 8 + ln)
                        continue
            i = buf.find(sig, i + 2)
    return [(f, d) for _s, f, d in sorted(out, key=lambda t: t[0])]


_INST = {"emf": (0x3D4, 0x3D5), "wmf": (0x216, 0x217), "pict": (0x542, 0x543), "jpeg": (0x46A, 0x46B, 0x6E2, 0x6E3), "png": (0x6E0, 0x6E1), "dib": (0x7A8, 0x7A9), "tiff": (0x6E4, 0x6E5)}


def _inst_ok(fmt: str, inst: int) -> bool:
    return inst in _INST.get(fmt, ())


def _plausible(fmt: str, d: bytes) -> bool:
    if fmt == "png":
        return d[:8] == b"\x89PNG\r\n\x1a\n"
    if fmt == "jpeg":
        return d[:2] == b"\xff\xd8"
    if fmt == "emf":
        return d[40:44] == b" EMF"
    if fmt == "tiff":
        return d[:2] in (b"II", b"MM")
    return bool(d)


def _biff_drawing_groups(wb: bytes) -> bytes:
    """Workbook：MsoDrawingGroup（0x00EB）と、続きの Continue（0x003C）をつなぐ。"""
    out = bytearray()
    pos = 0
    n = len(wb)
    collecting = False
    while pos + 4 <= n:
        rec, ln = struct.unpack_from("<HH", wb, pos)
        body = wb[pos + 4:pos + 4 + ln]
        if rec == 0x00EB:
            out += body
            collecting = True
        elif rec == 0x003C and collecting:
            out += body
        else:
            collecting = False
        pos += 4 + ln
    return bytes(out)


_RTF_PICT = re.compile(rb"\{\\(?:\*\\shppict\{\\)?pict\b")


def _rtf_pictures(blob: bytes) -> list[dict]:
    """RTF の \\pict の群：形式の印と、16 進（または \\bin）の中身。"""
    out = []
    i = 0
    n = len(blob)
    while True:
        i = blob.find(b"\\pict", i)
        if i < 0:
            break
        j = i + 5
        # 群の終わりまで
        depth = 1
        k = j
        while k < n and depth:
            c = blob[k]
            if c == 0x5C:
                k += 2
                continue
            if c == 0x7B:
                depth += 1
            elif c == 0x7D:
                depth -= 1
            k += 1
        body = blob[j:k - 1]
        hexpart = body
        while True:  # 入れ子の群（\picprop の {\sp{\sn ..}{\sv ..}}）を内側から消す
            hexpart, nsub = re.subn(rb"\{[^{}]*\}", b" ", hexpart)  # 空白に：直前の命令の数と、次の 16 進の字がつながらないように
            if not nsub:
                break
        words = re.findall(rb"\\[a-zA-Z]+-?\d*", hexpart)
        names = {w.rstrip(b"-0123456789") for w in words}
        fmt = ("png" if b"\\pngblip" in names else "jpeg" if b"\\jpegblip" in names else "emf" if b"\\emfblip" in names else "wmf" if b"\\wmetafile" in names
               else "dib" if b"\\dibitmap" in names else "")
        digits = re.sub(rb"\\[a-zA-Z]+-?\d* ?|\s", b"", hexpart)
        digits = re.sub(rb"[^0-9a-fA-F]", b"", digits)
        data = bytes.fromhex(digits[: len(digits) // 2 * 2].decode("ascii")) if digits else b""
        wg = re.search(rb"\\picwgoal(\d+)", body)
        hg = re.search(rb"\\pichgoal(\d+)", body)
        if data:
            num = len(out) + 1
            item = {"id": f"blip{num}", "part": f"\\pict の絵 {num}", "format": fmt or "不明", "where": "RTF の本文の中の絵", "alt": None, "decorative": False, "data": data}
            if wg and hg:
                item["size_mm"] = [round(int(wg[1]) / 1440 * 25.4, 1), round(int(hg[1]) / 1440 * 25.4, 1)]
            out.append(item)
        i = k
    return out


# ====================================================================== 窓に出す
def markdown(images: list[dict]) -> str:
    if not images:
        return ""
    lines = ["", f"### 書類の中の画像（{len(images)} 個。view_image で見られます）", ""]
    for im in images:
        size = f"約 {im['size_mm'][0]:g}×{im['size_mm'][1]:g} mm、" if im.get("size_mm") else ""
        alt = f"説明「{im['alt']}」" if im.get("alt") else ("飾り" if im.get("decorative") else "説明なし")
        lines.append(f"- {im['id']}（{im.get('where', '')}、{size}{im.get('format') or '形式不明'}、部品 {im.get('part') or '不明'}）：{alt}")
    return "\n".join(lines) + "\n"


# ====================================================================== 中身を渡す
def _as_output(data: bytes, fmt_hint: str, region: list | None = None) -> dict:
    """PNG・JPEG はそのまま（大きすぎるときは縮める）、ほかは描いて PNG に。{mime, data, width, height, not_drawn, notes}。"""
    from .formats import Frame, ImageError, decode, sniff

    sn = sniff(data)
    if fmt_hint == "dib" and sn is None:
        # 見出しのない DIB
        from .formats import dib_frame

        fr = dib_frame(data, 0, None)
        return _encode(fr, region)
    if sn is None and (fmt_hint in ("emf", "wmf") or data[40:44] == b" EMF" or data[:4] == b"\xd7\xcd\xc6\x9a" or data[:4] == b"\x01\x00\x09\x00"):
        from .metaraster import emf, wmf

        fr = emf(data) if fmt_hint == "emf" or data[40:44] == b" EMF" else wmf(data)
        return _encode(fr, region)
    if sn in ("png", "jpeg") and len(data) <= MAX_BYTES and region is None:
        try:
            from .formats import info

            _f, fr_i = info(data)
            w, h = fr_i[0].w, fr_i[0].h
        except Exception:  # noqa: BLE001
            w, h = 0, 0
        if max(w, h) <= MAX_SIDE:
            return {"mime": "image/png" if sn == "png" else "image/jpeg", "data": data, "width": w, "height": h, "not_drawn": [], "notes": []}
    if sn is None:
        raise ImageError(f"この形式（{fmt_hint or '不明'}）の絵は描けません")
    return _encode(decode(data, 0, MAX_SIDE * MAX_SIDE, time.monotonic() + 60.0), region)


def _encode(fr, region: list | None) -> dict:
    from ..pdf.pageimage import _png_rgb
    from ..pdf.render.image import area_down

    w, h, px = fr.w, fr.h, fr.rgb
    notes = list(fr.notes)
    if region is not None:
        if not (isinstance(region, list) and len(region) == 4):
            raise ValueError("region は [左, 上, 右, 下]（画素）で指定してください")
        x0, y0, x1, y1 = [int(v) for v in region]
        x0, x1 = sorted((max(0, x0), min(w, x1)))
        y0, y1 = sorted((max(0, y0), min(h, y1)))
        if x1 - x0 < 1 or y1 - y0 < 1:
            raise ValueError(f"範囲が画像（{w}×{h} 画素）の外です")
        px = b"".join(px[(y * w + x0) * 3:(y * w + x1) * 3] for y in range(y0, y1))
        w, h = x1 - x0, y1 - y0
    k = max(w, h)
    if k > MAX_SIDE:
        W, H = max(1, w * MAX_SIDE // k), max(1, h * MAX_SIDE // k)
        px = area_down(px, w, h, 3, W, H)
        notes.append(f"大きな画像なので {w}×{h} 画素を {W}×{H} 画素に縮めました（region で範囲を指定すると、元の細かさで見られます）")
        w, h = W, H
    png = _png_rgb(px, w, h)
    while len(png) > MAX_BYTES and max(w, h) > 400:
        W, H = max(1, w * 3 // 4), max(1, h * 3 // 4)
        px = area_down(px, w, h, 3, W, H)
        w, h = W, H
        png = _png_rgb(px, w, h)
        notes.append(f"大きな画像なので {w}×{h} 画素に縮めました（region で範囲を指定すると、元の細かさで見られます）")
    return {"mime": "image/png", "data": png, "width": w, "height": h, "not_drawn": list(fr.problems), "notes": notes}


def _part_bytes(k: str, blob: bytes, part: str) -> bytes:
    if k in ("docx", "xlsx", "xlsb", "pptx", "odt", "ods", "hwpx"):
        z = _zip(blob)
        data = z.read(part)
        if _ext(part) in ("emz", "wmz"):
            data = inflate(data, wbits=31, what="圧縮した絵", tolerant=True)
        return data
    if k == "hwp":
        from ..cfb import CFB

        c = CFB(blob)
        raw = c.read(part)
        d = zlib.decompressobj(-15)
        try:
            out = d.decompress(raw)
            if out:
                return out
        except zlib.error:
            pass
        return raw
    raise ImageNotFound(part)


def get_image(k: str, blob: bytes, images: list[dict], image_id: str, region: list | None = None) -> dict:
    """id の画像を渡す。{mime, data_base64, width, height, not_drawn, text}。"""
    item = next((im for im in images if im.get("id") == image_id or (im.get("part") and im.get("part") == image_id)), None)
    if item is None:
        ids = "・".join(str(im["id"]) for im in images[:12])
        raise ImageNotFound(f"画像「{image_id}」は書類にありません" + (f"（窓の一覧の id：{ids}）" if ids else "（この書類には画像の一覧がありません）"))
    if k in ("doc", "xls", "rtf", "ppt"):
        src = _legacy_blips(blob, k) if k != "rtf" else _rtf_pictures(blob)
        hit = next((b for b in src if b["id"] == item["id"]), None)
        if hit is None:
            raise ImageNotFound(f"画像「{image_id}」を読み出せません")
        data = hit["data"]
    else:
        if not item.get("part"):
            raise ImageNotFound(f"画像「{image_id}」の部品が見つかりません（外部のファイルへのリンクの絵は、書類の中にありません）")
        data = _part_bytes(k, blob, item["part"])
    return _deliver(data, item.get("format") or "", image_id, item, region)


def _deliver(data: bytes, fmt: str, image_id: str, item: dict, region) -> dict:
    out = _as_output(data, fmt, region)
    text = (f"画像 {image_id}（{item.get('where') or ''}"
            + (f"・部品 {item['part']}" if item.get("part") else "") + f"）。{out['width']}×{out['height']} 画素。"
            + (f"説明「{item['alt']}」。" if item.get("alt") else "")
            + "中身は画像の中にあります。読めない所は推測で埋めず、本人に伝えてください。")
    if region is None and out["mime"] != "image/png" or len(out["data"]) == len(data):
        text += " 元のファイルのバイトのままです。"
    if out["notes"]:
        text += " " + " ".join(out["notes"])
    if out["not_drawn"]:
        text += " 描けなかったもの：" + "／".join(out["not_drawn"][:8])
    return {"id": image_id, "mime": out["mime"], "data_base64": base64.b64encode(out["data"]).decode("ascii"), "width": out["width"], "height": out["height"],
            "not_drawn": list(out["not_drawn"]), "text": text}


# ====================================================================== PDF
def pdf_images(doc) -> list[dict]:
    """PDF の、ページの画像と、注釈の見た目の画像の一覧（id は p1.i1・p1.a1）。"""
    from ..pdf.content import page_content
    from ..pdf.layout import View
    from ..pdf.pageimage import _annot_pictures

    out = []
    for n, page in enumerate(doc.pages(), 1):
        pc = page_content(doc, page)
        view = View(doc, page)
        for k, p in enumerate(pc.pictures, 1):
            x0, y0, x1, y1 = view.rect(p.x0, p.y0, p.x1, p.y1)
            if x1 <= 0 or y1 <= 0 or x0 >= view.width or y0 >= view.height:
                continue
            out.append({"id": f"p{n}.i{k}", "page": n, "pic": p, "where": f"{n} ページ", "format": "pdf",
                        "size_mm": [round((x1 - x0) / view.mm, 1), round((y1 - y0) / view.mm, 1)], "at_mm": [round(x0 / view.mm, 1), round(y0 / view.mm, 1)]})
        stamps: list = []
        pics = _annot_pictures(doc, page, stamps)
        for j, p in enumerate(pics, 1):
            x0, y0, x1, y1 = view.rect(p.x0, p.y0, p.x1, p.y1)
            out.append({"id": f"p{n}.a{j}", "page": n, "pic": p, "where": f"{n} ページの注釈の見た目", "format": "pdf",
                        "size_mm": [round((x1 - x0) / view.mm, 1), round((y1 - y0) / view.mm, 1)], "at_mm": [round(x0 / view.mm, 1), round(y0 / view.mm, 1)]})
    return out


def pdf_annotation_images(doc) -> list[dict]:
    """窓の一覧に足す、注釈の見た目の画像（ページの画像は、窓のページごとの「画像」の節に p1.i1 の id で出ている）。"""
    from ..pdf.layout import View
    from ..pdf.pageimage import _annot_pictures

    out = []
    for n, page in enumerate(doc.pages(), 1):
        if not doc.resolve(page.get("Annots")):
            continue
        view = View(doc, page)
        for j, p in enumerate(_annot_pictures(doc, page, []), 1):
            x0, y0, x1, y1 = view.rect(p.x0, p.y0, p.x1, p.y1)
            out.append({"id": f"p{n}.a{j}", "part": "注釈の見た目（/AP）", "format": "PDF の画像", "where": f"{n} ページの左 {x0 / view.mm:.0f}・上 {y0 / view.mm:.0f} mm",
                        "alt": None, "decorative": False, "size_mm": [round((x1 - x0) / view.mm, 1), round((y1 - y0) / view.mm, 1)]})
    return out


def pdf_get(doc, image_id: str, region: list | None = None) -> dict:
    from .formats import Frame
    from ..pdf.objects import Stream
    from ..pdf.raster import RasterError
    from ..pdf.render.image import image_rgba

    items = pdf_images(doc)
    item = next((im for im in items if im["id"] == image_id), None)
    if item is None:
        ids = "・".join(im["id"] for im in items[:12])
        raise ImageNotFound(f"画像「{image_id}」は PDF にありません" + (f"（id：{ids}）" if ids else "（画像がありません）"))
    image = item["pic"].image
    # JPEG だけのものは元のバイトのまま
    if isinstance(image, Stream) and region is None:
        f = doc.resolve(image.dict.get("Filter"))
        names = [getattr(doc.resolve(x), "name", "") for x in (f if isinstance(f, list) else [f] if f is not None else [])]
        if names and names[-1] in ("DCTDecode", "DCT") and not image.dict.get("SMask") and not image.dict.get("Mask"):
            data = doc.decode(image)
            cs = doc.resolve(image.dict.get("ColorSpace"))
            if getattr(cs, "name", "") in ("DeviceRGB", "DeviceGray") and len(data) <= MAX_BYTES:
                return _deliver(data, "jpeg", image_id, {"where": item["where"]}, None)
    try:
        im = image_rgba(doc, image, 16_000_000, deadline=time.monotonic() + 60.0)
    except RasterError as e:
        raise ImageNotFound(f"画像「{image_id}」を開けません：{e}") from None
    if im.stencil:
        g = im.alpha.translate(bytes(255 - i for i in range(256)))
        rgb = bytearray(im.w * im.h * 3)
        rgb[0::3] = rgb[1::3] = rgb[2::3] = g
        px = bytes(rgb)
    else:
        px = im.px
        if im.alpha is not None:
            from .formats import _on_white

            px = _on_white(px, im.alpha)
    fr = Frame(im.w, im.h, px, None, list(im.notes), list(im.problems))
    out = _encode(fr, region)
    text = (f"画像 {image_id}（{item['where']}・左 {item['at_mm'][0]}・上 {item['at_mm'][1]} mm・約 {item['size_mm'][0]}×{item['size_mm'][1]} mm）。{out['width']}×{out['height']} 画素。"
            "読めない所は推測で埋めず、本人に伝えてください。")
    if out["notes"]:
        text += " " + " ".join(out["notes"])
    if out["not_drawn"]:
        text += " 描けなかったもの：" + "／".join(out["not_drawn"][:8])
    return {"id": image_id, "mime": out["mime"], "data_base64": base64.b64encode(out["data"]).decode("ascii"), "width": out["width"], "height": out["height"],
            "not_drawn": list(out["not_drawn"]), "text": text}
