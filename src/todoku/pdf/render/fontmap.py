"""PDF の書体の字の符号 → 字の形（ISO 32000-1 9.6.6 単純な書体の符号・9.7.4 CIDFont・9.9 埋め込み）。

- 埋め込んだ書体：Type 1（FontFile）・CFF（FontFile3 の Type1C・CIDFontType0C・OpenType）・TrueType（FontFile2）
  - 単純な書体の字の名前：Differences → 基の符号（Standard・WinAnsi・MacRoman）→ 書体の組み込みの符号
  - TrueType の単純な書体：9.6.6.4 の順（(3,1) の cmap を字の名前の Unicode で・(3,0) を符号・0xF000 で・(1,0) を符号で・post の名前）
  - CIDFont：CMap で CID に → CIDFontType2 は CIDToGIDMap、CFF の CID の書体は charset で字形の番号に
- 埋め込まない書体（標準の 14 書体・CJK の書体など）：このパソコンの書体で代わりに描く（字の Unicode で引く）。
  書体の記述（FontDescriptor の Flags：固定幅・セリフ、太さ、CIDSystemInfo）で、似た書体を選ぶ。見つからなければ描かない
  （paint.py は字の四角で示し、その数を伝える）
"""

from __future__ import annotations

import os
import struct
from functools import lru_cache

from ..fonts import decode_predefined
from ..objects import Name, Stream
from .fontprog import CFFOutlines, FontProgramError, TrueTypeOutlines, Type1Outlines, sfnt_tables

_UNKNOWN = "\ufffd"


# ====================================================================== このパソコンの書体
def _win(name: str) -> str:
    root = os.environ.get("WINDIR") or os.environ.get("SystemRoot") or "C:\\Windows"
    return root.rstrip("\\/") + "\\Fonts\\" + name


_MAC, _SUP = "/System/Library/Fonts/", "/System/Library/Fonts/Supplemental/"
_LIN = ("/usr/share/fonts/truetype/", "/usr/share/fonts/TTF/", "/usr/share/fonts/", "/usr/share/fonts/opentype/")

CANDIDATES: dict[str, list[str]] = {
    "sans": [_SUP + "Arial.ttf", "/Library/Fonts/Arial.ttf", _MAC + "Helvetica.ttc",
             *[d + "liberation/LiberationSans-Regular.ttf" for d in _LIN], *[d + "dejavu/DejaVuSans.ttf" for d in _LIN],
             _win("arial.ttf")],
    "sans-bold": [_SUP + "Arial Bold.ttf", _MAC + "Helvetica.ttc", *[d + "liberation/LiberationSans-Bold.ttf" for d in _LIN],
                  *[d + "dejavu/DejaVuSans-Bold.ttf" for d in _LIN], _win("arialbd.ttf")],
    "serif": [_SUP + "Times New Roman.ttf", _MAC + "Times.ttc", *[d + "liberation/LiberationSerif-Regular.ttf" for d in _LIN],
              *[d + "dejavu/DejaVuSerif.ttf" for d in _LIN], _win("times.ttf")],
    "serif-bold": [_SUP + "Times New Roman Bold.ttf", _MAC + "Times.ttc",
                   *[d + "liberation/LiberationSerif-Bold.ttf" for d in _LIN], _win("timesbd.ttf")],
    "mono": [_SUP + "Courier New.ttf", _MAC + "Courier.ttc", *[d + "liberation/LiberationMono-Regular.ttf" for d in _LIN],
             *[d + "dejavu/DejaVuSansMono.ttf" for d in _LIN], _win("cour.ttf")],
    "mono-bold": [_SUP + "Courier New Bold.ttf", _MAC + "Courier.ttc", *[d + "liberation/LiberationMono-Bold.ttf" for d in _LIN],
                  _win("courbd.ttf")],
    "cjk-sans": [_MAC + "ヒラギノ角ゴシック W3.ttc", *[d + "noto-cjk/NotoSansCJK-Regular.ttc" for d in _LIN],
                 *[d + "noto/NotoSansCJK-Regular.ttc" for d in _LIN], "/usr/share/fonts/opentype/ipaexfont-gothic/ipaexg.ttf",
                 *[d + "ipaexfont/ipaexg.ttf" for d in _LIN], *[d + "fonts-japanese-gothic.ttf" for d in _LIN],
                 _win("YuGothR.ttc"), _win("msgothic.ttc"), _SUP + "Arial Unicode.ttf"],
    "cjk-sans-bold": [_MAC + "ヒラギノ角ゴシック W6.ttc", *[d + "noto-cjk/NotoSansCJK-Bold.ttc" for d in _LIN],
                      _win("YuGothB.ttc")],
    "cjk-serif": [_MAC + "ヒラギノ明朝 ProN.ttc", *[d + "noto-cjk/NotoSerifCJK-Regular.ttc" for d in _LIN],
                  "/usr/share/fonts/opentype/ipaexfont-mincho/ipaexm.ttf", *[d + "ipaexfont/ipaexm.ttf" for d in _LIN],
                  *[d + "fonts-japanese-mincho.ttf" for d in _LIN], _win("yumin.ttf"), _win("msmincho.ttc")],
    "symbol": [_MAC + "Symbol.ttf", _win("symbol.ttf")],
    "dingbats": [_MAC + "ZapfDingbats.ttf"],
    "fallback": [_SUP + "Arial Unicode.ttf", "/Library/Fonts/Arial Unicode.ttf", _MAC + "ヒラギノ角ゴシック W3.ttc",
                 *[d + "dejavu/DejaVuSans.ttf" for d in _LIN], *[d + "noto-cjk/NotoSansCJK-Regular.ttc" for d in _LIN],
                 _win("seguisym.ttf"), _win("msgothic.ttc")],
}


class SystemFace:
    """このパソコンの書体の 1 つ（Unicode → 字形の番号 → 輪郭）。"""

    def __init__(self, path: str, outl, cmap: dict[int, int], name: str):
        self.path, self.outl, self.cmap, self.name = path, outl, cmap, name

    def lookup(self, text: str):
        if not text:
            return None
        cp = ord(text[0])
        g = self.cmap.get(cp)
        if g is None and 0xF000 <= cp <= 0xF0FF:
            return None
        if not g:
            return None
        return g


def _decoded_name(name: str) -> str:
    """書体の名前（PDF の名前はバイト列。日本の書体は Shift_JIS で書かれていることが多い：ＭＳ明朝 など）。"""
    raw = name.encode("latin-1", "replace")
    if raw.isascii():
        return name
    for enc in ("utf-8", "cp932"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return name


def _style(tabs) -> tuple[bool, bool]:
    os2 = tabs.get("OS/2", b"")
    if len(os2) >= 64:
        weight = struct.unpack_from(">H", os2, 4)[0]
        sel = struct.unpack_from(">H", os2, 62)[0]
        return weight >= 600 or bool(sel & 0x20), bool(sel & 1)
    return False, False


@lru_cache(maxsize=24)
def _open(path: str, bold: bool, italic: bool) -> SystemFace | None:
    try:
        with open(path, "rb") as fh:
            data = fh.read()
    except OSError:
        return None
    offs = [0]
    if data[:4] == b"ttcf":
        n = struct.unpack_from(">I", data, 8)[0]
        offs = [struct.unpack_from(">I", data, 12 + 4 * i)[0] for i in range(min(n, 64))]
    best = None
    for off in offs:
        try:
            tabs = sfnt_tables(data, off)
        except (struct.error, IndexError):
            continue
        b, it = _style(tabs)
        score = (b == bold) * 2 + (it == italic)
        if best is None or score > best[0]:
            best = (score, tabs)
    if best is None:
        return None
    tabs = best[1]
    try:
        outl = CFFOutlines(tabs["CFF "]) if "CFF " in tabs else TrueTypeOutlines(tabs)
    except (FontProgramError, KeyError, struct.error, IndexError, ValueError):
        return None
    cm = outl.cmaps if isinstance(outl, TrueTypeOutlines) else _cmaps_of(tabs)
    uni = cm.get((3, 10)) or cm.get((0, 4)) or cm.get((3, 1)) or cm.get((0, 3)) or {}
    if not uni and (3, 0) in cm:
        uni = {0xF000 | (c & 0xFF): g for c, g in cm[(3, 0)].items()}
    return SystemFace(path, outl, uni, os.path.basename(path))


def _cmaps_of(tabs):
    from .fontprog import _cmaps

    return _cmaps(tabs.get("cmap", b""))


def system_face(kind: str) -> SystemFace | None:
    """書体の種類（sans・serif・mono・cjk-sans など、末尾 -bold は太字）の、このパソコンにある書体。"""
    env = os.environ.get("TODOKU_RENDER_FONT_" + kind.upper().replace("-", "_"))
    paths = [env] if env else CANDIDATES.get(kind, [])
    bold = kind.endswith("-bold")
    for p in paths:
        if p and os.path.exists(p):
            f = _open(p, bold, False)
            if f is not None:
                return f
    if bold:
        return system_face(kind[:-5])
    return None


# ====================================================================== PDF の書体 → 字の形
class Glyphs:
    """1 つの PDF の書体の字の形。kind：embedded（埋め込み）・system（代わりの書体）・type3・none（描けない）。"""

    def __init__(self, doc, fd: dict, pf):
        self.doc, self.fd, self.pf = doc, fd, pf
        self.kind = "none"
        self.problem = ""
        self.outl = None
        self.subst: SystemFace | None = None
        self.subst_name = ""
        self.names: dict[int, str] = {}  # 単純な書体：符号 → 字の名前
        self.charprocs: dict = {}
        self.cache: dict = {}
        sub = getattr(fd.get("Subtype"), "name", "")
        self.subtype = sub
        self.desc: dict = {}
        self._uni_of: dict = {}
        desc: dict = {}
        try:
            if sub == "Type3":
                self.kind = "type3"
                self.charprocs = doc.resolve(fd.get("CharProcs")) or {}
                self.names = self._encoding_names(None)
                return
            desc, cfd = self._descriptor()
            self._load_program(desc, cfd)
        except Exception as e:  # noqa: BLE001
            self.outl = None
            self.problem = f"{type(e).__name__}"
        if self.outl is not None:
            self.kind = "embedded"
            if not pf.composite:
                self.names = self._encoding_names(self.outl)
            return
        self._choose_substitute(desc)

    # ------------------------------------------------------------------
    def _descriptor(self):
        doc, fd = self.doc, self.fd
        cfd = None
        if self.pf.composite:
            dfs = doc.resolve(fd.get("DescendantFonts")) or []
            cfd = doc.resolve(dfs[0]) if dfs else None
            cfd = cfd if isinstance(cfd, dict) else {}
            desc = doc.resolve(cfd.get("FontDescriptor")) or {}
        else:
            desc = doc.resolve(fd.get("FontDescriptor")) or {}
        self.desc = desc if isinstance(desc, dict) else {}
        return self.desc, cfd

    def _load_program(self, desc, cfd) -> None:
        doc = self.doc
        if not isinstance(desc, dict):
            return
        ff = doc.resolve(desc.get("FontFile"))
        ff2 = doc.resolve(desc.get("FontFile2"))
        ff3 = doc.resolve(desc.get("FontFile3"))
        if isinstance(ff2, Stream):
            data = doc.decode(ff2)
            self.outl = _sfnt_outlines(data)
        elif isinstance(ff3, Stream):
            data = doc.decode(ff3)
            st = getattr(doc.resolve(ff3.dict.get("Subtype")), "name", "")
            if st == "OpenType" or data[:4] in (b"OTTO", b"\x00\x01\x00\x00", b"true"):
                self.outl = _sfnt_outlines(data)
            else:
                self.outl = CFFOutlines(data)
        elif isinstance(ff, Stream):
            data = doc.decode(ff)
            if data[:4] in (b"\x00\x01\x00\x00", b"true", b"OTTO"):
                self.outl = _sfnt_outlines(data)
            else:
                l1 = doc.resolve(ff.dict.get("Length1"))
                self.outl = Type1Outlines(data, int(l1) if isinstance(l1, int) else None)

    def _choose_substitute(self, desc) -> None:
        pf = self.pf
        base = _decoded_name((pf.base or "").split("+")[-1])
        flags =int(self.doc.resolve((desc or {}).get("Flags")) or 0) if isinstance(desc, dict) else 0
        weight = self.doc.resolve((desc or {}).get("FontWeight")) if isinstance(desc, dict) else None
        low = base.lower()
        bold = bool(flags & 0x40000) or (isinstance(weight, (int, float)) and weight >= 600) or "bold" in low \
            or "black" in low or "heavy" in low
        ros = getattr(pf, "ros", "") or ""
        cjk = pf.composite and ros in ("Adobe-Japan1", "Adobe-GB1", "Adobe-CNS1", "Adobe-Korea1", "Adobe-KR", "Adobe-Identity", "")
        fam = base.split("-")[0].split(",")[0]
        if fam in ("Symbol",):
            kind = "symbol"
        elif fam in ("ZapfDingbats",):
            kind = "dingbats"
        elif pf.composite and cjk:
            # 名前で分かれば名前で（MS-Gothic の FontDescriptor に Serif の印が立っていることがある）。分からなければ Flags
            named_sans = any(k in low for k in ("gothic", "goth", "kaku", "sans", "ゴシック", "hei", "dotum", "gulim"))
            named_serif = any(k in low for k in ("mincho", "ming", "song", "serif", "明朝", "batang"))
            serif = named_serif or (bool(flags & 2) and not named_sans)
            kind = "cjk-serif" if serif else "cjk-sans"
        elif fam.startswith("Courier") or flags & 1:
            kind = "mono"
        elif fam.startswith("Times") or (flags & 2 and not fam.startswith("Helvetica")):
            kind = "serif"
        else:
            kind = "sans"
        face = system_face(kind + ("-bold" if bold and kind not in ("symbol", "dingbats") else ""))
        if face is None and kind.startswith("cjk"):
            face = system_face("cjk-sans")
        if face is None:
            face = system_face("fallback")
        if face is not None:
            self.kind = "system"
            self.subst = face
            self.subst_name = face.name
        else:
            self.kind = "none"
            self.problem = "代わりに描く書体が、このパソコンに見つかりません"

    # ------------------------------------------------------------------ 単純な書体の字の名前
    def _encoding_names(self, outl) -> dict[int, str]:
        from ..fonts import MACROMAN, STANDARD_ENCODING, WINANSI, glyph_to_unicode

        doc, fd = self.doc, self.fd
        enc = doc.resolve(fd.get("Encoding"))
        base_name = enc.name if isinstance(enc, Name) else None
        diffs: dict[int, str] = {}
        if isinstance(enc, dict):
            be = doc.resolve(enc.get("BaseEncoding"))
            base_name = be.name if isinstance(be, Name) else None
            code = 0
            for x in doc.resolve(enc.get("Differences")) or []:
                x = doc.resolve(x)
                if isinstance(x, (int, float)):
                    code = int(x)
                elif isinstance(x, Name):
                    diffs[code] = x.name
                    code += 1
        names: dict[int, str] = {}
        uni_table = None
        if base_name == "StandardEncoding":
            names = {i: n for i, n in enumerate(STANDARD_ENCODING) if n}
        elif base_name in ("WinAnsiEncoding", "MacRomanEncoding"):
            uni_table = WINANSI if base_name == "WinAnsiEncoding" else MACROMAN
        elif outl is not None:
            # 組み込みの符号（Type 1 の Encoding・CFF の encoding）
            if isinstance(outl, Type1Outlines):
                names = dict(outl.encoding_names)
            elif isinstance(outl, CFFOutlines):
                names = {c: outl.gid_name.get(g, "") for c, g in outl.encoding.items() if outl.gid_name.get(g)}
            if not names and self.subtype != "TrueType":
                names = {i: n for i, n in enumerate(STANDARD_ENCODING) if n}
        elif self.subtype != "Type3":
            names = {i: n for i, n in enumerate(STANDARD_ENCODING) if n}
        if uni_table is not None:
            rev = self._rev_names(outl)
            for c in range(256):
                u = uni_table[c]
                if u:
                    nm = rev.get(u)
                    if nm:
                        names[c] = nm
                    else:
                        names.setdefault(c, "")
        for c, n in diffs.items():
            names[c] = n
        self._uni_of = {c: (glyph_to_unicode(n) if n else (uni_table[c] if uni_table and uni_table[c] else "")) for c, n in names.items()}
        return names

    def _rev_names(self, outl) -> dict[str, str]:
        """Unicode → この書体の字の名前（WinAnsi・MacRoman の符号を字の名前にする）。"""
        from ..fonts import glyph_to_unicode
        from ..sfnt import MAC_GLYPHS

        pool = []
        if isinstance(outl, (Type1Outlines, CFFOutlines)):
            pool = list(outl.names)
        pool += [n for n in MAC_GLYPHS if n]
        rev: dict[str, str] = {}
        for n in pool:
            u = glyph_to_unicode(n)
            if u and u not in rev:
                rev[u] = n
        return rev

    # ------------------------------------------------------------------ 符号 → 字形
    def lookup(self, code: bytes, text: str):
        """(形の鍵, 輪郭, 字形の単位 → em の行列) か None。Type 3 は ("type3", 字の手続きのストリーム, None)。"""
        key = code
        hit = self.cache.get(key, False)
        if hit is not False:
            return hit
        try:
            r = self._lookup(code, text)
        except (IndexError, ValueError, KeyError, struct.error, TypeError):
            r = None
        if len(self.cache) < 20000:
            self.cache[key] = r
        return r

    def _lookup(self, code: bytes, text: str):
        pf = self.pf
        if self.kind == "type3":
            nm = self.names.get(code[0]) if code else None
            st = self.doc.resolve(self.charprocs.get(nm)) if nm and isinstance(self.charprocs, dict) else None
            return ("type3", st, None) if isinstance(st, Stream) else None
        if self.kind == "embedded":
            o = self.outl
            if pf.composite:
                cid = _cid(pf, code, text)
                if isinstance(o, CFFOutlines):
                    gid = o.gid_for_cid(cid) if o.is_cid else (pf._gid(cid) if pf.cid2gid is not None else cid)
                    return (gid, o.glyph(gid), o.glyph_matrix(gid))
                if isinstance(o, TrueTypeOutlines):
                    gid = pf._gid(cid)
                    return (gid, o.glyph(gid), o.matrix)
                return None
            c = code[0]
            if isinstance(o, Type1Outlines):
                nm = self.names.get(c) or ""
                if nm not in o.charstrings:
                    nm = self._name_by_text(o, text) or nm
                return (nm, o.glyph(nm), o.matrix) if nm in o.charstrings else None
            if isinstance(o, CFFOutlines):
                nm = self.names.get(c)
                g = o.names.get(nm) if nm else None
                if g is None and not self.names and c in o.encoding:
                    g = o.encoding[c]
                if g is None:
                    nm2 = self._name_by_text(o, text)
                    g = o.names.get(nm2) if nm2 else None
                if g is None and o.is_cid:
                    g = o.gid_for_cid(c)
                if g is None:
                    return None
                return (g, o.glyph(g), o.glyph_matrix(g))
            if isinstance(o, TrueTypeOutlines):
                g = self._tt_simple_gid(o, c, text)
                return (g, o.glyph(g), o.matrix) if g is not None else None
            return None
        if self.kind == "system":
            t = text or ""
            if not t or t == _UNKNOWN:
                return None
            f = self.subst
            g = f.lookup(t)
            face = f
            if g is None:
                for alt in ("cjk-sans", "fallback"):
                    fb = system_face(alt)
                    if fb is not None and fb is not f:
                        g = fb.lookup(t)
                        if g is not None:
                            face = fb
                            break
            if g is None:
                return None
            o = face.outl
            m = o.glyph_matrix(g) if isinstance(o, CFFOutlines) else o.matrix
            return ((face.path, g), o.glyph(g), m)
        return None

    def _name_by_text(self, o, text: str) -> str | None:
        if not text or text == _UNKNOWN:
            return None
        rev = self.__dict__.get("_rev")
        if rev is None:
            rev = self._rev = self._rev_names(o)
        return rev.get(text)

    def _tt_simple_gid(self, o: TrueTypeOutlines, c: int, text: str) -> int | None:
        from ..fonts import glyph_to_unicode

        cm = o.cmaps
        flags = int(self.doc.resolve(self.desc.get("Flags")) or 0) if isinstance(getattr(self, "desc", None), dict) else 0
        symbolic = bool(flags & 4)
        enc = self.doc.resolve(self.fd.get("Encoding"))
        nm = self.names.get(c) if (enc is not None or not symbolic) else None
        m31 = cm.get((3, 1)) or cm.get((0, 3)) or cm.get((0, 4)) or cm.get((3, 10))
        if m31 and nm and not symbolic:
            u = glyph_to_unicode(nm) or self._uni_of.get(c, "")
            if u and ord(u[0]) in m31:
                return m31[ord(u[0])]
        m30 = cm.get((3, 0))
        if m30:
            for v in (c, 0xF000 | c, 0xF100 | c, 0xF200 | c):
                if v in m30:
                    return m30[v]
        m10 = cm.get((1, 0))
        if m10 and c in m10:
            return m10[c]
        if m31:
            u = text if text and text != _UNKNOWN else (glyph_to_unicode(nm) if nm else "")
            if u and ord(u[0]) in m31:
                return m31[ord(u[0])]
            if c in m31:
                return m31[c]
        if nm:
            post = self.__dict__.get("_post")
            if post is None:
                post = self._post = _post_names(o)
            if nm in post:
                return post[nm]
        return c if c < o.num else None


def _post_names(o: TrueTypeOutlines) -> dict[str, int]:
    from ..sfnt import MAC_GLYPHS

    t = o.tabs.get("post", b"")
    out: dict[str, int] = {}
    try:
        if len(t) >= 34 and struct.unpack_from(">I", t, 0)[0] == 0x00020000:
            n = struct.unpack_from(">H", t, 32)[0]
            idx = struct.unpack_from(f">{n}H", t, 34)
            p = 34 + 2 * n
            names = []
            while p < len(t):
                ln = t[p]
                names.append(t[p + 1:p + 1 + ln].decode("latin-1"))
                p += 1 + ln
            for g, k in enumerate(idx):
                nm = MAC_GLYPHS[k] if k < 258 and k < len(MAC_GLYPHS) else names[k - 258] if 0 <= k - 258 < len(names) else ""
                if nm and nm not in out:
                    out[nm] = g
    except (struct.error, IndexError):
        pass
    return out


def _sfnt_outlines(data: bytes):
    off = 0
    if data[:4] == b"ttcf":
        off = struct.unpack_from(">I", data, 12)[0]
    tabs = sfnt_tables(data, off)
    if "CFF " in tabs:
        o = CFFOutlines(tabs["CFF "])
        o.cmaps = _cmaps_of(tabs)
        return o
    return TrueTypeOutlines(tabs)


def _cid(pf, code: bytes, text: str) -> int:
    cid = None
    if pf.cmap is not None:
        cid = pf.cmap.to_cid(code)
    elif pf.pre is not None and pf.pre[0] == "identity":
        cid = int.from_bytes(code, "big")
    elif pf.pre is not None:
        # 定義済みの CMap：符号から字に直して CID を引く（ToUnicode の字からではない）
        t = decode_predefined(code, pf.pre[0]) or (text if text != _UNKNOWN else "")
        cid = pf._cid_from_text(t, len(code)) if t else None
    if cid is None:
        cid = int.from_bytes(code, "big")
    return cid


def glyphs_for(doc, fd: dict, pf) -> Glyphs:
    cache = doc.__dict__.setdefault("_render_glyphs", {})
    k = id(pf)
    g = cache.get(k)
    if g is None:
        g = cache[k] = Glyphs(doc, fd, pf)
    return g
