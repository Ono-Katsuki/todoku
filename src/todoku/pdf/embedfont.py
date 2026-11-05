"""PDF に書く字を、書いた字の字形だけを埋め込んで描く（ISO 32000-1 9.7.4・9.9、CIDFontType2・Identity-H）。

- 同梱の書体（TdkB）：data/NotoSansJP-Regular.ttf.gz（SIL OFL 1.1。THIRD_PARTY_NOTICES.md）。字形のある字は、語や言語によらず
  これで描く（glyphs.choose）。最初に要るときに初めて読む。字形を 2 つ以上の字が分け合う（「〜」U+301C と「～」U+FF5E など）ので、
  字形の番号でなく字ごとに CID を決め（cmap の字を小さい順に並べた番号）、CIDToGIDMap のストリームで字形につなぐ。
  これで ToUnicode が書いた字そのものになる（読み上げ・検索・読み戻しで別の字にならない）
- 同梱の書体にない字のうち、埋め込まない書体（Helvetica・平成角ゴシックなど）の字形にもない字は、このパソコンの TrueType の書体（TdkE）で。
  書体のファイル：環境変数 TODOKU_PDF_FONT（空にすると、同梱の書体も含めて埋め込まない）、なければ OS の決まった場所の Arial・
  DejaVu Sans・Liberation Sans の順。見つからなければ、その字は理由を添えて断る（glyphs.choose）
- 書体の埋め込みの許し（OS/2 の fsType、OpenType の仕様）：Restricted License・ビットマップだけ・部分の埋め込み不可のものは使わない
- 部分の埋め込み：字形の番号（GID）は元のまま（CIDToGIDMap /Identity）、使う字形（と、組み合わせの字形の部品）の中身だけを残し、
  ほかは空にする。この書体（TdkE）では字の並べ替え・合字をしないので、結合の字（Mn・Mc・Me）・右から左の字は使わない
  （アラビア文字・インド系の文字などは、shape.py が字を組み立てて、文字ごとの書体（TdkX…）で描く。TrueType Collection
  （.ttc）の書体も、その中の 1 つの書体の表から埋め込める）
"""

from __future__ import annotations

import gzip
import hashlib
import os
import struct
import unicodedata
import zlib
from functools import lru_cache
from pathlib import Path

def windows_fonts_dir() -> str:
    """Windows の書体のフォルダ（%WINDIR%\\Fonts。Windows が C: 以外にあっても）。末尾に区切りの字をつける。"""
    root = os.environ.get("WINDIR") or os.environ.get("SystemRoot") or "C:\\Windows"
    return root.rstrip("\\/") + "\\Fonts\\"


CANDIDATES = [
    "/System/Library/Fonts/Supplemental/Arial.ttf", "/Library/Fonts/Arial.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/TTF/DejaVuSans.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    windows_fonts_dir() + "arial.ttf",
]


def face_tables(data: bytes, offset: int = 0) -> dict[str, bytes]:
    """書体のファイルの表（TrueType Collection（.ttc）なら、offset の位置の書体の表）。"""
    out: dict[str, bytes] = {}
    n = struct.unpack_from(">H", data, offset + 4)[0]
    for i in range(n):
        tag, _chk, off, ln = struct.unpack_from(">4sIII", data, offset + 12 + 16 * i)
        out[tag.decode("latin-1")] = data[off:off + ln]
    return out


def collection_faces(data: bytes) -> list[int]:
    """TrueType Collection の書体ごとの表の位置（ふつうの書体なら [0]）。"""
    if data[:4] == b"ttcf":
        n = struct.unpack_from(">I", data, 8)[0]
        return [struct.unpack_from(">I", data, 12 + 4 * i)[0] for i in range(n)]
    return [0]


def _cmap(t: bytes) -> dict[int, int]:
    from .sfnt import cmap_tables

    subs = cmap_tables(t)
    return subs.get((3, 10)) or subs.get((0, 4)) or subs.get((3, 1)) or subs.get((0, 3)) or {}


class TrueType:
    def __init__(self, path: str, data: bytes, offset: int = 0):
        self.path = path
        self.data = data
        self.offset = offset
        self.t = face_tables(data, offset)
        for need in ("head", "hhea", "maxp", "hmtx", "loca", "glyf"):
            if need not in self.t:
                raise ValueError(f"TrueType の {need} がありません")
        head = self.t["head"]
        self.upem = struct.unpack_from(">H", head, 18)[0]
        self.bbox = struct.unpack_from(">hhhh", head, 36)
        self.loc_long = struct.unpack_from(">h", head, 50)[0] == 1
        self.num = struct.unpack_from(">H", self.t["maxp"], 4)[0]
        hhea = self.t["hhea"]
        self.ascent, self.descent = struct.unpack_from(">hh", hhea, 4)
        n_hm = struct.unpack_from(">H", hhea, 34)[0]
        hm = self.t["hmtx"]
        adv = [struct.unpack_from(">H", hm, 4 * i)[0] for i in range(n_hm)]
        self.adv = adv + [adv[-1]] * (self.num - n_hm)
        loca = self.t["loca"]
        if self.loc_long:
            self.loca = list(struct.unpack_from(f">{self.num + 1}I", loca, 0))
        else:
            self.loca = [x * 2 for x in struct.unpack_from(f">{self.num + 1}H", loca, 0)]
        self.cmap = _cmap(self.t["cmap"]) if "cmap" in self.t else {}
        os2 = self.t.get("OS/2", b"")
        self.fstype = struct.unpack_from(">H", os2, 8)[0] if len(os2) >= 10 else 0
        self.cap = struct.unpack_from(">h", os2, 88)[0] if len(os2) >= 90 else self.ascent
        self.name = _ps_name(self.t.get("name", b"")) or os.path.splitext(os.path.basename(path))[0].replace(" ", "")

    @property
    def embeddable(self) -> bool:
        f = self.fstype
        if f & 0x0200 or f & 0x0100:  # ビットマップだけ・部分の埋め込み不可
            return False
        return (f & 0x000F) != 0x0002  # Restricted License（ほかの許しのビットがない）

    def glyph(self, gid: int) -> bytes:
        return self.t["glyf"][self.loca[gid]:self.loca[gid + 1]]

    def gid(self, ch: str) -> int | None:
        g = self.cmap.get(ord(ch))
        if not g or g >= self.num:
            return None
        if not self.glyph(g) and not ch.isspace():
            return None  # 字形の中身が空（形のない字）
        return g

    def width(self, gid: int) -> float:
        return self.adv[gid] * 1000.0 / self.upem

    def components(self, gid: int) -> list[int]:
        g = self.glyph(gid)
        if len(g) < 10 or struct.unpack_from(">h", g, 0)[0] >= 0:
            return []
        out, p = [], 10
        while p + 4 <= len(g):
            flags, comp = struct.unpack_from(">HH", g, p)
            out.append(comp)
            p += 4 + (4 if flags & 0x0001 else 2)
            if flags & 0x0008:
                p += 2
            elif flags & 0x0040:
                p += 4
            elif flags & 0x0080:
                p += 8
            if not flags & 0x0020:
                break
        return out

    def subset(self, gids: set[int]) -> bytes:
        keep = {0}
        todo = list(gids)
        while todo:
            g = todo.pop()
            if g in keep or g >= self.num:
                continue
            keep.add(g)
            todo += self.components(g)
        glyf, loca = bytearray(), []
        for g in range(self.num):
            loca.append(len(glyf))
            if g in keep:
                d = self.glyph(g)
                glyf += d + b"\0" * (-len(d) % 4)
        loca.append(len(glyf))
        head = bytearray(self.t["head"])
        struct.pack_into(">I", head, 8, 0)  # checkSumAdjustment（あとで決める）
        struct.pack_into(">h", head, 50, 1)  # indexToLocFormat：long
        post = struct.pack(">IIhhIIIII", 0x00030000, 0, 0, 0, 0, 0, 0, 0, 0)  # 字形の名前を持たない形（format 3）
        # 使わない字形の幅は 0 に（表の長さは字形の数のまま。圧縮でほとんど消える。PDF の幅は /W にある）
        n_hm = struct.unpack_from(">H", self.t["hhea"], 34)[0]
        hmtx = bytearray(len(self.t["hmtx"]))
        for g in keep:
            if g < n_hm:
                hmtx[4 * g:4 * g + 4] = self.t["hmtx"][4 * g:4 * g + 4]
            else:
                o = 4 * n_hm + 2 * (g - n_hm)
                hmtx[o:o + 2] = self.t["hmtx"][o:o + 2]
        tabs = {"head": bytes(head), "hhea": self.t["hhea"], "maxp": self.t["maxp"], "hmtx": bytes(hmtx),
                "loca": struct.pack(f">{len(loca)}I", *loca), "glyf": bytes(glyf), "post": post}
        for k in ("cvt ", "fpgm", "prep", "OS/2"):
            if k in self.t:
                tabs[k] = self.t[k]
        return _sfnt(tabs)


def _checksum(b: bytes) -> int:
    b = b + b"\0" * (-len(b) % 4)
    return sum(struct.unpack(f">{len(b) // 4}I", b)) & 0xFFFFFFFF


def _sfnt(tabs: dict[str, bytes]) -> bytes:
    tags = sorted(tabs)
    n = len(tags)
    es = max(k for k in range(16) if (1 << k) <= n)
    head = struct.pack(">IHHHH", 0x00010000, n, (1 << es) * 16, es, n * 16 - (1 << es) * 16)
    off = 12 + 16 * n
    recs, body = bytearray(), bytearray()
    for t in tags:
        d = tabs[t]
        recs += struct.pack(">4sIII", t.encode("latin-1"), _checksum(d), off + len(body), len(d))
        body += d + b"\0" * (-len(d) % 4)
    out = bytearray(head + recs + body)
    hpos = off + sum(len(tabs[t]) + (-len(tabs[t]) % 4) for t in tags[:tags.index("head")])
    struct.pack_into(">I", out, hpos + 8, (0xB1B0AFBA - _checksum(bytes(out))) & 0xFFFFFFFF)
    return bytes(out)


def _ps_name(t: bytes) -> str:
    try:
        count, so = struct.unpack_from(">HH", t, 2)
        for i in range(count):
            pid, eid, lid, nid, ln, off = struct.unpack_from(">HHHHHH", t, 6 + 12 * i)
            if nid != 6:
                continue
            raw = t[so + off:so + off + ln]
            s = raw.decode("utf-16-be") if pid in (0, 3) else raw.decode("latin-1")
            s = "".join(ch for ch in s if ch.isalnum() or ch in "-_")
            if s:
                return s
    except (struct.error, UnicodeDecodeError):
        pass
    return ""


@lru_cache(maxsize=4)
def _load(path: str) -> TrueType | None:
    try:
        with open(path, "rb") as fh:
            data = fh.read()
        f = TrueType(path, data)
    except (OSError, ValueError, struct.error):
        return None
    return f if f.embeddable else None


def system_font() -> TrueType | None:
    """埋め込みに使う書体（なければ None）。"""
    return _pick(os.environ.get("TODOKU_PDF_FONT"))


@lru_cache(maxsize=4)
def _pick(env: str | None) -> TrueType | None:
    if env is not None:
        return _load(env) if env else None
    for p in CANDIDATES:
        if os.path.exists(p):
            f = _load(p)
            if f is not None:
                return f
    return None


BUNDLED_FILE = Path(__file__).parent / "data" / "NotoSansJP-Regular.ttf.gz"
BUNDLED_KEY = "TdkB"  # 同梱の書体の資源の名前


@lru_cache(maxsize=1)
def _read_bundled() -> TrueType | None:
    try:
        return TrueType(BUNDLED_FILE.name[:-3], gzip.decompress(BUNDLED_FILE.read_bytes()))
    except (OSError, EOFError, ValueError, struct.error, zlib.error):
        return None


@lru_cache(maxsize=1)
def _bundled_cids() -> dict[int, int]:
    """同梱の書体の字 → CID（cmap の字を小さい順に並べた 1 からの番号。字ごとに違う番号）。"""
    font = _read_bundled()
    return {u: i + 1 for i, u in enumerate(sorted(font.cmap))} if font is not None else {}


def bundled_font() -> TrueType | None:
    """同梱の書体（TODOKU_PDF_FONT を空にしたときは、埋め込まない）。"""
    return None if os.environ.get("TODOKU_PDF_FONT") == "" else _read_bundled()


@lru_cache(maxsize=65536)
def _bundled_glyph(ch: str) -> tuple[bytes, float] | None:
    font = _read_bundled()
    if font is None or not usable(ch):
        return None
    g = font.gid(ch)
    if g is None:
        return None
    return _bundled_cids()[ord(ch)].to_bytes(2, "big"), font.width(g)


def bundled_glyph(ch: str) -> tuple[bytes, float] | None:
    """同梱の書体で描ける字なら (描く文字コード（CID の 2 バイト）, 幅（1000 分の 1）)。字形がなければ None。"""
    if os.environ.get("TODOKU_PDF_FONT") == "":
        return None
    return _bundled_glyph(ch)


def usable(ch: str) -> bool:
    """埋め込んだ書体で、並べ替え・合字なしに1字ずつ描ける字か（結合の字・右から左の字は除く）。"""
    if unicodedata.category(ch) in ("Mn", "Mc", "Me", "Cf", "Cc", "Co", "Cs", "Cn"):
        return False
    return unicodedata.bidirectional(ch) not in ("R", "AL", "RLE", "RLO", "RLI", "AN", "NSM")


def bundled_objects(cids: list[int]) -> tuple[dict, dict, dict, bytes, bytes, bytes, bytes]:
    """同梱の書体で、使った CID だけの書体。(Type0, CIDFontType2, FontDescriptor, 書体の中身, ToUnicode, CIDToGIDMap, CIDSet)。"""
    font = _read_bundled()
    by_cid = {c: u for u, c in _bundled_cids().items()}
    chars = {c: chr(by_cid[c]) for c in cids if c in by_cid}
    gid_of = {c: font.cmap[ord(ch)] for c, ch in chars.items()}
    t0, cid, desc, sub, tu = font_objects(font, chars, gid_of)
    last = max(chars, default=0)
    c2g = bytearray(2 * (last + 1))
    bits = bytearray(last // 8 + 1)
    bits[0] = 0x80  # CID 0（.notdef）
    for c, g in gid_of.items():
        c2g[2 * c:2 * c + 2] = g.to_bytes(2, "big")
        bits[c // 8] |= 0x80 >> (c % 8)
    return t0, cid, desc, sub, tu, bytes(c2g), bytes(bits)


def font_objects(font: TrueType, gids: dict[int, str], gid_of: dict[int, int] | None = None) -> tuple[dict, dict, dict, bytes, bytes]:
    """gids：使う CID → 字。gid_of：CID → 字形の番号（省くと CID が字形の番号。CIDToGIDMap /Identity）。
    (Type0 の辞書, CIDFontType2 の辞書, FontDescriptor, 埋め込む書体の中身, ToUnicode の CMap)。参照はつながずに返す
    （gid_of があるときの CIDToGIDMap は呼ぶ側が差し替える）。"""
    from .objects import Name, PdfString

    glyph = gid_of or {g: g for g in gids}
    sub = font.subset(set(glyph.values()))
    tag = "".join(chr(65 + b % 26) for b in hashlib.sha256(repr(sorted(gids)).encode() + font.name.encode()).digest()[:6])
    base = Name(f"{tag}+{font.name}")
    sc = 1000.0 / font.upem
    desc = {"Type": Name("FontDescriptor"), "FontName": base, "Flags": 32,
            "FontBBox": [round(v * sc) for v in font.bbox], "ItalicAngle": 0, "Ascent": round(font.ascent * sc),
            "Descent": round(font.descent * sc), "CapHeight": round(font.cap * sc), "StemV": 80}
    w: list = []
    for g in sorted(gids):
        w += [g, [round(font.width(glyph[g]), 1)]]
    cid = {"Type": Name("Font"), "Subtype": Name("CIDFontType2"), "BaseFont": base,
           "CIDSystemInfo": {"Registry": PdfString(b"Adobe"), "Ordering": PdfString(b"Identity"), "Supplement": 0},
           "DW": 1000, "W": w, "CIDToGIDMap": Name("Identity")}
    t0 = {"Type": Name("Font"), "Subtype": Name("Type0"), "BaseFont": base, "Encoding": Name("Identity-H")}
    lines = [f"<{g:04X}> <{ch.encode('utf-16-be').hex().upper()}>" for g, ch in sorted(gids.items()) if ch]
    body = ["/CIDInit /ProcSet findresource begin", "12 dict begin", "begincmap",
            "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def", "/CMapName /TdkE-UCS def",
            "/CMapType 2 def", "1 begincodespacerange", "<0000> <FFFF>", "endcodespacerange"]
    for i in range(0, len(lines), 100):
        chunk = lines[i:i + 100]
        body += [f"{len(chunk)} beginbfchar", *chunk, "endbfchar"]
    body += ["endcmap", "CMapName currentdict /CMap defineresource pop", "end", "end"]
    return t0, cid, desc, sub, ("\n".join(body) + "\n").encode("latin-1")


def compress(data: bytes) -> bytes:
    return zlib.compress(data, 9)
