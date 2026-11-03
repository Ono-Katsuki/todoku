"""テスト用の PDF を、手で書いたオブジェクトから作る（標準ライブラリだけ）。

- 書体のデータは埋め込まない（ライセンスの心配がない）。和文は定義済みの CMap（UniJIS-UTF16-H／-V）で、
  そのほかの文字（ヘブライ・アラビア・デーヴァナーガリー）は Identity-H と ToUnicode で、字を読めるようにする。
- 暗号化は、PDF 1.7（ISO 32000-1）7.6.3 の標準の暗号（RC4・128 ビット・R3）を、この中で書いて行う。
"""

from __future__ import annotations

import hashlib
import re
import struct
import zlib

PAD = bytes.fromhex("28BF4E5E4E758A4164004E56FFFA01082E2E00B6D0683E802F0CA9FE6453697A")


def utf16hex(s: str) -> str:
    return s.encode("utf-16-be").hex().upper()


def stream(d: str, data: bytes, compress: bool = True) -> bytes:
    if compress:
        data = zlib.compress(data, 9)
        d = (d + " /Filter /FlateDecode").strip()
    return f"<< {d} /Length {len(data)} >>\nstream\n".encode("latin-1") + data + b"\nendstream"


def rc4(key: bytes, data: bytes) -> bytes:
    s = list(range(256))
    j = 0
    for i in range(256):
        j = (j + s[i] + key[i % len(key)]) & 0xFF
        s[i], s[j] = s[j], s[i]
    out = bytearray()
    i = j = 0
    for b in data:
        i = (i + 1) & 0xFF
        j = (j + s[i]) & 0xFF
        s[i], s[j] = s[j], s[i]
        out.append(b ^ s[(s[i] + s[j]) & 0xFF])
    return bytes(out)


def _pad(pw: bytes) -> bytes:
    return (pw + PAD)[:32]


class Encryption:
    """標準の暗号（V2・R3・128 ビット）。7.6.3.3 Algorithm 2・3・5。"""

    def __init__(self, user: str, owner: str, file_id: bytes, perms: int = -4):
        self.id0, self.p = file_id, perms
        k = hashlib.md5(_pad(owner.encode("latin-1"))).digest()
        for _ in range(50):
            k = hashlib.md5(k).digest()
        o = rc4(k, _pad(user.encode("latin-1")))
        for i in range(1, 20):
            o = rc4(bytes(b ^ i for b in k), o)
        self.o = o
        h = hashlib.md5(_pad(user.encode("latin-1")) + o + struct.pack("<i", perms) + file_id).digest()
        for _ in range(50):
            h = hashlib.md5(h[:16]).digest()
        self.key = h[:16]
        u = rc4(self.key, hashlib.md5(PAD + file_id).digest())
        for i in range(1, 20):
            u = rc4(bytes(b ^ i for b in self.key), u)
        self.u = u + bytes(16)

    def dict(self) -> bytes:
        return (f"<< /Filter /Standard /V 2 /R 3 /Length 128 /P {self.p} /O <{self.o.hex()}> /U <{self.u.hex()}> >>").encode()

    def obj_key(self, num: int) -> bytes:
        return hashlib.md5(self.key + struct.pack("<I", num)[:3] + b"\x00\x00").digest()[:16]

    def encrypt_object(self, num: int, body: bytes) -> bytes:
        """オブジェクトの中の16進の文字列と、ストリームの中身を暗号にする（この生成器は、字の文字列を必ず16進で書く）。"""
        key = self.obj_key(num)
        m = re.search(rb">>\s*stream\n", body)
        head, data = (body[:m.end()], body[m.end():]) if m else (body, None)
        out = bytearray()
        i = 0
        while i < len(head):
            if head.startswith(b"<<", i) or head.startswith(b">>", i):
                out += head[i:i + 2]
                i += 2
            elif head[i:i + 1] == b"<":
                j = head.index(b">", i)
                raw = bytes.fromhex(head[i + 1:j].decode("latin-1"))
                out += b"<" + rc4(key, raw).hex().encode() + b">"
                i = j + 1
            elif head[i:i + 1] == b"(":
                raise ValueError("暗号にする PDF では、文字列を16進（<...>）で書く")
            else:
                out += head[i:i + 1]
                i += 1
        if data is not None:
            end = data.rindex(b"\nendstream")
            enc = rc4(key, data[:end])
            out = bytearray(re.sub(rb"/Length \d+", f"/Length {len(enc)}".encode(), bytes(out), count=1))
            out += enc + data[end:]
        return bytes(out)


class Pdf:
    def __init__(self):
        self.objs: list[bytes] = []

    def add(self, b: bytes | str) -> int:
        self.objs.append(b if isinstance(b, bytes) else b.encode("latin-1"))
        return len(self.objs)

    def reserve(self) -> int:
        return self.add(b"null")

    def set(self, num: int, b: bytes | str) -> None:
        self.objs[num - 1] = b if isinstance(b, bytes) else b.encode("latin-1")

    def build(self, root: int, enc: Encryption | None = None, info: int | None = None) -> bytes:
        file_id = enc.id0 if enc else hashlib.md5(b"".join(self.objs)).digest()
        objs = list(self.objs)
        enc_num = None
        if enc:
            objs = [enc.encrypt_object(i, o) for i, o in enumerate(objs, 1)]
            objs.append(enc.dict())
            enc_num = len(objs)
        out = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
        offs = []
        for i, o in enumerate(objs, 1):
            offs.append(len(out))
            out += f"{i} 0 obj\n".encode() + o + b"\nendobj\n"
        x = len(out)
        out += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode()
        for o in offs:
            out += f"{o:010d} 00000 n \n".encode()
        tr = f"/Size {len(objs) + 1} /Root {root} 0 R /ID [<{file_id.hex()}> <{file_id.hex()}>]"
        if enc_num:
            tr += f" /Encrypt {enc_num} 0 R"
        if info:
            tr += f" /Info {info} 0 R"
        out += f"trailer\n<< {tr} >>\nstartxref\n{x}\n%%EOF\n".encode()
        return bytes(out)


# ---------------------------------------------------------------------- 書体（埋め込みなし）
def helvetica(p: Pdf) -> int:
    return p.add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")


def japanese(p: Pdf, vertical: bool = False) -> int:
    """定義済みの CMap（UniJIS-UTF16-H／V）の和文の書体。字のコードが UTF-16BE そのもの。"""
    desc = p.add("<< /Type /FontDescriptor /FontName /KozMinPr6N-Regular /Flags 6 /FontBBox [-437 -340 1147 1317] "
                 "/ItalicAngle 0 /Ascent 1317 /Descent -349 /CapHeight 742 /StemV 80 >>")
    cid = p.add(f"<< /Type /Font /Subtype /CIDFontType0 /BaseFont /KozMinPr6N-Regular "
                f"/CIDSystemInfo << /Registry <{b'Adobe'.hex()}> /Ordering <{b'Japan1'.hex()}> /Supplement 6 >> "
                f"/FontDescriptor {desc} 0 R /DW 1000 >>")
    cmap = "UniJIS-UTF16-V" if vertical else "UniJIS-UTF16-H"
    return p.add(f"<< /Type /Font /Subtype /Type0 /BaseFont /KozMinPr6N-Regular-{cmap} /Encoding /{cmap} "
                 f"/DescendantFonts [{cid} 0 R] >>")


def tounicode_font(p: Pdf, name: str, glyphs: dict[int, str], width: int = 600) -> int:
    """Identity-H の書体（データは埋め込まない）。glyphs：字形の番号 → その字形が表す字（合字なら複数の字）。"""
    lines = ["/CIDInit /ProcSet findresource begin", "12 dict begin", "begincmap",
             "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def",
             "/CMapName /Adobe-Identity-UCS def", "/CMapType 2 def",
             "1 begincodespacerange", "<0000> <FFFF>", "endcodespacerange",
             f"{len(glyphs)} beginbfchar"]
    for g, s in sorted(glyphs.items()):
        lines.append(f"<{g:04X}> <{utf16hex(s)}>")
    lines += ["endbfchar", "endcmap", "CMapName currentdict /CMap defineresource pop", "end", "end"]
    tu = p.add(stream("", "\n".join(lines).encode("latin-1")))
    desc = p.add(f"<< /Type /FontDescriptor /FontName /{name} /Flags 4 /FontBBox [0 -250 1000 900] /ItalicAngle 0 "
                 f"/Ascent 900 /Descent -250 /CapHeight 700 /StemV 80 >>")
    cid = p.add(f"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /{name} "
                f"/CIDSystemInfo << /Registry <{b'Adobe'.hex()}> /Ordering <{b'Identity'.hex()}> /Supplement 0 >> "
                f"/FontDescriptor {desc} 0 R /DW {width} /CIDToGIDMap /Identity >>")
    return p.add(f"<< /Type /Font /Subtype /Type0 /BaseFont /{name} /Encoding /Identity-H "
                 f"/DescendantFonts [{cid} 0 R] /ToUnicode {tu} 0 R >>")


def glyph_hex(ids: list[int]) -> str:
    return "".join(f"{g:04X}" for g in ids)


# ---------------------------------------------------------------------- 入力欄
def text_field(p: Pdf, page: int, name: str, rect: tuple[float, float, float, float], tooltip: str = "",
               maxlen: int | None = None, comb: bool = False, multiline: bool = False) -> int:
    x0, y0, x1, y1 = rect
    w, h = x1 - x0, y1 - y0
    ap = p.add(stream(f"/Type /XObject /Subtype /Form /BBox [0 0 {w:g} {h:g}] /Resources << >>", b"/Tx BMC EMC", compress=False))
    ff = (1 << 12 if multiline else 0) | (1 << 24 if comb else 0)
    extra = f" /MaxLen {maxlen}" if maxlen else ""
    extra += f" /Ff {ff}" if ff else ""
    extra += f" /TU <FEFF{utf16hex(tooltip)}>" if tooltip else ""
    return p.add(f"<< /Type /Annot /Subtype /Widget /FT /Tx /T <{name.encode().hex()}> /Rect [{x0:g} {y0:g} {x1:g} {y1:g}] "
                 f"/P {page} 0 R /DA <{b'/Helv 10 Tf 0 g'.hex()}> /F 4 /MK << >>{extra} /AP << /N {ap} 0 R >> >>")


def checkbox(p: Pdf, page: int, name: str, rect: tuple[float, float, float, float], on: str = "Yes", tooltip: str = "") -> int:
    x0, y0, x1, y1 = rect
    w, h = x1 - x0, y1 - y0
    zadb = p.add("<< /Type /Font /Subtype /Type1 /BaseFont /ZapfDingbats >>")
    yes = p.add(stream(f"/Type /XObject /Subtype /Form /BBox [0 0 {w:g} {h:g}] /Resources << /Font << /ZaDb {zadb} 0 R >> >>",
                       f"q BT 0 g /ZaDb {h * 0.8:g} Tf 1 {h * 0.2:g} Td (4) Tj ET Q".encode(), compress=False))
    off = p.add(stream(f"/Type /XObject /Subtype /Form /BBox [0 0 {w:g} {h:g}]", b"", compress=False))
    tu = f" /TU <FEFF{utf16hex(tooltip)}>" if tooltip else ""
    return p.add(f"<< /Type /Annot /Subtype /Widget /FT /Btn /T <{name.encode().hex()}> /Rect [{x0:g} {y0:g} {x1:g} {y1:g}] "
                 f"/P {page} 0 R /V /Off /AS /Off /F 4 /MK << /CA <34> >> /DA <{b'/ZaDb 0 Tf 0 g'.hex()}>{tu} "
                 f"/AP << /N << /{on} {yes} 0 R /Off {off} 0 R >> >> >>")


def combo(p: Pdf, page: int, name: str, rect: tuple[float, float, float, float], options: list[str], tooltip: str = "") -> int:
    x0, y0, x1, y1 = rect
    w, h = x1 - x0, y1 - y0
    ap = p.add(stream(f"/Type /XObject /Subtype /Form /BBox [0 0 {w:g} {h:g}]", b"/Tx BMC EMC", compress=False))
    opt = " ".join(f"<{o.encode('latin-1').hex()}>" for o in options)
    tu = f" /TU <FEFF{utf16hex(tooltip)}>" if tooltip else ""
    return p.add(f"<< /Type /Annot /Subtype /Widget /FT /Ch /Ff {1 << 17} /T <{name.encode().hex()}> /Opt [{opt}] "
                 f"/Rect [{x0:g} {y0:g} {x1:g} {y1:g}] /P {page} 0 R /DA <{b'/Helv 10 Tf 0 g'.hex()}> /F 4{tu} "
                 f"/AP << /N {ap} 0 R >> >>")


def page(p: Pdf, num: int, pages: int, content: bytes, fonts: dict[str, int], annots: list[int] = (),
         size=(595, 842), rotate: int = 0, xobjects: dict[str, int] | None = None) -> None:
    cs = p.add(stream("", content))
    font = " ".join(f"/{k} {v} 0 R" for k, v in fonts.items())
    xo = (" /XObject << " + " ".join(f"/{k} {v} 0 R" for k, v in xobjects.items()) + " >>") if xobjects else ""
    an = f" /Annots [{' '.join(f'{a} 0 R' for a in annots)}]" if annots else ""
    rot = f" /Rotate {rotate}" if rotate else ""
    p.set(num, f"<< /Type /Page /Parent {pages} 0 R /MediaBox [0 0 {size[0]} {size[1]}]{rot} /Contents {cs} 0 R "
               f"/Resources << /Font << {font} >>{xo} >>{an} >>")


def catalog(p: Pdf, cat: int, pages: int, kids: list[int], fields: list[int] = (), helv: int | None = None,
            lang: str = "") -> None:
    p.set(pages, f"<< /Type /Pages /Kids [{' '.join(f'{k} 0 R' for k in kids)}] /Count {len(kids)} >>")
    acro = ""
    if fields:
        dr = f" /DR << /Font << /Helv {helv} 0 R >> >>" if helv else ""
        acro = (f" /AcroForm << /Fields [{' '.join(f'{f} 0 R' for f in fields)}]{dr} "
                f"/DA <{b'/Helv 0 Tf 0 g'.hex()}> >>")
    lg = f" /Lang <{lang.encode().hex()}>" if lang else ""
    p.set(cat, f"<< /Type /Catalog /Pages {pages} 0 R{acro}{lg} >>")


def box(x0: float, y0: float, x1: float, y1: float) -> str:
    return f"{x0:g} {y0:g} {x1 - x0:g} {y1 - y0:g} re S\n"
