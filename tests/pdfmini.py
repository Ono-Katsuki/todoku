"""試験用の小さな PDF を手で組み立てる（描画の命令と、フォント・画像などのオブジェクトを、そのまま書く）。

objects: 2番目以降のオブジェクトの中身（文字列）。"{n}" の形で、ほかのオブジェクトの番号を使える（1 はページ）。
"""

from __future__ import annotations


def tounicode(pairs: dict[int, str], width: int = 2) -> str:
    lines = ["/CIDInit /ProcSet findresource begin", "12 dict begin", "begincmap",
             "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def",
             "/CMapName /Adobe-Identity-UCS def", "/CMapType 2 def", "1 begincodespacerange",
             "<00> <FF>" if width == 1 else "<0000> <FFFF>", "endcodespacerange", f"{len(pairs)} beginbfchar"]
    for c, s in pairs.items():
        lines.append(f"<{c:0{2 * width}X}> <{s.encode('utf-16-be').hex().upper()}>")
    lines += ["endbfchar", "endcmap", "CMapName currentdict /CMap defineresource pop", "end", "end"]
    return "\n".join(lines)


def stream(data: str | bytes, extra: str = "") -> str | bytes:
    b = data.encode("latin-1") if isinstance(data, str) else data
    return b"<< /Length %d %s >>\nstream\n" % (len(b), extra.encode()) + b + b"\nendstream"


def build(content: str, resources: str, objects: list, annots: str = "", acroform: str = "") -> bytes:
    """content：ページの描画の命令。resources：ページの資源の辞書（"<< /Font << /F1 2 0 R >> >>" など）。"""
    page = (f"<< /Type /Page /Parent {len(objects) + 3} 0 R /MediaBox [0 0 595 842] /Contents {len(objects) + 2} 0 R "
            f"/Resources {resources} {annots} >>")
    body = [page.encode()] + [o.encode("latin-1") if isinstance(o, str) else o for o in objects]
    body.append(stream(content))
    body.append(f"<< /Type /Pages /Kids [1 0 R] /Count 1 >>".encode())
    body.append(f"<< /Type /Catalog /Pages {len(body)} 0 R {acroform} >>".encode())
    out = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
    offs = []
    for i, b in enumerate(body, 1):
        offs.append(len(out))
        out += f"{i} 0 obj\n".encode() + b + b"\nendobj\n"
    x = len(out)
    out += f"xref\n0 {len(body) + 1}\n0000000000 65535 f \n".encode()
    for o in offs:
        out += f"{o:010d} 00000 n \n".encode()
    out += f"trailer\n<< /Size {len(body) + 1} /Root {len(body)} 0 R >>\nstartxref\n{x}\n%%EOF\n".encode()
    return bytes(out)


def type0(base: str, tu_obj: int, widths: str = "", vertical: bool = False, ordering: str = "Identity") -> str:
    enc = "Identity-V" if vertical else "Identity-H"
    return (f"<< /Type /Font /Subtype /Type0 /BaseFont /{base} /Encoding /{enc} /ToUnicode {tu_obj} 0 R "
            f"/DescendantFonts [<< /Type /Font /Subtype /CIDFontType2 /BaseFont /{base} /CIDSystemInfo << /Registry (Adobe) "
            f"/Ordering ({ordering}) /Supplement 0 >> /DW 1000 {widths} >>] >>")


def glyph_row(codes: list[int], xs: list[float], y: float, size: float = 12, font: str = "F1") -> str:
    """字を1つずつ、決めた位置に置く（2バイトの符号）。"""
    return "\n".join(f"BT /{font} {size} Tf {x:.2f} {y:.2f} Td <{c:04X}> Tj ET" for c, x in zip(codes, xs))


def text_pdf(rows: list[tuple[str, float, float]], size: float = 12, extra_content: str = "", extra_objs=None) -> bytes:
    """Unicode の字の行を、Type0（Identity-H）＋ToUnicode で描く。rows: (字, x, y)。字の幅は 1em（全角）か /W で 0.5em（ASCII）。"""
    chars = sorted({ch for t, _x, _y in rows for ch in t})
    code = {ch: i + 1 for i, ch in enumerate(chars)}
    w = " ".join(f"{code[ch]} [500]" for ch in chars if ord(ch) < 0x2E80)
    parts = []
    for t, x, y in rows:
        hexs = "".join(f"{code[ch]:04X}" for ch in t)
        parts.append(f"BT /F1 {size} Tf {x} {y} Td <{hexs}> Tj ET")
    objs = [type0("TestFont", 3, f"/W [{w}]" if w else ""), stream(tounicode({c: ch for ch, c in code.items()}))]
    objs += extra_objs or []
    return build("\n".join(parts) + "\n" + extra_content, "<< /Font << /F1 2 0 R >> >>", objs)
