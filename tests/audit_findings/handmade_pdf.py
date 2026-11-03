"""手で書いた PDF：Type0（Identity-H）のフォントと、自分で決めた ToUnicode で、字を1つずつ指定の位置に置く。

字の見た目は確かめない（読む側が ToUnicode と位置から何を読むかだけを確かめる）。
- arabic_pf     : ToUnicode がアラビア文字の表示形（U+FExx）を指す。並びは見た目の順（右の字が後ろ）
- deva_visual   : ToUnicode は字ごと。前に出る母音（ि U+093F）を、見た目どおり子音の前に置く
- pua           : 外字（U+E000 台）を指す ToUnicode
- ivs           : 異体字セレクタつき（葛 + U+E0100）を1グリフに対応させる
- notu_identity : ToUnicode なし、埋め込みなしの Identity-H（文字集合 Adobe-Identity）
- notu_aj1      : ToUnicode なし、埋め込みなしの Identity-H（文字集合 Adobe-Japan1）。CID は AJ1 の番号
- vertical      : Identity-V の縦書き（ToUnicode あり）
使い方: python handmade_pdf.py OUTDIR
"""

from __future__ import annotations

import sys
import zlib
from pathlib import Path


def _cmap(pairs: dict[int, str]) -> bytes:
    lines = ["/CIDInit /ProcSet findresource begin", "12 dict begin", "begincmap",
             "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def",
             "/CMapName /Adobe-Identity-UCS def", "/CMapType 2 def",
             "1 begincodespacerange", "<0000> <FFFF>", "endcodespacerange", f"{len(pairs)} beginbfchar"]
    for c, s in pairs.items():
        u = "".join(f"{x:04X}" for x in s.encode("utf-16-be"))
        u = s.encode("utf-16-be").hex().upper()
        lines.append(f"<{c:04X}> <{u}>")
    lines += ["endbfchar", "endcmap", "CMapName currentdict /CMap defineresource pop", "end", "end"]
    return "\n".join(lines).encode()


def build(glyph_rows: list[list[tuple[float, int]]], tounicode: dict[int, str] | None, ordering: str = "Identity",
          vertical: bool = False, base: str = "ArialUnicodeMS", widths: str = "") -> bytes:
    """glyph_rows: 行ごとに [(x, 符号)]。行は上から 24pt ずつ下がる。縦書きは x が下への位置（列は右から）。"""
    objs: list[bytes] = []

    def add(b: bytes) -> int:
        objs.append(b)
        return len(objs)

    content = []
    for r, row in enumerate(glyph_rows):
        for x, code in row:
            if vertical:
                content.append(f"BT /F1 14 Tf {500 - 30 * r} {780 - x} Td <{code:04X}> Tj ET")
            else:
                content.append(f"BT /F1 14 Tf {40 + x} {800 - 24 * r} Td <{code:04X}> Tj ET")
    cs = zlib.compress("\n".join(content).encode())
    content_id = add(b"<< /Length %d /Filter /FlateDecode >>\nstream\n" % len(cs) + cs + b"\nendstream")
    tu_id = None
    if tounicode is not None:
        t = _cmap(tounicode)
        tu_id = add(b"<< /Length %d >>\nstream\n" % len(t) + t + b"\nendstream")
    desc_id = add(f"<< /Type /FontDescriptor /FontName /{base} /Flags 4 /FontBBox [0 -200 1000 900] /ItalicAngle 0 "
                  f"/Ascent 880 /Descent -120 /CapHeight 700 /StemV 80 >>".encode())
    cid_id = add(f"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /{base} /CIDSystemInfo << /Registry (Adobe) "
                 f"/Ordering ({ordering}) /Supplement 0 >> /FontDescriptor {desc_id} 0 R /DW 1000 {widths} >>".encode())
    enc = "Identity-V" if vertical else "Identity-H"
    font = f"<< /Type /Font /Subtype /Type0 /BaseFont /{base} /Encoding /{enc} /DescendantFonts [{cid_id} 0 R]"
    if tu_id:
        font += f" /ToUnicode {tu_id} 0 R"
    font_id = add((font + " >>").encode())
    page_id = len(objs) + 1
    pages_id = page_id + 1
    add(f"<< /Type /Page /Parent {pages_id} 0 R /MediaBox [0 0 595 842] /Contents {content_id} 0 R "
        f"/Resources << /Font << /F1 {font_id} 0 R >> >> >>".encode())
    add(f"<< /Type /Pages /Kids [{page_id} 0 R] /Count 1 >>".encode())
    cat_id = add(f"<< /Type /Catalog /Pages {pages_id} 0 R >>".encode())
    out = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
    offs = []
    for i, b in enumerate(objs, 1):
        offs.append(len(out))
        out += f"{i} 0 obj\n".encode() + b + b"\nendobj\n"
    x = len(out)
    out += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode()
    for o in offs:
        out += f"{o:010d} 00000 n \n".encode()
    out += f"trailer\n<< /Size {len(objs) + 1} /Root {cat_id} 0 R >>\nstartxref\n{x}\n%%EOF\n".encode()
    return bytes(out)


def row(text_codes: list[int], step: float = 9.0, rtl: bool = False) -> list[tuple[float, int]]:
    """符号を、見た目の左から右へ並べる。rtl のときは、渡した論理順の最初の字をいちばん右に置く。"""
    n = len(text_codes)
    if rtl:
        return [((n - 1 - i) * step, c) for i, c in enumerate(text_codes)]
    return [(i * step, c) for i, c in enumerate(text_codes)]


def samples() -> dict[str, bytes]:
    out = {}
    # 1. アラビア語「محمد」を表示形で：م(初) ح(中) م(中) د(終)、論理順で ToUnicode
    pf = ["ﻣ", "ﺤ", "ﻤ", "ﺪ"]  # meem initial, hah medial, meem medial, dal final
    # 「الاسم」: ا ل ا س م → ﺍ ﻟ(init) ﺎ? 簡単に: ﺍ ﻻ ﺳ ﻢ（لا の合字 U+FEFB）
    pf2 = ["ﺍ", "ﻻ", "ﺳ", "ﻢ"]
    tu = {i + 1: ch for i, ch in enumerate(pf + pf2)}
    out["arabic_pf.pdf"] = build([row([1, 2, 3, 4], rtl=True), row([5, 6, 7, 8], rtl=True)], tu)
    # 2. デーヴァナーガリー「किशोर」：見た目は ि क श ो र（ि が前）。字ごとの ToUnicode
    tu = {1: "ि", 2: "क", 3: "श", 4: "ो", 5: "र", 6: "क्ष"}
    out["deva_visual.pdf"] = build([[(0, 1), (5, 2), (14, 3), (20, 4), (23, 5)], [(0, 1), (5, 6)]], tu)
    # 3. 外字
    tu = {1: "山", 2: "", 3: "太", 4: "郎"}
    out["pua.pdf"] = build([row([1, 2, 3, 4], step=14)], tu)
    # 4. 異体字セレクタ
    tu = {1: "葛\U000E0100", 2: "飾", 3: "区"}
    out["ivs.pdf"] = build([row([1, 2, 3], step=14)], tu)
    # 5. ToUnicode なし、埋め込みなし（Identity）。符号はグリフ番号のつもり
    out["notu_identity.pdf"] = build([row([0x0123, 0x0456, 0x0789], step=14)], None)
    # 6. ToUnicode なし、埋め込みなし（Adobe-Japan1）：AJ1 の CID 1200 前後（漢字）
    out["notu_aj1.pdf"] = build([row([1125, 1200, 3000], step=14)], None, ordering="Japan1", base="MS-Mincho")
    # 7. 縦書き（Identity-V）「申請者」
    tu = {1: "申", 2: "請", 3: "者", 4: "（", 5: "）"}
    out["vertical.pdf"] = build([[(0, 1), (14, 2), (28, 3)], [(0, 4), (14, 1), (28, 5)]], tu, vertical=True)
    # 8. 分けて描いた結合文字（NFD）：「Họ tên:」の ọ・ê を、基の字と、上に重ねた記号のグリフで描く
    chars = ["H", "o", "̣", " ", "t", "e", "̂", "n", ":"]
    tu = {i + 1: c for i, c in enumerate(chars)}
    xs = [0, 9, 10, 18, 27, 36, 37, 45, 54]  # 記号は、基の字の 1pt 右から（同じマスの中）
    out["nfd_marks.pdf"] = build([[(x, i + 1) for i, x in enumerate(xs)]], tu)
    # 9. タイ語の上下の記号を、基の字と別のグリフで（Word の書き出しと同じ）。ชื่อ = ช ื ่ อ
    chars = ["ช", "ื", "่", "อ"]
    tu = {i + 1: c for i, c in enumerate(chars)}
    out["thai_marks.pdf"] = build([[(0, 1), (2, 2), (3, 3), (9, 4)]], tu)
    # 10. 幅の狭い字の重なり（Helvetica・Arial の l・i は 0.222em）：「Castillo Williams」を、字の幅どおりに並べる
    word = "Castillo Williams"
    adv = {"C": 722, "a": 556, "s": 500, "t": 278, "i": 222, "l": 222, "o": 556, " ": 278, "W": 944, "m": 833}
    codes = {ch: i + 1 for i, ch in enumerate(sorted(set(word)))}
    tu = {c: ch for ch, c in codes.items()}
    w = " ".join(f"{c} [{adv[ch]}]" for ch, c in codes.items())
    x, rowx = 0.0, []
    for ch in word:
        rowx.append((x, codes[ch]))
        x += adv[ch] * 14 / 1000
    out["narrow_double.pdf"] = build([rowx], tu, widths=f"/W [{w}]")
    return out


if __name__ == "__main__":
    d = Path(sys.argv[1])
    d.mkdir(parents=True, exist_ok=True)
    for name, blob in samples().items():
        (d / name).write_bytes(blob)
        print(d / name)
