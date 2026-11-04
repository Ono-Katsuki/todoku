"""Measure PDFKit: for each (value, font, size, Q), a ladder of field widths; which width shows the value whole?
Writes a synthetic PDF with widgets (no page content at all), renders with PDFKit, reports for each width the ink extent."""
import os, subprocess, sys, json
import numpy as np
from PIL import Image
PK = os.environ["TODOKU_PK_DIR"].rstrip("/") + "/"

def u16(s):
    return "<FEFF" + s.encode("utf-16-be").hex().upper() + ">"

objs = {}
def add(n, s):
    objs[n] = s

CASES = [("昭和60", "J", 10), ("昭和60", "J", 8), ("H20", "H", 10), ("H20", "J", 10), ("平成20", "J", 11), ("03-0000-0000", "H", 8), ("03-0000-0000", "J", 10.5), ("12", "H", 10), ("山田 花子", "J", 11)]
add(1, "<</Type/Catalog/Pages 2 0 R/AcroForm<</Fields[%FIELDS%]/DR<</Font<</H 3 0 R/J 4 0 R>>>>/DA(/H 0 Tf 0 g)>>>>")
add(3, "<</Type/Font/Subtype/Type1/BaseFont/Helvetica/Encoding/WinAnsiEncoding>>")
add(4, "<</Type/Font/Subtype/Type0/BaseFont/HeiseiKakuGo-W5-UniJIS-UCS2-HW-H/Encoding/UniJIS-UCS2-HW-H/DescendantFonts[5 0 R]>>")
add(5, "<</Type/Font/Subtype/CIDFontType0/BaseFont/HeiseiKakuGo-W5/CIDSystemInfo<</Registry(Adobe)/Ordering(Japan1)/Supplement 2>>/DW 1000>>")
n = 10
fields = []
layout = []
y = 780
HEIGHT = 16
for val, font, size in CASES:
    for q in (0, 2):
        x = 20
        for extra in (0, 2, 4, 6, 8, 10, 12, 16, 24):
            em = sum(1.0 if ord(c) > 0x7f else 0.56 for c in val) * size
            w = em + extra
            add(n, f"<</Type/Annot/Subtype/Widget/FT/Tx/T(f{n})/F 4/Rect[{x:.2f} {y} {x+w:.2f} {y+HEIGHT}]/V {u16(val)}/DA(/{font} {size} Tf 0 g)/Q {q}/P 6 0 R>>")
            fields.append(f"{n} 0 R")
            layout.append(dict(val=val, font=font, size=size, q=q, extra=extra, em=em, rect=[x, y, x + w, y + HEIGHT]))
            n += 1
            x += w + 14
        y -= HEIGHT + 6
add(6, "<</Type/Page/Parent 2 0 R/MediaBox[0 0 2000 842]/Resources<<>>/Annots[%ANN%]>>")
add(2, "<</Type/Pages/Kids[6 0 R]/Count 1>>")
out = bytearray(b"%PDF-1.7\n")
off = {}
for k in sorted(objs):
    off[k] = len(out)
    s = objs[k].replace("%FIELDS%", " ".join(fields)).replace("%ANN%", " ".join(fields))
    out += f"{k} 0 obj\n{s}\nendobj\n".encode()
xr = len(out)
mx = max(objs) + 1
out += f"xref\n0 {mx}\n0000000000 65535 f \n".encode()
for k in range(1, mx):
    out += (f"{off[k]:010d} 00000 n \n".encode() if k in off else b"0000000000 65535 f \n")
out += f"trailer\n<</Size {mx}/Root 1 0 R>>\nstartxref\n{xr}\n%%EOF\n".encode()
open(PK + "synth.pdf", "wb").write(out)
DPI = 300
subprocess.run([PK + "pkr", "render", PK + "synth.pdf", "1", PK + "synth.png", str(DPI), "pdfkit"], check=True)
im = np.array(Image.open(PK + "synth.png").convert("L"))
k = DPI / 72
H = im.shape[0]
res = {}
for c in layout:
    x0, y0, x1, y1 = c["rect"]
    r = im[int(H - y1 * k):int(H - y0 * k), int((x0 - 6) * k):int((x1 + 6) * k)]
    cols = np.where((r < 128).any(axis=0))[0]
    rows = np.where((r < 128).any(axis=1))[0]
    ink = int((r < 128).sum())
    ext = (0, 0) if not len(cols) else (cols[0] / k - 6, cols[-1] / k - 6)
    res.setdefault((c["val"], c["font"], c["size"], c["q"]), []).append((c["extra"], round(ext[0], 1), round(ext[1], 1), round(c["rect"][2] - c["rect"][0], 1), ink, 0 if not len(rows) else round((len(r) - rows[-1]) / k, 1)))
for key, rows in res.items():
    full = max(x[4] for x in rows)
    print(key)
    for ex, l, r_, w, ink, base in rows:
        print(f"   extra+{ex:2d} w={w:6.1f} ink x={l:5.1f}..{r_:5.1f} inkpx={ink:6d} {'FULL' if ink > 0.97 * full else 'CLIP'}  bottom={base}")
d = json.loads(subprocess.run([PK + "pkr", "dump", PK + "synth.pdf"], capture_output=True, text=True).stdout)
print({(a["font"], a["fontSize"]) for a in d["pages"][0]["annots"]})
