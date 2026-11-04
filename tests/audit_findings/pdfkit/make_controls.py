"""Controls: every (text, size, height) written in the dev forms, drawn with so much width that nothing can clip.
control_field.pdf: as a widget (/V + /DA + /AP) -- what each renderer shows for a field of the same height with unlimited width.
control_content.pdf: the same appearance as page content.  control_rows.json: position of each row."""
import json, re, sys, os
BASE = os.environ["TODOKU_PK_DIR"].rstrip("/") + "/"
from todoku.pdf.doc import PdfDocument
from todoku.pdf import fill as F
from todoku.pdf.layout import View
d = json.load(open(BASE + "var/field.json"))
rows = {}
for rel, v in d.items():
    for t in v["targets"]:
        if t["kind"] != "text":
            continue
        m = re.search(r"([\d.]+)\s+Tf", t["da"] or "")
        size = float(m[1]) if m else 0.0
        if size <= 0:
            continue
        rows[(t["text"].replace("\n", " "), size, round(t["rect"][3] - t["rect"][1], 1))] = None
W, PH = 2000, 1200
keys = sorted(rows)
place = []
pno, y = 1, 10.0
for k in keys:
    if y + k[2] + 12 > PH:
        pno, y = pno + 1, 10.0
    place.append((pno, y))
    y += k[2] + 14
npages = pno
objs = ["<</Type/Catalog/Pages 2 0 R>>", "<</Type/Pages/Kids[%s]/Count %d>>" % (" ".join(f"{3 + i} 0 R" for i in range(npages)), npages)]
for i in range(npages):
    objs.append(f"<</Type/Page/Parent 2 0 R/MediaBox[0 0 {W} {PH}]/Resources<<>>>>")
out = bytearray(b"%PDF-1.7\n")
off = []
for i, o in enumerate(objs):
    off.append(len(out))
    out += f"{i + 1} 0 obj\n{o}\nendobj\n".encode()
xr = len(out)
out += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode()
for x in off:
    out += f"{x:010d} 00000 n \n".encode()
out += f"trailer\n<</Size {len(objs) + 1}/Root 1 0 R>>\nstartxref\n{xr}\n%%EOF\n".encode()
blank = bytes(out)
meta = []
for style in ("field", "content"):
    doc = PdfDocument(blank)
    f = F.Filler(doc)
    f.draw_on_page = style == "content"
    pages = doc.pages()
    meta = []
    for i, ((text, size, h), (pno, y0)) in enumerate(zip(keys, place)):
        view = View(doc, pages[pno - 1])
        spot = {"id": f"c{i}", "kind": "text", "rect": [30.0, y0, W - 30.0, y0 + h], "size": size, "shape": "gap"}
        try:
            f.add_widget(pno, spot, text, view)
        except F.FillError as e:
            print("skip", text, size, h, str(e)[:80])
            continue
        meta.append({"text": text, "size": size, "h": h, "page": pno, "y0": y0, "y1": y0 + h})
    open(BASE + f"var/control_{style}.pdf", "wb").write(f.finish())
json.dump({"rows": meta, "W": W, "PH": PH, "npages": npages}, open(BASE + "var/control_rows.json", "w"), ensure_ascii=False)
print(len(keys), npages)
