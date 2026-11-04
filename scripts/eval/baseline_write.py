"""比べる道具で書く・書いた結果を確かめる。使える人が書くように書く（わざと弱くしない）。別の venv で python -I。
  baseline_write.py write <道具> <元> <出力> <kind> <欄の JSON> <値>
    openpyxl      読み込み（.xlsm は keep_vba）→セルに代入→保存
    python-docx   表のマスの段落に run を足す（空でなければ text に代入）・ほかの欄は、見出しの字のある段落に run を足す
    pypdf         フォームの欄は update_page_form_field_values（署名があれば incremental）。欄のない PDF は FreeText の注釈（incremental）
    pypdf_overlay reportlab で字を描いたページを merge_page で重ねる（欄のない PDF の、よくある書き方。日本語は CID フォント）
  baseline_write.py check <元> <出力> <kind> <欄の JSON> <値>
    とどくを使わずに、値が入ったか・何が消えたかを確かめる"""

import io
import json
import re
import sys
import time
import zipfile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _pdf_rect(page, rect):
    """とどくの rect は、表示の左上からの pt。PDF の座標（左下から。回転のないページ）に直す。"""
    box = page.cropbox
    x0, y0, x1, y1 = rect
    h, left, bottom = float(box.height), float(box.left), float(box.bottom)
    return (left + x0, bottom + h - y1, left + x1, bottom + h - y0)


def write_pdf(tool, p, out, tg, val):
    import pypdf

    rd = pypdf.PdfReader(p)
    fields = rd.get_fields() or {}
    signed = any(f.get("/FT") == "/Sig" and f.get("/V") is not None for f in fields.values())
    n = tg.get("page", 1) - 1
    if tool == "pypdf_overlay":
        if not tg.get("rect"):
            raise ValueError("位置が分からない欄")
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        from reportlab.pdfgen import canvas

        pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
        w = pypdf.PdfWriter(clone_from=rd)
        pg = w.pages[n]
        x0, y0, x1, y1 = _pdf_rect(pg, tg["rect"])
        buf = io.BytesIO()
        c = canvas.Canvas(buf, pagesize=(float(pg.mediabox.width), float(pg.mediabox.height)))
        size = min(max(y1 - y0 - 2, 5), 11)
        c.setFont("HeiseiKakuGo-W5", size)
        c.drawString(x0 + 1 - float(pg.mediabox.left), y0 + 1 - float(pg.mediabox.bottom), val)
        c.save()
        pg.merge_page(pypdf.PdfReader(io.BytesIO(buf.getvalue())).pages[0])
        w.write(out)
        return {"method": "overlay"}
    if tg.get("name"):
        w = pypdf.PdfWriter(p, incremental=True) if signed else pypdf.PdfWriter(clone_from=rd)
        w.update_page_form_field_values(w.pages[n], {tg["name"]: val}, auto_regenerate=False)
        w.write(out)
        return {"method": "form_field", "incremental": signed}
    if not tg.get("rect"):
        raise ValueError("位置が分からない欄")
    from pypdf.annotations import FreeText

    w = pypdf.PdfWriter(p, incremental=True)
    an = FreeText(text=val, rect=_pdf_rect(w.pages[n], tg["rect"]), font="Helvetica", font_size="10pt")
    w.add_annotation(page_number=n, annotation=an)
    w.write(out)
    return {"method": "freetext", "incremental": True}


def write_docx(p, out, tg, val):
    import docx

    d = docx.Document(p)
    m = re.match(r"T(\d+)\.R(\d+)\.C(\d+)$", tg.get("id") or "")
    if m:
        cell = d.tables[int(m.group(1)) - 1].rows[int(m.group(2)) - 1].cells[int(m.group(3)) - 1]
        if cell.text.strip():
            cell.text = val
        else:
            cell.paragraphs[0].add_run(val)
        d.save(out)
        return {"method": "table_cell"}
    label = (tg.get("label") or "").strip()
    paras = list(d.paragraphs) + [pp for t in d.tables for r in t.rows for c in r.cells for pp in c.paragraphs]
    for pp in paras:
        if label and label[:12] in pp.text:
            pp.add_run(val)
            d.save(out)
            return {"method": "after_label"}
    raise ValueError("同じ所を指せない欄（表のマスでなく、見出しの字も本文にない）")


def do_write(a):
    tool, p, out, kind, target, val = a
    tg = json.loads(target)
    o = {"tool": tool}
    t = time.time()
    try:
        if kind == "xlsx" and tool == "openpyxl":
            import openpyxl

            wb = openpyxl.load_workbook(p, keep_vba=p.lower().endswith("xlsm"))
            sh, addr = tg["id"].rsplit("!", 1)
            wb[sh][addr.split(":")[0]] = val
            wb.save(out)
        elif kind == "docx" and tool == "python-docx":
            o.update(write_docx(p, out, tg, val))
        elif kind == "pdf" and tool in ("pypdf", "pypdf_overlay"):
            o.update(write_pdf(tool, p, out, tg, val))
        else:
            raise ValueError(f"{tool} では {kind} に書けない")
    except Exception as e:  # noqa: BLE001
        o["error"] = f"{type(e).__name__}: {str(e)[:200]}"
    o["seconds"] = round(time.time() - t, 2)
    return o


COUNT_PATS = {"dataValidation": rb"<(?:\w+:)?dataValidation[ >]", "conditionalFormatting": rb"<(?:\w+:)?conditionalFormatting[ >]",
              "extLst": rb"<(?:\w+:)?extLst[ >]", "definedName": rb"<(?:\w+:)?definedName[ >]", "shapes": rb"<xdr:sp[ >]", "x14_dataValidation": rb"<x14:dataValidation[ >]",
              "sdt": rb"<w:sdt[ >]", "textbox": rb"<w:txbxContent[ >]", "field": rb"<w:fldChar[ >]", "section": rb"<w:sectPr[ >]",
              "alternate_content": rb"<mc:AlternateContent[ >]", "sheetProtection": rb"<sheetProtection[ >]", "legacyDrawing": rb"<legacyDrawing[ >]",
              "threadedComments": rb"threadedComment", "pivot": rb"pivotTable", "tablePart": rb"<tablePart[ >]", "cell_style": rb"<xf[ >]"}


def counts(z):
    names = z.namelist()
    c = {"parts": len(names), "images": sum("/media/" in n for n in names), "vba": sum(n.endswith("vbaProject.bin") for n in names),
         "charts": sum(bool(re.search(r"/charts/chart\d+\.xml$", n)) for n in names), "drawings": sum(bool(re.search(r"/drawings/drawing\d+\.xml$", n)) for n in names),
         "comments": sum(bool(re.search(r"comments\d*\.xml$", n)) for n in names), "form_controls": sum("/ctrlProps/" in n for n in names),
         "customXml": sum(n.startswith("customXml/") for n in names), "printer_settings": sum("printerSettings" in n for n in names),
         "embeddings": sum("/embeddings/" in n for n in names)}
    txt = b"".join(z.read(n) for n in names if n.endswith((".xml", ".rels")))
    for k, pat in COUNT_PATS.items():
        c[k] = len(re.findall(pat, txt))
    # 数式の保存された値（ビューアーで開いたときに見える値）
    fv = re.findall(rb"<c\b[^>]*>\s*<f\b[^>]*?(?:/>|>.*?</f>)\s*(<v>)?", b"".join(z.read(n) for n in names if re.fullmatch(r"xl/worksheets/[^/]+\.xml", n)), flags=re.S)
    c["formula_cached_values"] = sum(1 for x in fv if x)
    return c


def pdf_ink(a_path, b_path, page, rect):
    """Poppler で前後のページを描いて、変わった画素を数える（とどくの描画でない）。rect があれば、その中とそれ以外に分ける。"""
    import os
    import subprocess
    import tempfile

    from PIL import Image, ImageChops

    d = tempfile.mkdtemp(prefix="todoku_ink_")
    dpi = 100
    ims = []
    for name, src in (("a", a_path), ("b", b_path)):
        pre = os.path.join(d, name)
        subprocess.run(["pdftoppm", "-r", str(dpi), "-f", str(page), "-l", str(page), "-png", "-singlefile", src, pre], check=False, capture_output=True, timeout=120)
        ims.append(Image.open(pre + ".png").convert("L"))
    if ims[0].size != ims[1].size:
        return {"size_changed": True}
    diff = ImageChops.difference(*ims).point(lambda v: 255 if v > 40 else 0)
    total = sum(1 for v in diff.getdata() if v)
    inside = None
    if rect:
        k = dpi / 72
        box = (int(max(0, rect[0] * k - 4)), int(max(0, rect[1] * k - 4)), int(rect[2] * k + 4), int(rect[3] * k + 4))
        if box[2] > box[0] and box[3] > box[1]:
            inside = sum(1 for v in diff.crop(box).getdata() if v)
    return {"pixels_changed": total, "pixels_in_rect": inside}


def do_check(a):
    p, out, kind, target, val = a
    tg = json.loads(target)
    A, B = open(p, "rb").read(), open(out, "rb").read()
    o = {}
    try:
        if kind in ("xlsx", "docx"):
            za, zb = zipfile.ZipFile(io.BytesIO(A)), zipfile.ZipFile(io.BytesIO(B))
            na, nb = set(za.namelist()), set(zb.namelist())
            o["parts_removed"] = sorted(na - nb)[:30]
            o["parts_changed"] = sorted(n for n in na & nb if za.read(n) != zb.read(n))[:60]
            o["parts_total"] = len(na)
            o["parts_same"] = sum(1 for n in na & nb if za.read(n) == zb.read(n))
            ca, cb = counts(za), counts(zb)
            o["lost"] = {k: ca[k] - cb.get(k, 0) for k in ca if ca[k] - cb.get(k, 0) > 0}
            if kind == "xlsx":
                import openpyxl

                sh, addr = tg["id"].rsplit("!", 1)
                got = openpyxl.load_workbook(io.BytesIO(B))[sh][addr.split(":")[0]].value
                o["landed"] = val in str(got)
            else:
                body = zb.read("word/document.xml").decode("utf-8", "replace")
                o["landed"] = val in re.sub(r"<[^>]+>", "", body)
                import docx

                o["reopens"] = bool(docx.Document(io.BytesIO(B)))
        elif kind == "xls":
            import xlrd

            bk = xlrd.open_workbook(file_contents=B)
            sh, addr = tg["id"].rsplit("!", 1)
            m = re.match(r"([A-Z]+)(\d+)", addr)
            col = 0
            for ch in m.group(1):
                col = col * 26 + ord(ch) - 64
            got = bk.sheet_by_name(sh).cell_value(int(m.group(2)) - 1, col - 1)
            o["landed"] = val in str(got)
            o["lost"] = {}
        elif kind == "pdf":
            import pypdf

            o["original_bytes_kept"] = B[:len(A)] == A

            def info(r):
                root = r.trailer["/Root"]
                af = root.get("/AcroForm")
                af = af.get_object() if af else {}
                fl = r.get_fields() or {}
                return {"pages": len(r.pages), "fields": len(fl), "xfa": af.get("/XFA") is not None if af else False,
                        "metadata": root.get("/Metadata") is not None, "struct_tree": root.get("/StructTreeRoot") is not None, "outlines": root.get("/Outlines") is not None}

            ra, rb = pypdf.PdfReader(io.BytesIO(A)), pypdf.PdfReader(io.BytesIO(B))
            ia, ib = info(ra), info(rb)
            o["lost"] = {k: f"{ia[k]} → {ib[k]}" for k in ("xfa", "metadata", "struct_tree", "outlines") if ia[k] and not ib[k]}
            if ib["pages"] < ia["pages"]:
                o["lost"]["pages"] = f"{ia['pages']} → {ib['pages']}"
            o["ink"] = pdf_ink(p, out, tg.get("page", 1), tg.get("rect"))
            ink = o["ink"].get("pixels_in_rect") if tg.get("rect") else o["ink"].get("pixels_changed")
            o["visible"] = bool(ink)
            if tg.get("name"):
                v = (rb.get_fields() or {}).get(tg["name"], {}).get("/V")
                o["field_value_ok"] = str(v) == val
            o["landed"] = bool(o["visible"])
    except Exception as e:  # noqa: BLE001
        o["error"] = f"{type(e).__name__}: {str(e)[:200]}"
    return o


def do_crop(a):
    """Poppler で描いて、欄の周りを切り抜く（文字認識に渡す）。"""
    import os
    import subprocess
    import tempfile

    from PIL import Image

    pdf, page, rect, out = a
    rect = json.loads(rect)
    dpi = 300
    pre = os.path.join(tempfile.mkdtemp(prefix="todoku_crop_"), "p")
    subprocess.run(["pdftoppm", "-r", str(dpi), "-f", str(page), "-l", str(page), "-png", "-singlefile", pdf, pre], check=False, capture_output=True, timeout=120)
    im = Image.open(pre + ".png")
    if rect:
        k = dpi / 72
        im = im.crop((int(max(0, rect[0] * k - 40)), int(max(0, rect[1] * k - 40)), int(rect[2] * k + 40), int(rect[3] * k + 40)))
    im.save(out)
    return {"ok": True}


def main():
    if sys.argv[1] == "crop":
        print(json.dumps(do_crop(sys.argv[2:])))
    elif sys.argv[1] == "write":
        print(json.dumps(do_write(sys.argv[2:]), ensure_ascii=False, default=str))
    else:
        print(json.dumps(do_check(sys.argv[2:]), ensure_ascii=False, default=str))


if __name__ == "__main__":
    main()
