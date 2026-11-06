"""比べる道具で書く・書いた結果を確かめる（採点 v2。scripts/eval/baseline_write.py の写し）。使える人が書くように書く（わざと弱くしない）。別の venv で python -I。
v1 からの違い：「消えた」（lost）を、書類の中身だけで数える。どの道具にも同じ（理由と試験は docs/再評価の計画.md・tests/test_eval_v2.py）：
  W1 parts（部品の数）をやめ、中身でない名前を除いて消えた部品を数える other_parts に：ZIP のフォルダの項目・OPC の [trash]・.rels・calcChain（計算の順の控え）・printerSettings（機械ごとの設定）・docProps・OPC の中身の説明・sharedStrings。
  W2 extLst（中身を問わない数）をやめる：styles の slicerStyles・workbook の calcFeatures などは見え方に関わらない。シートの中身の拡張は x14 の条件付き書式・入力規則・スパークライン・スライサーとして数える。
  W3 alternate_content：workbook.xml の x15ac:absPath（作った人の PC のパスの控え）を数えない。シート・図形の中のものは数える。
  W4 printer_settings をやめ、ページの設定（pageSetup・pageMargins・printOptions）を数える。
  W5 formula_cached_values：<v></v>（空）を保存された値に数えない。v1 は空でも数え、openpyxl が値を全部消しても lost にならなかった（比べる道具に有利だった）。
  W7 .doc の確かめ：LibreOffice の .txt で値が見えないときは .docx に変えて本文の XML も見る。
  W6 .xls の確かめ：xlrd を ignore_workbook_corruption=True で開く（元のファイルからこの OLE の警告が出る書類がある）。
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
    mp = re.match(r"P(\d+)\.c(\d+)$", tg.get("id") or "")
    if mp:  # 段落の中の空欄（文書の順の段落の番号と、前の字数）。同じ所に字を入れる
        ps = list(d.element.body.iter(W + "p"))
        n, off = int(mp.group(1)) - 1, int(mp.group(2))
        if n >= len(ps):
            raise ValueError("同じ所を指せない欄（段落がない）")
        pos = 0
        ts = list(ps[n].iter(W + "t"))
        for t in ts:
            ln = len(t.text or "")
            if pos + ln >= off:
                t.text = (t.text or "")[: off - pos] + val + (t.text or "")[off - pos:]
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                d.save(out)
                return {"method": "inline"}
            pos += ln
        if off == pos and not ts:
            ps[n].append(docx.oxml.parse_xml('<w:r xmlns:w="%s"><w:t>%s</w:t></w:r>' % (W[1:-1], val)))
            d.save(out)
            return {"method": "inline_empty"}
        raise ValueError("同じ所を指せない欄（段落の字が足りない）")
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
            cell = wb[sh][addr.split(":")[0]]
            if tg.get("kind") == "inline":  # マスの字の中の最初の空欄に入れる
                old = str(cell.value or "")
                new = re.sub(r"　+|[ ]{2,}", lambda _m: val, old, count=1)
                if new == old:
                    raise ValueError("同じ所を指せない欄（空欄がない）")
                cell.value = new
            else:
                cell.value = val
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
              "x14_conditionalFormatting": rb"<x14:conditionalFormatting[ >]", "sparklineGroup": rb"<x14:sparklineGroup[ >]", "slicer": rb"<x14:slicer[ >]|<x15:timelineRef[ >]",
              "definedName": rb"<(?:\w+:)?definedName[ >]", "shapes": rb"<xdr:sp[ >]", "x14_dataValidation": rb"<x14:dataValidation[ >]",
              "sdt": rb"<w:sdt[ >]", "textbox": rb"<w:txbxContent[ >]", "field": rb"<w:fldChar[ >]", "section": rb"<w:sectPr[ >]",
              "sheetProtection": rb"<sheetProtection[ >]", "legacyDrawing": rb"<legacyDrawing[ >]", "pageSetup": rb"<pageSetup[ >]", "pageMargins": rb"<pageMargins[ >]", "printOptions": rb"<printOptions[ >]",
              "threadedComments": rb"threadedComment", "pivot": rb"pivotTable", "tablePart": rb"<tablePart[ >]", "cell_style": rb"<xf[ >]"}
# 消えても書類の中身が変わらない部品の名前（W1）
NOT_CONTENT_PART = re.compile(r"(/$|^\[trash\]/|\.rels$|(^|/)calcChain\.xml$|/printerSettings/|^docProps/|^package/services/metadata/|^xl/sharedStrings\.xml$|^\[Content_Types\]\.xml$)", re.I)
_CACHED_V = re.compile(rb"<c\b[^>]*>\s*<f\b[^>]*?(?:/>|>.*?</f>)\s*<v>\s*[^<\s]", re.S)


def counts(z):
    names = z.namelist()
    c = {"images": sum("/media/" in n for n in names), "vba": sum(n.endswith("vbaProject.bin") for n in names),
         "charts": sum(bool(re.search(r"/charts/chart\d+\.xml$", n)) for n in names), "drawings": sum(bool(re.search(r"/drawings/drawing\d+\.xml$", n)) for n in names),
         "comments": sum(bool(re.search(r"comments\d*\.xml$", n)) for n in names), "form_controls": sum("/ctrlProps/" in n for n in names),
         "customXml": sum(n.startswith("customXml/") for n in names),
         "embeddings": sum("/embeddings/" in n for n in names)}
    txt = b"".join(z.read(n) for n in names if n.endswith((".xml", ".rels")))
    for k, pat in COUNT_PATS.items():
        c[k] = len(re.findall(pat, txt))
    # 図形の入れ子の代わりの形。workbook.xml の x15ac:absPath（作った人の PC のパスの控え）は数えない（W3）
    c["alternate_content"] = sum(len(re.findall(rb"<mc:AlternateContent[ >]", z.read(n))) for n in names if n.endswith(".xml") and n != "xl/workbook.xml")
    # 数式の保存された値（ビューアーで開いたときに見える値）。空の <v></v> は値がない（W5）
    c["formula_cached_values"] = len(_CACHED_V.findall(b"".join(z.read(n) for n in names if re.fullmatch(r"xl/worksheets/[^/]+\.xml", n))))
    return c


def other_parts(na: set, nb: set) -> list:
    """消えた部品のうち、中身のあるもの（W1）。名前の大文字小文字だけが変わったものは消えていない。"""
    low = {x.lower() for x in nb}
    return sorted(x for x in na if x.lower() not in low and not NOT_CONTENT_PART.search(x))


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
            gone = other_parts(na, nb)
            if gone:
                o["lost"]["other_parts"] = len(gone)
                o["other_parts_removed"] = gone[:30]
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

            bk = xlrd.open_workbook(file_contents=B, ignore_workbook_corruption=True)
            sh, addr = tg["id"].rsplit("!", 1)
            m = re.match(r"([A-Z]+)(\d+)", addr)
            col = 0
            for ch in m.group(1):
                col = col * 26 + ord(ch) - 64
            got = bk.sheet_by_name(sh).cell_value(int(m.group(2)) - 1, col - 1)
            o["landed"] = val in str(got)
            o["lost"] = {}
        elif kind == "doc":
            # 古い Word：LibreOffice で字にして、値があるかを見る（比べる道具はこの形式に書けない）
            import os
            import subprocess
            import tempfile

            d = tempfile.mkdtemp(prefix="todoku_doc_")
            subprocess.run(["soffice", f"-env:UserInstallation=file://{d}/prof", "--headless", "--norestore", "--convert-to", "txt:Text (encoded):UTF8", "--outdir", d, out],
                           capture_output=True, timeout=180)
            f = os.path.join(d, os.path.splitext(os.path.basename(out))[0] + ".txt")
            o["landed"] = os.path.exists(f) and val in open(f, encoding="utf-8", errors="replace").read()
            if not o["landed"]:  # W7：LibreOffice の .txt の書き出しは、表の中身を落とす書類がある（KR-text-001）。.docx に変えて本文の XML も見る
                subprocess.run(["soffice", f"-env:UserInstallation=file://{d}/prof", "--headless", "--norestore", "--convert-to", "docx", "--outdir", d, out],
                               capture_output=True, timeout=180)
                f = os.path.join(d, os.path.splitext(os.path.basename(out))[0] + ".docx")
                if os.path.exists(f):
                    body = zipfile.ZipFile(f).read("word/document.xml").decode("utf-8", "replace")
                    o["landed"] = val in re.sub(r"<[^>]+>", "", body)
            o["lost"] = {}
            o["checked_by"] = "libreoffice"
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
            o["xfa"] = ia["xfa"]
            if ia["xfa"]:  # Poppler は XFA の様式を描かないので、値がデータにあるかで確かめる
                af = rb.trailer["/Root"]["/AcroForm"].get_object()
                x = af["/XFA"]
                x = x.get_object() if hasattr(x, "get_object") else x
                blobs = [e.get_object().get_data() for e in list(x)[1::2]] if isinstance(x, list) else [x.get_data()]
                o["value_in_xfa"] = any(val.encode("utf-8") in b for b in blobs)
                ink = o["value_in_xfa"] or (rb.get_fields() or {}).get(tg.get("name"), {}).get("/V") == val
            if tg.get("name"):
                v = (rb.get_fields() or {}).get(tg["name"], {}).get("/V")
                o["field_value_ok"] = str(v) == val
            o["landed"] = bool(ink)
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
