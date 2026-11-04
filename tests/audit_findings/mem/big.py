"""監査用：ふつうに大きな書類の常駐と時間（壊れていない書類）。大きさを変えて、線形かどうかを見る。

使い方：big.py 出力フォルダ [xlsx|csv|pdf|docx ...]
道具ごとに別のプロセス（窓の控えが効かないように）で、guard の上限（1.5 GB・60 秒）で動かす。
"""
import io
import json
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from guard import guarded  # noqa: E402


def make_xlsx(rows, blank=False, formulas=True):
    R = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=1) as z:
        z.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        z.writestr("_rels/.rels", R + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml", '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels", R + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        b = ['<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>',
             '<row r="1"><c r="A1" t="inlineStr"><is><t>品目</t></is></c><c r="B1" t="inlineStr"><is><t>数量</t></is></c><c r="C1" t="inlineStr"><is><t>単価</t></is></c><c r="D1" t="inlineStr"><is><t>金額</t></is></c><c r="E1" t="inlineStr"><is><t>備考</t></is></c></row>']
        for r in range(2, rows + 2):
            b.append(f'<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>品{r}</t></is></c><c r="B{r}"><v>{r % 97}</v></c><c r="C{r}"><v>{r % 1009}</v></c>' + (f'<c r="D{r}"><f>B{r}*C{r}</f><v>{(r % 97) * (r % 1009)}</v></c>' if formulas else "") + '' + (f'<c r="E{r}" s="0"/>' if blank else "") + "</row>")
        b.append("</sheetData></worksheet>")
        z.writestr("xl/worksheets/sheet1.xml", "".join(b))
    return out.getvalue()


def make_csv(rows):
    return b"id,name,qty,price,note\n" + b"".join(b"%d,item%d,%d,%d,\n" % (i, i, i % 97, i % 1009) for i in range(rows))


def make_pdf(pages):
    out = bytearray(b"%PDF-1.7\n")
    offs = {}

    def obj(n, body):
        offs[n] = len(out)
        out.extend(b"%d 0 obj\n" % n + body + b"\nendobj\n")
    kids = b" ".join(b"%d 0 R" % (10 + 2 * i) for i in range(pages))
    obj(1, b"<< /Type /Catalog /Pages 2 0 R >>")
    obj(2, b"<< /Type /Pages /Kids [" + kids + b"] /Count %d >>" % pages)
    obj(3, b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
    for i in range(pages):
        lines = b" ".join(b"1 0 0 1 72 %d Tm (Page %d line %d of the long report text) Tj" % (780 - 15 * k, i, k) for k in range(40))
        c = b"BT /F1 10 Tf " + lines + b" ET 72 100 400 50 re S"
        obj(10 + 2 * i, b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 3 0 R >> >> /Contents %d 0 R >>" % (11 + 2 * i))
        obj(11 + 2 * i, b"<< /Length %d >>\nstream\n" % len(c) + c + b"\nendstream")
    x = len(out)
    size = 10 + 2 * pages
    out += b"xref\n0 %d\n0000000000 65535 f \n" % size
    for n in range(1, size):
        out += (b"%010d 00000 n \n" % offs[n]) if n in offs else b"0000000000 00000 f \n"
    out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (size, x)
    return bytes(out)


def make_docx(pages):
    src = zipfile.ZipFile(os.path.abspath(os.path.join(HERE, "..", "..", "fixtures", "public", "form.docx")))
    head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
    b = [head]
    for p in range(pages):
        for k in range(30):
            b.append(f"<w:p><w:r><w:t>第{p}頁の{k}行目の本文です。申請の内容について、次のとおり説明します。氏名：＿＿＿＿</w:t></w:r></w:p>")
        b.append("<w:tbl>" + "".join("<w:tr>" + "".join(f"<w:tc><w:p><w:r><w:t>{'項目' if c == 0 else ''}</w:t></w:r></w:p></w:tc>" for c in range(4)) + "</w:tr>" for _ in range(5)) + "</w:tbl>")
        b.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    b.append("</w:body></w:document>")
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=1) as z:
        for i in src.infolist():
            z.writestr(i.filename, "".join(b) if i.filename == "word/document.xml" else src.read(i.filename))
    return out.getvalue()


SETS = {
    "xlsx": (make_xlsx, [1000, 2000, 4000, 8000], "xlsx", ["summary", "audit", "fill", "read_table", "calculate"], {"PROBE_CELL": "B2"}),
    "xlsxread": (make_xlsx, [100000], "xlsx", ["read_table"], {}),
    # 各行に 1 つ空のセルのある表（空の欄が行の数だけ「？」の所になる）
    "xlsxblank": (lambda n: make_xlsx(n, True), [500, 1000, 2000, 4000, 8000], "xlsx", ["summary"], {}),
    # 数式も空のセルもない表（「？」の所がない）
    "xlsxplain": (lambda n: make_xlsx(n, False, False), [10000, 30000, 100000], "xlsx", ["summary", "audit", "fill", "read_table"], {}),
    "pdfshape": (make_pdf, [100, 200, 400, 800], "pdf", ["summary", "grade", "audit", "fill", "view_page"], {}),
    "docxshape": (make_docx, [50, 100, 200], "docx", ["audit"], {}),
    "csvshape": (make_csv, [12500, 25000, 50000], "csv", ["audit"], {}),
    "csv": (make_csv, [100000, 300000, 1000000], "csv", ["summary", "grade", "audit", "fill"], {}),
    "pdf": (make_pdf, [500, 1500, 5000], "pdf", ["summary", "grade", "audit", "fill", "view_page"], {}),
    "docx": (make_docx, [30, 100, 300], "docx", ["summary", "grade", "audit", "fill"], {}),
}


def main():
    out = sys.argv[1]
    os.makedirs(out, exist_ok=True)
    want = sys.argv[2:] or list(SETS)
    for k in want:
        fn, sizes, ext, tools, env = SETS[k]
        for n in sizes:
            path = os.path.join(out, f"big_{k}_{n}.{ext}")
            if not os.path.exists(path):
                with open(path, "wb") as f:
                    f.write(fn(n))
            size = os.path.getsize(path)
            for t in tools:
                g = guarded(path, [t], limit_s=60, limit_gb=1.5, max_kills=1, env_extra=env, outdir=out)
                rec = next((x for x in g["recs"] if x.get("tool") == t), {})
                print(json.dumps({"set": k, "n": n, "bytes": size, "tool": t, "s": rec.get("s"), "rss_mb": rec.get("rss") or g["peak"],
                                  "ok": rec.get("ok"), "err": (rec.get("err") or "")[:70], "kill": g["kills"][0]["why"] if g["kills"] else ""}, ensure_ascii=False), flush=True)
                for fnm in os.listdir(out):
                    if fnm.startswith("out_"):
                        os.unlink(os.path.join(out, fnm))


if __name__ == "__main__":
    main()
