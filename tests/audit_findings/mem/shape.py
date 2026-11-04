"""監査用：大きさを倍々にして、時間とメモリの増え方（指数）を見る。壊れていない書類の、ありそうな形。

使い方：shape.py 出力フォルダ [名前の一部 ...]
n を 2 倍にして、前の 2 倍以上（指数 > 1.3）に増えるものを示す。1 つの道具 30 秒で次の大きさへは進まない。
"""
import json
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ.setdefault("CRAFT_ONLY", "1")
import crafted as C  # noqa: E402
from guard import guarded  # noqa: E402


def pdf_lines(n):
    k = n // 2
    return C.page(content=b"".join(b"%d 0 m %d 842 l S\n" % (i % 595, i % 595) for i in range(k)) + b"".join(b"0 %d m 595 %d l S\n" % (i % 842, i % 842) for i in range(k)))


def pdf_grid(n):
    # 本物の表：n 行の罫線と 10 列
    s = bytearray()
    for r in range(n):
        y = 800 - (r * 700 / n)
        s += b"50 %.2f m 550 %.2f l S\n" % (y, y)
    for c in range(11):
        s += b"%d 800 m %d 100 l S\n" % (50 + 50 * c, 50 + 50 * c)
    return C.page(content=bytes(s))


def pdf_rects(n):
    return C.page(content=b"".join(b"%d %d 5 5 re S\n" % (i % 590, (i // 590) % 840) for i in range(n)))


def pdf_spans(n):
    return C.page(content=b"BT /F1 6 Tf " + b"".join(b"1 0 0 1 %d %d Tm (ab) Tj\n" % (i % 590, 800 - (i // 100) % 800) for i in range(n)) + b"ET")


def pdf_spans_one_line(n):
    return C.page(content=b"BT /F1 1 Tf 0 700 Td " + b"(a) Tj 0.01 0 Td " * n + b"ET")


def pdf_widgets(n):
    more, ids = {}, []
    for i in range(n):
        more[10 + i] = b"<< /Type /Annot /Subtype /Widget /T (f%d) /FT /Tx /Rect [%d %d %d %d] /P 3 0 R /F 4 /DA (/Helv 10 Tf 0 g) >>" % (
            i, 50 + i % 40 * 12, 800 - (i // 40) * 12, 60 + i % 40 * 12, 810 - (i // 40) * 12)
        ids.append(b"%d 0 R" % (10 + i))
    arr = b" ".join(ids)
    return C.page(cat=b"/AcroForm << /Fields [" + arr + b"] /DR << /Font << /Helv 5 0 R >> >> >>", extra=b"/Annots [" + arr + b"]", more=more)


def docx_paras(n):
    return C.docx("".join(f"<w:p><w:r><w:t>第{i}段落の本文です。申請の内容。</w:t></w:r></w:p>" for i in range(n)))


def docx_runs(n):
    return C.docx("<w:p>" + "<w:r><w:t>あ</w:t></w:r>" * n + "</w:p>")


def docx_cells(n):
    row = "<w:tr>" + C.TC.format("") * 10 + "</w:tr>"
    return C.docx("<w:tbl>" + row * (n // 10) + "</w:tbl>")


def docx_vmerge(n):
    return C.docx("<w:tbl>" + f'<w:tr>{C.TC.format("""<w:vMerge w:val="restart"/>""")}</w:tr>' + f'<w:tr>{C.TC.format("<w:vMerge/>")}</w:tr>' * n + "</w:tbl>")


def xlsx_merges(n):
    m = "".join(f'<mergeCell ref="A{i + 3}:C{i + 3}"/>' for i in range(n))
    cells = "".join(f'<row r="{i + 3}"><c r="A{i + 3}" t="inlineStr"><is><t>見出し{i}</t></is></c></row>' for i in range(n))
    return C.xlsx(C.sheet(cells=cells, tail=f"<mergeCells>{m}</mergeCells>"))


def xlsx_formulas_cross(n):
    # 各行の数式が、上の列全部を読む（累計）：SUM(A$3:A{i})。範囲の広さの合計が n^2/2
    cells = "".join(f'<row r="{i}"><c r="A{i}"><v>1</v></c><c r="B{i}"><f>SUM(A$3:A{i})</f></c></row>' for i in range(3, n + 3))
    return C.xlsx(C.sheet(cells=cells))


def xlsx_blanks(n):
    cells = "".join(f'<row r="{i}"><c r="A{i}" t="inlineStr"><is><t>項目{i}</t></is></c><c r="B{i}" s="1"/></row>' for i in range(3, n + 3))
    return C.xlsx(C.sheet(cells=cells))


def xlsx_validations(n):
    dv = "".join(f'<dataValidation type="whole" sqref="B{i}"><formula1>1</formula1></dataValidation>' for i in range(3, n + 3))
    return C.xlsx(C.sheet(tail=f'<dataValidations count="{n}">{dv}</dataValidations>'))


def ods_rows(n):
    return C.ods("<table:table-column/>" + "".join(f'<table:table-row><table:table-cell office:value-type="string"><text:p>行{i}</text:p></table:table-cell></table:table-row>' for i in range(n)))


def rtf_rows(n):
    return C.RTF[:-1] + (b"\\trowd\\cellx1000\\cellx2000\\intbl a\\cell b\\cell\\row\n" * n) + b"}"


def json_items(n):
    return json.dumps({"items": [{"name": f"n{i}", "value": ""} for i in range(n)]}, ensure_ascii=False).encode()


def txt_lines(n):
    return ("".join(f"{i}行目の文です。\n" for i in range(n))).encode()


SETS = {
    "pdf_lines": (pdf_lines, "pdf", 100, 5, ["summary"]),
    "pdf_grid": (pdf_grid, "pdf", 100, 5, ["summary"]),
    "pdf_rects": (pdf_rects, "pdf", 2000, 5, ["summary"]),
    "pdf_spans": (pdf_spans, "pdf", 2000, 5, ["summary", "audit"]),
    "pdf_spans_one_line": (pdf_spans_one_line, "pdf", 5000, 5, ["summary", "grade"]),
    "pdf_widgets": (pdf_widgets, "pdf", 200, 5, ["summary", "audit", "fill"]),
    "docx_paras": (docx_paras, "docx", 2000, 5, ["summary", "audit", "fill"]),
    "docx_runs": (docx_runs, "docx", 5000, 5, ["summary", "audit", "fill"]),
    "docx_cells": (docx_cells, "docx", 5000, 5, ["summary", "audit", "fill"]),
    "docx_vmerge": (docx_vmerge, "docx", 2000, 5, ["summary", "audit", "fill"]),
    "xlsx_merges": (xlsx_merges, "xlsx", 500, 5, ["summary", "audit", "fill"]),
    "xlsx_formulas_cross": (xlsx_formulas_cross, "xlsx", 500, 5, ["summary", "audit", "calculate"]),
    "xlsx_blanks": (xlsx_blanks, "xlsx", 500, 5, ["summary", "audit", "fill"]),
    "xlsx_validations": (xlsx_validations, "xlsx", 500, 5, ["summary", "audit", "fill"]),
    "ods_rows": (ods_rows, "ods", 1000, 5, ["summary", "audit", "fill"]),
    "rtf_rows": (rtf_rows, "rtf", 1000, 5, ["summary", "audit", "fill"]),
    "json_items": (json_items, "json", 5000, 5, ["summary", "audit", "fill"]),
    "txt_lines": (txt_lines, "txt", 5000, 5, ["summary", "audit", "fill"]),
}


def main():
    out = sys.argv[1]
    os.makedirs(out, exist_ok=True)
    want = sys.argv[2:]
    for name, (fn, ext, n0, steps, tools) in SETS.items():
        if want and not any(w in name for w in want):
            continue
        for tool in tools:
            prev = None
            n = n0
            for _ in range(steps):
                path = os.path.join(out, f"{name}.{ext}")
                with open(path, "wb") as f:
                    f.write(fn(n))
                g = guarded(path, [tool], limit_s=30, limit_gb=1.5, max_kills=1, outdir=out,
                            env_extra={"PROBE_CELL": "A3"})
                rec = next((x for x in g["recs"] if x.get("tool") == tool), {})
                s = rec.get("s")
                rss = rec.get("rss") or g["peak"]
                ex = ""
                if prev and s and prev[0] and prev[0] > 0.2:
                    ex = f" exp={math.log2(s / prev[0]):.2f}"
                print(f"{name} {tool} n={n} {s}s rss={rss}MB{ex} {g['kills'][0]['why'] if g['kills'] else ''} {(rec.get('err') or '')[:50]}", flush=True)
                if g["kills"]:
                    break
                prev = (s, rss)
                n *= 2


if __name__ == "__main__":
    main()
