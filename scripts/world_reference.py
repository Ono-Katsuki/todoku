"""参照の読み：とどくを使わず、確立した別の道具で書類を読む（段階の検証・道具の比べの「参照」）。

とどくの依存には入れない別の環境（openpyxl・xlrd・pdfplumber・pypdf・python-docx）で、Python を -I で動かす。
書類は信頼しないデータとして扱う（マクロは動かさない。読むだけ）。

使い方: python -I scripts/world_reference.py <書類> <形式> <単位を書き出す JSON> [--textutil <textutil の出力>] [--pdf-fast]
最後の行に、構造の情報を JSON で出す。単位（セルの値・段落・行）は、別のファイルに {"units": [...]} で書き出す。
"""

from __future__ import annotations

import csv
import datetime as dt
import io
import json
import re
import sys
import time
import zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
EXCEL_EPOCH = dt.datetime(1899, 12, 30)


def serial(v):
    if isinstance(v, dt.datetime):
        return (v - EXCEL_EPOCH).total_seconds() / 86400
    if isinstance(v, dt.date):
        return (dt.datetime(v.year, v.month, v.day) - EXCEL_EPOCH).days
    if isinstance(v, dt.time):
        return (v.hour * 3600 + v.minute * 60 + v.second) / 86400
    return None


def alts(v) -> list:
    """そのセルの値として受け入れる形（生の値・日付の通し番号・ISO の日付）。"""
    if v is None:
        return []
    s = serial(v)
    if s is not None:
        return [s, v.isoformat()]
    if isinstance(v, bool):
        return ["TRUE" if v else "FALSE", v]
    return [v]


def decode(b: bytes) -> tuple[str, str]:
    for enc in ("utf-8-sig",):
        try:
            return b.decode(enc), enc
        except UnicodeDecodeError:
            pass
    try:
        from charset_normalizer import from_bytes

        best = from_bytes(b).best()
        if best is not None:
            return str(best), best.encoding
    except Exception:  # noqa: BLE001
        pass
    for enc in ("cp932", "cp1252", "latin-1"):
        try:
            return b.decode(enc), enc
        except UnicodeDecodeError:
            pass
    return b.decode("latin-1"), "latin-1"


# ---------------------------------------------------------------------- CSV
def ref_csv(b: bytes) -> tuple[list, dict]:
    text, enc = decode(b)
    try:
        dialect = csv.Sniffer().sniff(text[:65536], delimiters=",;\t|")
        delim = dialect.delimiter
    except csv.Error:
        delim = ","
    try:
        rows = list(csv.reader(io.StringIO(text, newline=""), delimiter=delim))
    except csv.Error as e:  # 行の区切りが混ざった CSV：行ごとに区切る
        rows = [ln.split(delim) for ln in text.splitlines()]
        delim += f" (csv.Error: {e})"
    units = [c for r in rows for c in r if c.strip()]
    return units, {"encoding": enc, "delimiter": delim, "rows": len(rows), "cols": max((len(r) for r in rows), default=0)}


# ---------------------------------------------------------------------- xlsx
def ref_xlsx_raw(b: bytes, why: str) -> tuple[list, dict]:
    """openpyxl が読めない xlsx：部品の XML を直に読む（共有文字列・シートのセルの値だけ）。"""
    S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    z = zipfile.ZipFile(io.BytesIO(b))
    names = z.namelist()
    sst = []
    if "xl/sharedStrings.xml" in names:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.iter(S + "si"):
            sst.append("".join(t.text or "" for t in si.iter(S + "t")))
    sheets, units = [], []
    sheet_name = {}
    try:  # 部品の名前 → シートの名前（workbook.xml と、その関係の部品）
        R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        tgt = {r.get("Id"): "xl/" + r.get("Target").lstrip("/").removeprefix("xl/") for r in rels}
        for sh in ET.fromstring(z.read("xl/workbook.xml")).iter(S + "sheet"):
            sheet_name[tgt.get(sh.get(R + "id"), "")] = sh.get("name")
    except (KeyError, ET.ParseError):
        pass
    for n in sorted(x for x in names if re.match(r"xl/worksheets/sheet\d+\.xml$", x)):
        root = ET.fromstring(z.read(n))
        values, formulas = {}, {}
        for c in root.iter(S + "c"):
            t = c.get("t")
            v = c.find(S + "v")
            f = c.find(S + "f")
            if f is not None:
                formulas[c.get("r")] = "=" + (f.text or "")
            if t == "s" and v is not None:
                val = sst[int(v.text)]
                units.append(val)
            elif t == "inlineStr":
                val = "".join(x.text or "" for x in c.iter(S + "t"))
                units.append(val)
            elif v is not None:
                try:
                    val = float(v.text)
                except (TypeError, ValueError):
                    val = v.text
            else:
                continue
            values[c.get("r")] = [val]
        sheets.append({"name": sheet_name.get(n, n.rsplit("/", 1)[1]), "state": "visible", "values": values, "formulas": formulas, "merged": []})
    return units, {"sheets": sheets, "reader": "raw-xml", "openpyxl_error": why}


def ref_xlsx(path: str, b: bytes) -> tuple[list, dict]:
    import openpyxl

    try:
        openpyxl.load_workbook(path, read_only=True).close()
    except Exception as e:  # noqa: BLE001
        return ref_xlsx_raw(b, f"{type(e).__name__}: {str(e)[:120]}")
    wb_f = openpyxl.load_workbook(path, data_only=False)
    wb_v = openpyxl.load_workbook(path, data_only=True)
    sheets, units = [], []
    for ws in wb_f.worksheets:
        wv = wb_v[ws.title]
        values, formulas, fmts = {}, {}, {}
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                v = wv[c.coordinate].value
                if isinstance(c.value, str) and c.value.startswith("="):
                    formulas[c.coordinate] = c.value
                elif getattr(c, "data_type", "") == "f":
                    formulas[c.coordinate] = str(c.value)
                if v is not None:
                    values[c.coordinate] = alts(v)
                    if isinstance(v, str):
                        units.append(v)
                    if c.number_format and c.number_format != "General":
                        fmts[c.coordinate] = c.number_format
        hidden_rows = [i for i, d in ws.row_dimensions.items() if d.hidden]
        hidden_cols = [k for k, d in ws.column_dimensions.items() if d.hidden]
        sheets.append({"name": ws.title, "state": ws.sheet_state, "values": values, "formulas": formulas,
                       "merged": [str(r) for r in ws.merged_cells.ranges], "hidden_rows": len(hidden_rows),
                       "hidden_cols": len(hidden_cols), "validations": len(ws.data_validations.dataValidation),
                       "protected": bool(ws.protection.sheet), "images": len(getattr(ws, "_images", [])),
                       "charts": len(getattr(ws, "_charts", [])), "cond_formats": len(ws.conditional_formatting),
                       "comments": sum(1 for r in ws.iter_rows() for c in r if c.comment), "number_formats": len(fmts),
                       "number_format_examples": dict(list(fmts.items())[:5])})
    for ws in wb_f.chartsheets:
        sheets.append({"name": ws.title, "state": ws.sheet_state, "values": {}, "formulas": {}, "merged": [], "chartsheet": True})
    z = zipfile.ZipFile(io.BytesIO(b))
    names = z.namelist()
    info = {"sheets": sheets, "parts": len(names), "vba": any(n.endswith("vbaProject.bin") for n in names),
            "drawings": sum(1 for n in names if n.startswith("xl/drawings/drawing")),
            "textbox_text": shapes_text(z), "defined_names": len(list(wb_f.defined_names)) if hasattr(wb_f, "defined_names") else None}
    units += info["textbox_text"]
    return units, info


def shapes_text(z: zipfile.ZipFile) -> list[str]:
    """図形（テキストボックス）の字：xl/drawings/*.xml の a:t。"""
    out = []
    for n in z.namelist():
        if re.match(r"xl/drawings/drawing\d+\.xml$", n):
            root = ET.fromstring(z.read(n))
            for sp in root.iter("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}sp"):
                t = "".join(x.text or "" for x in sp.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t"))
                if t.strip():
                    out.append(t)
    return out


# ---------------------------------------------------------------------- xls
def ref_xls(path: str) -> tuple[list, dict]:
    import xlrd

    try:
        bk = xlrd.open_workbook(path, formatting_info=True, on_demand=False)
        fmt = True
    except Exception:  # noqa: BLE001  （formatting_info を読めない形）
        bk = xlrd.open_workbook(path)
        fmt = False
    sheets, units = [], []
    for sh in bk.sheets():
        values = {}
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    continue
                addr = xlrd.formula.cellname(r, c).replace("$", "")
                v = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        d = xlrd.xldate.xldate_as_datetime(v, bk.datemode)
                        values[addr] = [v, d.isoformat()]
                    except Exception:  # noqa: BLE001
                        values[addr] = [v]
                elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    values[addr] = ["TRUE" if v else "FALSE", bool(v)]
                elif cell.ctype == xlrd.XL_CELL_ERROR:
                    values[addr] = [xlrd.error_text_from_code.get(v, "#ERR")]
                else:
                    if isinstance(v, str) and not v.strip():
                        continue
                    values[addr] = [v]
                    if isinstance(v, str):
                        units.append(v)
        merged = [f"{xlrd.formula.cellname(r0, c0)}:{xlrd.formula.cellname(r1 - 1, c1 - 1)}".replace("$", "")
                  for r0, r1, c0, c1 in (sh.merged_cells if fmt else [])]
        hidden_rows = sum(1 for ri in (sh.rowinfo_map.values() if fmt else []) if ri.hidden)
        hidden_cols = sum(1 for ci in (sh.colinfo_map.values() if fmt else []) if ci.hidden)
        sheets.append({"name": sh.name, "state": {0: "visible", 1: "hidden", 2: "veryHidden"}.get(sh.visibility, "visible"),
                       "values": values, "formulas": {}, "merged": merged, "hidden_rows": hidden_rows, "hidden_cols": hidden_cols})
    return units, {"sheets": sheets, "formatting_info": fmt, "biff": bk.biff_version}


# ---------------------------------------------------------------------- OpenDocument
TN = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
TX = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
OF = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"


def _col(n: int) -> str:
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def odf_text(el) -> str:
    out = []
    for p in el.iter():
        if p.tag in (TX + "p", TX + "h"):
            out.append("".join(p.itertext()))
    return "\n".join(out)


def ref_ods(b: bytes) -> tuple[list, dict]:
    z = zipfile.ZipFile(io.BytesIO(b))
    root = ET.fromstring(z.read("content.xml"))
    sheets, units = [], []
    for tb in root.iter(TN + "table"):
        values = {}
        r = 0
        for row in tb.iter(TN + "table-row"):
            rrep = min(int(row.get(TN + "number-rows-repeated", "1")), 2000)
            cells = [c for c in row if c.tag in (TN + "table-cell", TN + "covered-table-cell")]
            for _ in range(rrep):
                c_i = 0
                for c in cells:
                    crep = min(int(c.get(TN + "number-columns-repeated", "1")), 1000)
                    vt = c.get(OF + "value-type")
                    txt = odf_text(c)
                    v = None
                    if vt in ("float", "percentage", "currency"):
                        v = [float(c.get(OF + "value")), txt]
                    elif vt == "date":
                        v = [c.get(OF + "date-value"), txt]
                    elif vt == "boolean":
                        v = [c.get(OF + "boolean-value"), txt]
                    elif txt.strip():
                        v = [txt]
                    if v is not None:
                        for k in range(crep):
                            values[f"{_col(c_i + k)}{r + 1}"] = v
                            if vt in (None, "string") and k == 0:
                                units.append(txt)
                    c_i += crep
                r += 1
        sheets.append({"name": tb.get(TN + "name"), "state": "visible", "values": values, "formulas": {}, "merged": []})
    return units, {"sheets": sheets}


def ref_odt(b: bytes) -> tuple[list, dict]:
    z = zipfile.ZipFile(io.BytesIO(b))
    units = []
    for part in ("content.xml", "styles.xml"):
        if part in z.namelist():
            root = ET.fromstring(z.read(part))
            for p in root.iter():
                if p.tag in (TX + "p", TX + "h"):
                    # 入れ子の段落（枠の中）は、その段落として数える：直下の字だけ
                    t = "".join(_own_text(p))
                    if t.strip():
                        units.append(t)
    return units, {}


def _own_text(p):
    if p.text:
        yield p.text
    for ch in p:
        if ch.tag not in (TX + "p", TX + "h"):
            yield from _own_text(ch)
        if ch.tail:
            yield ch.tail


# ---------------------------------------------------------------------- docx
def ref_docx(b: bytes) -> tuple[list, dict]:
    z = zipfile.ZipFile(io.BytesIO(b))
    names = z.namelist()
    parts = [n for n in names if re.match(r"word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$", n)]
    units = []
    counts = {"sdt": 0, "ffData": 0, "txbx": 0, "drawing": 0, "pict": 0, "tables": 0, "checkbox_sdt": 0}
    for n in parts:
        root = ET.fromstring(z.read(n))
        counts["sdt"] += sum(1 for _ in root.iter(W + "sdt"))
        counts["ffData"] += sum(1 for _ in root.iter(W + "ffData"))
        counts["txbx"] += sum(1 for _ in root.iter(W + "txbxContent"))
        counts["drawing"] += sum(1 for _ in root.iter(W + "drawing"))
        counts["pict"] += sum(1 for _ in root.iter(W + "pict"))
        counts["tables"] += sum(1 for _ in root.iter(W + "tbl"))
        counts["checkbox_sdt"] += sum(1 for _ in root.iter("{http://schemas.microsoft.com/office/word/2010/wordml}checkbox"))
        _walk_docx(root, units)
    info = {"parts": parts, "counts": counts, "vba": any(n.endswith("vbaProject.bin") for n in names),
            "embeddings": [n for n in names if n.startswith("word/embeddings/")][:20]}
    try:
        import docx  # python-docx：段落と表（既定の読み方）

        d = docx.Document(io.BytesIO(b))
        info["python_docx_paragraphs"] = len(d.paragraphs)
    except Exception as e:  # noqa: BLE001
        info["python_docx_error"] = f"{type(e).__name__}: {e}"[:160]
    return units, info


def _walk_docx(el, units, cur=None):
    """段落ごとに、その段落の字（入れ子の段落の字は、入れ子の段落に）。mc:Fallback（同じ中身の写し）は飛ばす。"""
    for ch in el:
        if ch.tag == MC + "Fallback":
            continue
        if ch.tag == W + "p":
            buf: list[str] = []
            _walk_docx(ch, units, buf)
            t = "".join(buf)
            if t.strip():
                units.append(t)
        elif ch.tag == W + "t" and cur is not None:
            cur.append(ch.text or "")
        elif ch.tag in (W + "tab",) and cur is not None:
            cur.append(" ")
        elif ch.tag == W + "delText":
            continue
        else:
            _walk_docx(ch, units, cur)


# ---------------------------------------------------------------------- PDF
def ref_pdf(path: str, fast: bool) -> tuple[list, dict]:
    import pypdf

    info: dict = {}
    rd = pypdf.PdfReader(path)
    info["encrypted"] = rd.is_encrypted
    if rd.is_encrypted:
        try:
            rd.decrypt("")
        except Exception as e:  # noqa: BLE001
            info["decrypt_error"] = str(e)[:100]
    try:
        flds = rd.get_fields() or {}
    except Exception:  # noqa: BLE001
        flds = {}
    info["acroform_fields"] = len(flds)
    info["field_types"] = dict(__import__("collections").Counter(str(v.get("/FT")) for v in flds.values()))
    try:
        root = rd.trailer["/Root"]
        af = root.get("/AcroForm")
        info["xfa"] = bool(af and af.get_object().get("/XFA") is not None)
    except Exception:  # noqa: BLE001
        info["xfa"] = None
    widgets = 0
    pages = []
    units = []
    for i, pg in enumerate(rd.pages, 1):
        try:
            for an in pg.get("/Annots") or []:
                if an.get_object().get("/Subtype") == "/Widget":
                    widgets += 1
        except Exception:  # noqa: BLE001
            pass
        if fast:
            try:
                t = pg.extract_text() or ""
            except Exception:  # noqa: BLE001
                t = ""
            pages.append({"n": i, "chars": len(re.sub(r"\s", "", t)), "images": -1, "cid": t.count("(cid:")})
            units += [ln for ln in t.splitlines() if ln.strip()]
    info["widgets"] = widgets
    if not fast:
        import pdfplumber

        with pdfplumber.open(path) as pdf:
            for i, pg in enumerate(pdf.pages, 1):
                t = pg.extract_text() or ""
                pages.append({"n": i, "chars": len(pg.chars), "images": len(pg.images), "cid": t.count("(cid:")})
                units += [ln for ln in t.splitlines() if ln.strip()]
                pg.flush_cache()
    info["pages"] = pages
    info["reader"] = "pypdf" if fast else "pdfplumber"
    # (cid:n) は字に直せない字：単位からは除く（参照でも読めない字）
    units = [re.sub(r"\(cid:\d+\)", "", u) for u in units]
    return units, info


# ---------------------------------------------------------------------- 入口
def read(path: str, kind: str, textutil: str | None, fast: bool, depth: int = 0) -> tuple[list, dict]:
    b = open(path, "rb").read()
    if kind == "csv":
        return ref_csv(b)
    if kind == "xlsx":
        return ref_xlsx(path, b)
    if kind == "xls":
        return ref_xls(path)
    if kind == "ods":
        return ref_ods(b)
    if kind == "odt":
        return ref_odt(b)
    if kind == "docx":
        return ref_docx(b)
    if kind == "pdf":
        return ref_pdf(path, fast)
    if kind in ("doc", "rtf"):
        if textutil:
            t = open(textutil, encoding="utf-8", errors="replace").read()
            return [ln for ln in t.splitlines() if ln.strip()], {"reader": "textutil"}
        return [], {"reader": "none"}
    if kind == "zip" and depth == 0:
        import tempfile
        import os

        z = zipfile.ZipFile(io.BytesIO(b))
        units, members = [], []
        ext_kind = {".csv": "csv", ".xlsx": "xlsx", ".xls": "xls", ".pdf": "pdf", ".docx": "docx", ".ods": "ods",
                    ".odt": "odt", ".txt": "txt", ".tsv": "csv"}
        with tempfile.TemporaryDirectory() as td:
            for n in z.namelist()[:50]:
                k = ext_kind.get(os.path.splitext(n)[1].lower())
                if not k:
                    members.append({"name": n[-80:], "kind": None})
                    continue
                p = os.path.join(td, "m" + str(len(members)) + os.path.splitext(n)[1].lower())
                with open(p, "wb") as f:
                    f.write(z.read(n))
                if k == "txt":
                    t, _ = decode(open(p, "rb").read())
                    u = [ln for ln in t.splitlines() if ln.strip()]
                    inf = {}
                else:
                    u, inf = read(p, k, None, fast, depth + 1)
                units += u
                members.append({"name": n[-80:], "kind": k, "units": len(u),
                                **({"sheets": inf["sheets"]} if "sheets" in inf else {})})
        return units, {"members": members}
    return [], {"unsupported": kind}


def main() -> int:
    path, kind, out = sys.argv[1], sys.argv[2], sys.argv[3]
    textutil = sys.argv[sys.argv.index("--textutil") + 1] if "--textutil" in sys.argv else None
    fast = "--pdf-fast" in sys.argv
    t = time.time()
    try:
        units, info = read(path, kind, textutil, fast)
    except Exception as e:  # noqa: BLE001
        print(json.dumps({"error": f"ref:{type(e).__name__}: {str(e)[:200]}"}))
        return 0
    json.dump({"units": units}, open(out, "w", encoding="utf-8"), ensure_ascii=False, default=str)
    info["n_units"] = len(units)
    info["seconds"] = round(time.time() - t, 2)
    if "members" in info:  # zip：中のシートを外へ（名前に中のファイルを付ける）
        sh = []
        for m in info["members"]:
            for s in m.pop("sheets", []) or []:
                sh.append(s)
        if sh:
            info["sheets"] = sh
    print(json.dumps(info, ensure_ascii=False, default=str))
    return 0


if __name__ == "__main__":
    sys.exit(main())
