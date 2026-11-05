"""書類の作りを、とどくを使わずに調べる（比べる特徴の「ある・ない」の基準）。ファイルを XML と PDF の道具で直接読む。
使い方（別の venv で、python -I）: probes.py <ファイル> <kind: xlsx|docx|pdf> <出力 JSON>
.xls・.doc はこの前に LibreOffice で .xlsx・.docx に変えたものを渡す（truth.py）。"""

import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
XDR = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
MAX_CELLS = 300_000

DATE_IDS = set(range(14, 23)) | set(range(45, 48))
ERA_IDS = set(range(27, 37)) | set(range(50, 59))  # 日本語の版の既定：和暦を含む日付の形
BUILTIN_NUM = {1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%", 11: "0.00E+00", 37: "#,##0;(#,##0)", 38: "#,##0;[Red](#,##0)",
               39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 48: "##0.0E+0"}


def fmt_kind(code: str | None, fid: int) -> str | None:
    if fid in DATE_IDS:
        return "date"
    if fid in ERA_IDS:
        return "era"
    if not code or code.lower() in ("general", "@", "g/標準"):
        return None
    c = re.sub(r'"[^"]*"|\\.', "", code)
    bracket = re.findall(r"\[[^\]]*\]", c)
    c2 = re.sub(r"\[[^\]]*\]", "", c).lower()
    if re.search(r"[ge]", c2) and re.search(r"[ymd]|gg|ee", c2) or any(b.startswith("[$-411]") or "gannen" in b for b in bracket) and re.search(r"[ymdge]", c2):
        return "era" if re.search(r"[ge]", c2) else "date"
    if re.search(r"[ymdhs]", c2):
        return "date"
    if re.search(r"[0#?]", c2):
        return "number"
    return None


def rels_of(z, part: str) -> dict:
    d, f = part.rsplit("/", 1) if "/" in part else ("", part)
    rp = f"{d}/_rels/{f}.rels" if d else f"_rels/{f}.rels"
    if rp not in z.namelist():
        return {}
    out = {}
    for r in ET.fromstring(z.read(rp)):
        t = r.get("Target") or ""
        if r.get("TargetMode") == "External":
            continue
        if t.startswith("/"):
            t = t[1:]
        else:
            base = d.split("/") if d else []
            for seg in t.split("/"):
                if seg == "..":
                    base = base[:-1]
                elif seg != ".":
                    base.append(seg)
            t = "/".join(base)
        out[r.get("Id")] = (r.get("Type", "").rsplit("/", 1)[-1], t)
    return out


def text_no_rph(el) -> str:
    out = []
    for ch in el:
        if ch.tag == M + "t":
            out.append(ch.text or "")
        elif ch.tag == M + "r":
            t = ch.find(M + "t")
            out.append((t.text or "") if t is not None else "")
    return "".join(out)


def _color_ok(c) -> bool:
    if c is None:
        return False
    rgb = c.get("rgb")
    if rgb:
        return rgb[-6:].upper() not in ("000000", "FFFFFF")
    idx = c.get("indexed")
    if idx is not None:
        return idx not in ("8", "9", "64", "65")
    return False  # テーマの色は、既定の字の色のことが多いので数えない


def probe_xlsx(p: str) -> dict:
    z = zipfile.ZipFile(p)
    names = set(z.namelist())
    ss, ss_ph, ss_rich = [], [], []
    if "xl/sharedStrings.xml" in names:
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall(M + "si"):
            ss.append(text_no_rph(si))
            ss_ph.append("".join((t.text or "") for ph in si.findall(M + "rPh") for t in ph.findall(M + "t")))
            strike, color = False, False
            for r in si.findall(M + "r"):
                pr = r.find(M + "rPr")
                if pr is not None:
                    st = pr.find(M + "strike")
                    strike = strike or (st is not None and st.get("val", "1") not in ("0", "false"))
                    color = color or _color_ok(pr.find(M + "color"))
            ss_rich.append((strike, color))
    fonts, xfs, numfmts = [], [], {}
    if "xl/styles.xml" in names:
        st = ET.fromstring(z.read("xl/styles.xml"))
        for nf in st.iter(M + "numFmt"):
            numfmts[int(nf.get("numFmtId"))] = nf.get("formatCode")
        fs = st.find(M + "fonts")
        for f in fs if fs is not None else []:
            s = f.find(M + "strike")
            fonts.append((s is not None and s.get("val", "1") not in ("0", "false"), _color_ok(f.find(M + "color"))))
        cx = st.find(M + "cellXfs")
        for x in cx if cx is not None else []:
            fid = int(x.get("numFmtId", "0"))
            xfs.append((fid, numfmts.get(fid) or BUILTIN_NUM.get(fid), int(x.get("fontId", "0"))))
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    wrels = rels_of(z, "xl/workbook.xml")
    sheets, shapes, checks, comments = [], [], [], []
    total = 0
    for sh in wb.find(M + "sheets"):
        name, state = sh.get("name"), sh.get("state", "visible")
        _ty, part = wrels.get(sh.get(R + "id"), (None, None))
        if not part or part not in names:
            sheets.append({"name": name, "state": state, "cells": {}, "merges": [], "hidden_rows": [], "hidden_cols": [], "note": "部品なし"})
            continue
        root = ET.fromstring(z.read(part))
        hidden_rows, hidden_cols, merges = [], [], []
        cols = root.find(M + "cols")
        for c in cols if cols is not None else []:
            if c.get("hidden") in ("1", "true"):
                hidden_cols += list(range(int(c.get("min")), min(int(c.get("max")), 16384) + 1))[:2000]
        mc = root.find(M + "mergeCells")
        for m in mc if mc is not None else []:
            merges.append(m.get("ref"))
        cells = {}
        sd = root.find(M + "sheetData")
        for row in sd if sd is not None else []:
            rn = int(row.get("r", "0"))
            if row.get("hidden") in ("1", "true"):
                hidden_rows.append(rn)
            for c in row:
                if total >= MAX_CELLS:
                    break
                ty = c.get("t", "n")
                v, f, is_ = c.find(M + "v"), c.find(M + "f"), c.find(M + "is")
                raw, ph, rich = None, "", (False, False)
                if ty == "s" and v is not None and v.text is not None:
                    i = int(v.text)
                    if i < len(ss):
                        raw, ph, rich = ss[i], ss_ph[i], ss_rich[i]
                elif ty == "inlineStr" and is_ is not None:
                    raw = text_no_rph(is_)
                elif v is not None:
                    raw = v.text
                if raw is None and f is None:
                    continue
                si = int(c.get("s", "0"))
                fid, code, fo = xfs[si] if si < len(xfs) else (0, None, 0)
                font = fonts[fo] if fo < len(fonts) else (False, False)
                d = {"raw": raw, "t": ty}
                if f is not None:
                    d["f"] = f.text or "(shared)"
                k = fmt_kind(code, fid)
                if k and ty == "n" and raw is not None:
                    d["fmt"] = code or f"id{fid}"
                    d["fk"] = k
                if font[0] or rich[0]:
                    d["strike"] = True
                if font[1] or rich[1]:
                    d["color"] = True
                if ph:
                    d["ph"] = ph
                cells[c.get("r")] = d
                total += 1
        sheets.append({"name": name, "state": state, "cells": cells, "merges": merges, "hidden_rows": hidden_rows, "hidden_cols": hidden_cols})
        for _rid, (ty, tgt) in rels_of(z, part).items():
            if tgt not in names:
                continue
            if ty == "drawing":
                dr = drop_fallback(ET.fromstring(z.read(tgt)))
                for sp in dr.iter(XDR + "sp"):
                    t = "".join((x.text or "") for x in sp.iter(A + "t"))
                    if t.strip():
                        shapes.append({"sheet": name, "text": t})
            elif ty == "vmlDrawing":
                raw = z.read(tgt).decode("utf-8", "replace")
                for shape in re.findall(r"<v:shape\b.*?</v:shape>", raw, flags=re.S):
                    if re.search(r'ObjectType="Checkbox"', shape):
                        label = re.sub(r"<[^>]+>", "", " ".join(re.findall(r"<v:textbox\b.*?</v:textbox>", shape, flags=re.S)))
                        on = re.search(r"<x:Checked>\s*1", shape) is not None
                        checks.append({"sheet": name, "label": re.sub(r"\s+", " ", label).strip(), "checked": on})
            elif ty == "comments":
                for cm in ET.fromstring(z.read(tgt)).iter(M + "comment"):
                    comments.append("".join((x.text or "") for x in cm.iter(M + "t")))
    return {"kind": "xlsx", "sheets": sheets, "shapes": shapes, "checkboxes": checks, "comments": comments, "truncated_cells": total >= MAX_CELLS}


def drop_fallback(root):
    """mc:AlternateContent の Fallback（同じ図形の古い形）を除く。Choice と同じ字が 2 度数えられるのを防ぐ。"""
    for parent in list(root.iter()):
        for ch in list(parent):
            if ch.tag == MC + "Fallback":
                parent.remove(ch)
    return root


def probe_docx(p: str) -> dict:
    z = zipfile.ZipFile(p)
    names = set(z.namelist())
    paras, textbox, checks, strike, color, hidden, ruby = [], [], [], [], [], [], []
    parts = ["word/document.xml"] + sorted(n for n in names if re.fullmatch(r"word/(header|footer|footnotes|endnotes)\d*\.xml", n))
    for part in parts:
        if part not in names:
            continue
        root = drop_fallback(ET.fromstring(z.read(part)))
        in_tb = set()
        for tb in root.iter(W + "txbxContent"):
            for pp in tb.iter(W + "p"):
                in_tb.add(id(pp))
                t = "".join((x.text or "") for x in pp.iter(W + "t"))
                if t.strip():
                    textbox.append(t)
        for pp in root.iter(W + "p"):
            own = "".join((x.text or "") for x in pp.iter(W + "t"))
            if own.strip() and id(pp) not in in_tb:
                paras.append(own)
        for r in root.iter(W + "r"):
            pr = r.find(W + "rPr")
            t = "".join((x.text or "") for x in r.findall(W + "t"))
            if pr is None or not t.strip():
                continue
            if pr.find(W + "strike") is not None or pr.find(W + "dstrike") is not None:
                strike.append(t)
            c = pr.find(W + "color")
            if c is not None and c.get(W + "val") not in (None, "auto", "000000", "FFFFFF"):
                color.append(t)
            if pr.find(W + "vanish") is not None:
                hidden.append(t)
        for rb in root.iter(W + "rt"):
            t = "".join((x.text or "") for x in rb.iter(W + "t"))
            if t.strip():
                ruby.append(t)
        for ck in root.iter(W14 + "checkbox"):
            st = ck.find(W14 + "checked")
            checks.append({"label": "", "checked": st is not None and st.get(W14 + "val") in ("1", "true")})
        for ck in root.iter(W + "checkBox"):
            st = ck.find(W + "checked")
            if st is None:
                st = ck.find(W + "default")
            checks.append({"label": "", "checked": st is not None and st.get(W + "val") not in ("0", "false")})
    return {"kind": "docx", "paragraphs": paras, "textboxes": textbox, "checkboxes": checks, "strike": strike, "color": color, "hidden": hidden, "ruby": ruby}


def probe_pdf(p: str) -> dict:
    import pypdf

    rd = pypdf.PdfReader(p)
    o = {"kind": "pdf", "pages": len(rd.pages), "fields": [], "checkboxes": [], "signed": False, "encrypted": bool(rd.is_encrypted)}
    af = rd.trailer["/Root"].get("/AcroForm")
    af = af.get_object() if af is not None else None
    o["xfa"] = bool(af is not None and af.get("/XFA") is not None)
    for name, f in (rd.get_fields() or {}).items():
        ft, v = f.get("/FT"), f.get("/V")
        if ft == "/Sig":
            o["signed"] = o["signed"] or v is not None
        elif ft == "/Btn":
            o["checkboxes"].append({"label": name, "checked": v is not None and str(v) not in ("/Off", "")})
        elif v not in (None, ""):
            o["fields"].append({"name": name, "value": str(v)})
    return o


def main():
    p, kind, out = sys.argv[1:4]
    res = {"xlsx": probe_xlsx, "docx": probe_docx, "pdf": probe_pdf}[kind](p)
    with open(out, "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False)
    print(json.dumps({"ok": True}))


if __name__ == "__main__":
    main()
