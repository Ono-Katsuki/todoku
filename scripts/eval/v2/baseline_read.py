"""比べる道具で読む。使える人が書くように書く（わざと弱くしない）：
- openpyxl：data_only で保存された値・数式の字も並べる・日付は日付に・結合は全部のマスに広げる・非表示の印・取り消し線・色・メモ・字ごとの書式
- xlrd：formatting_info で結合・非表示・字の書式・メモ・日付
- python-docx：本文の順に段落と表・テキストボックス・コンテンツコントロール・チェックボックス・ヘッダーとフッター・字の書式・ルビ（XML をたどる）
- pdfplumber：字と表とフォームの欄（注釈）・画像の数
- pypdf：字とフォームの欄（get_fields）・画像の数
- MarkItDown：既定どおり（既定のほかに選ぶものがない）
使い方（別の venv で、python -I）: baseline_read.py <道具> <ファイル> <出力>   1 行目の JSON に結果"""

import datetime as dt
import json
import sys
import time

CHECK_ON, CHECK_OFF = "☑", "☐"


def fmt_val(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d") if (v.hour, v.minute, v.second) == (0, 0, 0) else v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, (dt.date, dt.time)):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def read_openpyxl(p):
    import openpyxl

    wbv = openpyxl.load_workbook(p, data_only=True, rich_text=True)
    wbf = openpyxl.load_workbook(p)
    lines = []
    for ws in wbv.worksheets:
        wf = wbf[ws.title]
        state = "" if ws.sheet_state == "visible" else f" [非表示のシート:{ws.sheet_state}]"
        lines.append(f"# {ws.title}{state}")
        fill = {}
        for rng in ws.merged_cells.ranges:
            tl = ws.cell(rng.min_row, rng.min_col).value
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    fill[(r, c)] = tl
        hidden_cols = set()
        for key, d in ws.column_dimensions.items():
            if d.hidden:
                for c in range(d.min or openpyxl.utils.column_index_from_string(key), (d.max or openpyxl.utils.column_index_from_string(key)) + 1):
                    hidden_cols.add(c)
        for row in ws.iter_rows():
            cells = []
            r = row[0].row if row else 0
            hid_row = bool(ws.row_dimensions[r].hidden) if r in ws.row_dimensions else False
            for c in row:
                v = c.value
                if v is None and (c.row, c.column) in fill:
                    v = fill[(c.row, c.column)]
                f = wf.cell(c.row, c.column).value
                tags = []
                if isinstance(v, openpyxl.cell.rich_text.CellRichText):
                    for part in v:
                        if hasattr(part, "font") and part.font is not None:
                            if part.font.strike:
                                tags.append("取消線")
                            col = part.font.color
                            if col is not None and isinstance(col.rgb, str) and col.rgb[-6:].upper() not in ("000000", "FFFFFF"):
                                tags.append(f"色:{col.rgb[-6:]}")
                    v = str(v)
                txt = fmt_val(v)
                if isinstance(f, str) and f.startswith("="):
                    txt += f" [数式 {f}]"
                if not txt:
                    continue
                if c.font is not None:
                    if c.font.strike:
                        tags.append("取消線")
                    col = c.font.color
                    if col is not None and isinstance(col.rgb, str) and col.rgb[-6:].upper() not in ("000000", "FFFFFF"):
                        tags.append(f"色:{col.rgb[-6:]}")
                if c.column in hidden_cols:
                    tags.append("非表示の列")
                if c.comment is not None:
                    txt += f" [メモ {c.comment.text}]"
                cells.append(("".join(f"[{t}]" for t in dict.fromkeys(tags))) + txt.replace("\n", " "))
            if cells:
                lines.append((f"[非表示の行 {r}]\t" if hid_row else "") + "\t".join(cells))
        if ws._images:
            lines.append(f"[画像 {len(ws._images)}]")
    return "\n".join(lines)


def read_xlrd(p):
    import xlrd

    bk = xlrd.open_workbook(p, formatting_info=True)
    lines = []
    for sh in bk.sheets():
        lines.append(f"# {sh.name}" + (f" [非表示のシート:{sh.visibility}]" if sh.visibility else ""))
        fill = {}
        for r0, r1, c0, c1 in sh.merged_cells:
            tl = sh.cell(r0, c0)
            for r in range(r0, r1):
                for c in range(c0, c1):
                    fill[(r, c)] = tl
        for r in range(sh.nrows):
            cells = []
            info = sh.rowinfo_map.get(r)
            hid_row = bool(info and info.hidden)
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK) and (r, c) in fill:
                    cell = fill[(r, c)]
                if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    continue
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        txt = fmt_val(xlrd.xldate.xldate_as_datetime(cell.value, bk.datemode))
                    except Exception:  # noqa: BLE001
                        txt = str(cell.value)
                elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    txt = "TRUE" if cell.value else "FALSE"
                else:
                    txt = fmt_val(cell.value)
                tags = []
                xf = bk.xf_list[sh.cell_xf_index(r, c)]
                font = bk.font_list[xf.font_index]
                if font.struck_out:
                    tags.append("取消線")
                rgb = bk.colour_map.get(font.colour_index)
                if rgb and tuple(rgb) not in ((0, 0, 0), (255, 255, 255)):
                    tags.append("色:%02X%02X%02X" % tuple(rgb))
                ci = sh.colinfo_map.get(c)
                if ci and ci.hidden:
                    tags.append("非表示の列")
                note = sh.cell_note_map.get((r, c))
                if note is not None:
                    txt += f" [メモ {note.text}]"
                cells.append("".join(f"[{t}]" for t in tags) + txt.replace("\n", " "))
            if cells:
                lines.append((f"[非表示の行 {r + 1}]\t" if hid_row else "") + "\t".join(cells))
    return "\n".join(lines)


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"


def _docx_para(p):
    out = []
    for r in p.iter():
        if r.tag != W + "r":
            continue
        txt = "".join((e.text or "") if e.tag == W + "t" else ("\t" if e.tag == W + "tab" else ("\n" if e.tag == W + "br" else "")) for e in r)
        rpr = r.find(W + "rPr")
        tags = []
        if rpr is not None:
            if rpr.find(W + "strike") is not None or rpr.find(W + "dstrike") is not None:
                tags.append("取消線")
            col = rpr.find(W + "color")
            if col is not None and col.get(W + "val") not in (None, "auto", "000000", "FFFFFF"):
                tags.append("色:" + col.get(W + "val"))
            if rpr.find(W + "vanish") is not None:
                tags.append("非表示")
        if txt:
            out.append("".join(f"[{t}]" for t in tags) + txt)
    return "".join(out)


def _sdt_checks(root):
    lines = []
    for ck in root.iter(W14 + "checkbox"):
        st = ck.find(W14 + "checked")
        on = st is not None and st.get(W14 + "val") in ("1", "true")
        lines.append(f"[チェックボックス {CHECK_ON if on else CHECK_OFF}]")
    for ck in root.iter(W + "checkBox"):
        st = ck.find(W + "checked")
        if st is None:
            st = ck.find(W + "default")
        on = st is not None and st.get(W + "val") not in ("0", "false")
        lines.append(f"[チェックボックス {CHECK_ON if on else CHECK_OFF}]")
    return lines


def read_docx(p):
    import docx

    d = docx.Document(p)
    lines = []
    for parent in list(d.element.body.iter()):  # 同じ図形の古い形（Fallback）は読まない
        for ch in list(parent):
            if ch.tag == "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback":
                parent.remove(ch)

    def walk(el, depth=0):
        for ch in el:
            if ch.tag == W + "p":
                t = _docx_para(ch)
                marks = _sdt_checks(ch)
                if t or marks:
                    lines.append(" ".join(marks + [t]) if marks else t)
            elif ch.tag == W + "tbl":
                for tr in ch.iter(W + "tr"):
                    row = []
                    for tc in tr.findall(W + "tc"):
                        sub = [_docx_para(pp) for pp in tc.iter(W + "p")]
                        marks = _sdt_checks(tc)
                        row.append(" ".join(marks + [" ".join(s for s in sub if s)]))
                    lines.append("\t".join(row))
            elif ch.tag == W + "sdt":
                c = ch.find(W + "sdtContent")
                if c is not None:
                    walk(c, depth + 1)

    walk(d.element.body)
    # 本文の外：テキストボックス（p の中の r の中にあるので、上の段落にも一部入る）・ヘッダー・フッター
    for tb in d.element.body.iter(W + "txbxContent"):
        for pp in tb.iter(W + "p"):
            t = _docx_para(pp)
            if t:
                lines.append("[テキストボックス] " + t)
    for sec in d.sections:
        for part, name in ((sec.header, "ヘッダー"), (sec.footer, "フッター")):
            for para in part.paragraphs:
                if para.text.strip():
                    lines.append(f"[{name}] {para.text}")
    return "\n".join(lines)


def _pdf_fields_pypdf(rd):
    out = []
    for name, f in (rd.get_fields() or {}).items():
        v = f.get("/V")
        if f.get("/FT") == "/Btn":
            on = v is not None and str(v) not in ("/Off", "")
            out.append(f"[フォームの欄] {name} = {CHECK_ON if on else CHECK_OFF}")
        else:
            out.append(f"[フォームの欄] {name} = {'' if v is None else v}")
    return out


def read_pypdf(p):
    import pypdf

    rd = pypdf.PdfReader(p)
    lines = []
    for i, pg in enumerate(rd.pages, 1):
        lines.append(f"# page {i}")
        lines.append(pg.extract_text() or "")
        try:
            n = len(pg.images)
        except Exception:  # noqa: BLE001
            n = 0
        if n:
            lines.append(f"[画像 {n}]")
    lines += _pdf_fields_pypdf(rd)
    return "\n".join(lines)


def read_pdfplumber(p):
    import pdfplumber

    lines = []
    with pdfplumber.open(p) as pdf:
        for i, pg in enumerate(pdf.pages, 1):
            lines.append(f"# page {i}")
            lines.append(pg.extract_text() or "")
            try:
                for tb in pg.extract_tables():
                    for row in tb:
                        lines.append("\t".join((c or "").replace("\n", " ") for c in row))
            except Exception:  # noqa: BLE001
                pass
            if pg.images:
                lines.append(f"[画像 {len(pg.images)}]")
            for an in pg.annots or []:
                data = an.get("data") or {}
                if str(data.get("FT")) in ("/'Tx'", "/'Btn'", "/'Ch'"):
                    v = data.get("V")
                    if isinstance(v, bytes):
                        try:
                            v = v.decode("utf-16") if v[:2] in (b"\xfe\xff", b"\xff\xfe") else v.decode("latin-1")
                        except UnicodeDecodeError:
                            v = v.decode("latin-1")
                    t = data.get("T")
                    t = t.decode("latin-1", "replace") if isinstance(t, bytes) else t
                    if str(data.get("FT")) == "/'Btn'":
                        s = str(data.get("AS") or data.get("V") or "")
                        lines.append(f"[フォームの欄] {t} = {CHECK_OFF if ('Off' in s or not s) else CHECK_ON}")
                    else:
                        lines.append(f"[フォームの欄] {t} = {'' if v is None else v}")
    return "\n".join(lines)


def read_csvish(p):
    from charset_normalizer import from_path

    best = from_path(p).best()
    if best is None:
        raise ValueError("文字コードを決められない")
    return str(best)


def main():
    tool, p, out = sys.argv[1:4]
    t = time.time()
    o = {"tool": tool}
    try:
        if tool == "markitdown":
            from markitdown import MarkItDown

            txt = MarkItDown().convert(p).text_content
        else:
            txt = {"openpyxl": read_openpyxl, "xlrd": read_xlrd, "python-docx": read_docx, "pypdf": read_pypdf,
                   "pdfplumber": read_pdfplumber, "charset": read_csvish}[tool](p)
        with open(out, "w", encoding="utf-8") as f:
            f.write(txt)
        o["chars"] = len(txt)
    except Exception as e:  # noqa: BLE001
        o["error"] = f"{type(e).__name__}: {str(e)[:200]}"
    o["seconds"] = round(time.time() - t, 2)
    print(json.dumps(o, ensure_ascii=False))


if __name__ == "__main__":
    main()
