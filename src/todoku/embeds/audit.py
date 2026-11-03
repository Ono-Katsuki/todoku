"""埋め込みの読み落とし：窓とは別の読み方で、グラフ・SmartArt・埋め込みの字を全部拾い、窓に出ているかを確かめる。

別の読み方：
- グラフ（.xlsx・.docx）：関係（.rels）で指されているグラフの部品を、名前から直接探し、XML の字（c:v・a:t・cx:v・cx:pt）を
  正規表現で拾う。数値の控え（c:numCache・c:numLit）は、控えの表示形式で見えている字にしてから比べる
- グラフ（.xls）：ブックのストリームを頭から読み、グラフの部分（BOF 0x0020）を数え、SeriesText の字と、値の控え（SIIndex の後ろの
  Label・Number）を拾う
- SmartArt：データの部品の a:t をすべて
- 埋め込み：入れ物を数え直し（.docx は本文の o:OLEObject、.doc は ObjectPool と Data の FOBJH、.xlsx はシートの関係の
  oleObject・package、.xls は MBD の入れ物）、中身の種類ごとに、Word は文書の字の要素（各形式の検査と同じ別の読み方）、
  Excel はセルの見えている字、PowerPoint は字の atom・a:t、PDF は文字の命令、数式は字、そのほかは種類の名前が窓に出ているか
"""

from __future__ import annotations

import re
import struct
from collections import Counter

from ..xlsx.audit import Check
from .chart import number_text


def _squash(s: str) -> str:
    return re.sub(r"[\s　]+", "", s).replace("\\|", "|")


def _unxml(s: str) -> str:
    s = re.sub(r"&#x([0-9A-Fa-f]+);", lambda m: chr(int(m[1], 16)), s)
    s = re.sub(r"&#(\d+);", lambda m: chr(int(m[1])), s)
    return s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&")


# ====================================================================== グラフの部品の字（XML から直接）
def chart_xml_texts(xml: bytes) -> list[str]:
    """グラフの部品の字：文字の控え・書式つきの字はそのまま、数値の控えは表示形式で見えている字にしたもの。"""
    s = xml.decode("utf-8", "replace")
    out = []
    num_spans = []
    for m in re.finditer(r"<(?:c:)?(numCache|numLit)\b[^>]*>(.*?)</(?:c:)?\1>", s, re.S):
        num_spans.append(m.span())
        fm = re.search(r"<(?:c:)?formatCode>(.*?)</(?:c:)?formatCode>", m[2], re.S)
        code = _unxml(fm[1]) if fm else "General"
        for pm in re.finditer(r"<(?:c:)?pt\b([^>]*)>\s*<(?:c:)?v>(.*?)</(?:c:)?v>", m[2], re.S):
            pc = re.search(r'formatCode="([^"]*)"', pm[1])
            out.append(number_text(_unxml(pm[2]), _unxml(pc[1]) if pc else code))
    for m in re.finditer(r"<(?:c:|cx:)?v>(.*?)</(?:c:|cx:)?v>|<a:t>(.*?)</a:t>|<cx:pt\b[^>]*>(.*?)</cx:pt>", s, re.S):
        if any(a <= m.start() < b for a, b in num_spans):
            continue
        t = _unxml(next(x for x in m.groups() if x is not None))
        out.append(t)
    return [t for t in out if t.strip()]


def _rel_targets(z) -> set[str]:
    """パッケージのすべての関係が指している部品の名前。"""
    import posixpath

    out = set()
    for n in z.names():
        if not n.endswith(".rels"):
            continue
        d = posixpath.dirname(posixpath.dirname(n))
        for m in re.finditer(rb'Target="([^"]+)"', z.read(n)):
            t = m[1].decode("utf-8", "replace")
            out.add(t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(d, t)))
    return out


def _missing(texts: list[str], shown: str) -> list[str]:
    """窓（shown）に出ていない字。改行で分けた行ごとに、空白を除いて比べる（窓では改行を「／」などで表すため）。"""
    flat = _squash(shown)
    return [t for t in texts if any(_squash(x) and _squash(x) not in flat for x in t.split("\n"))]


# ====================================================================== Excel
def excel_checks(wb, doc: dict) -> list[Check]:
    from .excel import sheet_markdown

    shown_charts = "\n".join("\n".join(sheet_markdown({"charts": sh.get("charts")})) for sh in doc["sheets"])
    n_win = sum(len(sh.get("charts") or []) for sh in doc["sheets"])
    out = []
    if getattr(wb, "kind", "xlsx") == "xls":
        texts, n_raw = _xls_chart_texts(wb)
    else:
        z = wb.zip
        refs = _rel_targets(z)
        parts = [n for n in z.names() if re.fullmatch(r"xl/charts/chart(Ex)?\d+\.xml", n) and n in refs]
        n_raw = len(parts)
        texts = [t for p in parts for t in chart_xml_texts(z.read(p))]
        dg = [n for n in z.names() if re.fullmatch(r"xl/diagrams/data\d+\.xml", n) and n in refs]
        if dg:
            out.append(_smartart_check(z, dg, "\n".join(_obj_md(sh) for sh in doc["sheets"])))
    if n_raw or n_win:
        miss = _missing(texts, shown_charts)
        ok = not miss and n_raw == n_win
        out.append(Check("読み落とし（グラフ）", ok,
                         f"グラフ {n_raw} 個の字と値 {len(texts):,} 個を、窓とは別の読み方で拾い、窓のグラフの表 {n_win} 個に"
                         + ("すべて出ていることを確かめました" if ok else "出ていないものがあります"),
                         ([f"グラフの数が違います（ファイル {n_raw}・窓 {n_win}）"] if n_raw != n_win else []) + miss[:30]))
    objs = [(sh["name"], o) for sh in doc["sheets"] for o in sh.get("objects") or [] if o.get("kind") != "smartart"]
    n_emb = _count_excel_embeds(wb)
    if objs or n_emb:
        out.append(_objects_check([(o, _excel_source(wb, o)) for _s, o in objs], n_emb))
    return out


def _obj_md(sh: dict) -> str:
    from .excel import sheet_markdown

    return "\n".join(sheet_markdown({"objects": sh.get("objects")}))


def _xls_chart_texts(wb) -> tuple[list[str], int]:
    """ブックのストリームを頭から：グラフの部分の数と、SeriesText・値の控えの字。"""
    from ..xls.book import xlstring

    data = wb.data
    pos, depth, n = 0, 0, 0
    in_chart: list[bool] = []
    si = 0
    texts = []
    while pos + 4 <= len(data):
        t, ln = struct.unpack_from("<HH", data, pos)
        b = data[pos + 4:pos + 4 + ln]
        pos += 4 + ln
        if t == 0x0809:
            dt = struct.unpack_from("<H", b, 2)[0] if len(b) >= 4 else 0
            in_chart.append(dt == 0x0020)
            n += dt == 0x0020
            si = 0
            continue
        if t == 0x000A:
            if in_chart:
                in_chart.pop()
            continue
        if not in_chart or not in_chart[-1]:
            continue
        if t == 0x100D and len(b) >= 4:
            cch, fl = b[2], b[3]
            raw = b[4:4 + cch * (2 if fl & 1 else 1)]
            texts.append(raw.decode("utf-16-le", "replace") if fl & 1 else raw.decode("latin-1"))
        elif t == 0x1065:
            si = struct.unpack_from("<H", b, 0)[0]
        elif si and t == 0x0204:
            texts.append(xlstring(b, 6)[0])
        elif si and t == 0x0203:
            _r, _c, x = struct.unpack_from("<HHH", b, 0)
            texts.append(number_text(repr(struct.unpack_from("<d", b, 6)[0]), wb.style(x).num_fmt))
    return [t for t in texts if t.strip()], n


def _count_excel_embeds(wb) -> int:
    """埋め込みの入れ物の数（ActiveX・リンクは除く）。"""
    if getattr(wb, "kind", "xlsx") == "xls":
        from .excel import _xls_objects

        names = {p.split("/")[0] for p in wb.cfb.streams() if re.match(r"MBD[0-9A-Fa-f]{8}/", p)}
        enc = next((p for p in wb.cfb.streams() if p.lower() == "encryption"), None)
        if enc is not None:
            from ..bincrypto import summary_entries

            names |= {n for n, is_stream, _d in summary_entries(wb.cfb.read(enc)) if not is_stream and n.startswith("MBD")}
        ctl = set()
        for sh in wb.sheets:
            ctl |= _xls_control_storages(wb, sh)
        del _xls_objects
        return len(names - ctl)
    from .chart import rel_types

    n = 0
    for sh in wb.sheets:
        if sh.path and wb.zip.has(sh.path):
            xml = wb.zip.read(sh.path)
            rels = rel_types(wb.zip, sh.path)
            ids = set(re.findall(rb'<(?:\w+:)?oleObject\b[^>]*\br:id="([^"]+)"', xml))
            n += sum(1 for i in ids if rels.get(i.decode(), ("", ""))[0] in ("oleObject", "package"))
    return n


def _xls_control_storages(wb, sh) -> set[str]:
    """ActiveX・カメラ・リンクの部品が使う MBD の入れ物（埋め込みとしては数えない）。"""
    from .biffchart import BOF, CONTINUE, EOF, substream
    from .excel import pict_info

    try:
        lb = int(str(sh.path).lstrip("#"))
    except ValueError:
        return set()
    out = set()
    recs = substream(wb.data, lb)
    depth = 0
    for i, (_p, t, b) in enumerate(recs):
        depth += (t == BOF) - (t == EOF)
        if t != 0x005D or depth != 1:
            continue
        j = i + 1
        while j < len(recs) and recs[j][1] == CONTINUE:
            b += recs[j][2]
            j += 1
        info = pict_info(b)
        if info and "pos" in info and info.get("flags", 0) & (0x02 | 0x10 | 0x20 | 0x80):
            out.add(f"MBD{info['pos']:08X}")
    return out


def _excel_source(wb, o: dict):
    if getattr(wb, "kind", "xlsx") == "xls":
        from .excel import storage_of

        return storage_of(wb, o["storage"])
    return wb.zip.read(o["storage"]) if wb.zip.has(o["storage"]) else None


# ====================================================================== Word
def word_checks(doc, win: dict) -> list[Check]:
    from .word import _source, item_markdown

    out = []
    objs = win.get("objects") or []
    charts = [o for o in objs if o["kind"] == "chart"]
    if getattr(doc, "kind", "") == "docx":
        z = doc.zip
        parts = []
        for _k, part, _b in doc.all_blocks():
            xml = z.read(part) if z.has(part) else b""
            rels = doc._rels(part)
            for m in re.finditer(rb'<(?:c|cx):chart\b[^>]*\br:id="([^"]+)"', re.sub(rb"<mc:Fallback\b.*?</mc:Fallback>", b"", xml, flags=re.S)):
                from ..docx.document import _join

                t = _join(part, rels.get(m[1].decode(), ("", ""))[1])
                if z.has(t):
                    parts.append(t)
        if parts or charts:
            texts = [t for p in parts for t in chart_xml_texts(z.read(p))]
            shown = "\n".join(x for o in charts for x in item_markdown(o))
            miss = _missing(texts, shown)
            ok = not miss and len(parts) == len(charts)
            out.append(Check("読み落とし（グラフ）", ok,
                             f"文書のグラフ {len(parts)} 個の字と値 {len(texts):,} 個を、窓とは別の読み方で拾い、窓のグラフの表 {len(charts)} 個に"
                             + ("すべて出ていることを確かめました" if ok else "出ていないものがあります"),
                             ([f"グラフの数が違います（ファイル {len(parts)}・窓 {len(charts)}）"] if len(parts) != len(charts) else [])
                             + miss[:30]))
        dg = sorted({o["part"] for o in objs if o["kind"] == "smartart"} |
                    {n for n in z.names() if re.fullmatch(r"word/diagrams/data\d+\.xml", n) and n in _rel_targets(z)})
        if dg:
            out.append(_smartart_check(z, dg, "\n".join(x for o in objs if o["kind"] == "smartart" for x in item_markdown(o))))
        n_emb = 0  # 本文などの部品の o:OLEObject（リンクではないもの）を、XML から直接数える
        for _k, p, _b in doc.all_blocks():
            xml = re.sub(rb"<mc:Fallback\b.*?</mc:Fallback>", b"", z.read(p) if z.has(p) else b"", flags=re.S)
            ids = {m[1] for m in re.finditer(rb'<o:OLEObject\b([^>]*)>', xml) if b'Type="Link"' not in m[1]
                   for m in [re.search(rb'r:id="([^"]+)"', m[1])] if m}
            n_emb += len(ids)
    else:  # ObjectPool の入れ物と、Data の FOBJH の入れ物を、複合ファイルから直接数える
        streams = list(doc.cfb.streams())
        pool = {p.split("/")[1] for p in streams if p.startswith("ObjectPool/") and p.count("/") >= 2}
        # ActiveX の部品（CONTROL のフィールド。入れ物に \x03OCXNAME がある）は、埋め込みとしては数えない
        pool -= {p.split("/")[1] for p in streams if p.startswith("ObjectPool/") and p.endswith("/\x03OCXNAME")}
        # リンクした部品（\x03LinkInfo がある。中身は文書の外のファイル）も数えない
        pool -= {p.split("/")[1] for p in streams if p.startswith("ObjectPool/") and p.endswith("/\x03LinkInfo")}
        n_emb = len(pool) + sum(1 for e in doc.embeds if e["storage"].startswith("Data@"))
    emb = [o for o in objs if o["kind"] == "embed"]
    books = {e["n"] for e in win.get("embeds") or []}  # Excel の表として出したもの（埋め込みの表の検査が確かめる）
    if emb or n_emb - len(books) > 0:
        pairs = []
        for o in emb:
            e = next((x for x in doc.embeds if x["n"] == o["n"]), None)
            try:
                src = _source(doc, e) if e is not None else None
            except Exception:  # noqa: BLE001
                src = None
            pairs.append((o, src))
        out.append(_objects_check(pairs, n_emb - len(books)))
    return out


def _smartart_check(z, parts: list[str], shown: str) -> Check:
    from .smartart import smartart_texts

    texts = [t for p in parts for t in smartart_texts(z.read(p)) if t.strip()]
    miss = _missing(texts, shown)
    return Check("読み落とし（SmartArt）", not miss,
                 f"SmartArt {len(parts)} 個の字 {len(texts):,} 個を、データの部品から直接拾い、"
                 + ("すべて窓に出ていることを確かめました" if not miss else "窓に出ていないものがあります"), miss[:30])


# ====================================================================== 埋め込みの中身
def _objects_check(pairs: list[tuple[dict, object]], n_raw: int) -> Check:
    from .ole import object_markdown

    problems = []
    kinds = Counter()
    for o, src in pairs:
        kind = o.get("kind_of", o.get("kind"))
        kinds[kind] += 1
        md = "\n".join(object_markdown(o, "###"))
        try:
            problems += [f"埋め込み{o['n']}：{x}" for x in content_problems(kind, src, md, o)]
        except Exception as e:  # noqa: BLE001
            problems.append(f"埋め込み{o['n']}：検査のために開けませんでした（{e}）")
        if o.get("error"):
            problems.append(f"埋め込み{o['n']}：{o['error']}")
    if n_raw != len(pairs):
        problems.insert(0, f"埋め込みの数が違います（ファイル {n_raw}・窓 {len(pairs)}）")
    from .ole import KIND_LABEL

    what = "・".join(f"{KIND_LABEL.get(k, k)} {n}" for k, n in kinds.items())
    unreadable = kinds.get("other", 0)
    return Check("読み落とし（埋め込み）", not problems,
                 f"埋め込み {len(pairs)} 個（{what}）の中身を、窓とは別の読み方で拾い直し、"
                 + ("すべて窓に出ていることを確かめました" if not problems else "窓に出ていないものがあります")
                 + (f"。中身の形を読めない部品 {unreadable} 個は、種類の名前と代替テキストを出しました" if unreadable else ""),
                 problems[:30])


def _count(texts) -> Counter:
    import unicodedata

    c = Counter()
    for t in texts:
        for ch in unicodedata.normalize("NFD", t):  # 窓で1字にまとめた母音（ে + া → ো）も、元と同じに数える
            if not ch.isspace() and ch not in "□☐☑☒■":
                c[ch] += 1
    return c


def content_problems(kind: str, src, md: str, o: dict) -> list[str]:
    """埋め込み1つの中身の字が、窓（md）に出ているか。出ていないものの説明の並び。"""
    from ..cfb import CFB

    if src is None:
        return ["入れ物が見つかりません"]
    st = CFB(src) if isinstance(src, bytes) and src[:4] == b"\xd0\xcf\x11\xe0" else src
    if kind == "package":
        return [] if (o.get("file") or "") in md else ["ファイル名が窓に出ていません"]
    if kind == "other":
        return [] if _squash(o.get("class") or "") in _squash(md) else ["種類の名前が窓に出ていません"]
    if kind == "equation":
        return [] if o.get("equation") else ["数式の字を取り出せませんでした"]
    inner = _inner(st)
    if inner is None:
        return ["中身を開き直せません"]
    k, obj = inner
    if k == "word":
        if getattr(obj, "kind", "") == "docx":
            from ..docx.audit import raw_texts

            raw = _count(raw_texts(obj))
        else:
            from ..doc.audit import raw_chars

            raw = _count(raw_chars(obj))
        miss = raw - _count([md])
        return [f"「{ch}」が {n} 字、窓に出ていません" for ch, n in miss.items()]
    if k == "excel":
        from ..xlsx.window import display

        out = []
        for sh in obj.sheets:
            for (r, c), cell in sh.cells.items():
                if cell.value in (None, ""):
                    continue
                try:
                    t = display(obj, cell)
                except Exception:  # noqa: BLE001
                    t = str(cell.value)
                if _squash(str(t)) and _squash(str(t)) not in _squash(md):
                    out.append(f"{sh.name} のセル（{r}行・{c}列）「{str(cell.value)[:20]}」")
        return out
    if k == "pptx":
        from .ppt import pptx_texts

        return [f"「{t[:30]}」" for t in _missing(pptx_texts(obj), md)]
    if k == "ppt":
        from .ppt import text_atoms

        return [f"「{t[:30]}」" for t in _missing(text_atoms(obj), md)]
    if k == "pdf":
        from ..pdf.audit import raw_chars
        from ..pdf.window import build

        win = build(obj)
        out = []
        for pg, page in zip(win["pages"], obj.pages()):
            raw = _count(raw_chars(obj, page)) - _count(t for _k, t in pg["_dropped"])
            for ch, n in (raw - _count([md])).items():
                out.append(f"ページ {pg['number']}：「{ch}」が {n} 字")
        return out
    if k == "visio":
        texts = []
        import xml.etree.ElementTree as ET

        for n in obj.names():
            if re.fullmatch(r"visio/pages/page\d+\.xml", n):
                texts += ["".join(t.itertext()) for t in ET.fromstring(obj.read(n)).iter() if t.tag.endswith("}Text")]
        return [f"「{t[:30]}」" for t in _missing(texts, md)]
    return []


def _inner(st):
    """入れ物・部品のバイト列を、種類ごとの読み手で開き直す：(種類, 読んだもの)。"""
    from ..zipraw import RawZip

    if isinstance(st, (bytes, bytearray)):
        b = bytes(st)
        if b[:4] == b"%PDF":
            from ..pdf.doc import PdfDocument

            return "pdf", PdfDocument(b, "")
        if b[:2] == b"PK":
            z = RawZip(b)
            low = [n.lower() for n in z.names()]
            if any(n.startswith("word/") for n in low):
                from ..docx.document import DocxDocument

                return "word", DocxDocument(b, "")
            if any(n.startswith("xl/") for n in low):
                from ..xlsx.workbook import Workbook

                return "excel", Workbook(b, "")
            if any(n.startswith("ppt/") for n in low):
                return "pptx", z
            if any(n.startswith("visio/") for n in low):
                return "visio", z
        return None
    names = {n.lower(): n for n in st.streams()}
    outer = getattr(st, "blob", b"")
    if "workbook" in names or "book" in names:
        from ..xls.book import XlsWorkbook

        return "excel", XlsWorkbook(outer, "", cfb=st)
    if "worddocument" in names:
        from ..doc.book import DocFile

        return "word", DocFile(outer, "", cfb=st)
    if "powerpoint document" in names:
        return "ppt", st
    if "package" in names:
        return _inner(st.read(names["package"]))
    if "contents" in names:
        return _inner(st.read(names["contents"]))
    if "\x01ole10native" in names:
        from .ole import ole10native

        _name, body = ole10native(st.read(names["\x01ole10native"]))
        if body is not None and body[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            from ..cfb import CFB

            return _inner(CFB(body))
        return _inner(body) if body is not None else None
    return None
