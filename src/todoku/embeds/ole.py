"""埋め込んだ部品（OLE の入れ物・パッケージ）を見分けて、中身を平文にする（読むだけ）。

入れ物は、複合ファイルの中の入れ物（MS-OLEDS：\\x01CompObj に表示名とクリップボードの形式・ProgID、\\x01Ole10Native に
OLE1 の中身）か、部品のバイト列（.xlsx などの ZIP・複合ファイル・PDF）。どの種類かは、中身（ストリームの名前・先頭のバイト）で
決める（名前や拡張子だけでは決めない）：
- Workbook・Book → Excel の表（.xls の読み取り）。Package（ZIP）→ 中の形で決める（xl/ Excel・word/ Word・ppt/ PowerPoint・
  visio/ Visio）
- WordDocument → Word の文書（.doc の読み取り）
- PowerPoint Document → PowerPoint（ppt.py）
- Equation Native → 数式（mtef.py）
- CONTENTS が %PDF で始まる → PDF（Adobe の PDF の部品）。\\x01Ole10Native の中身が %PDF・ZIP・複合ファイルなら、その形で読む
  （「パッケージ」：ファイルをそのまま埋め込んだもの。ファイル名も読む）
- それ以外 → 種類の名前（CompObj の表示名・ProgID・入れ物の CLSID）と、図としての代替テキストだけ（中身は読めない、と書く）
"""

from __future__ import annotations

import re
import struct

from ..cfb import CFB, SubStorage

MAGIC_CFB = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
LABELS = {"Excel.Sheet": "Excel の表", "Excel.Chart": "Excel のグラフ", "Word.Document": "Word の文書",
          "PowerPoint.Show": "PowerPoint のスライド", "PowerPoint.Slide": "PowerPoint のスライド",
          "Equation.3": "数式", "Equation.DSMT4": "数式（MathType）", "Package": "ファイル（パッケージ）",
          "AcroExch.Document": "PDF", "Acrobat.Document": "PDF", "MSGraph.Chart": "グラフ（Microsoft Graph）",
          "Visio.Drawing": "Visio の図面", "Paint.Picture": "画像（ペイント）"}
KIND_LABEL = {"excel": "Excel の表", "word": "Word の文書", "powerpoint": "PowerPoint のスライド", "pdf": "PDF",
              "equation": "数式", "visio": "Visio の図面", "package": "ファイル（パッケージ）", "other": "中身を読めない形の部品"}
MAX_DEPTH = 3


def label_of(prog: str) -> str:
    """ProgID（「Excel.Sheet.8」など）→ 本文の〔埋め込み1：…〕に出す名前。"""
    head = ".".join(prog.split(".")[:2])
    if head in LABELS:
        return LABELS[head]
    for k, v in LABELS.items():
        if prog.startswith(k):
            return v
    return prog or "埋め込み"


# ====================================================================== 入れ物の情報（MS-OLEDS 2.3.8 CompObjStream）
def _lp_ansi(b: bytes, pos: int) -> tuple[str, int]:
    n = struct.unpack_from("<I", b, pos)[0]
    if n > 0x10000 or pos + 4 + n > len(b):
        raise ValueError("長さが合いません")
    raw = b[pos + 4:pos + 4 + n].rstrip(b"\x00")
    try:
        s = raw.decode("cp932")
    except UnicodeDecodeError:
        s = raw.decode("latin-1")
    return s, pos + 4 + n


def compobj(data: bytes) -> dict:
    """\\x01CompObj：{"user_type": 表示名, "clipboard": 形式, "prog": ProgID}（読めたところまで）。"""
    out: dict = {}
    try:
        pos = 28
        out["user_type"], pos = _lp_ansi(data, pos)
        marker = struct.unpack_from("<I", data, pos)[0]
        if marker in (0xFFFFFFFF, 0xFFFFFFFE):
            pos += 8
        elif marker == 0:
            pos += 4
        else:
            out["clipboard"], pos = _lp_ansi(data, pos)
        if pos + 4 <= len(data):
            n = struct.unpack_from("<I", data, pos)[0]
            if 0 < n <= 0x28:  # Reserved1（実際には ProgID が入る。40 字を超えるものは仕様どおり無視する）
                out["prog"], pos = _lp_ansi(data, pos)
                if pos + 4 <= len(data) and struct.unpack_from("<I", data, pos)[0] == 0x71B239F4:
                    pos += 4
                    cch = struct.unpack_from("<I", data, pos)[0]
                    if 0 < cch < 0x1000:
                        uni = data[pos + 4:pos + 4 + 2 * cch].decode("utf-16-le", errors="replace").rstrip("\x00")
                        if uni:
                            out["user_type"] = uni
    except (struct.error, ValueError):
        pass
    return out


def clsid_of(cfb, path: str = "") -> str:
    """入れ物の CLSID（ディレクトリの項目の 80 バイト目から 16 バイト）。path が空ならルート。"""
    base, prefix = cfb, ""
    while isinstance(base, SubStorage):
        prefix = base.prefix + ("/" + prefix if prefix else "")
        base = base.parent
    full = "/".join(x for x in (prefix, path) if x)
    try:
        e = base.entries[0] if not full else next(x for x in base.entries if x.path == full)
        raw = base._chain_bytes(base.dir_start)[e.index * 128 + 80:e.index * 128 + 96]
    except (StopIteration, AttributeError, IndexError, ValueError):
        return ""
    if len(raw) != 16 or not any(raw):
        return ""
    a, b, c = struct.unpack_from("<IHH", raw, 0)
    d = raw[8:].hex().upper()
    return f"{{{a:08X}-{b:04X}-{c:04X}-{d[:4]}-{d[4:]}}}"


def ole10native(data: bytes) -> tuple[str, bytes | None]:
    """\\x01Ole10Native（OLE1 の中身。NativeDataSize の後ろが中身）。パッケージ（ファイルを埋め込んだもの）なら
    (ファイル名, ファイルの中身)。中身の形は、OLE のパッケージの記録の並び（2 バイトの印・表示名・元のパス・…・中身の長さ・中身）。
    並びを読めなければ、中身の中の既知の先頭のバイト（%PDF・ZIP・複合ファイル）から後ろを中身とみなす。"""
    if len(data) < 4:
        return "", None
    n = struct.unpack_from("<I", data, 0)[0]
    native = data[4:4 + n] if 0 < n <= len(data) - 4 else data[4:]
    try:
        pos = 2
        end = native.index(b"\x00", pos)
        label = native[pos:end]
        pos = native.index(b"\x00", end + 1) + 1  # 元のパス
        pos += 4  # 2 つの 2 バイトの値
        ln = struct.unpack_from("<I", native, pos)[0]
        pos += 4 + ln  # 一時ファイルのパス
        size = struct.unpack_from("<I", native, pos)[0]
        body = native[pos + 4:pos + 4 + size]
        if len(body) == size:
            try:
                name = label.decode("cp932")
            except UnicodeDecodeError:
                name = label.decode("latin-1")
            return name, body
    except (ValueError, struct.error):
        pass
    for sig in (b"%PDF", MAGIC_CFB, b"PK\x03\x04"):
        i = native.find(sig)
        if i >= 0:
            return "", native[i:]
    return "", native


# ====================================================================== 見分けて読む
def describe(src, name: str = "", prog: str = "", alt: str = "", depth: int = 0) -> dict:
    """埋め込み1つを読む。src は部品のバイト列か、入れ物（CFB・SubStorage）。
    返す形：{"kind", "label", "class", "markdown": [Markdown の行], "error"（読めなかったときの理由）, ...}。"""
    d = {"kind": "other", "label": label_of(prog) if prog else "埋め込み", "class": prog, "markdown": []}
    try:
        _describe(src, name, prog, d, depth)
    except Exception as e:  # noqa: BLE001
        d["error"] = f"中身を読めませんでした（{type(e).__name__}: {e}）"
    if alt:
        d["alt"] = alt
    return d


def _describe(src, name: str, prog: str, d: dict, depth: int) -> None:
    if isinstance(src, (bytes, bytearray)):
        blob = bytes(src)
        if blob[:8] == MAGIC_CFB:
            _storage(CFB(blob), blob, name, prog, d, depth)
        elif blob[:4] == b"PK\x03\x04":
            _package(blob, name, d, depth)
        elif blob[:4] == b"%PDF":
            _pdf(blob, name, d)
        else:
            d["error"] = "中身の形が分かりません（ZIP・複合ファイル・PDF のどれでもありません）"
        return
    _storage(src, getattr(src, "blob", b""), name, prog, d, depth)


def _storage(st, outer: bytes, name: str, prog: str, d: dict, depth: int) -> None:
    names = {n.lower(): n for n in st.streams()}
    info = compobj(st.read(names["\x01compobj"])) if "\x01compobj" in names else {}
    prog = prog or info.get("prog", "")
    cls = clsid_of(st)
    parts = [x for x in (info.get("user_type", ""), prog, cls if not (info.get("user_type") or prog) else "") if x]
    d["class"] = "・".join(dict.fromkeys(parts))
    if cls:
        d["clsid"] = cls
    if prog:
        d["label"] = label_of(prog)
    if "workbook" in names or "book" in names:
        from ..xls.book import XlsWorkbook

        _excel(XlsWorkbook(outer, name, cfb=st), d)
    elif "worddocument" in names:
        from ..doc.book import DocFile

        _word(DocFile(outer, name, cfb=st), d, depth)
    elif "powerpoint document" in names:
        from .ppt import ppt_slides

        _slides(ppt_slides(st), d)
    elif "package" in names:
        _package(st.read(names["package"]), name, d, depth)
    elif "equation native" in names:
        from .mtef import equation_text

        text, err = equation_text(st.read(names["equation native"]))
        d["kind"], d["label"] = "equation", d["label"] if prog.startswith("Equation.DSMT") else "数式"
        d["equation"] = text
        d["markdown"] = [f"- 数式：{_md(text) if text else '（字を取り出せませんでした）'}"]
        if err:
            d["markdown"].append(f"- {err}（取り出せたところまでを出しています）")
    elif "contents" in names and st.read(names["contents"])[:4] == b"%PDF":
        _pdf(st.read(names["contents"]), name, d)
    elif "\x01ole10native" in names:
        fname, body = ole10native(st.read(names["\x01ole10native"]))
        d["kind"] = "package"
        if fname:
            d["file"] = fname
        if body is not None and depth < MAX_DEPTH and body[:4] in (b"%PDF", MAGIC_CFB[:4], b"PK\x03\x04"):
            inner = describe(body, fname or name, "", "", depth + 1)
            d["markdown"] = ([f"- ファイル名：{_md(fname)}"] if fname else []) + inner["markdown"]
            d["inner"] = inner["kind"]
            d["label"] = f"{KIND_LABEL.get(inner['kind'], '埋め込み')}（パッケージ）"
            if inner.get("error"):
                d["error"] = inner["error"]
        else:
            d["markdown"] = ([f"- ファイル名：{_md(fname)}"] if fname else []) + [
                "- 埋め込んだファイルの中身は、この形（" + _kind_of_bytes(body) + "）では読めません"]
    else:
        d["kind"] = "other"
        d["markdown"] = [f"- 種類：{_md(d['class']) or '（種類の名前がありません）'}",
                         "- この部品の中身は、とどくでは読めない形です（種類の名前と、図としての代替テキストだけを出しています）"]


def _kind_of_bytes(b: bytes | None) -> str:
    if not b:
        return "空"
    if b[:4] == b"%PDF":
        return "PDF"
    if b[:8] == MAGIC_CFB:
        return "複合ファイル"
    if b[:2] == b"PK":
        return "ZIP"
    return f"先頭が {b[:4].hex()} のファイル"


def _package(blob: bytes, name: str, d: dict, depth: int) -> None:
    from ..zipraw import RawZip

    z = RawZip(blob)
    names = z.names()
    low = [n.lower() for n in names]
    if any(n.startswith("word/") for n in low):
        from ..docx.document import DocxDocument

        _word(DocxDocument(blob, name), d, depth)
    elif any(n.startswith("xl/") for n in low):
        from ..xlsx.workbook import Workbook

        _excel(Workbook(blob, name), d)
    elif any(n.startswith("ppt/") for n in low):
        from .ppt import pptx_slides

        _slides(pptx_slides(z), d)
    elif any(n.startswith("visio/") for n in low):
        _visio(z, d)
    else:  # ふつうの ZIP の書庫：中のファイルの名前を出す（中身は開かない）
        d["kind"] = "package"
        d["label"] = "ZIP の書庫"
        d["markdown"] = ["- ZIP の書庫の中のファイル：" + "、".join(_md(n) for n in names)]


def _md(s) -> str:
    return str(s).replace("|", "\\|").replace("\r", " ").replace("\n", " ")


def _excel(wb, d: dict) -> None:
    from ..xlsx import window as xw

    d["kind"] = "excel"
    if d["label"] == "埋め込み" or not d["class"]:
        d["label"] = "Excel の表"
    doc = xw.build(wb)
    lines = []
    for sh in doc["sheets"]:
        for ln in xw._md_sheet(sh):
            if ln.startswith("## "):
                ln = "#### " + ln[3:]
            elif ln.startswith("### "):
                ln = "##### " + ln[4:]
            lines.append(ln)
    d["sheets"] = [sh["name"] for sh in doc["sheets"]]
    d["markdown"] = lines


class _NoSpots:
    """埋め込みの文書の中は書き込めないので、欄の印（［P3.1］など）を付けずに読む。"""

    items: dict = {}
    marks: dict = {}


def word_lines(doc) -> list[str]:
    """Word の文書（DocxDocument・DocFile）の字の行：段落は1行、表は行ごとにマスを ｜ で区切る。欄の印は付けない。"""
    from ..docx.window import STORY_JA, _render_blocks

    out = []
    for kind, _part, blocks in doc.all_blocks():
        lines: list[str] = []
        _render_blocks(blocks, _NoSpots(), {}, lines, [])
        if not lines:
            continue
        if kind != "body":
            out.append(f"（{STORY_JA.get(kind, kind)}）")
        out += lines
    return out


def _word(doc, d: dict, depth: int) -> None:
    d["kind"] = "word"
    if d["label"] == "埋め込み" or not d["class"]:
        d["label"] = "Word の文書"
    d["markdown"] = [f"- {_md(x)}" for x in word_lines(doc)]
    if depth < MAX_DEPTH:
        from .word import item_markdown, objects as word_objects

        for o in word_objects(doc, depth + 1):
            d["markdown"] += [""] + item_markdown(o, "#####")


def _slides(slides: list[tuple[str, list[str]]], d: dict) -> None:
    d["kind"] = "powerpoint"
    if d["label"] == "埋め込み" or not d["class"]:
        d["label"] = "PowerPoint のスライド"
    lines = []
    for title, rows in slides:
        lines.append(f"- {title}：" + ("" if rows else "（字なし）"))
        lines += [f"  - {_md(r)}" for r in rows]
    d["markdown"] = lines


def _pdf(blob: bytes, name: str, d: dict) -> None:
    from ..pdf.doc import PdfDocument
    from ..pdf.window import build

    d["kind"] = "pdf"
    d["label"] = "PDF" if d["label"] in ("埋め込み", "") else d["label"]
    win = build(PdfDocument(blob, name))
    lines = []
    for pg in win["pages"]:
        lines.append(f"- ページ {pg['number']}：" + ("（画像だけのページ。字の情報がありません）" if pg.get("image_only") else ""))
        lines += [f"  - {_md(t)}" for t in pg["_plain"] if t.strip()]
    d["markdown"] = lines
    d["pages"] = len(win["pages"])


def _visio(z, d: dict) -> None:
    """Visio の図面（.vsdx）：ページの部品の図形の字（Text 要素）。"""
    import xml.etree.ElementTree as ET

    from .chart import _ln

    d["kind"] = "visio"
    d["label"] = "Visio の図面"
    pages = sorted((n for n in z.names() if re.fullmatch(r"visio/pages/page\d+\.xml", n)),
                   key=lambda n: int(re.search(r"(\d+)\.xml$", n)[1]))
    lines = []
    for k, p in enumerate(pages, 1):
        root = ET.fromstring(z.read(p))
        texts = ["".join(t.itertext()).strip() for t in root.iter() if _ln(t.tag) == "Text"]
        lines.append(f"- ページ {k}：")
        lines += [f"  - {_md(t)}" for t in texts if t]
    d["markdown"] = lines


def object_markdown(o: dict, level: str = "##") -> list[str]:
    """埋め込み1つの節（見出しと中身）。o は describe の結果に n・storage・where などを足したもの。"""
    head = f"{level} 埋め込み{o['n']}：{_md(o['label'])}（{_md(o.get('class') or o.get('prog') or '種類の名前なし')}"
    head += f"・{_md(o['storage'])}）" if o.get("storage") else "）"
    out = [head, ""]
    note = []
    if o.get("where"):
        note.append(f"位置：{o['where']}")
    if o.get("file"):
        note.append(f"ファイル名「{_md(o['file'])}」")
    if o.get("alt"):
        note.append(f"代替テキスト「{_md(o['alt'])}」")
    if note:
        out.append("- " + "、".join(note))
    out.append("- 埋め込みの中は読むだけです（書き込めません）。")
    if o.get("error"):
        out.append(f"- {o['error']}")
    out.append("")
    out += o.get("markdown") or []
    if out[-1] != "":
        out.append("")
    return out
