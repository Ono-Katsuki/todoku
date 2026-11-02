"""Word の文書（.doc・.docx）に埋め込んだ Excel の表に書く：ブックのセルと、Word が表示・印刷に使う表の絵の両方。

欄の名前は「E{埋め込みの番号}!{シート}!{セル}」（例 E1!Sheet1!C4）。Word の窓の埋め込みの節に、この名前で出す。

書き方（どれも、元のファイルを正本にした差分）：
1. ブック：埋め込みの入れ物を、それだけの Excel のファイルとして開き（.doc は ObjectPool/_… の入れ物を
   todoku.cfb.build_cfb で .xls に包む。.docx は word/embeddings の .xlsx、または .bin の中の Workbook・Package）、
   Excel への書き込み（todoku.xlsx.edit.fill：入力の解釈・入力規則・数式の再計算）をそのまま使う。
   書き戻すのは、変わったストリーム（Workbook・Package）か部品だけ。
2. 表の絵：Word は、埋め込みの表を Excel が描いた絵で表示・印刷する。
   - .doc：印刷用の \\003EPRINT（EMF）と、表示用の絵（EMBED のフィールドの結果の絵の字 0x01 の sprmCPicLocation が指す
     Data の PICF → OfficeArtInlineSpContainer → OfficeArtFBSE → OfficeArtBlipEMF。EMF は DEFLATE で縮めてある）。
     OfficeArtInlineSpContainer の絵の入れ物（rgfb）は、図形の性質の中の BLIP の性質（pib など）の現れる順に並ぶ。
     pib の値（文書全体の絵の番号）は使わない（MS-ODRAW 2.2.15：その値は無視する）。
   - .docx：w:object の v:imagedata（と a:blip）が指す絵の部品（.emf）。
   どの絵も todoku.sheetpic で、変わったセルの字だけを消して描き足す（元のレコードはバイト列のまま）。
   .doc の表示用の絵は大きさが変わるので、元の PICF はそのまま残し、新しい PICF を Data の後ろに足して、
   絵の字の sprmCPicLocation（4バイト、CHPX の中）だけを新しい位置に書き換える（ほかの絵や欄の位置は動かない）。
   OfficeArt の絵の識別子（rgbUid）は、縮める前の EMF の MD4（RFC 1320）。元の絵でも、そうなっていることを確かめた。
   ブックで選ばれているシートがグラフのシートなら、絵はそのグラフ：グラフ（と、グラフの中の図形）が使うセルの値が
   変わらない書き込みだけを受け付け、絵はそのまま残す（使うセルの値が変わるなら、グラフを描き直せないので書かない）。
3. 確かめられないことは書かない：WMF の絵・写真（ビットマップ）の絵・印刷用の \\003PRINT・再計算できない数式・
   表の絵の中でセルの位置が決まらないとき・セルの値を変えたときに動くマクロのあるブックなどは、何が足りないかを
   日本語で示して断る（ボタンやブックを開いたときに動くマクロは、値を書くことでは動かないので、知らせて書く）。

証明（verify）は、書き込み処理の内部の情報を使わず、前後の2つのファイルから確かめる：
- ブック：Excel の証明（todoku.xlsx.proof.prove）。変わったのは書いたセルとその計算結果だけ
- 表の絵：消したレコードは、変わったセルの元の字だけ。足したレコードは、状態の保存・切り抜き・字体・字の色・寄せ方・
  字だけで、描いた字は変わったセルの新しい字と同じ。そのほかのレコードと見出しは、数（バイト数・レコード数・
  オブジェクトの表の大きさ）を除いてバイト一致
- .doc の表示用の絵：元の PICF は Data にそのまま残り、新しい PICF は、絵の識別子・大きさ・縮めた中身のほかは元と同じ。
  文字の書式（CHPX）で変わったのは、その絵を指す sprmCPicLocation の値だけ
- グラフを表示している埋め込み：グラフが使うセルの値が前後で同じで、絵（印刷用・表示用）は1バイトも変わっていない
"""

from __future__ import annotations

import re
import struct
import zlib
from dataclasses import dataclass, field

from .cfb import CFB, SubStorage, build_cfb
from .metafile import Emf, MetafileError, md4
from .sheetpic import Layout, PictureError, SheetModel, changed_cells, edit_picture

EMBED_KEY = re.compile(r"^E(\d+)!(.+)$")
PRESENTATION = ("\x03EPRINT", "\x03PRINT", "\x02OlePres")


class EmbedError(ValueError):
    pass


def is_embed_key(key: str) -> bool:
    return bool(EMBED_KEY.match(str(key)))


# ====================================================================== 埋め込みのブックを開く
@dataclass
class Book:
    """埋め込み1つのブック。how は、書いたブックを外側に戻すやり方。"""

    e: dict
    wb: object
    how: str  # doc-xls / doc-pkg / part-xlsx / part-xls / part-pkg
    raw: bytes  # 書き込みに渡す、それだけのファイル（.xls に包んだもの、または .xlsx）
    stream: str = ""  # doc-xls・part-xls：ブックのストリームの名前


def open_book(doc, e: dict) -> Book:
    from .xls.book import XlsWorkbook
    from .xlsx.workbook import Workbook

    name = e["storage"]
    if getattr(doc, "kind", "") == "doc":
        items = {p: d for p, d in SubStorage(doc.cfb, name).contents().items() if not p.startswith(PRESENTATION)}
        low = {p.lower(): p for p in items}
        if "workbook" in low or "book" in low:
            raw = build_cfb(items)
            wb = XlsWorkbook(raw, name)
            return Book(e, wb, "doc-xls", raw, wb.stream_name)
        if "package" in low:
            raw = items[low["package"]]
            return Book(e, Workbook(raw, name), "doc-pkg", raw, low["package"])
        raise EmbedError(f"埋め込み{e['n']}：Excel のブックが入っていません")
    data = doc.zip.read(name)
    if data[:4] == b"PK\x03\x04":
        return Book(e, Workbook(data, name), "part-xlsx", data)
    if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        c = CFB(data)
        low = {p.lower(): p for p in c.streams()}
        if "workbook" in low or "book" in low:
            # .doc と同じく、表の絵のストリームを除いた入れ物だけを .xls に包んで書き、ブックのストリームだけを戻す
            raw = build_cfb({p: c.read(p) for p in c.streams() if not p.startswith(PRESENTATION)})
            wb = XlsWorkbook(raw, name)
            return Book(e, wb, "part-xls", raw, wb.stream_name)
        if "package" in low:
            raw = c.read(low["package"])
            return Book(e, Workbook(raw, name), "part-pkg", raw, low["package"])
    raise EmbedError(f"埋め込み{e['n']}：Excel のブックが入っていません")


def macro_block(wb) -> str:
    """マクロのある埋め込みのブックに書けない理由（書けるなら空）。セルの値を変えたときに動くマクロ（Worksheet_Change など）が
    あると、Excel で表を編集したときの結果を確かめられないので書かない。それ以外のマクロ（ボタン・開いたとき）は、
    値を書くことでは動かない。"""
    if not getattr(wb, "has_macros", False):
        return ""
    from .macros import change_event_sheets, macro_info

    try:
        project = wb.vba
    except Exception:  # noqa: BLE001
        project = None
    if project is None:
        return "マクロの入れ物を読めないので、値を変えたときに動く処理があるかを確かめられず、書けません"
    info = macro_info(project, True, {sh.code_name: sh.name for sh in wb.sheets if getattr(sh, "code_name", None)})
    ev = change_event_sheets(info)
    if ev:
        procs = sorted({p for ps in ev.values() for p in ps})
        return ("セルの値を変えたときに動くマクロ（" + "、".join(procs) + "）があるので、まだ書けません"
                "（Excel で表を編集したときにマクロが変えるものを、表の絵に描けないため）")
    return ""


def reopen(book: Book, raw: bytes):
    from .xls.book import XlsWorkbook
    from .xlsx.workbook import Workbook

    if book.how in ("doc-xls", "part-xls"):
        return XlsWorkbook(raw, book.e["storage"])
    return Workbook(raw, book.e["storage"])


def is_icon(doc, e: dict) -> bool:
    """アイコンで表示している埋め込み（表の絵ではなくアイコンなので、セルに書いても絵は変わらない）。
    .doc は ObjInfo（ODT）の ODTPersist1 の fIcon（0x0040）、.docx は o:OLEObject の DrawAspect="Icon"。"""
    if getattr(doc, "kind", "") == "doc":
        path = e["storage"] + "/\x03ObjInfo"
        if doc.cfb.has(path):
            odt = doc.cfb.read(path)
            return len(odt) >= 2 and bool(struct.unpack_from("<H", odt, 0)[0] & 0x0040)
        return False
    return e.get("aspect") == "Icon"


# ====================================================================== グラフのシートを映す埋め込み
_AREA = re.compile(r"^(?:'((?:[^']|'')+)'|([^!'\s(),\[\]]+))!(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$")
# 部品の中の、セルを指す式を持つ属性（DrawingML の図形の textlink、フォームの部品の fmlaLink など）
_FORMULA_ATTRS = {"textlink", "fmlalink", "fmlarange", "fmlatxbx", "fmlamacro", "fmlagroup"}


def _areas(wb, text: str, where: str) -> list[tuple[str, int, int, int, int]]:
    """参照の字（「Sheet1!$A$1:$A$5」「(Sheet1!$A$1,Sheet1!$C$1)」）→ (シート, 行1, 列1, 行2, 列2)。
    シートとセルの範囲だけでできていない式（名前・関数・ほかのブック）は、どのセルを使うか決められないので断る。"""
    from .embeds.chart import _split_top
    from .formula.parser import parse_cell

    t = text.strip()
    if t.startswith("(") and t.endswith(")"):
        t = t[1:-1]
    out = []
    for part in _split_top(t):
        m = _AREA.match(part.strip())
        sheet = (m[1].replace("''", "'") if m[1] is not None else m[2]) if m else None
        if not m or not wb.has_sheet(sheet):
            raise EmbedError(f"表の絵のグラフ（{where}）の参照「{text}」が、ブックのシートとセルの範囲だけの形ではないので、"
                             "グラフがどのセルを使うかを決められません")
        a, _, b = m[3].replace("$", "").partition(":")
        r1, c1 = parse_cell(a)
        r2, c2 = parse_cell(b) if b else (r1, c1)
        out.append((sheet, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)))
    return out


def chart_areas(wb) -> list[tuple[str, int, int, int, int]] | None:
    """表の絵に映るシートがグラフのシートなら、そのグラフ（と、グラフの中の図形）が使うセルの範囲。グラフのシートでなければ None。
    Word は、埋め込みのブックを選ばれているシートの絵で表示するので、グラフのシートが選ばれていれば、絵はそのグラフ。
    グラフが使うセルの値が変わらなければ、絵は変わらない。使うセルを決められないときは EmbedError。"""
    from .sheetpic import shown_sheet

    sh = shown_sheet(wb)
    if sh is None or not sh.state.startswith("non-worksheet:chartsheet"):
        return None
    if getattr(wb, "kind", "xlsx") == "xls":
        return _biff_chart_areas(wb, sh)
    return _ooxml_chart_areas(wb, sh)


def _biff_chart_areas(wb, sh) -> list[tuple[str, int, int, int, int]]:
    """グラフのシートの部分（MS-XLS 2.1.7.20.1）：セルを指すのは、系列・題名などの BRAI（0x1051）の formula。"""
    from .embeds.biffchart import BRAI, substream

    try:
        recs = substream(wb.data, int(str(sh.path).lstrip("#")))
    except (ValueError, struct.error, IndexError) as ex:
        raise EmbedError(f"表の絵のグラフ（{sh.name}）のレコードを読めません（{ex}）") from None
    out = []
    for _pos, t, b in recs[1:]:
        if t == 0x0809:
            raise EmbedError(f"表の絵のグラフ（{sh.name}）の中に、さらに別の部分があるので、グラフが使うセルを決められません")
        if t in (0x005D, 0x00EC, 0x01B6):  # OBJ・MsoDrawing・TxO：グラフの上の図形（セルにつないだ字などがありうる）
            raise EmbedError(f"表の絵のグラフ（{sh.name}）の上に図形があるので、グラフの絵が使うセルを決められません")
        if t == BRAI:
            if len(b) < 8:
                raise EmbedError(f"表の絵のグラフ（{sh.name}）の参照（BRAI）が壊れています")
            cce = struct.unpack_from("<H", b, 6)[0]
            if not cce:
                continue
            try:
                text = wb.decoder.decode(b[8:8 + cce])
            except Exception as ex:  # noqa: BLE001
                raise EmbedError(f"表の絵のグラフ（{sh.name}）の参照の式を読めません（{ex}）") from None
            out += _areas(wb, text, sh.name)
    return out


def _ooxml_chart_areas(wb, sh) -> list[tuple[str, int, int, int, int]]:
    """グラフのシート（xl/chartsheets）から関係でたどれる部品（図・グラフ・グラフの中の図形）の、セルを指す式すべて
    （c:f・cx:f の要素と、textlink などの属性）。"""
    import xml.etree.ElementTree as ET

    from .embeds.chart import rel_types

    z = wb.zip
    seen, todo, out = set(), [sh.path], []
    while todo:
        part = todo.pop()
        if part in seen or not z.has(part):
            continue
        seen.add(part)
        if not part.lower().endswith(".xml"):
            continue  # 絵（.png など）
        try:
            root = ET.fromstring(z.read(part))
        except ET.ParseError:
            raise EmbedError(f"表の絵のグラフの部品 {part} を読めません") from None
        for el in root.iter():
            if el.tag.rsplit("}", 1)[-1] == "f" and (el.text or "").strip():
                out += _areas(wb, el.text, part)
            for k, v in el.attrib.items():
                if k.rsplit("}", 1)[-1].lower() in _FORMULA_ATTRS and v.strip():
                    out += _areas(wb, v.lstrip("="), part)
        for _typ, target in rel_types(z, part).values():
            if "://" not in target:
                todo.append(target)
    return out


def chart_changes(old_wb, new_wb, areas) -> list[str]:
    """グラフが使うセルのうち、値が前後で変わったもの（「シート!セル」）。"""
    from .formula.parser import addr

    out = []
    for sheet, r1, c1, r2, c2 in areas:
        a, b = old_wb.sheet(sheet), new_wb.sheet(sheet)
        cells = {rc for rc in list(a.cells) + list(b.cells) if r1 <= rc[0] <= r2 and c1 <= rc[1] <= c2}
        for r, c in sorted(cells):
            x, y = a.cells.get((r, c)), b.cells.get((r, c))
            vx = x.value if x is not None else None
            vy = y.value if y is not None else None
            if vx != vy or type(vx) is not type(vy):
                out.append(f"{sheet}!{addr(r, c)}")
    return out


def find_embed(doc, n: int) -> dict:
    for e in getattr(doc, "embeds", []):
        if e["n"] == n:
            return e
    raise EmbedError(f"埋め込み{n} はありません")


# ====================================================================== 書く
@dataclass
class EmbedWrite:
    """埋め込みの表に書いた結果：外側のファイルに戻す材料と、書いた欄。"""

    streams: dict[str, bytes] = field(default_factory=dict)  # .doc：差し替えるストリーム
    parts: dict[str, bytes] = field(default_factory=dict)  # .docx：差し替える部品
    data_tail: bytes = b""  # .doc：Data の後ろに足す、新しい表示用の絵
    wd_patches: list[tuple[int, bytes]] = field(default_factory=list)  # .doc：WordDocument の同じ大きさの書き換え
    written: list[tuple[str, object]] = field(default_factory=list)  # (欄の名前, 書いた値)
    pictures: int = 0  # 描き直した表の絵の数
    warnings: list[str] = field(default_factory=list)


def write(doc, values: dict[str, object], data_len: int | None = None) -> EmbedWrite:
    """values：{"E1!Sheet1!C4": 値, …}。data_len：.doc の Data のいまの長さ（新しい絵を、その後ろに置く）。"""
    from .formula.parser import addr
    from .xlsx.edit import EditError, fill, resolve_target

    out = EmbedWrite()
    groups: dict[int, dict[str, object]] = {}
    for key, v in values.items():
        m = EMBED_KEY.match(str(key))
        if not m:
            raise EmbedError(f"{key}: 埋め込みの表の欄の名前ではありません（E1!Sheet1!C4 のように書きます）")
        groups.setdefault(int(m[1]), {})[m[2]] = (key, v)
    tail = bytearray()
    base = data_len if data_len is not None else len(getattr(doc, "data", b""))
    for n, cells in sorted(groups.items()):
        e = find_embed(doc, n)
        if not str(e.get("prog", "")).startswith(("Excel.Sheet", "Excel.Chart")):
            raise EmbedError(f"埋め込み{n} は Excel の表ではありません（{e.get('prog') or '種類不明'}）")
        book = open_book(doc, e)
        wb = book.wb
        why = macro_block(wb)
        if why:
            raise EmbedError(f"埋め込み{n}：{why}")
        targets = {}
        for rest, (key, v) in cells.items():
            if "!" in rest:
                sheet, cell = rest.rsplit("!", 1)
                if wb.has_sheet(sheet.strip("'")) and any(c.name == cell for c in wb.sheet(sheet.strip("'")).controls):
                    raise EmbedError(f"{key}: 埋め込みの表のフォームの部品には、まだ書けません（表の絵に部品の状態を描けないため）")
            try:
                s, r, c = resolve_target(wb, rest)
            except EditError as ex:
                raise EmbedError(f"{key}: {ex}") from None
            targets[f"{s}!{addr(r, c)}"] = (key, v)
        try:
            res = fill(book.wb, {k: v for k, (_key, v) in targets.items()}, run_macros=False)
        except EditError as ex:
            raise EmbedError("\n".join(f"埋め込み{n}：{ln}" for ln in str(ex).splitlines())) from None
        if res.unverified:
            raise EmbedError(f"埋め込み{n}：書いた値に関わる数式のうち、計算し直せないものがあるので、表の絵を正しく描けません（"
                             + "、".join(res.unverified[:5]) + "）")
        out.warnings += [f"埋め込み{n}：{w}" for w in res.warnings]
        if getattr(wb, "has_macros", False):
            out.warnings.append(f"埋め込み{n}：この Excel の表にはマクロがあります（セルの値を変えたときに動くものはありません）。"
                                "Word で表を開いて編集すると、Excel がブックを開いたときに動くマクロなどが動くことがあります")
        after = {(p.sheet, p.address): p.after for p in res.patches if p.kind == "input"}
        for k, (key, v) in targets.items():
            s, a = k.split("!", 1)
            out.written.append((key, after.get((s, a), v)))
        if not res.patches:
            continue  # 同じ値を書いた：何も変わらない
        new_wb = reopen(book, res.blob)
        areas = None if is_icon(doc, e) else chart_areas(wb)
        if areas is not None:
            # 表の絵はグラフ：グラフが使うセルの値が変わらなければ、絵はそのまま（変わるなら、グラフを描き直せないので書かない）
            moved = chart_changes(wb, new_wb, areas)
            if moved:
                raise EmbedError(f"埋め込み{n}：Word に表示しているのはグラフで、書いた値でグラフが使うセル（"
                                 + "、".join(moved[:5]) + "）の値が変わります。グラフの絵は、まだ描き直せません")
        # ブックを外側に戻す
        name = e["storage"]
        if book.how == "doc-xls":
            out.streams[f"{name}/{book.stream}"] = CFB(res.blob).read(book.stream)
        elif book.how == "doc-pkg":
            out.streams[f"{name}/{book.stream}"] = res.blob
        elif book.how == "part-xlsx":
            out.parts[name] = res.blob
        elif book.how == "part-xls":
            out.parts[name] = _cfb_replace(doc.zip.read(name), book.stream, CFB(res.blob).read(book.stream))
        else:
            out.parts[name] = _cfb_replace(doc.zip.read(name), book.stream, res.blob)
        # 表の絵を描き直す
        if areas is not None or is_icon(doc, e):
            continue  # グラフの絵（使うセルは変わらない）・アイコンの絵は変わらない
        old_model, new_model = SheetModel(wb), SheetModel(new_wb)
        if not changed_cells(old_model, new_model):
            continue  # 絵に映る範囲の字は変わらない
        if getattr(doc, "kind", "") == "doc":
            n_pics = _doc_pictures(doc, e, old_model, new_model, out, base, tail)
        else:
            n_pics = _docx_pictures(doc, e, old_model, new_model, out)
        if not n_pics:
            raise EmbedError(f"埋め込み{n}：Word が表示・印刷に使う表の絵が見つかりません")
        out.pictures += n_pics
    out.data_tail = bytes(tail)
    return out


def _cfb_replace(blob: bytes, stream: str, data: bytes) -> bytes:
    """入れ物（複合ファイル）の1つのストリームを差し替える。小さなストリーム（ミニストリーム）の大きさが変わってもよい
    （todoku.cfbwrite。ほかのストリームのバイトは変えない）。"""
    from .cfbwrite import replace_stream

    return replace_stream(blob, stream, data)


def _edit(data: bytes, old: SheetModel, new: SheetModel, what: str) -> bytes:
    _emf_or_fail(data, what)
    try:
        return edit_picture(data, old, new, what).data
    except PictureError as ex:
        raise EmbedError(str(ex)) from None


def _emf_or_fail(data: bytes, what: str) -> None:
    if data[:4] == b"\x01\x00\x00\x00" and data[40:44] == b" EMF":
        return
    if data[:4] in (b"\xd7\xcd\xc6\x9a", b"\x01\x00\x09\x00", b"\x02\x00\x09\x00"):
        raise EmbedError(f"{what}が Windows メタファイル（WMF）なので、まだ描き直せません（EMF の絵だけ描き直せます）")
    raise EmbedError(f"{what}が、字を描き直せる形（EMF）ではありません（写真などの絵です）")


# ---------------------------------------------------------------------- .docx
def _docx_pictures(doc, e: dict, old: SheetModel, new: SheetModel, out: EmbedWrite) -> int:
    n = 0
    name = e["storage"]
    part = out.parts.get(name) or doc.zip.read(name)
    if part[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        # OLE の入れ物（.bin）の中に、印刷用の絵が入っていることがある
        c = CFB(part)
        names = c.streams()
        if any(p.startswith(("\x03PRINT", "\x02OlePres")) for p in names):
            raise EmbedError(f"埋め込み{e['n']}：入れ物の中の絵（\\003PRINT・OlePres）が、まだ描き直せない形です")
        if "\x03EPRINT" in names:
            out.parts[name] = _cfb_replace(part, "\x03EPRINT", _edit(c.read("\x03EPRINT"), old, new, "印刷用の表の絵（EPRINT）"))
            n += 1
    for img in e.get("images", []):
        data = doc.zip.read(img)
        out.parts[img] = _edit(data, old, new, f"表の絵（{img}）")
        n += 1
    return n


# ---------------------------------------------------------------------- .doc
def _doc_pictures(doc, e: dict, old: SheetModel, new: SheetModel, out: EmbedWrite, base: int, tail: bytearray) -> int:
    name = e["storage"]
    n = 0
    sub = SubStorage(doc.cfb, name)
    names = sub.streams()
    if any(p.startswith("\x03PRINT") or p.startswith("\x02OlePres") for p in names):
        raise EmbedError(f"埋め込み{e['n']}：印刷用の絵（\\003PRINT・OlePres）が Windows メタファイルなどで、まだ描き直せません")
    if "\x03EPRINT" in names:
        out.streams[f"{name}/\x03EPRINT"] = _edit(sub.read("\x03EPRINT"), old, new, "印刷用の表の絵（EPRINT）")
        n += 1
    done = set()
    for _cp, loc in e.get("pictures", []):
        if loc in done:
            continue
        done.add(loc)
        pic = Picture(doc.data, loc)
        emf = pic.metafile()
        new_emf = _edit(emf, old, new, "表示用の表の絵（Data の PICF）")
        at = base + len(tail)
        tail += pic.rebuild(new_emf, at)
        out.wd_patches += location_patches(doc, loc, at)
        n += 1
    return n


class Picture:
    """Data の中の絵：PICF（68 バイト）＋ OfficeArtInlineSpContainer（図形と、絵の入れ物 OfficeArtFBSE・BLIP）。"""

    def __init__(self, data: bytes, pos: int):
        if pos + 68 > len(data):
            raise EmbedError("表示用の表の絵（PICF）の位置が、Data の外を指しています")
        self.data, self.pos = data, pos
        self.lcb, self.cbh = struct.unpack_from("<iH", data, pos)
        self.mm = struct.unpack_from("<h", data, pos + 6)[0]
        if self.cbh != 0x44 or self.lcb < self.cbh or pos + self.lcb > len(data):
            raise EmbedError("表示用の表の絵（PICF）の大きさが壊れています")
        q = pos + self.cbh
        if self.mm == 0x66:
            q += 1 + data[q]  # 絵のファイル名（MM_SHAPEFILE）
        elif self.mm != 0x64:
            raise EmbedError(f"表示用の表の絵が OfficeArt の形でない古い形（mm={self.mm}）なので、まだ描き直せません")
        self.head_end = q
        end = pos + self.lcb
        recs = []
        while q + 8 <= end:
            vi, rt, ln = struct.unpack_from("<HHI", data, q)
            recs.append((q, vi, rt, ln))
            q += 8 + ln
        if q != end or not recs or recs[0][2] != 0xF004:
            raise EmbedError("表示用の表の絵（OfficeArt）の入れ物が壊れています")
        self.recs = recs
        # 絵の入れ物（rgfb）は、図形の性質の中の BLIP の性質の現れる順に並ぶ。pib などの値（番号）は使わない（MS-ODRAW 2.2.15）
        order = _blip_props(data, recs[0][0], recs[0][0] + 8 + recs[0][3])
        blips = recs[1:]
        if 0x0104 not in order:
            raise EmbedError("表示用の表の絵の、表示する絵（pib）が見つかりません")
        if len(order) != len(blips):
            raise EmbedError(f"表示用の表の絵の、絵の性質の数（{len(order)}）と絵の入れ物の数（{len(blips)}）が合わないので、"
                             "どの絵を表示しているか決まりません")
        self.slot = 1 + order.index(0x0104)  # recs の中の番号（0 は図形）

    def _blip_at(self) -> tuple[int, int, int, int]:
        """(BLIP の位置, ver/inst, 種類, 長さ)。FBSE の中に埋め込まれていれば、その中。"""
        q, vi, rt, ln = self.recs[self.slot]
        if rt == 0xF007:
            cb_name = self.data[q + 8 + 33]
            b = q + 8 + 36 + cb_name
            if b + 8 > q + 8 + ln:
                raise EmbedError("表示用の表の絵が、文書の中に埋め込まれていません（別の所を指しています）")
            vi2, rt2, ln2 = struct.unpack_from("<HHI", self.data, b)
            return b, vi2, rt2, ln2
        return q, vi, rt, ln

    def metafile(self) -> bytes:
        b, vi, rt, ln = self._blip_at()
        if rt == 0xF01B:
            raise EmbedError("表示用の表の絵が Windows メタファイル（WMF）なので、まだ描き直せません（EMF の絵だけ描き直せます）")
        if rt != 0xF01A:
            raise EmbedError("表示用の表の絵が写真などのビットマップなので、字を描き直せません")
        inst = vi >> 4
        h = b + 8 + (16 if inst == 0x3D4 else 32)
        cb_size, cb_save, comp = struct.unpack_from("<I", self.data, h)[0], struct.unpack_from("<I", self.data, h + 28)[0], \
            self.data[h + 32]
        raw = self.data[h + 34:h + 34 + cb_save]
        if comp == 0:
            emf = zlib.decompress(raw)
        elif comp == 0xFE:
            emf = raw
        else:
            raise EmbedError("表示用の表の絵の縮め方が分かりません")
        if len(emf) != cb_size:
            raise EmbedError("表示用の表の絵の大きさが、記録と合いません")
        return emf

    def rebuild(self, emf: bytes, at: int) -> bytes:
        """新しい EMF の絵にした PICF 全体（Data の at の位置に置く）。図形と、ほかの絵はバイト列のまま。"""
        d = self.data
        b, vi, rt, ln = self._blip_at()
        inst = vi >> 4
        n_uid = 1 if inst == 0x3D4 else 2
        h = b + 8 + 16 * n_uid
        comp = d[h + 32]
        packed = zlib.compress(emf, 9) if comp == 0 else emf
        uid = md4(emf)
        meta = struct.pack("<I", len(emf)) + d[h + 4:h + 28] + struct.pack("<I", len(packed)) + d[h + 32:h + 34]
        blip = struct.pack("<HHI", vi, rt, 16 * n_uid + 34 + len(packed)) + uid * n_uid + meta + packed
        q, fvi, frt, fln = self.recs[self.slot]
        if frt == 0xF007:
            cb_name = d[q + 8 + 33]
            fo = struct.unpack_from("<I", d, q + 8 + 28)[0]
            fo_new = at + self.head_end - self.pos if fo == self.head_end else fo  # Word は図形の位置を書く
            fixed = bytearray(d[q + 8:q + 8 + 36 + cb_name])
            fixed[2:18] = uid
            struct.pack_into("<I", fixed, 20, len(blip))
            struct.pack_into("<I", fixed, 28, fo_new)
            item = struct.pack("<HHI", fvi, frt, 36 + cb_name + len(blip)) + bytes(fixed) + blip
        else:
            item = blip
        body = bytearray(d[self.pos:self.head_end])
        for k, (rq, _v, _t, rl) in enumerate(self.recs):
            body += item if k == self.slot else d[rq:rq + 8 + rl]
        struct.pack_into("<i", body, 0, len(body))
        return bytes(body)


# BLIP の性質（MS-ODRAW 2.3.23 pib・pibPrint、2.3.7.7 fillBlip、2.3.8.6 lineFillBlip、2.3.9〜2.3.12 の四辺の lineFillBlip）
BLIP_PROPS = (0x0104, 0x010F, 0x0186, 0x01C5, 0x0545, 0x0585, 0x05C5, 0x0605)


def _blip_props(data: bytes, a: int, b: int) -> list[int]:
    """図形の入れ物（a〜b）の shapePrimaryOptions（0xF00B）と shapeTertiaryOptions1（0xF122）にある BLIP の性質を、
    現れる順に。OfficeArtInlineSpContainer の rgfb は、この順に並ぶ（MS-ODRAW 2.2.15）。"""
    q = a + 8
    out = []
    while q + 8 <= b:
        vi, rt, ln = struct.unpack_from("<HHI", data, q)
        if rt in (0xF00B, 0xF122):
            n = vi >> 4
            if 6 * n > ln:
                raise EmbedError("表示用の表の絵の、図形の性質の並びが壊れています")
            for k in range(n):
                pid = struct.unpack_from("<H", data, q + 8 + 6 * k)[0] & 0x3FFF
                if pid in BLIP_PROPS:
                    out.append(pid)
        q += 8 + ln
    return out


def _prc_grpprls(doc) -> list[bytes]:
    """ピースの表の前の Prc（ピース全体に掛かる書式）の grpprl の並び。"""
    out = []
    t = doc.table
    i, end = doc.fc["Clx"][0], doc.clx_pcdt
    while i < end and t[i] == 1:
        cb = struct.unpack_from("<H", t, i + 1)[0]
        out.append(t[i + 3:i + 3 + cb])
        i += 3 + cb
    return out


def location_patches(doc, old: int, new: int) -> list[tuple[int, bytes]]:
    """値が old の sprmCPicLocation（0x6A03）を、すべて new に書き換える、WordDocument の中の4バイトの書き換え。"""
    from .doc.book import sprm_spans

    for g in _prc_grpprls(doc):
        for sp, at, size in sprm_spans(g):
            if sp == 0x6A03 and size == 4 and struct.unpack_from("<I", g, at)[0] == old:
                raise EmbedError("表示用の表の絵の位置が、ピース全体の書式（Prc）に書かれているので、書き換えられません")
    fcp, lcb = doc.fc["PlcfBteChpx"]
    n = (lcb - 4) // 8 if lcb else 0
    pns = struct.unpack_from(f"<{n}I", doc.table, fcp + 4 * (n + 1)) if n else ()
    out = []
    for pn in pns:
        page = doc.wd[pn * 512:(pn + 1) * 512]
        if len(page) < 512:
            continue
        crun = page[511]
        for boff in {page[4 * (crun + 1) + k] * 2 for k in range(crun)}:
            if not boff:
                continue
            cb = page[boff]
            g = page[boff + 1:boff + 1 + cb]
            for sp, at, size in sprm_spans(g):
                if sp == 0x6A03 and size == 4 and struct.unpack_from("<I", g, at)[0] == old:
                    out.append((pn * 512 + boff + 1 + at, struct.pack("<I", new)))
    if not out:
        raise EmbedError("表示用の表の絵を指す、文字の書式（sprmCPicLocation）が見つかりません")
    return out


# ====================================================================== 証明
def verify_picture(old_data: bytes, new_data: bytes, old: SheetModel, new: SheetModel, what: str,
                   stats: dict | None = None) -> list[str]:
    """表の絵の差分が「変わったセルの元の字を消し、新しい字を足しただけ」であることを確かめる（問題の一覧を返す）。"""
    from collections import Counter

    from .formula.parser import addr
    from .metafile import (EMR_DELETEOBJECT, EMR_EXTCREATEFONTINDIRECTW, EMR_EXTTEXTOUTA, EMR_EXTTEXTOUTW,
                           EMR_INTERSECTCLIPRECT, EMR_RESTOREDC, EMR_SAVEDC, EMR_SELECTOBJECT, EMR_SETBKMODE,
                           EMR_SETTEXTALIGN, EMR_SETTEXTCOLOR)

    try:
        a, b = Emf(old_data), Emf(new_data)
    except MetafileError as ex:
        return [f"{what}：EMF として読めません（{ex}）"]
    probs = []
    if a.data[:48] + a.data[58:a.records[0][2]] != b.data[:48] + b.data[58:b.records[0][2]]:
        probs.append(f"{what}：見出しの、数のほかの所が変わっています")
    ins = a.insertion_point()
    ra = [a.rec(i) for i in range(len(a.records))]
    rb = [b.rec(i) for i in range(len(b.records))]
    k = len(ra) - ins
    if rb[len(rb) - k:] != ra[ins:]:
        probs.append(f"{what}：終わりの部分（EMF+ の終わり・EOF）が変わっています")
        return probs
    head_a, head_b = ra[:ins], rb[:len(rb) - k]
    removed, i, j = [], 1, 1  # 0 番目は見出し（上で、数のほかが同じことを確かめた）
    while i < len(head_a):
        if j < len(head_b) and head_b[j] == head_a[i]:
            i, j = i + 1, j + 1
        elif a.records[i][1] in (EMR_EXTTEXTOUTW, EMR_EXTTEXTOUTA):
            removed.append(i)
            i += 1
        else:
            probs.append(f"{what}：字でないレコード（{i} 番目）が変わっています")
            return probs
    added = head_b[j:]
    allowed = {EMR_SAVEDC, EMR_RESTOREDC, EMR_INTERSECTCLIPRECT, EMR_EXTCREATEFONTINDIRECTW, EMR_SELECTOBJECT,
               EMR_SETTEXTCOLOR, EMR_SETBKMODE, EMR_SETTEXTALIGN, EMR_EXTTEXTOUTW, EMR_DELETEOBJECT}
    depth, created, drawn = 0, set(), []
    base_idx = len(head_b) - len(added)
    for n_, r in enumerate(added):
        t = struct.unpack_from("<I", r, 0)[0]
        if t not in allowed:
            probs.append(f"{what}：足したレコードに、字を描くのに要らない種類（0x{t:02X}）があります")
            continue
        if t == EMR_SAVEDC:
            depth += 1
        elif t == EMR_RESTOREDC:
            depth -= 1
        elif t == EMR_EXTCREATEFONTINDIRECTW:
            created.add(struct.unpack_from("<I", r, 8)[0])
        elif t == EMR_DELETEOBJECT:
            created.discard(struct.unpack_from("<I", r, 8)[0])
        elif t == EMR_EXTTEXTOUTW:
            run = next(x for x in b.runs if x.index == base_idx + n_)
            drawn.append(run)
        if depth < 0:
            probs.append(f"{what}：足したレコードの状態の保存と戻しが合いません")
    if depth or created:
        probs.append(f"{what}：足したレコードが、状態を元に戻していないか、作った字体を消していません")
    changes = changed_cells(old, new)
    want_removed = Counter(ln for _rc, x, _y in changes for ln in x.lines)
    want_added = Counter(y.text for _rc, _x, y in changes if y.text)
    got_removed = Counter(next(r.text for r in a.runs if r.index == i) for i in removed)
    got_added = Counter(r.text for r in drawn)
    if got_removed != want_removed:
        probs.append(f"{what}：消した字（{dict(got_removed)}）が、変わったセルの元の字（{dict(want_removed)}）と違います")
    if got_added != want_added:
        probs.append(f"{what}：足した字（{dict(got_added)}）が、変わったセルの新しい字（{dict(want_added)}）と違います")
    if probs:
        return probs
    try:
        lay = Layout(a, old)
    except PictureError as ex:
        return [f"{what}：元の絵のセルの位置が決まりません（{ex}）"]
    for i in removed:
        rc = lay.anchors.get(i)
        if rc is None or rc not in {c for c, _x, _y in changes}:
            probs.append(f"{what}：消した字（{i} 番目）が、変わったセルの字と結び付きません")
    for run in drawn:
        ok = False
        for rc, _x, y in changes:
            if y.text != run.text:
                continue
            L, _L1, _R0, R = lay.x_range(*rc)  # 境が決まらないときは、ありうる中でいちばん広いセル
            _L, T, _R, B = lay.rect(*rc)
            if L - 1 <= run.x <= R + 1 and (not lay.y_fixed or T - 1 <= run.y <= B + 1):
                ok = True
                break
        if not ok:
            probs.append(f"{what}：足した字「{run.text}」が、そのセル（" + "・".join(addr(*rc) for rc, _x, y in changes if y.text == run.text)
                         + "）の中に描かれていません")
    if stats is not None and not probs:
        stats["pictures"] = stats.get("pictures", 0) + 1
        stats["removed"] = stats.get("removed", 0) + len(removed)
        stats["added"] = stats.get("added", 0) + len(drawn)
    return probs


def _chart_shown(wa, wb_, doc, e: dict, n: int, probs: list, stats: dict) -> bool:
    """表の絵がグラフのシートの絵なら、グラフが使うセルの値が前後で同じことを確かめて True（問題は probs に足す）。"""
    if is_icon(doc, e):
        return False
    try:
        areas = chart_areas(wa)
    except EmbedError as ex:
        probs.append(f"埋め込み{n}：{ex}")
        return True
    if areas is None:
        return False
    moved = chart_changes(wa, wb_, areas)
    if moved:
        probs.append(f"埋め込み{n}：表示しているグラフが使うセル（" + "、".join(moved[:5]) + "）の値が変わっていますが、グラフの絵は描き直していません")
    else:
        stats["charts"] = stats.get("charts", 0) + 1
    return True


def _note(stats: dict) -> str:
    return (f"埋め込みの表：変わったセルは入力 {stats.get('inputs', 0)}・計算結果 {stats.get('formulas', 0)}、"
            f"表の絵 {stats.get('pictures', 0)} 枚は字の変わったセルの所だけを描き直し（消した字 {stats.get('removed', 0)}・"
            f"描いた字 {stats.get('added', 0)}）"
            + (f"。グラフを表示している埋め込み {stats['charts']} 個は、グラフが使うセルの値が前後で同じ（絵はそのまま）"
               if stats.get("charts") else ""))


def _count(stats: dict, pr) -> None:
    for c in pr.cell_changes:
        k = "formulas" if c["kind"] == "formula" else "inputs"
        stats[k] = stats.get(k, 0) + 1


def written_cells(doc, written: dict) -> dict[int, set[tuple[str, str]]]:
    """書いた欄（E…）→ 埋め込みの番号ごとの (シート, セル)。"""
    from .formula.parser import addr
    from .xlsx.edit import resolve_target

    out: dict[int, set] = {}
    books: dict[int, object] = {}
    for key in written:
        m = EMBED_KEY.match(str(key))
        if not m:
            continue
        n = int(m[1])
        if n not in books:
            books[n] = open_book(doc, find_embed(doc, n)).wb
        s, r, c = resolve_target(books[n], m[2])
        out.setdefault(n, set()).add((s, addr(r, c)))
    return out


def verify_docx(a, b, written: dict) -> tuple[set[str], list[str], str]:
    """.docx：書いた埋め込みの部品（ブックと表の絵）のうち、変わってよいものと、確かめた結果の問題。"""
    from .xlsx.proof import prove as prove_book

    allowed, probs, stats = set(), [], {}
    for n, cells in written_cells(a, written).items():
        ea, eb = find_embed(a, n), find_embed(b, n)
        name = ea["storage"]
        allowed.add(name)
        ba, bb = open_book(a, ea), open_book(b, eb)
        pr = prove_book(ba.raw, bb.raw, allowed_inputs=cells)
        probs += [f"埋め込み{n}（{name}）：{x}" for x in pr.unexpected + pr.lint_errors]
        _count(stats, pr)
        if _chart_shown(ba.wb, bb.wb, a, ea, n, probs, stats):
            continue  # 表の絵（グラフ）は変わらない：絵の部品は、変わってよいものに入れない
        old_m, new_m = SheetModel(ba.wb), SheetModel(bb.wb)
        if ba.how in ("part-pkg", "part-xls"):
            ca, cb = CFB(a.zip.read(name)), CFB(b.zip.read(name))
            if set(ca.streams()) != set(cb.streams()):
                probs.append(f"埋め込み{n}（{name}）：入れ物のストリームの構成が変わっています")
            for p in ca.streams():
                if p == ba.stream or not cb.has(p) or ca.read(p) == cb.read(p):
                    continue
                if p == "\x03EPRINT":
                    probs += verify_picture(ca.read(p), cb.read(p), old_m, new_m, f"埋め込み{n} の印刷用の表の絵（EPRINT）", stats)
                else:
                    probs.append(f"埋め込み{n}（{name}）：入れ物のほかのストリーム {p!r} が変わっています")
        for img in ea.get("images", []):
            allowed.add(img)
            da, db = a.zip.read(img), b.zip.read(img)
            if da != db:
                probs += verify_picture(da, db, old_m, new_m, f"埋め込み{n} の表の絵（{img}）", stats)
            elif changed_cells(old_m, new_m) and not is_icon(a, ea):
                probs.append(f"埋め込み{n} の表の絵（{img}）が、書いた値に合わせて描き直されていません")
    return allowed, probs, _note(stats)


def verify_doc(a, b, written: dict) -> tuple[set[str], list[str], str]:
    """.doc：書いた埋め込みの入れ物のストリームのうち、変わってよいものと、確かめた結果の問題。
    Data と WordDocument は、表示用の絵を足したことと、その位置の書き換えだけであることを確かめる。"""
    from .doc.book import sprms
    from .xls.proof import prove as prove_xls

    allowed, probs, stats = set(), [], {}
    moves: dict[int, int] = {}
    for n, cells in written_cells(a, written).items():
        ea = find_embed(a, n)
        eb = next((x for x in b.embeds if x["storage"] == ea["storage"]), None)
        if eb is None:
            probs.append(f"埋め込み{n}：書いた後の文書に、同じ埋め込みがありません")
            continue
        name = ea["storage"]
        ba, bb = open_book(a, ea), open_book(b, eb)
        allowed.add(f"{name}/{ba.stream}")
        if ba.how == "doc-xls":
            pr = prove_xls(ba.raw, bb.raw, cells)
        else:
            from .xlsx.proof import prove as prove_book

            pr = prove_book(ba.raw, bb.raw, allowed_inputs=cells)
        probs += [f"埋め込み{n}（{name}）：{x}" for x in pr.unexpected + pr.lint_errors]
        _count(stats, pr)
        if _chart_shown(ba.wb, bb.wb, a, ea, n, probs, stats):
            # 表の絵（グラフ）は変わらない：印刷用の絵は変わってよいものに入れず、表示用の絵（Data の PICF）も同じ所のまま
            pa = [loc for _cp, loc in ea.get("pictures", [])]
            pb = [loc for _cp, loc in eb.get("pictures", [])]
            if pa != pb or any(b.data[x:x + Picture(a.data, x).lcb] != a.data[x:x + Picture(a.data, x).lcb] for x in pa):
                probs.append(f"埋め込み{n}：グラフを表示している埋め込みの表示用の絵が変わっています")
            continue
        old_m, new_m = SheetModel(ba.wb), SheetModel(bb.wb)
        changed = bool(changed_cells(old_m, new_m)) and not is_icon(a, ea)
        ep = f"{name}/\x03EPRINT"
        if a.cfb.has(ep):
            allowed.add(ep)
            da, db = a.cfb.read(ep), b.cfb.read(ep)
            if da != db:
                probs += verify_picture(da, db, old_m, new_m, f"埋め込み{n} の印刷用の表の絵（EPRINT）", stats)
            elif changed:
                probs.append(f"埋め込み{n} の印刷用の表の絵（EPRINT）が、書いた値に合わせて描き直されていません")
        pa = [loc for _cp, loc in ea.get("pictures", [])]
        pb = [loc for _cp, loc in eb.get("pictures", [])]
        if len(pa) != len(pb):
            probs.append(f"埋め込み{n}：表示用の絵の数が変わっています")
            continue
        for la, lb in zip(pa, pb):
            if la == lb:
                if changed:
                    probs.append(f"埋め込み{n} の表示用の表の絵が、書いた値に合わせて描き直されていません")
                continue
            moves[la] = lb
            try:
                xa, xb = Picture(a.data, la), Picture(b.data, lb)
                if lb < len(a.data):
                    probs.append(f"埋め込み{n}：新しい表示用の絵が、Data の元の部分に重ねて書かれています")
                if b.data[la:la + xa.lcb] != a.data[la:la + xa.lcb]:
                    probs.append(f"埋め込み{n}：元の表示用の絵（PICF）が変わっています")
                if a.data[la + 4:xa.head_end] != b.data[lb + 4:xb.head_end] or len(xa.recs) != len(xb.recs) or \
                        any(a.data[q:q + 8 + ln] != b.data[q2:q2 + 8 + ln2]
                            for k, ((q, _v, _t, ln), (q2, _v2, _t2, ln2)) in enumerate(zip(xa.recs, xb.recs)) if k != xa.slot):
                    probs.append(f"埋め込み{n}：表示用の絵の、絵の中身のほか（見出し・図形）が変わっています")
                probs += verify_picture(xa.metafile(), xb.metafile(), old_m, new_m, f"埋め込み{n} の表示用の表の絵（Data）", stats)
                emf_b = xb.metafile()
                bpos = xb._blip_at()[0]
                if b.data[bpos + 8:bpos + 24] != md4(emf_b):
                    probs.append(f"埋め込み{n}：表示用の絵の識別子（MD4）が、絵の中身と合いません")
            except EmbedError as ex:
                probs.append(f"埋め込み{n}：{ex}")
    # 文字の書式：元の書式の並びが、絵の位置の値のほかは同じであること
    if moves:
        runs_b = {(r.fc0, r.fc1): r.grpprl for r in b.chpx}
        for r in a.chpx:
            g = runs_b.get((r.fc0, r.fc1))
            if g is None:
                probs.append("文字の書式（CHPX）の並びが変わっています")
                break
            if g != r.grpprl:
                pa_ = list(sprms(r.grpprl))
                pb_ = list(sprms(g))
                same = len(pa_) == len(pb_) and all(
                    s1 == s2 and (v1 == v2 or (s1 == 0x6A03 and len(v1) == 4 and moves.get(struct.unpack("<I", v1)[0])
                                               == struct.unpack("<I", v2)[0]))
                    for (s1, v1), (s2, v2) in zip(pa_, pb_))
                if not same:
                    probs.append("文字の書式（CHPX）が、表示用の絵の位置のほかにも変わっています")
                    break
    return allowed, probs, _note(stats)


# ====================================================================== 窓
def window_info(doc, e: dict, xdoc: dict) -> dict:
    """窓の埋め込みの節に出すもの：書くときの欄の名前（E1!Sheet1!C4）と、表の絵を描き直せるか。"""
    fields = []
    for sh in xdoc.get("sheets", []):
        for f in sh.get("fields", []):
            if not f.get("cell") or f.get("kind") == "checkbox":
                continue
            fields.append({"id": f"E{e['n']}!{f['id']}", "sheet": sh["name"], "cell": f["cell"], "kind": f.get("kind"),
                           "label": f.get("label", ""), "value": f.get("value")})
    k = _writable_key(doc, e)
    if k is not None and k in _WRITABLE:
        ok, why = _WRITABLE[k]
    else:
        ok, why = writable(doc, e)
        if k is not None:
            while len(_WRITABLE) >= 64:
                _WRITABLE.pop(next(iter(_WRITABLE)))
            _WRITABLE[k] = (ok, why)
    return {"fields": fields, "writable": ok, "why": why}


# 書けるかの判断は、埋め込みの入れ物と表の絵のバイト列だけで決まるので、同じ中身なら覚えておいたものを使う
_WRITABLE: dict[tuple, tuple[bool, str]] = {}


def _writable_key(doc, e: dict) -> tuple | None:
    import hashlib

    h = hashlib.sha256()
    try:
        if getattr(doc, "kind", "") == "doc":
            for path, data in sorted(SubStorage(doc.cfb, e["storage"]).contents().items()):
                h.update(path.encode("utf-8") + b"\0" + len(data).to_bytes(8, "little") + data)
            for _cp, loc in e.get("pictures", []):
                lcb = struct.unpack_from("<i", doc.data, loc)[0] if 0 <= loc <= len(doc.data) - 4 else 0
                h.update(b"\1" + doc.data[loc:loc + max(lcb, 0)])
        else:
            h.update(doc.zip.read(e["storage"]))
            for img in e.get("images", []):
                h.update(b"\1" + img.encode("utf-8") + b"\0" + doc.zip.read(img))
    except Exception:  # noqa: BLE001
        return None
    return (getattr(doc, "kind", "docx"), e["storage"], e.get("prog"), e.get("aspect"),
            tuple(loc for _cp, loc in e.get("pictures", [])), tuple(e.get("images", [])), h.hexdigest())


def writable(doc, e: dict) -> tuple[bool, str]:
    """この埋め込みに書いたとき、Word が表示・印刷に使う表の絵を描き直せるか（セルの位置が絵から決まるか）。"""
    try:
        if not str(e.get("prog", "")).startswith(("Excel.Sheet", "Excel.Chart")):
            return False, "Excel の表ではありません"
        book = open_book(doc, e)
        why = macro_block(book.wb)
        if why:
            return False, why
        if is_icon(doc, e):
            return True, ""  # アイコンで表示している：書いても絵は変わらない
        if chart_areas(book.wb) is not None:
            return True, "Word に表示しているのはグラフなので、グラフが使うセルの値が変わる書き込みはできません"
        model = SheetModel(book.wb)
        pics: list[tuple[str, bytes]] = []
        if getattr(doc, "kind", "") == "doc":
            sub = SubStorage(doc.cfb, e["storage"])
            names = sub.streams()
            if any(p.startswith(("\x03PRINT", "\x02OlePres")) for p in names):
                return False, "印刷用の絵（\\003PRINT・OlePres）が、描き直せない形です"
            if "\x03EPRINT" in names:
                pics.append(("印刷用の表の絵（EPRINT）", sub.read("\x03EPRINT")))
            for _cp, loc in e.get("pictures", []):
                pics.append(("表示用の表の絵（Data の PICF）", Picture(doc.data, loc).metafile()))
        else:
            part = doc.zip.read(e["storage"])
            if part[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
                c = CFB(part)
                if any(p.startswith(("\x03PRINT", "\x02OlePres")) for p in c.streams()):
                    return False, "入れ物の中の絵（\\003PRINT・OlePres）が、描き直せない形です"
                if c.has("\x03EPRINT"):
                    pics.append(("印刷用の表の絵（EPRINT）", c.read("\x03EPRINT")))
            pics += [(f"表の絵（{img}）", doc.zip.read(img)) for img in e.get("images", [])]
        if not pics:
            return False, "Word が表示・印刷に使う表の絵が見つかりません"
        for what, data in pics:
            _emf_or_fail(data, what)
            try:
                Layout(Emf(data), model)
            except (PictureError, MetafileError) as ex:
                return False, f"{what}：{ex}"
        return True, ""
    except (EmbedError, PictureError, MetafileError) as ex:
        return False, str(ex)
    except Exception as ex:  # noqa: BLE001
        return False, f"埋め込みを読めません（{ex}）"
