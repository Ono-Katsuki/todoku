"""窓を作らずに、ファイルの字を順に読む（read_window の raw）：窓を作るのが上限に当たる大きな書類でも、中身に届くための最後の道。

- ZIP の書類（Word・PowerPoint・OpenDocument・HWPX・Excel）：XML の部品を順に展開して、タグを除いた字を出す（メモリは一定）
- 文字のファイル（JSON・RTF・HTML・1 つの XML の OpenDocument）：字のまま（文字コードは textstream で決める）
入力欄は見つけない（窓ではない）。続きは cursor。どの部品のどこを読んでいるかを示す。
"""

from __future__ import annotations

import codecs
import re
import zipfile

from .base import Page, RangeError, comma
from .textstream import Source, detect

_TAG = re.compile(r"<[^>]*>")
_BLOCK_END = re.compile(r"</(?:\w+:)?(?:p|tr|row|si|text-p|h)\s*>")


def _xml_names(z: zipfile.ZipFile) -> list[str]:
    return [i.filename for i in z.infolist() if i.filename.lower().endswith((".xml", ".rels")) is True and not i.filename.endswith(".rels")
            and not i.filename.startswith(("_rels/", "[Content_Types]"))]


def _plain(xml_piece: str) -> str:
    s = _BLOCK_END.sub("\n", xml_piece)
    s = _TAG.sub("", s)
    return s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&")


def read_raw(path: str, cursor: str | None, max_chars: int, encoding: str | None = None) -> Page:
    with open(path, "rb") as f:
        head = f.read(8)
    if head == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return _read_items(_cfb_items(path), cursor, max_chars)
    if head[:2] == b"PK":
        with zipfile.ZipFile(path) as z:
            xlsb = "xl/workbook.bin" in z.namelist()
        if xlsb:
            return _read_items(_xlsb_items(path), cursor, max_chars)
        return _read_zip(path, cursor, max_chars)
    return _read_text(path, cursor, max_chars, encoding)


# ---------------------------------------------------------------------- 字の並びで読む（複合ファイルの書類・.xlsb）
def _read_items(items, cursor: str | None, max_chars: int) -> Page:
    """items(いまの位置) は (次の位置, 字) を順に返す。位置は字（数字と : ）。1 つの字が大きいときは、字の位置で分ける。"""
    state = None
    skip = 0
    if cursor:
        m = re.fullmatch(r"raw:([\w:.-]*):s(\d+)", str(cursor))
        if not m:
            raise RangeError(f"cursor「{cursor}」はこの書類の続きの印ではありません")
        state, skip = m[1] or None, int(m[2])
    out: list[str] = []
    size = 0
    for here, text in items(state):
        if skip:
            text = text[skip:]
        if len(text) > max_chars - size:
            if out:
                return Page("\n".join(out), "字のまま（窓ではなく。ここまで）", {"raw": True, "cursor": f"raw:{here}:s{skip}"}, "残りの字", out)
            out.append(text[:max_chars])
            return Page("\n".join(out), "字のまま（窓ではなく。長い字を分けています）", {"raw": True, "cursor": f"raw:{here}:s{skip + max_chars}"}, "この字の残りと、後ろの字", out)
        skip = 0
        out.append(text)
        size += len(text) + 1
    return Page("\n".join(out), "字のまま（窓ではなく。終わりまで）", None, "", out)


def _sheet_items(meta, book_of):
    """シートのセルを、2048 行ごとに読む。位置は シート:行。"""

    def items(state):
        from ..formula.parser import addr

        si, row = (int(x) for x in state.split(":")) if state else (0, 0)
        for s in range(si, len(meta)):
            d = meta[s]
            ix = d["index"]
            if ix is None or not ix.rows:
                continue
            k = 0 if s > si or not state else __import__("bisect").bisect_left(ix.rows, row)
            if s > si or not state:
                yield f"{s}:0", f"〔シート {d['name']}〕"
            while k < len(ix.rows):
                e = min(len(ix.rows), k + 2048)
                ra, rb = ix.rows[k], ix.rows[e - 1]
                sh = book_of({s: [(ra, rb)]}).sheets[s]
                by_row: dict[int, list[str]] = {}
                for (r, c), cell in sorted(sh.cells.items()):
                    if cell.value is not None and cell.value != "":
                        by_row.setdefault(r, []).append(f"{addr(r, c)}＝{cell.value}")
                for r in sorted(by_row):
                    yield f"{s}:{r}", "\t".join(by_row[r])
                k = e

    return items


def _xlsb_items(path: str):
    import os

    from .sheetdocs import XlsbLazy

    doc = XlsbLazy(path, os.path.basename(path))
    return _sheet_items(doc._meta, doc._book)


def _hwp_items(cf, name: str):
    from ..hwpx.hwp5 import PARA_TEXT
    from .seqdocs import HwpLazy, _StreamMember
    from .xlsxdoc import SeekableInflate

    doc = HwpLazy.__new__(HwpLazy)  # 本文を読まないで、節と圧縮の有無だけを見る
    from ..hwpx.hwp5 import Hwp5Document

    skel = Hwp5Document(cf.fb, name, partial={"cfb": cf.cfb, "records": None})
    del doc

    def items(state):
        p0, at0 = (int(x) for x in state.split(":")) if state else (0, 0)
        for pi in range(p0, len(skel.parts)):
            part = skel.parts[pi]
            reader = SeekableInflate(_StreamMember(cf.stream(part), skel.compressed))
            carry = bytearray()
            at = 0
            pos = at0 if pi == p0 else 0
            for c in reader.scan():
                carry += c
                while True:
                    k = pos - at
                    if k < 0 or k + 4 > len(carry):
                        break
                    h = int.from_bytes(carry[k:k + 4], "little")
                    size = h >> 20
                    head = 4
                    if size == 0xFFF:
                        if k + 8 > len(carry):
                            break
                        size = int.from_bytes(carry[k + 4:k + 8], "little")
                        head = 8
                    if k + head + size > len(carry):
                        break
                    if h & 0x3FF == PARA_TEXT:
                        t = _hwp_text(bytes(carry[k + head:k + head + size]))
                        if t.strip():
                            yield f"{pi}:{pos}", t
                    pos += head + size
                drop = pos - at
                if drop > 0:
                    del carry[:drop]
                    at = pos
            at0 = 0

    return items


def _hwp_text(b: bytes) -> str:
    """PARA_TEXT の字：制御の字を除く（字の制御は 1 字、ほかは 8 字）。"""
    from ..hwpx.hwp5 import CHAR_CTRL

    out = []
    i = 0
    n = len(b) // 2
    while i < n:
        c = b[2 * i] | (b[2 * i + 1] << 8)
        if c < 32:
            i += 1 if c in CHAR_CTRL else 8
            if c in (10, 13):
                out.append("\n")
            continue
        out.append(chr(c))
        i += 1
    return "".join(out)


def _doc_items(cf, name: str):
    from ..doc.book import DocFile
    from .cfbfile import SpoolCfb

    d = DocFile(cf.fb, name, cfb=SpoolCfb(cf), partial={"range": ("body", 0, 0), "no_text": True})

    def items(state):
        p0, at0 = (int(x) for x in state.split(":")) if state else (0, 0)
        for pi in range(p0, len(d.pieces)):
            pc = d.pieces[pi]
            n = pc.cp1 - pc.cp0
            width = 1 if pc.compressed else 2
            k = at0 if pi == p0 else 0
            while k < n:
                e = min(n, k + (1 << 14))
                raw = d.wd[pc.fc + k * width:pc.fc + e * width]
                t = raw.decode("cp1252" if pc.compressed else "utf-16-le", "replace")
                t = re.sub("[\\x13\\x14\\x15\\x01\\x08]", "", t).replace("\r", "\n").replace("\x07", "\t")
                if t:
                    yield f"{pi}:{k}", t
                k = e
            at0 = 0

    return items


def _parsed_items(doc):
    """一太郎・Word 6.0/95：読んだ段落の字。位置は段落の番号。"""
    from ..docx.document import iter_paras

    paras = [p for _k, _part, bl in doc.all_blocks() for p in iter_paras(bl)]

    def items(state):
        for i in range(int(state) if state else 0, len(paras)):
            t = paras[i].text
            if t.strip():
                yield str(i), t

    return items


def _cfb_items(path: str):
    import os

    from ..cfb import CfbError
    from .cfbfile import CfbFile
    from .parts import _kind_of_cfb

    name = os.path.basename(path)
    try:
        cf = CfbFile(path)
        limits_kind = _kind_of_cfb(cf)
    except (CfbError, ValueError) as e:
        raise RangeError(f"字のまま読めません：{e}") from None
    if limits_kind == "xls":
        from .. import limits
        from .xlsdoc import XlsLazy

        keep = limits.LAZY_MIN_BYTES
        doc = XlsLazy(path, name, cf, cf.find("Workbook") or cf.find("Book"))
        del keep
        return _sheet_items(doc.sheet_meta, doc.book_of)
    if limits_kind == "doc":
        return _doc_items(cf, name)
    if limits_kind == "hwp":
        return _hwp_items(cf, name)
    if limits_kind == "jtd":
        from ..jtd.book import JtdFile

        return _parsed_items(JtdFile(cf.fb, name))
    if limits_kind == "word6":
        from ..doc.word6 import Word6File

        return _parsed_items(Word6File(cf.fb, name))
    raise RangeError("この複合ファイルは、パスワードで保護されているか、字のまま読める形ではありません（open_document で理由が分かります）")


def _read_zip(path: str, cursor: str | None, max_chars: int) -> Page:
    from .. import limits
    from .textstream import CHUNK  # noqa: F401
    from .xlsxdoc import Member

    part, skip = 0, 0
    if cursor:
        m = re.fullmatch(r"raw:(\d+):(\d+)", str(cursor))
        if not m:
            raise RangeError(f"cursor「{cursor}」はこの書類の続きの印ではありません")
        part, skip = int(m[1]), int(m[2])
    src = Source(path)
    try:
        with zipfile.ZipFile(path) as z:
            names = _xml_names(z)
            out: list[str] = []
            size = 0
            start = (part, skip)
            while part < len(names):
                m_ = Member(src, z.getinfo(names[part]))
                dec = codecs.getincrementaldecoder("utf-8")("replace")
                seen = 0  # この部品で、ここまでに出した字
                carry = ""
                piece_all: list[str] = []
                done = False
                for c in m_.chunks():
                    carry += dec.decode(c)
                    k = carry.rfind(">")
                    if k < 0:
                        continue
                    text = _plain(carry[:k + 1])
                    carry = carry[k + 1:]
                    if seen + len(text) <= skip:
                        seen += len(text)
                        continue
                    text = text[max(0, skip - seen):] if skip > seen else text
                    seen = max(seen, skip)
                    room = max_chars - size
                    if len(text) > room:
                        piece_all.append(text[:room])
                        seen += room
                        size += room
                        done = True
                        break
                    piece_all.append(text)
                    size += len(text)
                    seen += len(text)
                    limits.progress(f"{names[part]} の字を読んでいます")
                got = "".join(piece_all)
                if got.strip():
                    out.append((f"〔部品 {names[part]}〕\n" if not skip else "") + got)
                if done:
                    return Page("".join(out), f"部品 {start[0] + 1}〜{part + 1}／全 {len(names)} 部品（窓ではなく字のまま）",
                                {"raw": True, "cursor": f"raw:{part}:{seen}"}, f"部品 {comma(len(names) - part)} つの残り", out)
                part += 1
                skip = 0
            return Page("".join(out), f"部品 {start[0] + 1}〜{len(names)}／全 {len(names)} 部品（窓ではなく字のまま）", None, "", out)
    finally:
        src.close()


def _read_text(path: str, cursor: str | None, max_chars: int, encoding: str | None) -> Page:
    src = Source(path)
    try:
        dec_info = detect(src, encoding)
        at = len(dec_info.bom)
        chars = 0
        if cursor:
            m = re.fullmatch(r"raw:(\d+):(\d+)", str(cursor))
            if not m:
                raise RangeError(f"cursor「{cursor}」はこの書類の続きの印ではありません")
            at, chars = int(m[1]), int(m[2])
        dec = codecs.getincrementaldecoder(dec_info.encoding)("strict")
        out: list[str] = []
        size = 0
        pos = at
        for c in src.chunks(at, size=1 << 16):
            # 字の境目で止められるよう、少しずつ。読みすぎた分は捨てて、次は続きの位置から読み直す
            t = dec.decode(c)
            pos += len(c)
            if size + len(t) >= max_chars:
                keep = max_chars - size
                out.append(t[:keep])
                size += keep
                rest = t[keep:]
                pending = dec.getstate()[0]
                back = len(rest.encode(dec_info.encoding)) + len(pending)
                return Page("".join(out), f"{chars + 1:,}〜{chars + size:,} 字（窓ではなく字のまま）", {"raw": True, "cursor": f"raw:{pos - back}:{chars + size}"},
                            f"{comma(src.size - (pos - back))} バイト", ["".join(out)])
            out.append(t)
            size += len(t)
        out.append(dec.decode(b"", True))
        text = "".join(out)
        return Page(text, f"{chars + 1:,}〜{chars + len(text):,} 字（窓ではなく字のまま）", None, "", [text])
    finally:
        src.close()
