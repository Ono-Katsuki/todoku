"""形式を見分けて、形式ごとの分けて読む道を選ぶ（小さな書類・暗号の書類・見分けられない書類は None：窓を作ってから分ける）。"""

from __future__ import annotations

import os
import re
import struct
import zipfile

from .. import limits

ODF_ODT = b"application/vnd.oasis.opendocument.text"
ODF_ODS = b"application/vnd.oasis.opendocument.spreadsheet"


def parts_kind(path: str) -> str | None:
    """ZIP の書類の、分けて読む形式（pptx・odt・ods・hwpx・xlsb）。小さなものは None。"""
    try:
        with zipfile.ZipFile(path) as z:
            infos = z.infolist()
            names = {i.filename for i in infos}
            if len({n.lower() for n in names}) != len(infos) or any(i.flag_bits & 1 for i in infos):
                return None
            size = sum(i.file_size for i in infos if i.filename.lower().endswith((".xml", ".bin")))
            if size < limits.LAZY_MIN_BYTES:
                return None
            if "xl/workbook.bin" in names:
                return "xlsb"
            if any(n.startswith("ppt/slides/") for n in names) and "ppt/presentation.xml" in names:
                return "pptx"
            if "mimetype" in names:
                mt = z.read("mimetype").strip()
                if mt == b"application/hwp+zip":
                    return "hwpx"
                if "content.xml" in names and mt == ODF_ODT:
                    return "odt"
                if "content.xml" in names and mt == ODF_ODS:
                    return "ods"
    except (zipfile.BadZipFile, OSError, ValueError, KeyError):
        return None
    return None


def cfb_kind(path: str) -> str | None:
    """複合ファイルの書類の、分けて読む形式（xls・doc・word6・hwp・jtd）。暗号・小さなファイルは None。"""
    if os.path.getsize(path) < limits.LAZY_MIN_BYTES // 16:  # 圧縮した本文（HWP・一太郎）は小さなファイルでも大きい
        return None
    from ..cfb import CfbError
    from .cfbfile import CfbFile

    try:
        cf = CfbFile(path)
    except (CfbError, OSError, ValueError, struct.error):
        return None
    try:
        got = _kind_of_cfb(cf)
        if got not in ("hwp", "jtd") and cf.size < limits.LAZY_MIN_BYTES:
            return None
        return got
    except (CfbError, OSError, ValueError, struct.error, IndexError):
        return None
    finally:
        cf.close()


def _kind_of_cfb(cf) -> str | None:
    if cf.find("EncryptedPackage"):
        return None
    wb = cf.find("Workbook") or cf.find("Book")
    if wb:
        head = cf.stream(wb).read(0, 1 << 12)
        pos = 0
        while pos + 4 <= len(head):
            t, ln = struct.unpack_from("<HH", head, pos)
            if t == 0x002F:
                return None  # パスワードで保護されたブック
            if t == 0x0085:
                break
            pos += 4 + ln
        return "xls"
    if cf.find("WordDocument"):
        wd = cf.stream(cf.find("WordDocument")).read(0, 64)
        ident, nfib = struct.unpack_from("<HH", wd, 0)
        flags = struct.unpack_from("<H", wd, 0x0A)[0]
        if flags & 0x0100:
            return None  # パスワードで暗号化された文書
        if ident == 0xA5EC and nfib >= 105:
            return "doc"
        return "word6"
    if cf.find("FileHeader"):
        head = cf.stream(cf.find("FileHeader")).read(0, 64)
        if head[:17] == b"HWP Document File" and not struct.unpack_from("<I", head, 36)[0] & 0x16:
            return "hwp"
        return None
    if cf.find("DocumentText") or cf.find("JSCompDocument"):
        return "jtd"
    return None


def json_lazy_ok(path: str) -> bool:
    if os.path.getsize(path) < limits.LAZY_MIN_BYTES:
        return False
    return True


def make_lazy(k: str | None, path: str, name: str, password: str | None):
    """分けて読む書類（作れなければ None：窓を作ってから分ける）。"""
    try:
        if k == "pk":
            kind = parts_kind(path)
            if kind == "pptx":
                from .seqdocs import PptxLazy

                return PptxLazy(path, name)
            if kind == "odt":
                from .seqdocs import OdtLazy

                return OdtLazy(path, name)
            if kind == "hwpx":
                from .seqdocs import HwpxLazy

                return HwpxLazy(path, name)
            if kind == "ods":
                from .sheetdocs import OdsLazy

                return OdsLazy(path, name)
            if kind == "xlsb":
                from .sheetdocs import XlsbLazy

                return XlsbLazy(path, name)
            return None
        if k == "cfb":
            return _make_cfb(path, name, password)
        if k == "rtf" and os.path.getsize(path) >= limits.LAZY_MIN_BYTES:
            return _make_rtf(path, name)
        if k == "json" and json_lazy_ok(path):
            from .jsondoc import JsonLazy

            return JsonLazy(path, name)
    except limits.LimitExceeded:
        raise
    except (ValueError, KeyError, IndexError, struct.error, OSError):
        return None  # 分けて読めない作り：窓を作る道で読み、読めないなら、その理由を伝える
    return None


def _make_cfb(path: str, name: str, password: str | None):
    from .cfbfile import CfbFile

    kind = cfb_kind(path)
    if kind is None:
        return None
    cf = CfbFile(path)
    if kind == "xls":
        from .xlsdoc import XlsLazy

        return XlsLazy(path, name, cf, cf.find("Workbook") or cf.find("Book"), password)
    if kind == "doc":
        from .seqdocs import DocLazy

        return DocLazy(path, name, cf)
    if kind == "hwp":
        from .seqdocs import HwpLazy

        return HwpLazy(path, name, cf)
    from .blocks import ViewLazy

    if kind == "jtd":
        from ..jtd.book import JtdFile
        from ..jtd.window import build, markdown_parts

        return ViewLazy(path, name, "jtd", JtdFile(cf.fb, name), build, markdown_parts, cf.size, password)
    if kind == "word6":
        from ..doc.word6 import Word6File, build, is_word6
        from ..docx.window import markdown_parts

        if not is_word6(cf.fb):
            return None
        return ViewLazy(path, name, "doc", Word6File(cf.fb, name), build, markdown_parts, cf.size, password)
    return None


def _make_rtf(path: str, name: str):
    from ..rtf.document import RtfDocument
    from ..rtf.window import build, markdown_parts
    from .blocks import ViewLazy
    from .filebytes import FileBytes

    fb = FileBytes(path)
    return ViewLazy(path, name, "rtf", RtfDocument(fb, name), build, markdown_parts, len(fb))


del re
