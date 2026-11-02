"""書類のファイルを開く：形式を中身（先頭のバイト）で見分けて、同じ形の読み取り結果を返す。"""

from __future__ import annotations


class FormatError(ValueError):
    pass


def kind_of(blob: bytes) -> str:
    if blob[:2] == b"PK":
        return "xlsx"
    if blob[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        from .cfb import CFB

        names = {n.lower() for n in CFB(blob).streams()}
        if "encryptedpackage" in names:
            return "encrypted"
        if "workbook" in names or "book" in names:
            return "xls"
        if "worddocument" in names:
            return "doc"
        return "cfb"
    if blob[:4] == b"%PDF":
        return "pdf"
    return "unknown"


def open_book(blob: bytes, filename: str = ""):
    k = kind_of(blob)
    if k == "xlsx":
        from .xlsx.workbook import Workbook

        return Workbook(blob, filename)
    if k == "xls":
        from .xls.book import XlsWorkbook

        return XlsWorkbook(blob, filename)
    if k == "encrypted":
        raise FormatError("パスワードで暗号化されたファイルです")
    raise FormatError(f"この形式（{k}）は、Excelのブックとして開けません")
