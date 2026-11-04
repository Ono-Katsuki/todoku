"""大きな書類を分けて読み書きする。「大きいから読み減らす」はよいが、「読めない」は作らない。

入口は open_doc：書類の窓の全体の大きさ（字）を見積もり、limits.CONTEXT_BUDGET を超えるとき（または範囲を頼まれたとき）に、
形式ごとの LazyDoc を返す。小さな書類には None を返す（いままでどおり、窓の全体を作る）。
形式ごとの作り：
- CSV・TSV・JSON Lines・.txt（textdocs）：ファイルをメモリに読まない。行の位置の索引と、頼まれた範囲のバイトだけ
- .xlsx（xlsxdoc）：シートの XML を流し読みして行の位置を覚え、頼まれた行だけの小さなブックを作って、いままでの窓の読み手で読む
- PDF（pdfdoc）：ページの木と xref だけを読み、頼まれたページだけを配置する
- ほか（builtdoc）：窓を作ってから、測った字の数で分ける（Word は段落の範囲でも読める）
"""

from __future__ import annotations

import os
from pathlib import Path

from .base import LazyDoc, Page, RangeError, clamp_chars, parse_range  # noqa: F401

_CACHE: dict = {}
_CACHE_MAX = 4


def _kind_for(path: str, name: str) -> str | None:
    """丸ごと読まずに、形式を見分ける（頭のバイトと、ZIP の部品の名前。見分けられなければ None）。"""
    from ..book import _text_kind

    with open(path, "rb") as f:
        head = f.read(4096)
    if head[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return None  # 古い形式・暗号化：丸ごと読む（api.window）
    if head[:4] == b"%PDF":
        return "pdf"
    if head[:2] == b"PK":
        return "pk"
    return _text_kind(head, name)


def open_doc(path: str, password: str | None = None, encoding: str | None = None, force: bool = False) -> LazyDoc | None:
    """分けて読む書類。force でなければ、窓の全体が予算に入らないときだけ返す（入るなら None：いままでどおり）。"""
    from ..api import _member

    if _member(path) is not None:
        from .builtdoc import BuiltDoc

        doc = BuiltDoc(path, password, encoding)
        return doc if force or doc.large else None
    p = Path(path).expanduser()
    if not p.is_file():
        return None
    st = p.stat()
    if st.st_size == 0:
        return None
    key = (str(p), st.st_size, st.st_mtime_ns, encoding, password)
    doc = _CACHE.get(key)
    if doc is None:
        doc = _make(str(p), p.name, password, encoding)
        if doc is None:
            return None
        _CACHE[key] = doc
        while len(_CACHE) > _CACHE_MAX:
            old = _CACHE.pop(next(iter(_CACHE)))
            if hasattr(old, "close"):
                old.close()
    return doc if force or doc.large else None


def _make(path: str, name: str, password: str | None, encoding: str | None) -> LazyDoc | None:
    k = _kind_for(path, name)
    if k in ("csv", "tsv"):
        from .textdocs import CsvDoc

        return CsvDoc(path, name, k, encoding)
    if k == "jsonl":
        from .textdocs import JsonlDoc

        return JsonlDoc(path, name, k, encoding)
    if k == "txt":
        from .textdocs import TxtLazy

        return TxtLazy(path, name, k, encoding)
    if k == "pdf":
        from .pdfdoc import PdfLazy

        return PdfLazy(path, name, password)
    if k == "pk":
        from .xlsxdoc import package_kind

        if package_kind(path) == "xlsx":
            from .xlsxdoc import XlsxLazy

            return XlsxLazy(path, name)
    from .builtdoc import BuiltDoc

    return BuiltDoc(path, password, encoding)


def clear() -> None:
    _CACHE.clear()


def smaller_request(tool: str, args: dict) -> str:
    """止まった呼び出しの、代わりに頼める小さな頼み方（行き止まりにしない）。"""
    path = str(args.get("path", ""))
    ext = os.path.splitext(path.split("!/")[-1])[1].lower()
    what = {".pdf": ('pages: "1-20"', "ページを分けて"), ".xlsx": ('sheet: "シート名", rows: "1-500"', "シートと行の範囲を分けて"),
            ".xlsm": ('sheet: "シート名", rows: "1-500"', "シートと行の範囲を分けて"),
            ".csv": ('rows: "1-500"', "行の範囲を分けて"), ".tsv": ('rows: "1-500"', "行の範囲を分けて"),
            ".jsonl": ('records: "1-500"', "行の範囲を分けて"), ".ndjson": ('records: "1-500"', "行の範囲を分けて"),
            ".txt": ('paragraphs: "1-500"', "行の範囲を分けて"), ".docx": ('paragraphs: "1-200"', "段落の範囲を分けて"),
            ".docm": ('paragraphs: "1-200"', "段落の範囲を分けて")}.get(ext, ('max_chars: 20000', "小さく分けて"))
    if tool == "audit_document":
        return "検査は audit_document の長い仕事（job）として動かせます。もう一度 audit_document を呼ぶと、控えから続けます"
    if tool == "fill_document":
        return (f"一度に書く所を減らして、何回かに分けて書いてください（書いたファイルを次の path に）。読むだけなら read_window に {what[0]} を渡せます。"
                "TODOKU_MEMORY_LIMIT_MB・TODOKU_TIME_LIMIT を大きくしても動きます")
    if tool == "read_table":
        return f"表の見出しまでの行を read_window に {what[0]} で読み、要る行の範囲を頼んでください"
    return (f"{what[1]}読んでください：read_window に {what[0]} を渡します（読む量を減らせば読めます）。"
            "止まらないようにするには TODOKU_MEMORY_LIMIT_MB・TODOKU_TIME_LIMIT も大きくできます")
