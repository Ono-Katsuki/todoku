"""大きな書類を分けて読み書きする。「大きいから読み減らす」はよいが、「読めない」は作らない。

入口は open_doc：書類の窓の全体の大きさ（字）を見積もり、limits.CONTEXT_BUDGET を超えるとき（または範囲を頼まれたとき）に、
形式ごとの LazyDoc を返す。小さな書類には None を返す（いままでどおり、窓の全体を作る）。
形式ごとの作り：
- CSV・TSV・JSON Lines・.txt（textdocs）：ファイルをメモリに読まない。行の位置の索引と、頼まれた範囲のバイトだけ
- .xlsx（xlsxdoc）：シートの XML を流し読みして行の位置を覚え、頼まれた行だけの小さなブックを作って、いままでの窓の読み手で読む
- PDF（pdfdoc）：ページの木と xref だけを読み、頼まれたページだけを配置する
- Access（jetdoc）：表ごとにページの表だけを作り、頼まれた行のあるページだけを解く
- ほか（builtdoc）：窓を作ってから、測った字の数で分ける（Word は段落の範囲でも読める）
"""

from __future__ import annotations

import os
import struct
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
        return "cfb"  # 古い形式・暗号化：lazy/legacy で見分ける
    if head[:4] == b"%PDF":
        return "pdf"
    from ..legacy.jet import is_jet

    if is_jet(head):
        return "mdb"
    if head[:2] == b"PK":
        return "pk"
    k = _text_kind(head, name)
    if k is None and name:
        from ..book import _known_kind, content_kind
        from .textstream import Source

        if _known_kind(head, name) == "unknown":  # 印でも名前でも決まらない：中身が字だけか（book.kind_of と同じ決まり）
            src = Source(path)
            try:
                k = content_kind(src)
            finally:
                src.close()
    return k


def open_doc(path: str, password: str | None = None, encoding: str | None = None, force: bool = False,
             built_ok: bool = True) -> LazyDoc | None:
    """分けて読む書類。force でなければ、窓の全体が予算に入らないときだけ返す（入るなら None：いままでどおり）。
    built_ok：窓を作って測る形式（BuiltDoc）も見るか（段階・書き込みは、その形式なら、いままでの道で全体を確かめる）。"""
    from ..api import PasswordNeeded, PasswordWrong, _member

    if not force:
        try:
            return _open_doc(path, password, encoding, False, built_ok, _member)
        except (PasswordNeeded, PasswordWrong):
            raise
        except Exception:  # noqa: BLE001  見分けられない・読めない書類は、いままでの道で、その理由を伝える
            return None
    return _open_doc(path, password, encoding, True, built_ok, _member)


def _open_doc(path, password, encoding, force, built_ok, _member) -> LazyDoc | None:
    if _member(path) is not None:
        from .jetdoc import jet_member

        jet = jet_member(path, password)
        if jet is not None:
            return jet if force or jet.large else None
        if not built_ok and not force:
            return None
        from .builtdoc import BuiltDoc

        doc = BuiltDoc(path, password, encoding)
        return doc if force or doc.large else None
    p = Path(path).expanduser()
    if not p.is_file():
        return None
    st = p.stat()
    if st.st_size == 0:
        return None
    # パスワードのある書類は控えない（パスワードを控えに残さない）
    from ..datafile.csvfile import chosen_delimiter

    key = (str(p), st.st_size, st.st_mtime_ns, encoding, chosen_delimiter()) if password is None else None
    doc = _CACHE.get(key) if key is not None else None
    if doc is None:
        if not force and not built_ok and _is_built(str(p), p.name):
            return None
        doc = _make(str(p), p.name, password, encoding)
        if doc is None:
            return None
        if key is not None:
            _CACHE[key] = doc
        while len(_CACHE) > _CACHE_MAX:
            old = _CACHE.pop(next(iter(_CACHE)))
            if hasattr(old, "close"):
                old.close()
    if not (force or doc.large):
        return None
    if not force and not built_ok and getattr(doc, "recoded", False):
        return None  # 切り替えの印のある文字コードの文字のファイル：書き込み・検査・段階は、全体を読む道で
    return doc


def _is_built(path: str, name: str) -> bool:
    k = _kind_for(path, name)
    if k in ("csv", "tsv", "jsonl", "txt", "pdf", "mdb"):
        return False
    if k == "pk":
        from .docxdoc import word_main
        from .parts import parts_kind
        from .xlsxdoc import package_kind

        return package_kind(path) != "xlsx" and word_main(path) is None and parts_kind(path) is None
    if k == "cfb":
        from .parts import cfb_kind

        return cfb_kind(path) is None
    if k == "json" or name.lower().endswith(".json"):
        from .parts import json_lazy_ok

        return not json_lazy_ok(path)
    return True


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
    if k == "mdb":
        from ..legacy.jet import JetError
        from .jetdoc import JetLazy

        try:
            return JetLazy(path, name, password)
        except (JetError, struct.error, IndexError):
            from .builtdoc import BuiltDoc  # 表の一覧が読めない：窓を作る道で、その理由を伝える

            return BuiltDoc(path, password, encoding)
    if k == "pk":
        from .xlsxdoc import package_kind

        if package_kind(path) == "xlsx":
            from .xlsxdoc import XlsxLazy

            return XlsxLazy(path, name)
        from .docxdoc import DocxLazy, word_main

        main = word_main(path)
        if main is not None:
            try:
                return DocxLazy(path, name, main)
            except ValueError:  # 分けて読めない作り：窓を作る道で読む（読めないなら、その理由を伝える）
                pass
    from .parts import make_lazy

    got = make_lazy(k, path, name, password)
    if got is not None:
        return got
    from .builtdoc import BuiltDoc

    return BuiltDoc(path, password, encoding)


def clear() -> None:
    """控えを捨てる。開いたファイルも閉じる（Windows は、開いたままのファイルを消すことも名前を変えることもできない）。"""
    for doc in list(_CACHE.values()):
        if hasattr(doc, "close"):
            try:
                doc.close()
            except Exception:  # noqa: BLE001  閉じられなくても、控えは捨てる
                pass
    _CACHE.clear()


_SHEET_REQ = ('sheet: "シート名", rows: "1-500"', "シートと行の範囲を分けて")
_PARA_REQ = ('paragraphs: "1-200"', "段落の範囲を分けて")


def smaller_request(tool: str, args: dict) -> str:
    """止まった呼び出しの、代わりに頼める小さな頼み方（行き止まりにしない）。"""
    path = str(args.get("path", ""))
    ext = os.path.splitext(path.split("!/")[-1])[1].lower()
    what = {".pdf": ('pages: "1-20"', "ページを分けて"), ".xlsx": _SHEET_REQ, ".xlsm": _SHEET_REQ, ".xls": _SHEET_REQ, ".xlsb": _SHEET_REQ, ".ods": _SHEET_REQ, ".mdb": _SHEET_REQ, ".accdb": _SHEET_REQ,
            ".csv": ('rows: "1-500"', "行の範囲を分けて"), ".tsv": ('rows: "1-500"', "行の範囲を分けて"),
            ".jsonl": ('records: "1-500"', "行の範囲を分けて"), ".ndjson": ('records: "1-500"', "行の範囲を分けて"),
            ".txt": ('paragraphs: "1-500"', "行の範囲を分けて"), ".docx": _PARA_REQ, ".docm": _PARA_REQ, ".doc": _PARA_REQ, ".jtd": _PARA_REQ, ".jtt": _PARA_REQ,
            ".hwp": _PARA_REQ, ".hwpx": _PARA_REQ, ".odt": _PARA_REQ, ".pptx": _PARA_REQ, ".rtf": _PARA_REQ,
            ".json": ('path: "/items/0"', "JSON Pointer で値を指して")}.get(ext, ('raw: true', "窓を作らずに、字のまま少しずつ"))
    if tool == "audit_document":
        return "検査は長い仕事として動かせます：audit_document に job: true を渡してください（1 回の呼び出しの時間に縛られず、止まっても控えの続きから進みます）"
    if tool == "fill_document":
        return (f"一度に書く所を減らして、何回かに分けて書いてください（書いたファイルを次の path に）。読むだけなら read_window に {what[0]} を渡せます。"
                "TODOKU_MEMORY_LIMIT_MB・TODOKU_TIME_LIMIT を大きくしても動きます")
    if tool == "view_page":
        return "ページの一部だけを頼んでください：view_page に region（[左, 上, 右, 下] の mm）を渡すと、その範囲の画像だけを開きます。dpi を下げても軽くなります"
    if tool == "read_table":
        return f"表の見出しまでの行を read_window に {what[0]} で読み、要る行の範囲を頼んでください"
    if what[0] != "raw: true" and ext not in (".csv", ".tsv", ".jsonl", ".ndjson", ".txt"):
        return (f"{what[1]}読んでください：read_window に {what[0]} を渡します。それでも窓を作るところで止まるときは、raw: true なら窓を作らずに"
                "字のまま少しずつ読めます（欄は見つけません）。TODOKU_MEMORY_LIMIT_MB・TODOKU_TIME_LIMIT を大きくすれば、窓も作れます")
    return (f"{what[1]}読んでください：read_window に {what[0]} を渡します（読む量を減らせば読めます）。"
            "止まらないようにするには TODOKU_MEMORY_LIMIT_MB・TODOKU_TIME_LIMIT も大きくできます")
