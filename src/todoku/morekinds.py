"""追加の形式（古い・役所が今も配っている形式）を、形式を問わない入口につなぐ：openkinds の各関数の先頭から呼ぶ。"""

from __future__ import annotations

from .openkinds_base import OpenFormatError

KIND_JA = {"xlsb": "Excel のバイナリのブック（.xlsb）"}
WORKBOOKS = frozenset({"xlsb"})  # xlsx と同じ窓・計算・表の取り出しを使うもの
# 読むだけの形式：形式 → legacy の部品（build・grade・audit・KIND_JA を持つ。窓は readonlydoc の形）
READONLY = {"ppt": "ppt", "xps": "xps", "epub": "epub"}
KIND_JA.update({"ppt": "古い PowerPoint（.ppt。読むだけ）", "xps": "XPS（.xps・.oxps。読むだけ）", "epub": "EPUB（.epub。読むだけ）"})
KINDS = frozenset(KIND_JA)


def _mod(k: str):
    import importlib

    return importlib.import_module(f".legacy.{READONLY[k]}", __package__)


def open_workbook(k: str, blob: bytes, name: str):
    if k == "xlsb":
        from .xlsb.book import XlsbWorkbook

        return XlsbWorkbook(blob, name)
    raise OpenFormatError(f"この形式（{k}）はブックとして開けません")


def window(k: str, blob: bytes, name: str):
    if k in WORKBOOKS:
        from .xlsx import window as xw

        doc = xw.build(open_workbook(k, blob, name))
        doc["kind"] = k
        doc["kind_ja"] = KIND_JA[k]
        if k == "xlsb":
            _mark_xlsb_fields(doc)
        return doc, xw.markdown_parts(doc)
    if k in READONLY:
        from . import readonlydoc as rd

        doc = _mod(k).build(blob, name)
        return doc, rd.markdown_parts(doc)
    raise OpenFormatError(f"この形式（{k}）は扱えません")


def _mark_xlsb_fields(doc: dict) -> None:
    from .xlsb.fill import CANT_GROW

    for sh in doc["sheets"]:
        for f in sh.get("fields", []):
            if f.get("value") in (None, ""):
                f["readonly"] = True
                f["why"] = CANT_GROW


def field_rows(k: str, doc: dict):
    if k in WORKBOOKS:
        for sh in doc["sheets"]:
            for f in sh.get("fields", []):
                yield f, {"sheet": sh["name"]}
        return
    for f in doc.get("fields", []):
        yield f, {}


def summary(k: str, doc: dict) -> dict:
    if k in WORKBOOKS:
        return {"sheets": [s["name"] for s in doc["sheets"]], "macros": bool(doc.get("macros"))}
    return {}


def fill(k: str, blob: bytes, name: str, values: dict, force: bool = False) -> dict:
    if k == "xlsb":
        from .xlsb.fill import XlsbWriteError, api_fill

        try:
            return api_fill(blob, name, values, force)
        except XlsbWriteError as e:
            raise OpenFormatError(str(e)) from None
    raise OpenFormatError(f"この形式（{KIND_JA.get(k, k)}）には書けません（読むだけの形式です）")


def read_back(k: str, blob: bytes, name: str, keys: dict) -> dict:
    if k == "xlsb":
        from .xlsb.fill import read_cells

        return read_cells(blob, name, keys)
    return {}


def audit_runner(k: str):
    if k == "xlsb":
        from .xlsb.audit import run

        return run
    if k in READONLY:
        return _mod(k).audit
    raise OpenFormatError(f"この形式（{k}）は検査できません")


def grade(k: str, blob: bytes, name: str):
    from .grade import Reason, _grade_excel

    if k == "xlsb":
        g = _grade_excel(blob, name)
        for r in g.reasons:
            if r.title == "広げていない範囲":
                r.title = "読んでいない部分"
        g.write_reasons = [Reason("D", ".xlsb の空のセルには書けない",
                                  "すでに値のあるセルの上書き（大きさの変わらないもの）だけ書けます。空のセルに書くとレコードの大きさが変わり、"
                                  "シートの部品の中の位置を指す索引がずれるので、書きません。本人か発行元に .xlsx で保存し直してもらえば書けます")]
        g.write = "D"
        return g
    if k in READONLY:
        return _mod(k).grade(blob, name)
    raise OpenFormatError(f"この形式（{k}）の段階は決められません")
