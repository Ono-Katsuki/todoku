"""追加の形式（古い・役所が今も配っている形式）を、形式を問わない入口につなぐ：openkinds の各関数の先頭から呼ぶ。"""

from __future__ import annotations

from .openkinds_base import OpenFormatError

KIND_JA = {"xlsb": "Excel のバイナリのブック（.xlsb）"}
KIND_JA.update({"xls2": "古い Excel（Excel 2〜4 の .xls。読むだけ）", "wk1": "Lotus 1-2-3（.wk1。読むだけ）", "dbf": "dBase（.dbf。読むだけ）",
                "mdb": "Access のデータベース（.mdb・.accdb。読むだけ）"})
WORKBOOKS = frozenset({"xlsb", "xls2", "wk1", "dbf", "mdb"})  # xlsx と同じ窓・計算・表の取り出しを使うもの
# 読むだけの形式：形式 → legacy の部品（build・grade・audit・KIND_JA を持つ。窓は readonlydoc の形）
READONLY = {"ppt": "ppt", "xps": "xps", "epub": "epub", "word2": "word2", "iwork": "iwork"}
KIND_JA.update({"ppt": "古い PowerPoint（.ppt。読むだけ）", "xps": "XPS（.xps・.oxps。読むだけ）", "epub": "EPUB（.epub。読むだけ）",
                "word2": "古い Word（Word 1.x・2.0 の .doc。読むだけ）", "iwork": "Apple の iWork（.pages・.numbers・.key。読むだけ）"})
SXO = frozenset({"sxw", "sxc"})  # OpenOffice.org 1.x：OpenDocument に写して読む
KIND_JA.update({"sxw": "OpenOffice.org 1.x の文書（.sxw。読むだけ）", "sxc": "OpenOffice.org 1.x の表計算（.sxc。読むだけ）"})
KINDS = frozenset(KIND_JA)


def _mod(k: str):
    import importlib

    return importlib.import_module(f".legacy.{READONLY[k]}", __package__)


def open_workbook(k: str, blob: bytes, name: str):
    if k == "xlsb":
        from .xlsb.book import XlsbWorkbook

        return XlsbWorkbook(blob, name)
    if k == "xls2":
        from .legacy.biff234 import Biff234Workbook

        return Biff234Workbook(blob, name)
    if k == "wk1":
        from .legacy.lotus import Wk1Workbook

        return Wk1Workbook(blob, name)
    if k == "dbf":
        from .legacy.dbf import DbfWorkbook

        return DbfWorkbook(blob, name)
    if k == "mdb":
        from .legacy.jet import JetWorkbook

        return JetWorkbook(blob, name)
    raise OpenFormatError(f"この形式（{k}）はブックとして開けません")


def _write_note(k: str) -> str:
    return ("すでに値のあるセルの上書き（レコードの大きさの変わらないもの）だけ書けます。空のセルには書けません（.xlsx で保存し直してもらえば書けます）"
            if k == "xlsb" else WRITE_NOTE[k])


def _raw_values(k: str, wb) -> dict:
    """窓を作る処理とは別に数えた、値のあるセル（検査用）。"""
    if k == "xlsb":
        from .xlsb.audit import raw_values
    elif k == "xls2":
        from .legacy.biff234 import raw_values
    elif k == "wk1":
        from .legacy.lotus import raw_values
    elif k == "dbf":
        from .legacy.dbf import raw_values
    else:
        from .legacy.jet import raw_values
    return raw_values(wb)


def reading_check(k: str, wb, doc: dict):
    """窓の読み落とし：ファイルのレコードから別に数えた、値のあるセルが、すべて窓のセルにあるか。"""
    from .xlsb.audit import cells_check

    return cells_check(doc, _raw_values(k, wb))


def window(k: str, blob: bytes, name: str):
    if k in WORKBOOKS:
        from .xlsx import window as xw

        doc = xw.build(open_workbook(k, blob, name))
        doc["kind"] = k
        doc["kind_ja"] = KIND_JA[k]
        if k == "xlsb":
            _mark_xlsb_fields(doc)
        else:
            _mark_readonly(doc, k)
        return doc, xw.markdown_parts(doc)
    if k in SXO:
        return _sxo_window(k, blob, name)
    if k in READONLY:
        from . import readonlydoc as rd

        doc = _mod(k).build(blob, name)
        return doc, rd.markdown_parts(doc)
    raise OpenFormatError(f"この形式（{k}）は扱えません")


WRITE_NOTE = {"xls2": "Excel 2〜4 の形には書きません。Excel で .xlsx に保存し直してもらえば書けます",
              "wk1": "Lotus 1-2-3 の形には書きません。.xlsx か CSV に書き出してもらえば書けます",
              "dbf": "dBase の形には書きません。CSV か .xlsx に書き出してもらえば書けます",
              "mdb": "Access のデータベースには書きません。表を .xlsx か CSV に書き出してもらえば書けます"}


def _mark_readonly(doc: dict, k: str) -> None:
    for sh in doc["sheets"]:
        for f in sh.get("fields", []):
            f["readonly"] = True
            f["why"] = WRITE_NOTE[k]


def _sxo_window(k: str, blob: bytes, name: str):
    from .legacy.sxo import WRITE, as_odf

    odf, _k = as_odf(blob)
    if k == "sxw":
        from .odf.text import OdtDocument
        from .odf.textwindow import build, markdown_parts

        doc = build(OdtDocument(odf, name))
        parts = markdown_parts(doc)
    else:
        from .odf.sheetwindow import build, markdown_parts

        doc = build(odf, name)
        parts = markdown_parts(doc)
    doc["kind"] = k
    doc["kind_ja"] = KIND_JA[k]
    doc["write_note"] = WRITE
    fields = doc.get("fields", []) if k == "sxw" else [f for sh in doc["sheets"] for f in sh.get("fields", [])]
    for f in fields:
        f["readonly"] = True
        f["why"] = WRITE
    return doc, parts


def _mark_xlsb_fields(doc: dict) -> None:
    from .xlsb.fill import CANT_GROW

    for sh in doc["sheets"]:
        for f in sh.get("fields", []):
            if f.get("value") in (None, ""):
                f["readonly"] = True
                f["why"] = CANT_GROW


def field_rows(k: str, doc: dict):
    if k in WORKBOOKS or k == "sxc":
        for sh in doc["sheets"]:
            for f in sh.get("fields", []):
                yield f, {"sheet": sh["name"]}
        return
    for f in doc.get("fields", []):
        yield f, {}


def summary(k: str, doc: dict) -> dict:
    if k in WORKBOOKS or k == "sxc":
        return {"sheets": [s["name"] for s in doc["sheets"]], "macros": bool(doc.get("macros"))}
    return {}


def fill(k: str, blob: bytes, name: str, values: dict, force: bool = False) -> dict:
    if k == "xlsb":
        from .xlsb.fill import XlsbWriteError, api_fill

        try:
            return api_fill(blob, name, values, force)
        except XlsbWriteError as e:
            raise OpenFormatError(str(e)) from None
    if k in SXO:
        from .legacy.sxo import WRITE

        raise OpenFormatError(WRITE)
    raise OpenFormatError(WRITE_NOTE.get(k) or _mod(k).WRITE)


def read_back(k: str, blob: bytes, name: str, keys: dict) -> dict:
    if k == "xlsb":
        from .xlsb.fill import read_cells

        return read_cells(blob, name, keys)
    return {}


def audit_runner(k: str):
    if k in WORKBOOKS:
        from .xlsb.audit import workbook_audit

        return lambda blob, name: workbook_audit(k, open_workbook(k, blob, name), _raw_values, _write_note(k))
    if k in SXO:
        return lambda blob, name: _sxo_audit(k, blob, name)
    if k in READONLY:
        return _mod(k).audit
    raise OpenFormatError(f"この形式（{k}）は検査できません")


def _sxo_audit(k: str, blob: bytes, name: str):
    from . import readonlydoc as rd
    from .legacy.sxo import WRITE, raw_text
    from .xlsx.audit import Audit, Check

    _doc, parts = _sxo_window(k, blob, name)
    au = Audit(name)
    shown = rd.chars("".join(t for _n, t in parts))
    raw = rd.chars(raw_text(blob))
    if k == "sxc":  # 表計算の数は、窓では元の値の形で出るので、数字と区切りの字は比べない
        for ch in "0123456789,.%-+¥$":
            raw.pop(ch, None)
    missing = raw - shown
    au.checks.append(Check("読み落とし（字）", not missing, f"content.xml の段落の字 {sum(raw.values()):,} 字を、窓を作るのとは別に拾い、すべて窓に出ていることを確かめました"
                           if not missing else f"{sum(missing.values())} 字が窓に出ていません", [str(c) for c in list(missing)[:20]]))
    au.checks.append(Check("書き込み", True, WRITE, agent=True))
    return au


def grade(k: str, blob: bytes, name: str):
    from .grade import Reason, _final, _grade_excel

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
    if k in WORKBOOKS:
        g = _grade_excel(blob, name)
        for r in g.reasons:
            if r.title == "広げていない範囲":
                r.title = "読んでいない部分"
        g.write_reasons = [Reason("D", f"{KIND_JA[k]}には書かない", WRITE_NOTE[k])]
        g.write = "D"
        return g
    if k in SXO:
        from .legacy.sxo import WRITE

        _sxo_window(k, blob, name)
        return _final([Reason("B", "OpenDocument の形に写して読む", "OpenOffice.org 1.x の XML を OpenDocument の名前に写して読みます（字と表の形は読めます。"
                              "数式は、OpenDocument と同じ形のものだけ確かめます）")],
                      [Reason("D", "OpenOffice.org 1.x には書かない", WRITE)])
    if k in READONLY:
        return _mod(k).grade(blob, name)
    raise OpenFormatError(f"この形式（{k}）の段階は決められません")
