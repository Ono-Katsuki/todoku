"""形式を問わない入口（api・grade）から呼ぶ、追加の形式：CSV・TSV・JSON・JSON Lines・RTF・OpenDocument（.odt・.ods）。

api.py と grade.py には、この部品を呼ぶ短い区切りだけを足す（ほかの形式の処理には手を加えない）。
- 窓：window（窓の JSON と Markdown の部分）
- 入力欄：field_rows（api.fields が、欄ごとに _field を通す）
- 書き込み：fill（書いたファイル・証明・読み戻しに使う id）と read_back（書いた後のファイルから、書いた所の値を読む）
- 検査：audit_runner、段階：grade
"""

from __future__ import annotations

KIND_JA = {"csv": "CSV", "tsv": "TSV（タブ区切り）", "json": "JSON", "jsonl": "JSON Lines", "rtf": "RTF（リッチテキスト）",
           "odt": "OpenDocument の文書（.odt）", "ods": "OpenDocument の表計算（.ods）",
           "fodt": "OpenDocument の文書（.fodt・1つの XML）", "fods": "OpenDocument の表計算（.fods・1つの XML）"}
KINDS = frozenset(KIND_JA)
DATA = frozenset({"csv", "tsv", "json", "jsonl"})
TEXT_DOCS = frozenset({"odt", "fodt"})
SHEETS = frozenset({"ods", "fods"})


class OpenFormatError(ValueError):
    pass


def _err(e: Exception) -> OpenFormatError:
    return OpenFormatError(str(e))


# ====================================================================== 窓
def window(k: str, blob: bytes, name: str) -> tuple[dict, list[tuple[str, str]]]:
    try:
        if k in DATA:
            from .datafile import window as dw

            doc = dw.build(blob, name, k)
            return doc, dw.markdown_parts(doc)
        if k in TEXT_DOCS:
            from .odf.text import OdtDocument
            from .odf.textwindow import build, markdown_parts

            win = build(OdtDocument(blob, name))
            return win, markdown_parts(win)
        if k in SHEETS:
            from .odf.sheetwindow import build, markdown_parts

            win = build(blob, name)
            return win, markdown_parts(win)
        if k == "rtf":
            from .rtf.document import RtfDocument
            from .rtf.window import build, markdown_parts

            win = build(RtfDocument(blob, name))
            return win, markdown_parts(win)
    except (ValueError, KeyError) as e:
        raise _err(e) from None
    raise OpenFormatError(f"この形式（{k}）は扱えません")


def field_rows(k: str, doc: dict):
    """(欄, どこか)。表計算はシートの名前を添える。"""
    if k in SHEETS:
        for sh in doc["sheets"]:
            for f in sh.get("fields", []):
                yield f, {"sheet": sh["name"]}
        return
    for f in doc.get("fields", []):
        yield f, {}


def summary(k: str, doc: dict) -> dict:
    d: dict = {}
    if k in SHEETS:
        d["sheets"] = [s["name"] for s in doc["sheets"]]
        m = doc.get("macros")
        d["macros"] = bool(m)
    elif k in DATA:
        d["format"] = doc.get("format")
        if k in ("csv", "tsv"):
            d["rows"] = doc.get("row_count")
            d["columns"] = doc.get("width")
    else:
        d["protection"] = doc.get("protection")
    return d


# ====================================================================== 書き込み
def fill(k: str, blob: bytes, name: str, values: dict, force: bool = False) -> dict:
    """{"blob": 書いた中身, "result": api の結果に足すもの, "readback": 読み戻しで照らす id → 値, "labels": id → 見出し}。"""
    if k in ("csv", "tsv"):
        return _fill_csv(k, blob, name, values)
    if k in ("json", "jsonl"):
        return _fill_json(k, blob, name, values)
    if k in TEXT_DOCS:
        return _fill_odt(blob, name, values, force)
    if k in SHEETS:
        from .odf.sheetfill import api_fill

        return api_fill(blob, name, values, force)
    if k == "rtf":
        from .rtf.fill import api_fill

        return api_fill(blob, name, values, force)
    raise OpenFormatError(f"この形式（{k}）には書けません")


def _fill_csv(k: str, blob: bytes, name: str, values: dict) -> dict:
    from .datafile.csvfile import CsvError, CsvFile, fill, prove

    f = CsvFile(blob, name, k)
    try:
        res = fill(f, values)
    except CsvError as e:
        raise _err(e) from None
    written = {(p.row, p.col): p.value for p in res.patches}
    pr = prove(blob, res.blob, written, name, k)
    head = f.header
    labels = {}
    for p in res.patches:
        col = head[p.col - 1].strip() if p.col <= len(head) and head[p.col - 1].strip() else f"C{p.col}"
        labels[p.key] = f"{col}（{p.row}行目" + ("・足した行" if p.row > len(f.records) else "") + "）"
    result = {"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings), "unexpected": pr.unexpected[:20],
              "structure_errors": pr.lint_errors[:20],
              "cells": {p.key: p.id for p in res.patches if p.key != p.id}}
    if res.appended_rows:
        result["appended_rows"] = res.appended_rows
    return {"blob": res.blob, "result": result, "readback": {p.key: p.value for p in res.patches}, "labels": labels}


def _fill_json(k: str, blob: bytes, name: str, values: dict) -> dict:
    from .datafile.jsonfile import JsonError, JsonFile, fill, prove

    f = JsonFile(blob, name, k)
    try:
        res = fill(f, values)
    except JsonError as e:
        raise _err(e) from None
    pr = prove(blob, res.blob, {p.id: p.value for p in res.patches}, name, k)
    result = {"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings), "unexpected": pr.unexpected[:20],
              "structure_errors": pr.lint_errors[:20], "written_as": {p.key: p.text for p in res.patches}}
    added = {p.key: p.id for p in res.patches if p.added}
    if added:
        result["added"] = added  # 足した名前・値：書いた所の id
    labels = {p.key: (p.id.rsplit("/", 1)[-1] or p.id) for p in res.patches}
    return {"blob": res.blob, "result": result, "readback": {p.key: p.value for p in res.patches}, "labels": labels,
            "ids": {p.key: p.id for p in res.patches}}


def _fill_odt(blob: bytes, name: str, values: dict, force: bool) -> dict:
    from .docx.fill import FillError
    from .odf.text import OdtDocument
    from .odf.textfill import fill, written_of
    from .odf.textproof import prove

    try:
        res = fill(OdtDocument(blob, name), values, force=force)
    except FillError as e:
        raise _err(e) from None
    pr = prove(blob, res.blob, written_of(res.patches))
    result = {"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings), "unexpected": pr.unexpected[:20],
              "structure_errors": pr.lint_errors[:20]}
    readback = {key: v for key, v in values.items() if not key.startswith("@")}
    anchored = {p.key: p.id for p in res.patches if p.anchor}
    if anchored:
        result["anchored"] = anchored
        readback.update({p.id: p.value for p in res.patches if p.anchor})
    return {"blob": res.blob, "result": result, "readback": readback, "labels": {}}


def read_back(k: str, blob: bytes, name: str, keys: dict) -> dict:
    """書いた後のファイルから、書いた所（keys の id）の値を読む（欄の一覧に出ないマス・値のため）。"""
    out: dict = {}
    try:
        if k in ("csv", "tsv"):
            from .datafile.csvfile import CsvFile

            f = CsvFile(blob, name, k)
            for key in keys:
                try:
                    r, c = f.resolve(key)
                except ValueError:
                    continue
                out[key] = f.value(r, c)
        elif k in ("json", "jsonl"):
            from .datafile.jsonfile import JsonFile, _plain

            f = JsonFile(blob, name, k)
            for key in keys:
                try:
                    node, _parent, _last, _d = f.find(key)
                except ValueError:
                    continue
                if node is not None:
                    v = _plain(node)
                    out[key] = v if not isinstance(v, (dict, list)) else keys[key] if v == keys[key] else v
        elif k in SHEETS:
            from .odf.sheetfill import read_cells

            out.update(read_cells(blob, name, keys))
    except ValueError:
        return out
    return out


# ====================================================================== 検査・段階
def audit_runner(k: str):
    if k in DATA:
        from .datafile.audit import run

        return lambda blob, name: run(blob, name, k)
    if k in TEXT_DOCS:
        from .odf.textaudit import run

        return run
    if k in SHEETS:
        from .odf.sheetaudit import run

        return run
    if k == "rtf":
        from .rtf.audit import run

        return run
    raise OpenFormatError(f"この形式（{k}）は検査できません")


def grade(k: str, blob: bytes, name: str):
    from .grade import _locked

    try:
        if k in DATA:
            return _grade_data(k, blob, name)
        if k in TEXT_DOCS:
            from .odf.text import OdtDocument
            from .odf.textwindow import build

            return _grade_textdoc(build(OdtDocument(blob, name)), "OpenDocument の文書", "フォームの部品・入力の欄")
        if k in SHEETS:
            from .odf.sheetwindow import grade_sheets

            return grade_sheets(blob, name)
        if k == "rtf":
            from .rtf.document import RtfDocument
            from .rtf.window import build

            return _grade_textdoc(build(RtfDocument(blob, name)), "RTF", "フォームの欄")
    except Exception as e:  # noqa: BLE001
        from .odf.package import OdfEncrypted

        if isinstance(e, OdfEncrypted):
            return _locked("パスワード", str(e), "パスワードがなければ、書けません")
        return _locked("開けない", f"ファイルを読めませんでした（{type(e).__name__}: {str(e)[:80]}）", "読めないので、書けません")
    return _locked("対応していない形式", f"この形式（{k}）は、まだ読めません", "読めないので、書けません")


def _grade_data(k: str, blob: bytes, name: str):
    from .grade import Grade, Reason, _locked

    direct = [Reason("A", "データそのもの", "行と列（か、名前と値）のデータに、そのまま書けます")]
    if k in ("csv", "tsv"):
        from .datafile.csvfile import CsvFile

        f = CsvFile(blob, name, k)
        if "\x00" in f.text:
            return _locked("読めない", "文字のファイルではありません（バイナリのデータが入っています）", "読めないので、書けません")
        reasons = [Reason("S", "データそのもの", f"表のデータ（{KIND_JA[k]}・{f.dec.label}・{len(f.records):,} 行）で、変換しなくても機械が読めます")]
        if f.problems:
            reasons.append(Reason("S", "書き方の問題", "RFC 4180 に合わない書き方があります（どう読んだかは窓に書いてあります）", f.problems[:5]))
        return Grade("S", reasons, write="A", write_reasons=direct)
    from .datafile.jsonfile import JsonFile

    f = JsonFile(blob, name, k)
    reasons = [Reason("S", "データそのもの", f"構造のあるデータ（{KIND_JA[k]}・{f.dec.label}）で、変換しなくても機械が読めます")]
    if f.duplicates:
        reasons.append(Reason("C", "同じ名前が2つ以上ある", "オブジェクトに同じ名前が2つ以上あり、どの値を使うかは読むソフトによって違います。"
                              "JSON Pointer では1つに決まらないので、エージェントが読んで判断します", f.duplicates[:12]))
        return Grade("C", reasons, write="B", write_reasons=[Reason("B", "同じ名前が2つ以上ある",
                                                                  "どの値に書くかを、エージェントが読んで決めます", f.duplicates[:12])])
    return Grade("S", reasons, write="A", write_reasons=direct)


def _grade_textdoc(win: dict, kind_ja: str, explicit: str):
    """Word と同じ決め方：規則では決められない所は C、配置から推定した欄は B、作り手の印（部品・入力の欄）だけなら A。"""
    from .grade import Reason, _agent_reason, _final, _macro_reasons, _write_fields, _write_macros

    reasons: list = []
    agent = [f for f in win["fields"] if f.get("agent")]
    fields = [f for f in win["fields"] if not f.get("agent")]
    layout = [f for f in fields if f.get("source") == "layout"]
    if agent:
        reasons.append(_agent_reason(agent))
    if layout:
        reasons.append(Reason("B", "配置から推定した入力欄",
                              f"入力欄 {len(fields)} か所のうち {len(layout)} か所は、{explicit}ではなく、空のマスや空白・下線から推定しました",
                              [f["id"] for f in layout[:12]]))
    elif fields:
        reasons.append(Reason("A", "作り手の印のある入力欄", f"入力欄 {len(fields)} か所は、すべて{kind_ja}の{explicit}です"))
    reasons += _macro_reasons(win.get("macros"))
    for e in win.get("unread") or []:
        reasons.append(Reason("F", "読めない部分", e))
    write = _write_fields(fields, layout, agent, f"{kind_ja}の{explicit}", "欄のない所は、前後の字を指定して書き換えます")
    write += _write_macros(win.get("macros"))
    return _final(reasons, write)
