"""形式を問わない入口：開く・窓・入力欄・書き込み（証明つき）・検査。コマンドと MCP サーバーの両方が使う。

形式はファイルの中身で見分ける（Excel .xlsx/.xlsm・古い Excel .xls・PDF・Word .docx・古い Word .doc）。
書き込みは、元のファイルを決して上書きしない（別の名前で書き出す）。
"""

from __future__ import annotations

import hashlib
from pathlib import Path

from .book import FormatError, kind_of

KIND_JA = {"xlsx": "Excel", "xls": "古い Excel（.xls）", "pdf": "PDF", "docx": "Word", "doc": "古い Word（.doc）"}


class ApiError(ValueError):
    pass


def _read(path: str) -> tuple[bytes, str, str]:
    p = Path(path).expanduser()
    if not p.is_file():
        raise ApiError(f"ファイルが見つかりません: {path}")
    blob = p.read_bytes()
    k = kind_of(blob)
    if k not in KIND_JA:
        raise ApiError(f"この形式（{k}）は扱えません。Excel・PDF・Word のファイルを指定してください")
    return blob, p.name, k


def window(path: str) -> tuple[str, dict, list[tuple[str, str]]]:
    """(形式, 窓の JSON, Markdown の部分)。"""
    blob, name, k = _read(path)
    if k in ("xlsx", "xls"):
        from .book import open_book
        from .xlsx import window as xw

        try:
            wb = open_book(blob, name)
        except FormatError as e:
            raise ApiError(str(e)) from None
        doc = xw.build(wb)
        return k, doc, xw.markdown_parts(doc)
    if k == "pdf":
        from .pdf.doc import PdfDocument
        from .pdf.window import build, markdown_parts

        doc = build(PdfDocument(blob, name))
        return k, doc, markdown_parts(doc)
    from .docx.window import build, markdown_parts

    doc = build(_word(blob, name, k))
    return k, doc, markdown_parts(doc)


def _word(blob: bytes, name: str, k: str):
    if k == "doc":
        from .doc.book import DocFile

        return DocFile(blob, name)
    from .docx.document import DocxDocument

    return DocxDocument(blob, name)


def fields(path: str) -> list[dict]:
    """入力欄の一覧（名前・見出し・種類・選択肢・いまの値・書けるか・その行の文脈）。"""
    k, doc, _ = window(path)
    out = []
    if k in ("xlsx", "xls"):
        for sh in doc["sheets"]:
            for f in sh.get("fields", []):
                out.append(_field(f, sheet=sh["name"]))
    elif k == "pdf":
        for pg in doc["pages"]:
            for f in pg["fields"]:
                out.append(_field(f, page=pg["number"]))
    else:
        for f in doc["fields"]:
            out.append(_field(f))
    return out


def _field(f: dict, **where) -> dict:
    d = {"id": f["id"], "label": f.get("label", ""), "kind": f.get("kind", ""), "value": f.get("value")}
    for key in ("options", "maxlen", "format", "readonly", "protected", "confidence", "context", "boxes", "write_as",
                "options_depend_on", "unit", "rules"):
        if f.get(key) not in (None, "", [], False):
            d[key] = f[key]
    d.update({k: v for k, v in where.items() if v is not None})
    return d


def fill(path: str, values: dict, out_path: str, force: bool = False) -> dict:
    """書き込み、証明し、書き込み後のファイルを読み直して、書いた値がそのまま出ているかを確かめる。"""
    src = Path(path).expanduser().resolve()
    dst = Path(out_path).expanduser().resolve()
    if src == dst:
        raise ApiError("元のファイルは上書きしません。別の名前（out_path）を指定してください")
    blob, name, k = _read(path)
    result: dict = {"kind": k, "out_path": str(dst), "original_sha256": hashlib.sha256(blob).hexdigest()}
    if k in ("xlsx", "xls"):
        from .book import open_book
        from .xlsx.edit import EditError
        from .xlsx.edit import fill as xfill
        from .xlsx.proof import prove

        wb = open_book(blob, name)
        try:
            res = xfill(wb, values, force=force)
        except EditError as e:
            raise ApiError(str(e)) from None
        allowed = {(p.sheet, p.address) for p in res.patches if p.kind == "input"}
        allowed_c = {(p.sheet, p.address) for p in res.patches if p.kind == "control"}
        pr = prove(blob, res.blob, allowed, allowed_c)
        new_blob = res.blob
        result.update({"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings),
                       "unexpected": pr.unexpected[:20], "structure_errors": pr.lint_errors[:20],
                       "recalculated": {f"{p.sheet}!{p.address}": p.after for p in res.patches if p.kind == "formula"}})
    elif k == "pdf":
        from .pdf.doc import PdfDocument
        from .pdf.fill import FillError
        from .pdf.fill import fill as pfill
        from .pdf.proof import prove

        try:
            res = pfill(PdfDocument(blob, name), values, force=force)
        except FillError as e:
            raise ApiError(str(e)) from None
        pr = prove(blob, res.blob, {p.id for p in res.patches})
        new_blob = res.blob
        result.update({"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings),
                       "unexpected": pr.unexpected[:20], "structure_errors": pr.lint_errors[:20],
                       "written_as": {p.id: p.value for p in res.patches if str(p.value) != str(values.get(p.id))}})
    else:
        from .docx.fill import FillError

        try:
            if k == "doc":
                from .doc.audit import prove
                from .doc.fill import fill as wfill
            else:
                from .docx.fill import fill as wfill
                from .docx.proof import prove
            res = wfill(_word(blob, name, k), values, force=force)
        except FillError as e:
            raise ApiError(str(e)) from None
        pr = prove(blob, res.blob, {p.id: p.value for p in res.patches})
        new_blob = res.blob
        result.update({"proof_ok": pr.ok, "proof": pr.headline(), "unexpected": pr.unexpected[:20],
                       "structure_errors": pr.lint_errors[:20]})
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_bytes(new_blob)
    result["new_sha256"] = hashlib.sha256(new_blob).hexdigest()
    # 読み戻し：書き込み後のファイルの窓に、書いた値が出ているか
    after = {f["id"]: f.get("value") for f in fields(str(dst))}
    mismatch = []
    for key, v in values.items():
        g = after.get(key)
        if key not in after:
            continue
        if isinstance(v, bool) or isinstance(g, bool):
            ok = bool(g) == (v if isinstance(v, bool) else str(v).lower() in ("true", "1", "yes", "はい", "あり"))
        else:
            ok = str(g) == str(v) or _num_equal(g, v)
        if not ok:
            mismatch.append({"id": key, "written": v, "read_back": g})
    result["read_back_ok"] = not mismatch
    result["read_back_differences"] = mismatch[:30]
    return result


def _num_equal(a, b) -> bool:
    try:
        return abs(float(a) - float(b)) < 1e-9
    except (TypeError, ValueError):
        return False


def audit(path: str) -> dict:
    blob, name, k = _read(path)
    if k == "pdf":
        from .pdf.audit import run
    elif k == "docx":
        from .docx.audit import run
    elif k == "doc":
        from .doc.audit import run
    else:
        from .xlsx.audit import run
    au = run(blob, name)
    return au.to_dict()


def summary(path: str) -> dict:
    k, doc, parts = window(path)
    fl = fields(path)
    d = {"file": Path(path).name, "kind": k, "kind_ja": KIND_JA[k], "fields": len(fl),
         "writable_fields": sum(1 for f in fl if not f.get("readonly")), "window_parts": len(parts),
         "window_chars": sum(len(t) for _, t in parts)}
    if k == "pdf":
        d["pages"] = len(doc["pages"])
        d["form_fields"] = doc["form"]["acroform_fields"]
    elif k in ("xlsx", "xls"):
        d["sheets"] = [s["name"] for s in doc["sheets"]]
        d["macros"] = doc.get("macros", False)
    else:
        d["protection"] = doc.get("protection")
    return d
