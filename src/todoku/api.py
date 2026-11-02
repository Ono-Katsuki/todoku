"""形式を問わない入口：開く・窓・入力欄・書き込み（証明つき）・検査。コマンドと MCP サーバーの両方が使う。

形式はファイルの中身で見分ける（Excel .xlsx/.xlsm・古い Excel .xls・PDF・Word .docx・古い Word .doc）。
書き込みは、元のファイルを決して上書きしない（別の名前で書き出す）。
"""

from __future__ import annotations

import hashlib
import re
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
                "options_depend_on", "unit", "rules", "agent", "why", "placeholder", "cover"):
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
    readback_values = dict(values)  # 読み戻しで確かめる、欄の id → 書いた値
    try:
        labels = {f["id"]: f.get("label", "") for f in fields(path)}  # 白紙のときの見出し（対話で伝える一覧に使う）
    except ApiError:
        labels = {}
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
        allowed = res.allowed_inputs
        allowed_c = {(p.sheet, p.address) for p in res.patches if p.kind == "control"}
        pr = prove(blob, res.blob, allowed, allowed_c, {(p.sheet, p.address) for p in res.patches if p.kind == "macro"})
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
        anchored = {p.id: p.anchor for p in res.patches if getattr(p, "anchor", None)}
        pr = prove(blob, res.blob, {p.id for p in res.patches}, anchored)
        new_blob = res.blob
        result.update({"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings),
                       "unexpected": pr.unexpected[:20], "structure_errors": pr.lint_errors[:20],
                       "written_as": {p.id: p.value for p in res.patches
                                      if not p.anchor and str(p.value) != str(values.get(p.id))}})
        if anchored:
            # 見出しの字を目印にした書き込み：書いた欄の id（読み直した窓では、その名前のフォームの欄）
            result["anchored"] = {p.key: p.id for p in res.patches if p.anchor}
            readback_values = {k_: v for k_, v in values.items() if not k_.startswith("@")}
            readback_values.update({p.id: p.value for p in res.patches if p.anchor})
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
        from .docx.fill import written_of

        pr = prove(blob, res.blob, written_of(res.patches))
        new_blob = res.blob
        result.update({"proof_ok": pr.ok, "proof": pr.headline(), "unexpected": pr.unexpected[:20],
                       "structure_errors": pr.lint_errors[:20]})
        anchored = {p.key: p.id for p in res.patches if getattr(p, "anchor", None)}
        if anchored:
            # 字を目印にした書き込み：書いた所の id（読み直した窓では［P段落.w番号＝値］）
            result["anchored"] = anchored
            readback_values = {k: v for k, v in values.items() if not k.startswith("@")}
            readback_values.update({p.id: p.value for p in res.patches if getattr(p, "anchor", None)})
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_bytes(new_blob)
    result["new_sha256"] = hashlib.sha256(new_blob).hexdigest()
    # 読み戻し：書き込み後のファイルの窓に、書いた値が出ているか
    after_list = fields(str(dst))
    after = {f["id"]: f.get("value") for f in after_list}
    mismatch = []
    for key, v in readback_values.items():
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
    result.update(_to_tell(after_list, readback_values, result, labels))
    return result


def _to_tell(after_list: list[dict], written: dict, result: dict, labels: dict | None = None) -> dict:
    """本人が Word などで開いて確かめなくても、エージェントが対話で伝えられるように：書いた所（見出しと、書いた後の
    ファイルを読み直した値）・まだ空いている欄・本人に伝えること。"""
    by_id = {f["id"]: f for f in after_list}
    told = []
    for key, v in written.items():
        f = by_id.get(key, {})
        d = {"id": key, "label": (labels or {}).get(key) or f.get("label", ""), "value": f.get("value", v) if f else v}
        for k_ in ("sheet", "page"):
            if k_ in f:
                d[k_] = f[k_]
        told.append(d)
    empty = [f for f in after_list if f.get("value") in (None, "") and not f.get("readonly") and not f.get("agent")
             and f.get("kind") not in ("checkbox", "radio", "button", "picture")]
    notes = list(result.get("warnings") or [])
    if not result.get("proof_ok", True):
        notes.insert(0, "書き込みの証明で、想定外の変更が見つかりました：" + "／".join(map(str, result.get("unexpected") or []))[:300])
    if not result.get("read_back_ok", True):
        notes.insert(0, "書いた後のファイルを読み直すと、書いた値と違う欄があります")

    def place(d: dict) -> str:
        where = d.get("sheet") or (f"{d['page']}ページ" if d.get("page") else "")
        return f"{where}・" if where else ""

    def shown(d: dict) -> str:
        v = "☑" if d["value"] is True else "☐" if d["value"] is False else str(d["value"])
        lab = d["label"] or d["id"]
        m = re.search(r"［[ 　]*］", lab)
        if m:
            return f"{lab[:m.start()]}［{v}］{lab[m.end():]}"  # 「令和［8］年 月 日」：空欄に値を入れた形で
        return f"{lab}：{v}"

    lines = [f"書いた所（{len(told)} か所。書いた後のファイルを読み直した値）："]
    lines += [f"- {place(d)}{shown(d)}" for d in told[:60]]
    if len(told) > 60:
        lines.append(f"- ほか {len(told) - 60} か所")
    if empty:
        lines.append(f"まだ空いている欄（{len(empty)} か所）：" + "、".join(f"{place(f)}{f.get('label') or f['id']}" for f in empty[:20])
                     + ("ほか" if len(empty) > 20 else ""))
    if notes:
        lines.append("本人に伝えること：" + "／".join(notes[:10]))
    return {"written": told, "still_empty": [{"id": f["id"], "label": f.get("label", ""), **({"sheet": f["sheet"]} if "sheet" in f else {}),
                                              **({"page": f["page"]} if "page" in f else {})} for f in empty[:50]],
            "still_empty_count": len(empty), "say_to_user": "\n".join(lines)}


def run_macro(path: str, macro: str, out_path: str, answers: list | None = None, files: dict | None = None) -> dict:
    """マクロ（ボタンの処理など）を動かし、マクロが変えたセルを証明つきで out_path に書く。
    answers：MsgBox・InputBox の答え（順に）。files：マクロが読むファイル（マクロの中のパス → 手元のファイル）。
    人の答えやファイルが要るときは、勝手に決めずに、その問いを返す。"""
    from .book import open_book
    from .formula.parser import addr
    from .vbarun import NeedsInput, Unsupported
    from .vbarun import run_macro as run
    from .vbarun.parse import VbaSyntaxError
    from .xlsx.edit import EditError
    from .xlsx.edit import fill as xfill
    from .xlsx.proof import prove

    src = Path(path).expanduser().resolve()
    dst = Path(out_path).expanduser().resolve()
    if src == dst:
        raise ApiError("元のファイルは上書きしません。別の名前（out_path）を指定してください")
    blob, name, k = _read(path)
    if k not in ("xlsx", "xls"):
        raise ApiError("マクロを動かせるのは、いまは Excel の書類です")
    wb = open_book(blob, name)
    inputs = {key: Path(v).expanduser().read_bytes() for key, v in (files or {}).items()}
    try:
        r = run(wb, macro, answers, inputs)
    except NeedsInput as e:
        what = {"msgbox": "はい・いいえなどの答え", "inputbox": "入力する値", "file": "ファイル"}.get(e.kind, e.kind)
        return {"done": False, "needs": e.kind, "prompt": e.prompt, "line": e.line,
                "note": f"マクロが{what}を求めています。本人に確かめて、answers（または files）に入れてもう一度呼んでください"}
    except (Unsupported, VbaSyntaxError) as e:
        return {"done": False, "unsupported": str(e), "note": "この命令は、規則では再現できません。ソースを読んで判断してください"}
    out: dict = {"done": True, "macro": macro, "messages": r.messages, "printed": r.printed, "error": r.error,
                 "volatile": r.volatile, "changed": {f"{s_}!{addr(rr, cc)}": v for (s_, rr, cc), v in r.changed.items()}}
    if r.files:
        fdir = dst.parent / f"{dst.stem}_マクロのファイル"
        fdir.mkdir(parents=True, exist_ok=True)
        saved = []
        for fname, data in r.files.items():
            target = fdir / Path(fname.replace("\\", "/")).name
            target.write_bytes(data)
            saved.append(str(target))
        out["files"] = saved
    if not r.changed:
        out["note"] = "マクロは、セルの値を変えませんでした（ファイルは書き出していません）"
        return out
    values = {f"{s_}!{addr(rr, cc)}": v for (s_, rr, cc), v in r.changed.items()}
    try:
        res = xfill(open_book(blob, name), values, allow_formula_overwrite=True, force=True, by_macro=set(r.changed),
                    run_macros=False)
    except EditError as e:
        raise ApiError(str(e)) from None
    pr = prove(blob, res.blob, res.allowed_inputs, {(p.sheet, p.address) for p in res.patches if p.kind == "control"},
               {(p.sheet, p.address) for p in res.patches if p.kind in ("macro", "input")})
    dst.write_bytes(res.blob)
    out.update({"out_path": str(dst), "proof_ok": pr.ok, "proof": pr.headline(), "unexpected": pr.unexpected[:20],
                "recalculated": {f"{p.sheet}!{p.address}": p.after for p in res.patches if p.kind == "formula"},
                "warnings": res.warnings})
    return out


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
    d = au.to_dict()
    d["grade"] = grade(path)
    return d


def grade(path: str, findings: list[dict] | None = None, answered: list[str] | None = None) -> dict:
    """段階（S〜F）：誰が扱えば正確に読んで入れられるか。規則で決めた段階と、エージェントが確かめること。
    findings（エージェントが窓を読んで見つけたこと：{grade, title, detail, where}）を渡すと、段階に合わせる。"""
    from .grade import combine, grade_blob

    blob, name, _k = _read(path)
    g = grade_blob(blob, name)
    if findings or answered:
        g = combine(g, findings or [], answered)
    return g.to_dict()


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
        m = doc.get("macros")
        d["macros"] = bool(m)
        if isinstance(m, dict) and m.get("readable"):
            d["macro_procedures"] = sum(len(x["procedures"]) for x in m["modules"])
            d["macro_runs_on_change"] = any(e["on_change"] for e in m["events"])
    else:
        d["protection"] = doc.get("protection")
    return d
