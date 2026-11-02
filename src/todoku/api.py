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


class PasswordNeeded(ApiError):
    """開くのにパスワードが要る：本人に聞いて、password に入れてもう一度呼ぶ。"""


class PasswordWrong(ApiError):
    """パスワードが違う：本人にもう一度確かめる。"""


def _read(path: str, password: str | None = None):
    """(読む中身, ファイルの名前, 形式, 暗号の情報)。パスワードで保護された書類は、解いた中身を返す（暗号の情報は、
    書いた後に同じパスワードで保護し直すのに使う。パスワードそのものは結果に出さない）。"""
    from .unlock import NeedsPassword, WrongPassword, unlock

    p = Path(path).expanduser()
    if not p.is_file():
        raise ApiError(f"ファイルが見つかりません: {path}")
    blob = p.read_bytes()
    try:
        plain, lock = unlock(blob, password)
    except NeedsPassword as e:
        raise PasswordNeeded(f"{e}。本人にパスワードを聞いて、password に入れてもう一度呼んでください"
                             "（パスワードは、結果にも記録にも残しません）") from None
    except WrongPassword as e:
        raise PasswordWrong(f"{e}。本人にもう一度確かめてください") from None
    k = kind_of(plain)
    if k not in KIND_JA:
        raise ApiError(f"この形式（{k}）は扱えません。Excel・PDF・Word のファイルを指定してください")
    return plain, p.name, k, lock


def _pdf(blob: bytes, name: str, lock):
    from .pdf.doc import PdfDocument
    from .unlock import pdf_password

    pw = pdf_password(lock)
    return PdfDocument(blob, name, password=pw) if pw else PdfDocument(blob, name)


def window(path: str, password: str | None = None) -> tuple[str, dict, list[tuple[str, str]]]:
    """(形式, 窓の JSON, Markdown の部分)。"""
    blob, name, k, lock = _read(path, password)
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
        from .pdf.window import build, markdown_parts

        doc = build(_pdf(blob, name, lock))
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


def fields(path: str, password: str | None = None) -> list[dict]:
    """入力欄の一覧（名前・見出し・種類・選択肢・いまの値・書けるか・その行の文脈）。"""
    k, doc, _ = window(path, password)
    out = []
    if k in ("xlsx", "xls"):
        for sh in doc["sheets"]:
            for f in sh.get("fields", []):
                out.append(_field(f, sheet=sh["name"]))
    elif k == "pdf":
        for pg in doc["pages"]:
            for f in pg["fields"]:
                out.append(_field(f, page=pg["number"]))
        # XFA のフォーム：ページに置かれていない XFA の欄（動的なフォームの欄など）も、同じ形で出す
        seen = {f["id"] for f in out}
        for f in (doc.get("xfa") or {}).get("fields", []):
            if f["id"] not in seen:
                d = _field(f)
                if f.get("why_not"):
                    d["why"] = f["why_not"]  # 書けない理由
                out.append(d)
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


def fill(path: str, values: dict, out_path: str, force: bool = False, password: str | None = None) -> dict:
    """書き込み、証明し、書き込み後のファイルを読み直して、書いた値がそのまま出ているかを確かめる。
    パスワードで保護された書類は、解いた中身に書いて証明し、元と同じパスワードで保護し直して書き出す。"""
    src = Path(path).expanduser().resolve()
    dst = Path(out_path).expanduser().resolve()
    if src == dst:
        raise ApiError("元のファイルは上書きしません。別の名前（out_path）を指定してください")
    blob, name, k, lock = _read(path, password)
    result: dict = {"kind": k, "out_path": str(dst), "original_sha256": hashlib.sha256(src.read_bytes()).hexdigest()}
    readback_values = dict(values)  # 読み戻しで確かめる、欄の id → 書いた値
    try:
        labels = {f["id"]: f.get("label", "") for f in fields(path, password)}  # 白紙のときの見出し（対話で伝える一覧に使う）
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
        allowed_c = {(p.sheet, p.address): p.after for p in res.patches if p.kind == "control"}  # 部品 → 書いた状態
        pr = prove(blob, res.blob, allowed, allowed_c, {(p.sheet, p.address) for p in res.patches if p.kind == "macro"})
        new_blob = res.blob
        result.update({"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings),
                       "unexpected": pr.unexpected[:20], "structure_errors": pr.lint_errors[:20],
                       "recalculated": {f"{p.sheet}!{p.address}": p.after for p in res.patches if p.kind == "formula"}})
    elif k == "pdf":
        from .pdf.fill import FillError
        from .pdf.fill import fill as pfill
        from .pdf.proof import prove

        try:
            res = pfill(_pdf(blob, name, lock), values, force=force)
        except FillError as e:
            raise ApiError(str(e)) from None
        anchored = {p.id: p.anchor for p in res.patches if getattr(p, "anchor", None)}
        pr = _prove_pdf(prove, blob, res.blob, {p.id for p in res.patches}, anchored, lock)
        new_blob = res.blob
        result.update({"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings),
                       "unexpected": pr.unexpected[:20], "structure_errors": pr.lint_errors[:20],
                       "written_as": {p.id: p.value for p in res.patches
                                      if not p.anchor and str(p.value) != str(values.get(p.id))}})
        xfa_written = {p.id: p.value for p in res.patches if p.kind == "xfa"}
        if xfa_written:
            # XFA の欄：窓に出る値（日付・数は決まった形、選択肢はデータの値）で読み戻す
            readback_values.update(xfa_written)
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
    written_plain = new_blob  # 書いた中身（保護し直す前）
    if lock is not None:
        # 元と同じパスワードで保護し直し、それを同じパスワードで解いて、書いた中身とバイト一致かを確かめてから書き出す
        from .unlock import relock, unlock

        sealed = relock(lock, new_blob)
        if lock.method != "pdf":
            again, _l = unlock(sealed, password)
            if again != new_blob:
                raise ApiError("保護し直したファイルを同じパスワードで解くと、書いた中身と一致しません（書き出していません）")
        new_blob = sealed
        result["password_protected"] = f"元と同じパスワードで保護したまま書き出しました（{lock.label}）"
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_bytes(new_blob)
    result["new_sha256"] = hashlib.sha256(new_blob).hexdigest()
    # 読み戻し：書き込み後のファイルの窓に、書いた値が出ているか
    after_list = fields(str(dst), password)
    after = {f["id"]: f.get("value") for f in after_list}
    if k in ("xlsx", "xls"):
        # 欄の一覧に出ないセル（番地で書いたセル）も、書いた後のブックのセルの値を直接読んで照らす
        from .book import open_book
        from .xlsx.edit import resolve_target

        try:
            wb_after = open_book(written_plain, name)
            for key in readback_values:
                if key in after or key.startswith("@"):
                    continue
                try:
                    sh_, r_, c_ = resolve_target(wb_after, key)
                except Exception:  # noqa: BLE001
                    continue  # 部品の名前など、セルでない指定
                cell = wb_after.sheet(sh_).cells.get((r_, c_))
                after[key] = cell.value if cell is not None else None
        except Exception:  # noqa: BLE001
            pass
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


def _prove_pdf(prove, original: bytes, new: bytes, allowed: set, anchored: dict, lock):
    """PDF の証明。パスワードで保護された PDF は、同じパスワードで開いて確かめる。"""
    import inspect

    from .unlock import pdf_password

    pw = pdf_password(lock)
    if pw and "password" in inspect.signature(prove).parameters:
        return prove(original, new, allowed, anchored, password=pw)
    return prove(original, new, allowed, anchored)


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


def run_macro(path: str, macro: str, out_path: str, answers: list | None = None, files: dict | None = None,
              password: str | None = None) -> dict:
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
    blob, name, k, lock = _read(path, password)
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
    out_blob = res.blob
    if lock is not None:
        from .unlock import relock, unlock

        out_blob = relock(lock, res.blob)
        if unlock(out_blob, password)[0] != res.blob:
            raise ApiError("保護し直したファイルを同じパスワードで解くと、書いた中身と一致しません（書き出していません）")
        out["password_protected"] = f"元と同じパスワードで保護したまま書き出しました（{lock.label}）"
    dst.write_bytes(out_blob)
    out.update({"out_path": str(dst), "proof_ok": pr.ok, "proof": pr.headline(), "unexpected": pr.unexpected[:20],
                "recalculated": {f"{p.sheet}!{p.address}": p.after for p in res.patches if p.kind == "formula"},
                "warnings": res.warnings})
    return out


def _num_equal(a, b) -> bool:
    try:
        return abs(float(a) - float(b)) < 1e-9
    except (TypeError, ValueError):
        return False


def audit(path: str, password: str | None = None) -> dict:
    import inspect

    from .unlock import pdf_password

    blob, name, k, lock = _read(path, password)
    if k == "pdf":
        from .pdf.audit import run
    elif k == "docx":
        from .docx.audit import run
    elif k == "doc":
        from .doc.audit import run
    else:
        from .xlsx.audit import run
    pw = pdf_password(lock)
    au = run(blob, name, password=pw) if pw and "password" in inspect.signature(run).parameters else run(blob, name)
    d = au.to_dict()
    d["grade"] = grade(path, password=password)
    if lock is not None:
        d["password_protected"] = f"パスワードで保護された書類です（{lock.label}）。解いた中身を確かめました"
    return d


def grade(path: str, findings: list[dict] | None = None, answered: list[str] | None = None,
          password: str | None = None) -> dict:
    """段階（S〜F）：誰が扱えば正確に読んで入れられるか。規則で決めた段階と、エージェントが確かめること。
    findings（エージェントが窓を読んで見つけたこと：{grade, title, detail, where}）を渡すと、段階に合わせる。
    パスワードで保護された書類は、本人に聞いたパスワードで解いた中身で決める（パスワードがなければ PasswordNeeded）。"""
    from .grade import combine, grade_blob
    from .unlock import pdf_password

    blob, name, _k, lock = _read(path, password)
    g = grade_blob(blob, name, password=pdf_password(lock))
    if findings or answered:
        g = combine(g, findings or [], answered)
    return g.to_dict()


def summary(path: str, password: str | None = None) -> dict:
    k, doc, parts = window(path, password)
    fl = fields(path, password)
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
