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
# 一太郎（.jtd）：読む・書く（同じ字数の上書きと、本文の字数を変える書き込み。圧縮した文書は縮め直す）
KIND_JA["jtd"] = "一太郎（.jtd）"

# ---------------------------------------------------------------------- 追加の形式（CSV・TSV・JSON・JSON Lines・RTF・OpenDocument）
# 形式ごとの処理は openkinds が受け持つ。この入口には、それを呼ぶ短い区切りだけを足す
from . import openkinds as _open  # noqa: E402

KIND_JA.update(_open.KIND_JA)
# ZIP：中のファイルを「a.zip!/中の名前」で指して、その形式のまま読む・書く（ziparc）
KIND_JA["zip"] = "ZIP（ファイルをまとめたもの）"


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
    if not p.is_file() and _member(path) is not None:
        return _read_member(path, password)
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
    k = kind_of(plain, p.name)  # 名前は、中身だけでは決まらない形式（CSV・JSON など）を見分けるのに使う
    if k not in KIND_JA:
        raise ApiError(f"この形式（{k}）は扱えません。Excel・PDF・Word のファイルを指定してください")
    return plain, p.name, k, lock


def _member(path: str) -> tuple[str, list[str]] | None:
    """「a.zip!/b.csv」の形の指定（a.zip がファイルとしてあるもの）。"""
    from . import ziparc

    sp = ziparc.split(path)
    if sp is None or not Path(sp[0]).expanduser().is_file():
        return None
    return sp


def _member_blob(path: str) -> tuple[bytes, str]:
    """ZIP の中のファイルの中身と名前（ZIP の中の ZIP も、続けて「!/」で指す）。"""
    from . import ziparc

    outer, chain = _member(path)
    blob = Path(outer).expanduser().read_bytes()
    for m in chain:
        if kind_of(blob) != "zip":
            raise ApiError(f"「{m}」の外側のファイルは ZIP ではありません")
        try:
            blob = ziparc.extract(blob, m)
        except ziparc.ArchiveError as e:
            raise ApiError(str(e)) from None
    return blob, ziparc.base_name(chain[-1])


def _read_member(path: str, password: str | None):
    from .unlock import NeedsPassword, WrongPassword, unlock

    blob, name = _member_blob(path)
    try:
        plain, lock = unlock(blob, password)
    except NeedsPassword as e:
        raise PasswordNeeded(f"{e}。本人にパスワードを聞いて、password に入れてもう一度呼んでください"
                             "（パスワードは、結果にも記録にも残しません）") from None
    except WrongPassword as e:
        raise PasswordWrong(f"{e}。本人にもう一度確かめてください") from None
    k = kind_of(plain, name)
    if k not in KIND_JA:
        raise ApiError(f"ZIP の中の「{name}」の形式（{k}）は扱えません")
    return plain, name, k, lock


def _encoding(blob: bytes, k: str, encoding: str | None) -> str | None:
    """エージェントが選んだ文字コード（encoding）を確かめる：文字だけのファイル（CSV・TSV・JSON・JSON Lines・.txt）にだけ使え、
    ファイルがその文字コードとして厳密に読めること。読めなければ ApiError（黙ってほかの文字コードで読まない）。"""
    if not encoding:
        return None
    if k not in _open.TEXT_FILES:
        raise ApiError(f"encoding（文字コード）は、文字だけのファイル（CSV・TSV・JSON・JSON Lines・.txt）にだけ指定できます"
                       f"（この書類は {KIND_JA.get(k, k)}。文字コードはファイルの作りで決まります）")
    from .datafile.textenc import EncodingError, decode

    try:
        return decode(blob, encoding).encoding
    except EncodingError as e:
        raise ApiError(str(e)) from None


def _pdf(blob: bytes, name: str, lock):
    from .pdf.doc import PdfDocument
    from .unlock import pdf_password

    pw = pdf_password(lock)
    return PdfDocument(blob, name, password=pw) if pw else PdfDocument(blob, name)


# 窓の控え：対話では、同じファイルを開いてから書く（open_document → fill_document）ので、中身が同じなら作り直さない。
# 鍵はファイルの中身のハッシュと名前だけ（パスワードで保護された書類は控えない。パスワードは控えにも残さない）
_WINDOWS: "dict[tuple, tuple[str, bytes]]" = {}
_WINDOWS_MAX = 8


def window(path: str, password: str | None = None, encoding: str | None = None) -> tuple[str, dict, list[tuple[str, str]]]:
    """(形式, 窓の JSON, Markdown の部分)。encoding：文字だけのファイルの文字コードを、エージェントが候補から選んだとき。"""
    import hashlib
    import pickle

    blob, name, k, lock = _read(path, password)
    encoding = _encoding(blob, k, encoding)
    key = (hashlib.sha256(blob).hexdigest(), name, encoding) if lock is None else None
    if key is not None and key in _WINDOWS:
        _WINDOWS[key] = _WINDOWS.pop(key)  # いちばん新しく使ったものにする
        return pickle.loads(_WINDOWS[key][1])  # 呼んだ側が変えても控えは変わらないよう、写しを返す
    out = _window(blob, name, k, lock, encoding)
    if key is not None:
        try:
            _WINDOWS[key] = (k, pickle.dumps(out, protocol=pickle.HIGHEST_PROTOCOL))
        except Exception:  # noqa: BLE001  写せないものが入っていれば、控えない
            return out
        while len(_WINDOWS) > _WINDOWS_MAX:
            _WINDOWS.pop(next(iter(_WINDOWS)))
    return out


def _window(blob: bytes, name: str, k: str, lock, encoding: str | None = None) -> tuple[str, dict, list[tuple[str, str]]]:
    if k == "zip":
        from .ziparc import ArchiveError
        from .ziparc import window as zwindow

        try:
            doc, parts = zwindow(blob, name, KIND_JA)
        except ArchiveError as e:
            raise ApiError(str(e)) from None
        return k, doc, parts
    if k in _open.KINDS:  # 追加の形式
        try:
            doc, parts = _open.window(k, blob, name, encoding)
        except _open.OpenFormatError as e:
            raise ApiError(str(e)) from None
        return k, doc, parts
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
    if k == "jtd":
        # 一太郎：Word と同じ形の窓に、上書きできる字数を添える
        from .jtd.book import JtdError, JtdFile
        from .jtd.window import build as jbuild
        from .jtd.window import markdown_parts as jparts

        try:
            doc = jbuild(JtdFile(blob, name))
        except JtdError as e:
            raise ApiError(f"一太郎の文書として読めません：{e}") from None
        return k, doc, jparts(doc)
    from .docx.window import build, markdown_parts

    d = _word(blob, name, k)
    if getattr(d, "word6", False):
        # Word 6.0/95：上書きできるバイト数（窓では全角の字数）と、字のない欄（変更できない欄）を添える
        from .doc.word6 import build as build6

        doc = build6(d)
    else:
        doc = build(d)
    return k, doc, markdown_parts(doc)


def _is_word6(blob: bytes) -> bool:
    from .doc.word6 import is_word6

    return is_word6(blob)


def _word(blob: bytes, name: str, k: str):
    if k == "doc":
        from .doc.book import DocFile

        return DocFile(blob, name)
    from .docx.document import DocxDocument

    return DocxDocument(blob, name)


def fields(path: str, password: str | None = None, encoding: str | None = None) -> list[dict]:
    """入力欄の一覧（名前・見出し・種類・選択肢・いまの値・書けるか・その行の文脈）。"""
    k, doc, _ = window(path, password, encoding)
    if k == "zip":
        return []  # ZIP そのものに欄はない（中のファイルを「a.zip!/名前」で開く）
    if k in _open.KINDS:  # 追加の形式
        return [_field(f, **w) for f, w in _open.field_rows(k, doc)]
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
                "options_depend_on", "unit", "rules", "agent", "why", "placeholder", "cover", "kind_candidates",
                "alternative_to", "group", "why_not"):
        if f.get(key) not in (None, "", [], False):
            d[key] = f[key]
    d.update({k: v for k, v in where.items() if v is not None})
    return d


def fill(path: str, values: dict, out_path: str, force: bool = False, password: str | None = None,
         encoding: str | None = None) -> dict:
    """書き込み、証明し、書き込み後のファイルを読み直して、書いた値がそのまま出ているかを確かめる。
    パスワードで保護された書類は、解いた中身に書いて証明し、元と同じパスワードで保護し直して書き出す。
    encoding：文字だけのファイルの文字コードを、エージェントが候補から選んだとき（その文字コードで読み、書いた所だけを
    その文字コードのバイトにして置き換える。ほかのバイトは元のまま）。"""
    if _member(path) is not None:
        return _fill_member(path, values, out_path, force, password, encoding)
    src = Path(path).expanduser().resolve()
    dst = Path(out_path).expanduser().resolve()
    if src == dst:
        raise ApiError("元のファイルは上書きしません。別の名前（out_path）を指定してください")
    blob, name, k, lock = _read(path, password)
    encoding = _encoding(blob, k, encoding)
    if k == "zip":
        raise ApiError("ZIP そのものには書けません。中のファイルを path に「" + path + "!/中の名前」で指定してください")
    _seal_supported(k, values)
    result: dict = {"kind": k, "out_path": str(dst), "original_sha256": hashlib.sha256(src.read_bytes()).hexdigest()}
    if encoding:
        result["encoding"] = {"encoding": encoding, "chosen_by": "agent",
                              "detail": "エージェントが選んで指定した文字コードで読み、書いた所だけをその文字コードのバイトにしました"}
    readback_values = dict(values)  # 読み戻しで確かめる、欄の id → 書いた値
    told = None  # 本人に伝える、書いた所（ほかの形式は読み戻しの値と同じ）
    rest_fields = None  # まだ空いている欄を数える、欄の一覧（なければ、書いた後のファイルの欄）
    try:
        labels = {f["id"]: f.get("label", "") for f in fields(path, password, encoding)}  # 白紙のときの見出し（対話で伝える一覧に使う）
    except ApiError:
        labels = {}
    if k in ("xlsx", "xls"):
        from .book import open_book
        from .xlsx.edit import EditError
        from .xlsx.edit import fill as xfill
        from .xlsx.proof import prove

        circ = {k_: v for k_, v in values.items() if _is_circle(v)}
        cell_values = {k_: v for k_, v in values.items() if k_ not in circ}
        readback_values = {k_: v for k_, v in readback_values.items() if k_ not in circ}
        wb = open_book(blob, name)
        if cell_values or not circ:
            try:
                res = xfill(wb, cell_values, force=force)
            except EditError as e:
                raise ApiError(str(e)) from None
            allowed = res.allowed_inputs
            allowed_c = {(p.sheet, p.address): p.after for p in res.patches if p.kind == "control"}  # 部品 → 書いた状態
            pr = prove(blob, res.blob, allowed, allowed_c, {(p.sheet, p.address) for p in res.patches if p.kind == "macro"})
            new_blob = res.blob
            result.update({"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings),
                           "unexpected": pr.unexpected[:20], "structure_errors": pr.lint_errors[:20],
                           "recalculated": {f"{p.sheet}!{p.address}": p.after for p in res.patches if p.kind == "formula"}})
            # 文字の中の空き（セルの中の［　］）に値の並び（リスト）で書いたセル：読み戻しと本人に伝える値は、書いたセルの字
            from .formula.parser import addr as _addr
            from .xlsx.edit import resolve_target

            after_of = {(p.sheet, p.address): p.after for p in res.patches if p.kind == "input"}
            for key_, v_ in cell_values.items():
                if isinstance(v_, (list, tuple)):
                    try:
                        sn_, r_, c_ = resolve_target(wb, key_)
                    except EditError:
                        continue
                    if (sn_, _addr(r_, c_)) in after_of:
                        readback_values[key_] = after_of[(sn_, _addr(r_, c_))]
        else:
            new_blob = blob
            result.update({"proof_ok": True, "proof": "", "warnings": [], "unexpected": [], "structure_errors": [], "recalculated": {}})
        if circ:
            new_blob = _circle_excel(new_blob, name, k, circ, result)
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
        # 値（/V）と見た目の字が違う欄（選択肢の書き出しの値と見せる字・書式で整えた数）：読み戻しは値で、本人に伝えるのは見た目の字で
        pdf_v = {p.id: p.value for p in res.patches if p.kind in ("input", "control") and not p.anchor}
        pdf_shown = {p.id: (p.value, p.shown) for p in res.patches
                     if p.kind in ("input", "calc") and getattr(p, "shown", None) not in (None, p.value)}
        from .pdf.fill import appearance_readback

        pdf_ap_diff = appearance_readback(_pdf(res.blob, name, lock), res.patches)
        xfa_written = {p.id: p.value for p in res.patches if p.kind == "xfa"}
        if xfa_written:
            # XFA の欄：窓に出る値（日付・数は決まった形、選択肢はデータの値）で読み戻す
            readback_values.update(xfa_written)
        if anchored:
            # 見出しの字を目印にした書き込み：書いた欄の id（読み直した窓では、その名前のフォームの欄）
            result["anchored"] = {p.key: p.id for p in res.patches if p.anchor}
            readback_values = {k_: v for k_, v in values.items() if not k_.startswith("@")}
            readback_values.update({p.id: p.value for p in res.patches if p.anchor and p.kind not in ("seal", "circle")})
        circled = [p for p in res.patches if p.kind == "circle"]
        for p in circled:
            if not p.key.startswith("@") and p.circle.get("option"):
                readback_values[p.key] = p.circle["option"]  # 窓の選択肢に書いた所：書いた後の窓では、選んだ選択肢の名前
        if circled:
            # 印刷された字を○で囲んだ所：読み戻しは、書いた後のファイルのページを読み直して、下で確かめる
            result["circled"] = [{"key": p.key, "id": p.id, "page": p.page, "text": p.circle["text"],
                                  "label": p.circle["label"], "context": p.circle["context"], "how": "ページに楕円を描いた",
                                  "ellipse_pt": p.circle["ellipse"]} for p in circled]
        placed = [p for p in res.patches if p.kind == "seal"]
        if placed:
            # 本人の印影・署名の画像（本人が同意したときだけ）：何を、どこに置いたか。読み戻しは、書いた後のファイルで下で確かめる
            result["placed"] = [{"key": p.key, "id": p.id, "page": p.page, "what": p.seal["what"], "say": p.seal["told"],
                                 "box_mm": [round(v / 72 * 25.4, 1) for v in p.seal["rect"]], "rect_pt": p.seal["rect"], "image_sha256": p.seal["sha256"]}
                                for p in placed]
    elif k == "jtd" or (k == "doc" and _is_word6(blob)):
        # 一太郎・Word 6.0/95：同じ字数（バイト数）の上書きは、ファイル全体で変わったのがその字のバイトだけかを証明する。
        # 一太郎の字数を変える書き込み・圧縮した文書は、直すと決めた構造のほかが同じで、直した構造がそれぞれ思ったとおりかを証明する（jtd/proof.py）。
        # 書いた所の印（しおり）を足さないので、読み戻しは証明の中の「読み直した段落の文字」で確かめる
        from .docx.fill import FillError, written_of

        if k == "jtd":
            from .jtd.audit import prove as jprove
            from .jtd.book import JtdFile
            from .jtd.fill import fill as jfill
        else:
            from .doc.word6 import Word6File as JtdFile
            from .doc.word6 import fill as jfill
            from .doc.word6 import prove as jprove

        try:
            res = jfill(JtdFile(blob, name), values, force=force)
        except FillError as e:
            raise ApiError(str(e)) from None
        pr = jprove(blob, res.blob, written_of(res.patches))
        new_blob = res.blob
        result.update({"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings),
                       "unexpected": (pr.unexpected + pr.readback)[:20], "structure_errors": pr.lint_errors[:20]})
        anchored = {p.key: p.id for p in res.patches if p.anchor}
        if anchored:
            result["anchored"] = anchored
        readback_values = {}
        told = {p.id: p.value for p in res.patches}  # 一太郎：本人に伝える、書いた所と値
        result["jtd_read_back"] = pr.readback[:30]
    elif k in _open.KINDS:  # 追加の形式：書き込みと証明は、形式ごとの部品で
        try:
            res_ = _open.fill(k, blob, name, values, force, encoding)
        except _open.OpenFormatError as e:
            raise ApiError(str(e)) from None
        new_blob = res_["blob"]
        result.update(res_["result"])
        readback_values = res_["readback"]
        labels.update(res_.get("labels") or {})
        if res_.get("told") is not None:
            # 書いた所の印を置けない形式（.txt・PowerPoint）：本人に伝えるのは、証明で読み直した書いた値と、元のファイルの残りの欄
            told = res_["told"]
            rest_fields = res_.get("rest")
    else:
        from .docx.fill import FillError

        circ = {k_: v for k_, v in values.items() if _is_circle(v)}
        if circ and k == "doc":
            raise ApiError("古い Word（.doc）には、○の図形を足す書き方をまだ作っていません（図形を足したことを証明できないため）。"
                           "代わりに、字の前に「○」を足す（例：「○有・無」）か、選ばない字を消して書き直す書き方があります。"
                           "どちらにするか本人に確かめてから、字を目印にした書き込み（old・new）で書いてください")
        values = {k_: v for k_, v in values.items() if k_ not in circ}
        readback_values = {k_: v for k_, v in readback_values.items() if k_ not in circ}
        if values or not circ:
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
        else:
            new_blob = blob
            anchored = {}
            result.update({"proof_ok": True, "proof": "", "unexpected": [], "structure_errors": []})
        if anchored:
            # 字を目印にした書き込み：書いた所の id（読み直した窓では［P段落.w番号＝値］）
            result["anchored"] = anchored
            readback_values = {k: v for k, v in values.items() if not k.startswith("@")}
            readback_values.update({p.id: p.value for p in res.patches if getattr(p, "anchor", None)})
        if circ:
            new_blob = _circle_word(new_blob, name, circ, result)
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
    after_list = fields(str(dst), password, encoding)
    after = {f["id"]: f.get("value") for f in after_list}
    if k in _open.KINDS:  # 追加の形式：欄の一覧に出ないマス・値も、書いた後のファイルを読み直して照らす
        after.update(_open.read_back(k, written_plain, name, readback_values, encoding))
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
    if k == "pdf":
        readback_values.update({i: v for i, v in pdf_v.items() if i in readback_values})
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
    if k == "pdf":
        mismatch += pdf_ap_diff  # 見た目（/AP）の字：描ける字形で、見せるはずの字か
        for f_ in after_list:
            if f_["id"] in pdf_shown and str(f_.get("value")) == str(pdf_shown[f_["id"]][0]):
                f_["value"] = pdf_shown[f_["id"]][1]  # 本人に伝えるのは、書類に見える字（選択肢の名前・書式で整えた数）
    if result.get("placed"):
        mismatch += _read_back_placed(k, written_plain, name, lock, values, result["placed"])
    if result.get("circled"):
        mismatch += _read_back_circled(k, written_plain, name, lock, result["circled"])
    if "jtd_read_back" in result:
        mismatch += [{"id": "段落", "written": "", "read_back": x} for x in result.pop("jtd_read_back", [])]
    if "text_read_back" in result:  # .txt：証明の中で、書いた後のファイルを読み直した行
        mismatch += [{"id": "行", "written": "", "read_back": x} for x in result.pop("text_read_back", [])]
    if "para_read_back" in result:  # PowerPoint：証明の中で、書いた後のファイルを読み直した段落
        mismatch += [{"id": "段落", "written": "", "read_back": x} for x in result.pop("para_read_back", [])]
    result["read_back_ok"] = not mismatch
    result["read_back_differences"] = mismatch[:30]
    if told is not None:
        # 一太郎：書いた所の印がないので、読み直した窓の同じ id（残りの空白）ではなく、証明で確かめた書いた値を伝える
        rest = rest_fields if rest_fields is not None else [f for f in after_list if f["id"] not in told]
        result.update(_to_tell(rest, told, result, labels))
    else:
        result.update(_to_tell(after_list, readback_values, result, labels))
    return result


def _fill_member(path: str, values: dict, out_path: str, force: bool, password: str | None,
                 encoding: str | None = None) -> dict:
    """ZIP の中のファイルに書く：取り出したファイルに書いて証明し、ZIP はそのファイルの部品だけを入れ替えて書き出す。
    書き出した ZIP から取り出し直したファイルが、証明した書き込み後のファイルとバイト一致し、ほかの部品が元とバイト一致するかを確かめる。"""
    import tempfile

    from . import ziparc

    outer, chain = _member(path)
    src = Path(outer).expanduser().resolve()
    dst = Path(out_path).expanduser().resolve()
    if src == dst:
        raise ApiError("元のファイルは上書きしません。別の名前（out_path）を指定してください")
    original = src.read_bytes()
    layers = [original]  # 外側から順の ZIP
    try:
        for m in chain[:-1]:
            layers.append(ziparc.extract(layers[-1], m))
    except ziparc.ArchiveError as e:
        raise ApiError(str(e)) from None
    member_blob, name = _member_blob(path)
    with tempfile.TemporaryDirectory() as d:
        a, b = Path(d) / "a" / name, Path(d) / "b" / name
        a.parent.mkdir()
        b.parent.mkdir()
        a.write_bytes(member_blob)
        result = fill(str(a), values, str(b), force, password, encoding)
        written = b.read_bytes()
    # 内側から順に、書いた部品だけを入れ替える
    new = written
    for zblob, m in reversed(list(zip(layers, chain))):
        replaced = ziparc.replace(zblob, m, new)
        problems = ziparc.unchanged_except(zblob, replaced, m)
        if problems:
            raise ApiError("ZIP に入れ直すと、書いたファイルのほかの部品も変わります（書き出していません）：" + "、".join(problems[:5]))
        new = replaced
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_bytes(new)
    inner = ziparc.SEP.join(chain)
    back, _n = _member_blob(str(dst) + ziparc.SEP + inner)
    if back != written:
        dst.unlink(missing_ok=True)
        raise ApiError("書き出した ZIP から取り出したファイルが、書き込んで証明したファイルと一致しません（書き出していません）")
    result.update({"out_path": str(dst), "original_sha256": hashlib.sha256(original).hexdigest(),
                   "new_sha256": hashlib.sha256(new).hexdigest(), "member": inner,
                   "archive": "ZIP の中のファイルに書き、ZIP はそのファイルの部品だけを入れ替えました（ほかの部品は元とバイト一致）。"
                              f"書いたファイルは「{dst}{ziparc.SEP}{inner}」で開けます"})
    return result


def _is_circle(v) -> bool:
    """○で囲む指定（{"circle": 囲む字, …}）か。"""
    return isinstance(v, dict) and "circle" in v


def _circle_excel(blob: bytes, name: str, k: str, circ: dict, result: dict) -> bytes:
    """Excel：セルの字を○の図形で囲む。書いた後のブックで○を決め、図形を足し、足しただけであることを証明する。"""
    from .book import open_book
    from .xlsx import circle as xc

    if k == "xls":
        raise ApiError("古い Excel（.xls）には、○の図形を足す書き方をまだ作っていません（図形を足したことを証明できないため）。"
                       "代わりに、字の前に「○」を足す（例：「○有・無」）か、選ばない字を消して書き直す書き方があります。"
                       "どちらにするか本人に確かめてから、そのセルに字として書いてください")
    wb = open_book(blob, name)
    used: dict = {}
    sds = []
    problems = []
    for key, spec in circ.items():
        try:
            sds.append((key, xc.resolve(wb, spec, used)))
        except xc.CircleError as e:
            problems.append(f"{key}: {e}")
    if problems:
        raise ApiError("\n".join(problems) + "\n（代わりに、字の前に「○」を足す・選ばない字を消して書き直す書き方があります。"
                       "本人に確かめてから、そのセルに字として書いてください）")
    try:
        out = xc.apply(open_book(blob, name), [sd for _k, sd in sds])
    except xc.CircleError as e:
        raise ApiError(str(e)) from None
    bad = xc.prove(blob, out, name, [(sd["id"], circ[key_]) for key_, sd in sds])
    result["unexpected"] = list(result.get("unexpected") or []) + bad
    result["proof_ok"] = bool(result.get("proof_ok", True)) and not bad
    result["proof"] = (result.get("proof") or "") + (f"○ {len(sds)} つは、シートの図形に楕円を足しただけです（ほかの部品はバイト一致）"
                                                     if not bad else f"○の書き込みで想定外の変更 {len(bad)} 件")
    result["circled"] = [{"key": key_, "id": sd["id"], "sheet": sd["sheet"], "cell": sd["cell"], "text": sd["text"],
                          "label": sd["label"], "context": sd["context"], "how": "シートの図形に楕円を足した",
                          "uncertain": "位置は見積もり。印刷か画面で確かめが要る" + (f"：{'・'.join(sd['notes'])}" if sd["notes"] else "")}
                         for key_, sd in sds]
    w = list(result.get("warnings") or [])
    if xc.UNCERTAIN not in w:
        w.append(xc.UNCERTAIN)
    result["warnings"] = w
    result["anchored"] = {**(result.get("anchored") or {}), **{key_: sd["id"] for key_, sd in sds}}
    return out


def _circle_word(blob: bytes, name: str, circ: dict, result: dict) -> bytes:
    """Word（.docx）：本文の字を○で囲む（字に結び付けた楕円の図形、または字の囲み線）。足しただけであることを証明する。"""
    from .docx import circle as dc
    from .docx.document import DocxDocument

    doc = DocxDocument(blob, name)
    used: dict = {}
    sds, problems = [], []
    for key, spec in circ.items():
        try:
            sds.append((key, dc.resolve(doc, spec, used)))
        except dc.CircleError as e:
            problems.append(f"{key}: {e}")
    if problems:
        raise ApiError("\n".join(problems))
    out = dc.apply(doc, [sd for _k, sd in sds])
    bad = dc.prove(blob, out, name, [(sd["id"], circ[key_]) for key_, sd in sds])
    result["unexpected"] = list(result.get("unexpected") or []) + bad
    result["proof_ok"] = bool(result.get("proof_ok", True)) and not bad
    result["proof"] = (result.get("proof") or "") + (f"○ {len(sds)} つは、字の前に図形（または字の囲み線）を足しただけです"
                                                     "（ほかの部品はバイト一致）" if not bad else f"○の書き込みで想定外の変更 {len(bad)} 件")
    result["circled"] = [{"key": key_, "id": sd["id"], "text": sd["text"], "label": sd["label"], "context": sd["context"],
                          "style": sd["style"], "size": sd["size"],
                          "how": "字に結び付けた楕円の図形を足した" if sd["style"] == "ellipse" else "字に四角い囲み線を付けた（○ではなく□）",
                          "uncertain": "位置は見積もり。印刷か画面で確かめが要る" if sd["style"] == "ellipse" else ""}
                         for key_, sd in sds]
    if any(sd["style"] == "ellipse" for _k, sd in sds):
        w = list(result.get("warnings") or [])
        if dc.UNCERTAIN not in w:
            w.append(dc.UNCERTAIN)
        result["warnings"] = w
    result["anchored"] = {**(result.get("anchored") or {}), **{key_: sd["id"] for key_, sd in sds}}
    return out


def _read_back_circled(k: str, blob: bytes, name: str, lock, circled: list[dict]) -> list[dict]:
    """○で囲んだ所の読み戻し：書いた後のファイルを開き直し、○がその字だけを囲んでいるか（ほかの字に触れていないか）。"""
    if k == "pdf":
        from .pdf import circle

        return circle.read_back(_pdf(blob, name, lock), circled)
    if k in ("xlsx", "xls"):
        from .xlsx import circle as xcircle

        return xcircle.read_back(blob, name, circled)
    if k in ("docx",):
        from .docx import circle as dcircle

        return dcircle.read_back(blob, name, circled)
    return []


def _seal_supported(k: str, values: dict) -> None:
    """本人の印影・署名の画像（seal・signature）を置けるのは、いまは PDF だけ。ほかの形式は、証明つきで置く方法を
    作るまで、書かずに断る（証明のない書き込みはしない）。"""
    from .sealimage import KINDS, SealError, seal_key

    if k == "pdf":
        return
    for key, v in values.items():
        try:
            sk = seal_key(v)
        except SealError as e:
            raise ApiError(f"{key}: {e}") from None
        if sk is not None:
            raise ApiError(f"{key}: 本人の{KINDS[sk]}の画像を置けるのは、いまは PDF の書類だけです（{KIND_JA.get(k, k)}には、"
                           "変わったのが置いた画像だけであることを証明して置く方法を、まだ作っていません）。"
                           "書いていません。PDF の様式があればそれを使うか、このことを本人に伝えてください")


def _read_back_placed(k: str, blob: bytes, name: str, lock, values: dict, placed: list[dict]) -> list[dict]:
    """置いた画像の読み戻し：書いた後のファイルを開き直し、その名前の画像が、伝えた位置に、本人の画像のとおりにあるか。"""
    from .sealimage import SealError, decode

    out = []
    if k != "pdf":
        return out
    from .pdf import seal

    doc = _pdf(blob, name, lock)
    for pl in placed:
        spec = values.get(pl["key"]) or {}
        key = "seal" if "seal" in spec else "signature"
        try:
            img = decode(Path(str(spec.get(key))).expanduser().read_bytes())
        except (SealError, OSError) as e:
            out.append({"id": pl["id"], "written": pl["what"], "read_back": f"画像を読み直せません：{e}"})
            continue
        # 位置は、置いたときに決めた矩形（見た目の座標、pt）。ファイルの注釈の位置と照らす
        probs, _nums = seal.verify(doc, pl["page"], {"id": pl["id"], "rect": pl["rect_pt"]}, key, img)
        if img.sha256 != pl["image_sha256"]:
            probs.append("画像のファイルが、置いたときと変わっています")
        if probs:
            out.append({"id": pl["id"], "written": pl["what"], "read_back": "／".join(probs)[:300]})
    return out


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

    def bare_id(d: dict) -> str:
        i = d["id"]
        return i.split("!", 1)[1] if d.get("sheet") and i.startswith(d["sheet"] + "!") else i

    def name(d: dict) -> str:
        """読み上げても分かる欄の名前：窓の印（［▶］・［　］）は「何番目の欄」に、見出しのない欄はセルの番地にする。"""
        lab = d.get("label") or ""
        if not lab or lab == d["id"]:
            return f"{bare_id(d)}（見出しなし）"
        slots = list(re.finditer(r"［[^］]*］", lab))
        if not slots:
            return lab
        k = next((n for n, m in enumerate(slots, 1) if "▶" in m.group()), 0)
        base = re.sub(r"［[^］]*］", "", lab).rstrip("-－ー：:　 ")
        return f"{base}（{k}つ目の欄）" if k and len(slots) > 1 else base

    def shown(d: dict) -> str:
        v = "☑" if d["value"] is True else "☐" if d["value"] is False else str(d["value"])
        lab = d["label"] or ""
        m = re.search(r"［▶[^］]*］", lab) or re.search(r"［[ 　]*］", lab)
        if lab and m:
            rest = re.sub(r"［[^］]*］", "［ ］", lab[m.end():])
            head = re.sub(r"［[^］]*］", "［ ］", lab[:m.start()])
            return f"{head}［{v}］{rest}"  # 「令和［8］年［ ］月」：空欄に値を入れた形で
        return f"{name(d)}：{v}"

    # 記入した内容の一覧：書いた所を、書類の中の順（シート・ページごと）に、全部。値は書いた後のファイルを読み直したもの
    order = {f["id"]: i for i, f in enumerate(after_list)}
    told.sort(key=lambda d: order.get(d["id"], len(order)))
    def joined(run: list[dict]) -> str:
        """1つの項目が分かれた欄（〒［▶］-［　］と〒［　］-［▶］）は、1行に：〒［123］-［4567］。"""
        lab = run[0]["label"]
        vals: dict[int, str] = {}
        for d in run:
            k = next((n for n, m in enumerate(re.finditer(r"［[^］]*］", d["label"])) if "▶" in m.group()), None)
            vals[k] = "☑" if d["value"] is True else "☐" if d["value"] is False else str(d["value"])
        it = iter(range(10 ** 6))
        return re.sub(r"［[^］]*］", lambda m: f"［{vals.get(next(it), ' ')}］", lab)

    def stem(d: dict) -> str | None:
        lab = d.get("label") or ""
        return place(d) + re.sub(r"［[^］]*］", "［］", lab) if "［▶" in lab and lab.count("［") > 1 else None

    lines = [f"記入した内容（{len(told)} か所。書いた後のファイルを読み直した値）："] if told or not (result.get("placed") or result.get("circled")) else []
    group = None
    i = 0
    while i < len(told):
        d = told[i]
        g = place(d).rstrip("・")
        if g and g != group:
            lines.append(f"【{g}】")
            group = g
        j = i + 1
        while stem(d) and j < len(told) and stem(told[j]) == stem(d):
            j += 1
        lines.append(f"- {joined(told[i:j]) if j - i > 1 else shown(d)}")
        i = j
    for pl in result.get("placed") or []:
        lines.append(pl["say"])  # 置いた画像（本人の印影・署名）：何を、どこに
    told_ids = {d["id"] for d in told}
    for c in result.get("circled") or []:
        if c.get("key") in told_ids:
            continue  # 窓の「○で囲む選択肢」に書いた所：上の一覧に、選んだ字として出ている
        where = c.get("sheet") or (f"{c['page']}ページ" if c.get("page") else "")
        lab = c.get("label") or ""
        verb = "□で囲んだ（字の囲み線）" if c.get("style") == "border" else "○で囲んだ"
        lines.append(f"- {verb}：{where + '・' if where else ''}{lab + '：' if lab and lab != c['text'] else ''}「{c['text']}」"
                     + (f"（{c['uncertain']}）" if c.get("uncertain") else ""))
    if empty:
        # 次に書くとよい欄：書いた所の続き（同じシート・ページで、書いた所より後ろの空欄）から順に
        order = {f["id"]: i for i, f in enumerate(after_list)}
        last = max((order[d["id"]] for d in told if d["id"] in order), default=-1)
        where_last = next((after_list[last].get("sheet") or after_list[last].get("page") for _ in [0] if last >= 0), None)
        nxt = [f for f in empty if order.get(f["id"], -1) > last and (f.get("sheet") or f.get("page")) == where_last] or empty
        counts: dict = {}
        for f in empty:
            w = f.get("sheet") or (f"{f['page']}ページ" if f.get("page") else "")
            counts[w] = counts.get(w, 0) + 1
        by = "、".join(f"{w} {n} か所" for w, n in list(counts.items())[:6] if w) + ("ほか" if len(counts) > 6 else "")
        lines.append(f"まだ空いている欄：{len(empty)} か所" + (f"（{by}）" if by and len(counts) > 1 else ""))
        same = len({place(f) for f in nxt[:10]}) == 1
        head = f"続きの空欄（{place(nxt[0]).rstrip('・')}）：" if same and place(nxt[0]) else "続きの空欄："
        lines.append(head + "、".join(name(f) if same else f"{place(f)}{name(f)}" for f in nxt[:10]) + ("ほか" if len(nxt) > 10 else ""))
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


def audit(path: str, password: str | None = None, encoding: str | None = None) -> dict:
    import inspect

    from .unlock import pdf_password

    blob, name, k, lock = _read(path, password)
    encoding = _encoding(blob, k, encoding)
    if k == "zip":
        return _audit_zip(path, blob, password)
    if k == "pdf":
        from .pdf.audit import run
    elif k == "docx":
        from .docx.audit import run
    elif k == "doc":
        from .doc.audit import run
    elif k == "jtd":
        from .jtd.audit import run
    elif k in _open.KINDS:  # 追加の形式
        run = _open.audit_runner(k, encoding)
    else:
        from .xlsx.audit import run
    pw = pdf_password(lock)
    au = run(blob, name, password=pw) if pw and "password" in inspect.signature(run).parameters else run(blob, name)
    d = au.to_dict()
    d["grade"] = grade(path, password=password, encoding=encoding)
    if lock is not None:
        d["password_protected"] = f"パスワードで保護された書類です（{lock.label}）。解いた中身を確かめました"
    return d


def read_table(path: str, header: str, sheet: str | None = None, header_rows: int | None = None,
               out_path: str | None = None, password: str | None = None) -> dict:
    """Excel の表を、見出しの列で行ごとに取り出す（結合セルは画面どおりに範囲へ広げ、値ごとに元のセルを添える）。
    sheet を省くと、header のあるすべてのシートから。out_path（.csv か .json）を指すと、すべての行を書き出す。"""
    import csv
    import io
    import json

    from .book import open_book
    from .xlsx.tables import TableError
    from .xlsx.tables import read as tread

    blob, name, k, lock = _read(path, password)
    if k not in ("xlsx", "xls"):
        raise ApiError("表の取り出しは Excel（.xlsx・.xlsm・.xls）です（ほかの形式は窓で読めます）")
    try:
        t = tread(open_book(blob, name), header, sheet, header_rows)
    except TableError as e:
        raise ApiError(str(e)) from None
    out = {"columns": t["columns"], "rows": len(t["rows"]), "sheets": t["sheets"][:50], "sheet_count": len(t["sheets"]),
           "first_rows": t["rows"][:20], "check": t["check"]}
    if out_path:
        dst = Path(out_path).expanduser().resolve()
        src = Path(_member(path)[0] if _member(path) else path).expanduser().resolve()
        if dst == src:
            raise ApiError("元のファイルは上書きしません。別の名前（out_path）を指定してください")
        dst.parent.mkdir(parents=True, exist_ok=True)
        if dst.suffix.lower() == ".json":
            dst.write_text(json.dumps(t, ensure_ascii=False, indent=1), encoding="utf-8")
        else:
            buf = io.StringIO()
            w = csv.writer(buf, lineterminator="\n")
            head = ["シート", "行"] + t["columns"]
            w.writerow(head)
            for r in t["rows"]:
                w.writerow([r["_sheet"], r["_row"]] + [r.get(c, "") for c in t["columns"]])
            dst.write_text(buf.getvalue(), encoding="utf-8")
        out["out_path"] = str(dst)
    else:
        out["all_rows"] = t["rows"]
    return out


def calc_model(path: str, cell: str | None = None, password: str | None = None) -> dict:
    """Excel の様式の「データ」（入力欄）と「処理」（入力につながる数式）。cell を指すと、その計算を入力欄までさかのぼる。
    式は、番地を見出しに置き換えて示す（in_words）。"""
    from .book import open_book
    from .xlsx import model as xm

    blob, name, k, lock = _read(path, password)
    if k not in ("xlsx", "xls"):
        raise ApiError("計算の取り出しは Excel の様式（.xlsx・.xlsm・.xls）です")
    _k, doc, _ = window(path, password)
    m = xm.model(open_book(blob, name), doc)
    if cell:
        try:
            return {"explain": xm.explain(m, cell)}
        except ValueError as e:
            raise ApiError(str(e)) from None
    return {"inputs": len(m["inputs"]), "steps": len(m["steps"]), "results": len(m["results"]),
            "dynamic": m["dynamic"], "result_groups": xm.summary(m)}


def calculate(path: str, values: dict, password: str | None = None) -> dict:
    """値の組から、ファイルを書かずに結果を計算する（試算）。書き込みと同じ受け取り方・入力規則・計算エンジン。"""
    from .book import open_book
    from .xlsx import model as xm
    from .xlsx.edit import EditError

    blob, name, k, lock = _read(path, password)
    if k not in ("xlsx", "xls"):
        raise ApiError("試算は Excel の様式（.xlsx・.xlsm・.xls）です")
    _k, doc, _ = window(path, password)
    try:
        r = xm.calculate(open_book(blob, name), doc, values)
    except EditError as e:
        raise ApiError(str(e)) from None
    lines = [f"試算（ファイルは書いていません）：入れた値 {len(r['inputs'])} か所で、計算が変わる所 {len(r['changed'])} か所。"]
    for c in r["changed"][:40]:
        lines.append(f"- {c['label'] or c['cell']}：{c['before'] or '（空）'} → {c['after'] or '（空）'}")
    if r["not_computed"]:
        lines.append(f"計算し直せなかった所 {len(r['not_computed'])} か所（Excel で開いたときに計算されます）："
                     + "、".join(x["cell"] for x in r["not_computed"][:10]))
    r["say_to_user"] = "\n".join(lines)
    return r


def view_page(path: str, page: int, region: list | None = None, marks: list | None = None,
              password: str | None = None) -> dict:
    """PDF のページの画像（スキャンした紙のページを、エージェントが見て読み、書く所を mm の枠で決めるため）。
    todoku が書き込んだ欄（p ページ.w 番号など）は、青い枠で描き、id と値を添える。"""
    from .pdf.pageimage import view_page as pview

    blob, name, k, lock = _read(path, password)
    if k != "pdf":
        raise ApiError("view_page は PDF のページの画像です（ほかの形式は窓で読めます）")
    doc = _pdf(blob, name, lock)
    try:
        return pview(doc, int(page), region, marks, written=_written_boxes(path, password, int(page)))
    except ValueError as e:
        raise ApiError(str(e)) from None


def _written_boxes(path: str, password: str | None, page: int) -> list[dict]:
    """ページの、値の入った入力欄（フォームの欄・todoku が書いた欄）の枠（mm）と値。"""
    from .pdf.pageimage import MM

    out = []
    k, doc, _ = window(path, password)
    for pg in doc.get("pages", []):
        if pg["number"] != page:
            continue
        for f in pg["fields"]:
            v = f.get("value")
            r = f.get("rect")
            if v in (None, "", False) or not r:
                continue
            out.append({"box": [r[0] / MM, r[1] / MM, r[2] / MM, r[3] / MM], "id": f["id"], "value": v})
    return out


def _audit_zip(path: str, blob: bytes, password: str | None) -> dict:
    """ZIP：中の、扱える形式のファイルを1つずつ検査する（検査の名前の頭に、中の名前を付ける）。"""
    from . import ziparc

    checks = []
    for m in ziparc.files(blob):
        try:
            d = audit(path + ziparc.SEP + m.name, password)
        except ApiError as e:
            checks.append({"name": f"{m.name}：開けない", "ok": False, "agent": False, "detail": str(e), "items": []})
            continue
        checks += [{**c, "name": f"{m.name}：{c['name']}"} for c in d["checks"]]
    return {"file": Path(path).name, "ok": all(c["ok"] for c in checks), "needs_agent": any(c["agent"] for c in checks),
            "checks": checks, "grade": grade(path, password=password)}


def grade(path: str, findings: list[dict] | None = None, answered: list[str] | None = None,
          password: str | None = None, encoding: str | None = None) -> dict:
    """段階（S〜F）：誰が扱えば正確に読んで入れられるか。規則で決めた段階と、エージェントが確かめること。
    findings（エージェントが窓を読んで見つけたこと：{grade, title, detail, where}）を渡すと、段階に合わせる。
    パスワードで保護された書類は、本人に聞いたパスワードで解いた中身で決める（パスワードがなければ PasswordNeeded）。"""
    from .grade import combine, grade_blob
    from .unlock import pdf_password

    blob, name, k, lock = _read(path, password)
    encoding = _encoding(blob, k, encoding)
    g = grade_blob(blob, name, password=pdf_password(lock), encoding=encoding)
    if findings or answered:
        g = combine(g, findings or [], answered)
    return g.to_dict()


def summary(path: str, password: str | None = None, encoding: str | None = None) -> dict:
    k, doc, parts = window(path, password, encoding)
    fl = fields(path, password, encoding)
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
    elif k == "zip":
        d["members"] = [{"name": m["name"], "kind": m["kind"], "size": m["size"]} for m in doc["members"]]
    else:
        d["protection"] = doc.get("protection")
    if k in _open.KINDS:  # 追加の形式
        d.pop("protection", None)
        d.update(_open.summary(k, doc))
    return d
