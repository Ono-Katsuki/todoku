"""電子署名された PDF：署名が許す変更（DocMDP・FieldMDP・/Lock）を読み、書き込みをその範囲に限る。

- 認証の署名（Catalog の /Perms /DocMDP）の P（ISO 32000-1 12.8.2.2）
  P=1：変更を許さない → 書かない（書く段階 D）
  P=2：フォームの記入と署名だけ → 今ある欄に書くだけ。新しい欄・注釈（字を目印にした書き込み・枠・印影）・ページの描画（○）は足さない
  P=3：P=2 に加えて注釈の作成・変更 → 印影（スタンプの注釈）は置ける。新しい欄とページの描画は足さない
- 署名した欄の /Lock（SigFieldLock）と、署名の /Reference の FieldMDP（12.8.2.4・12.7.4.5）：固定された欄は読み取り専用として断る
- 署名があれば（どの種類でも）、書き込みは増分で足すので署名のバイトは変わらないが、ビューアーでは「署名のあとの変更」として表示される。
  そのことを本人に伝える
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .objects import PdfString, Ref, decode_text


@dataclass
class Signatures:
    signed: list[str] = field(default_factory=list)  # 署名済みの署名の欄の名前
    docmdp: int | None = None  # 認証の署名の P（なければ None）
    locks: list[tuple[str, str, list[str]]] = field(default_factory=list)  # (署名の欄, Action, Fields)

    def locked_by(self, name: str) -> str | None:
        """欄 name を固定している署名の欄の名前（固定されていなければ None）。"""
        for sig, action, fields in self.locks:
            hit = any(name == f or name.startswith(f + ".") for f in fields)
            if action == "All" or (action == "Include" and hit) or (action == "Exclude" and not hit):
                if name != sig:
                    return sig
        return None

    def note(self) -> str | None:
        if not self.signed:
            return None
        who = "・".join(self.signed[:3]) + ("ほか" if len(self.signed) > 3 else "")
        return (f"この PDF には電子署名（{who}）があります。書き込みは元のバイトの後ろに足すので署名そのものは壊しませんが、"
                "Acrobat などでは「署名のあとの変更」として表示されます。署名した人・出す先が、署名のあとの記入を認めているか確かめてください")


def _names(doc, arr) -> list[str]:
    out = []
    for x in doc.resolve(arr) or []:
        x = doc.resolve(x)
        if isinstance(x, (bytes, PdfString)):
            out.append(decode_text(x))
    return out


def read(doc) -> Signatures:
    from .form import read_fields

    s = Signatures()
    perms = doc.resolve(doc.root.get("Perms"))
    if isinstance(perms, dict):
        sig = doc.resolve(perms.get("DocMDP"))
        if isinstance(sig, dict):
            for ref in doc.resolve(sig.get("Reference")) or []:
                r = doc.resolve(ref)
                if not isinstance(r, dict) or getattr(doc.resolve(r.get("TransformMethod")), "name", "") != "DocMDP":
                    continue
                tp = doc.resolve(r.get("TransformParams")) or {}
                p = doc.resolve(tp.get("P")) if isinstance(tp, dict) else None
                s.docmdp = int(p) if isinstance(p, (int, float)) and int(p) in (1, 2, 3) else 2  # 既定は 2
    for f in read_fields(doc):
        if f.kind != "signature" or f.ref is None:
            continue
        v = doc.resolve(f.value) if isinstance(f.value, Ref) else f.value
        node = doc.resolve(f.ref)
        if not isinstance(v, dict):
            continue  # 空の署名の欄：まだ署名していないので、固定もない
        s.signed.append(f.name)
        lock = doc.resolve(node.get("Lock")) if isinstance(node, dict) else None
        if isinstance(lock, dict):
            act = getattr(doc.resolve(lock.get("Action")), "name", "")
            s.locks.append((f.name, act, _names(doc, lock.get("Fields"))))
        for ref in doc.resolve(v.get("Reference")) or []:
            r = doc.resolve(ref)
            if not isinstance(r, dict) or getattr(doc.resolve(r.get("TransformMethod")), "name", "") != "FieldMDP":
                continue
            tp = doc.resolve(r.get("TransformParams")) or {}
            if isinstance(tp, dict):
                act = getattr(doc.resolve(tp.get("Action")), "name", "")
                s.locks.append((f.name, act, _names(doc, tp.get("Fields"))))
    if s.docmdp is not None and not s.signed:
        s.signed.append("認証の署名")
    return s


def refuse_new(sigs: Signatures, what: str, annotation: bool = False) -> str | None:
    """新しい欄・注釈・ページの描画を足す書き込みを、認証の署名が許すか。許さなければ理由。
    what：何を足すか（「新しい欄」など）。annotation：足すのが（欄でない）注釈か（P=3 なら許す）。"""
    p = sigs.docmdp
    if p is None:
        return None
    if p == 3 and annotation:
        return None
    allowed = {1: "どの変更も", 2: "今ある欄への記入と署名のほか", 3: "今ある欄への記入・署名・注釈のほか"}[p]
    return (f"この PDF は認証の電子署名（DocMDP P={p}）がされていて、{allowed}は許されていません。"
            f"{what}を足すと、署名が無効（許されない変更）と表示されるので、書きません。今ある欄（窓のフォームの欄）にだけ書けます")


def write_reasons(doc):
    """書く段階の理由（grade）：変更を許さない認証の署名は D。"""
    from ..grade import Reason

    try:
        s = read(doc)
    except Exception as e:  # noqa: BLE001  確かめられないものを「署名なし」にしない
        return [Reason("C", "署名を確かめられない",
                       f"電子署名と、署名が許す変更（DocMDP・FieldMDP）を読めませんでした（{type(e).__name__}: {str(e)[:60]}）。"
                       "署名があれば、書き込みで署名が無効と表示されるかもしれません。書く前に、署名のある書類かを本人に確かめる")]
    if s.docmdp == 1:
        return [Reason("D", "変更を許さない電子署名", "認証の電子署名（DocMDP P=1）がされていて、どの書き込みも署名を無効にするので、書けません")]
    return []
