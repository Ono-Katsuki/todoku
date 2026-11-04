"""配置から見つけた所に書いた字を、ページの描画の後ろに足す（増分の保存）。どのビューアーでも同じ字形が、そのまま見える。

なぜ欄の見た目（/AP）だけにしないか（実物で測った。tests/audit_findings/pdfkit/ に測り方）：
- Apple の PDFKit（Mac のプレビュー・iPhone などの標準の表示）は、文字の欄の /AP を使わず、/V から描き直す。字の幅に余白を足した幅が
  ないと語の切れ目で折り返し、2 行目は欄の外なので描かれない（昭和60 → 昭和、H20 → H2）。字の大きさも欄の高さで変わる
- Quartz（CGContext.drawPDFPage）は注釈をどれも描かない
  ページの描画なら、どちらも、Poppler・pdf.js も、同じように描く

書き方：
- 欄の字の見た目（フォームの XObject）を、ページの描画の後ろのストリームから `Do` で描く（○と同じ、前後に足すストリーム）。
  元のストリームは 1 バイトも変えない。ページの /Resources には、足した XObject の名前（TdkW…）だけを足す
- 値の控えとして、同じ所に隠した欄（/F の Hidden、値 /V と見た目 /AP）を残す。見た目は描かないが、値を機械で読め、窓・読み戻し・
  直して書き直すときに使う。/TdkDrawn が、その値を描いた XObject を指す
- 読むときは、値と描いた字が一致する欄の XObject だけを、印刷された字として読まない（verified_refs）。一致しないものは、
  書類の字として読む（書類の作り手が、この印で字を読み手から隠すことはできない）
"""

from __future__ import annotations

import re

from .objects import PdfString, Ref, Stream, decode_text, dumps, text_string

KEY = "TdkDrawn"  # 欄の部品の辞書：その値を描いた、ページの描画の XObject
PREFIX = "TdkW"
ANNOT_HIDDEN = 2


def draw_ops(items: list[tuple[str, Ref, float, float]], taken) -> tuple[bytes, dict[str, Ref]]:
    """items：(欄の名前, 描く XObject, 欄の左下 x, y)。taken：ページの /XObject にすでにある名前。
    返す：(足すストリームの命令, 足す /XObject の項目)。証明が同じ決め方で照らすので、決まった順に、決まった形で。"""
    out = bytearray()
    names: dict[str, Ref] = {}
    used = set(taken)
    n = 0
    for name, ref, x, y in items:
        n += 1
        key = f"{PREFIX}{n}"
        while key in used:
            n += 1
            key = f"{PREFIX}{n}"
        used.add(key)
        names[key] = ref
        out += (b"/Span <</TdkField " + dumps(text_string(name)) + b">> BDC\nq\n"
                + f"1 0 0 1 {x:.3f} {y:.3f} cm\n/{key} Do\nQ\nEMC\n".encode())
    return bytes(out), names


def widget_drawn(doc, wd) -> Ref | None:
    """部品の辞書の、描いた XObject（なければ None）。"""
    v = wd.get(KEY) if isinstance(wd, dict) else None
    return v if isinstance(v, Ref) else None


def origin(doc, wd) -> tuple[float, float]:
    r = [float(doc.resolve(v)) for v in doc.resolve(wd.get("Rect"))]
    return min(r[0], r[2]), min(r[1], r[3])


def items_of(doc, page: dict, nums: set[int] | None = None) -> list[tuple[str, Ref, float, float]]:
    """ページの注釈の順に、値を描いた欄の (名前, XObject, x, y)。nums があれば、その番号の部品だけ。"""
    out = []
    for a in doc.resolve(page.get("Annots")) or []:
        if nums is not None and not (isinstance(a, Ref) and a.num in nums):
            continue
        wd = doc.resolve(a)
        ref = widget_drawn(doc, wd)
        if ref is None:
            continue
        t = doc.resolve(wd.get("T"))
        out.append((decode_text(t) if isinstance(t, (bytes, PdfString)) else "", ref, *origin(doc, wd)))
    return out


def value_problem(doc, wd) -> str | None:
    """欄の値と、ページに描いた字が合うか。合わなければ理由。文字の欄は /V と描いた字、チェックは「入」のときに状態の見た目と同じ印。"""
    from .glyphs import appearance_text

    ref = widget_drawn(doc, wd)
    st = doc.resolve(ref) if ref is not None else None
    if not isinstance(st, Stream):
        return "描いた見た目（フォームの XObject）がありません"
    ft = getattr(doc.resolve(wd.get("FT")), "name", "")
    if ft == "Btn":
        state = getattr(doc.resolve(wd.get("AS")), "name", "Off")
        ap = doc.resolve(wd.get("AP"))
        n = doc.resolve(ap.get("N")) if isinstance(ap, dict) else None
        want = doc.resolve(n.get(state)) if isinstance(n, dict) else None
        if not isinstance(want, Stream):
            return f"状態 {state} の見た目がありません"
        if doc.decode(want) != doc.decode(st):
            return f"ページに描いた印が、状態 {state} の見た目と違います"
        return None
    v = doc.resolve(wd.get("V"))
    want = decode_text(v) if isinstance(v, (bytes, PdfString)) else ""
    got, probs = appearance_text(doc, st)
    if probs:
        return "描いた字：" + "／".join(probs)[:200]
    if re.sub(r"\s", "", got) != re.sub(r"\s", "", want):
        return f"値は「{want}」なのに、ページに描いた字は「{got}」です"
    return None


def verified_refs(doc, page: dict) -> set[int]:
    """印刷された字として読まない XObject の番号：このページの、値と描いた字が一致する欄が指すもの。"""
    if not isinstance(page, dict) or page.get("Annots") is None:
        return set()
    out: set[int] = set()
    try:
        annots = doc.resolve(page.get("Annots")) or []
    except Exception:  # noqa: BLE001
        return out
    for a in annots:
        try:
            wd = doc.resolve(a)
            ref = widget_drawn(doc, wd)
            if ref is not None and value_problem(doc, wd) is None:
                out.add(ref.num)
        except Exception:  # noqa: BLE001  読めない欄は、一致を確かめられない：書類の字として読む
            continue
    return out
