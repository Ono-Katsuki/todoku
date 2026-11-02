"""PDF：エージェントがこの書類を「全て正確に読み書きできるか」を、決定的な処理で確かめる。

読む: ページの描画の命令から、窓とは別の読み方（並べずに、文字の命令だけを順に）で全部の字を拾い、窓に1字残らず出ているかを確かめる。
      文字に直せない字・画像だけのページ・注釈の文字も確かめる。
書く: 窓が示したすべての入力欄に試しの値を一度に書き込み、元のバイト列がそのまま残っていること、
      変わったのが書いた欄とその見た目だけであること、修復なしに開けること、書いた値が窓に読み戻せることを確かめる。
"""

from __future__ import annotations

from collections import Counter

from ..xlsx.audit import Audit, Check
from .content import tokens
from .doc import PdfDocument
from .fill import FillError, fill
from .fonts import load_font
from .form import read_fields
from .objects import Name, PdfString, Stream, decode_text
from .proof import prove
from .window import build, markdown_parts

BOXES = set("□☐☑☒◻▢❏❐❑❒⬜")


def raw_chars(doc: PdfDocument, page: dict) -> list[str]:
    """ページの文字の命令を、並べずに順に読む（窓とは別の読み方）。ActualText の範囲は、その文字にする。"""
    out: list[str] = []

    def run(data: bytes, res: dict, depth: int):
        if depth > 12:
            return
        fonts = doc.resolve(res.get("Font")) or {}
        xobjs = doc.resolve(res.get("XObject")) or {}
        props = doc.resolve(res.get("Properties")) or {}
        font = None
        stack: list = []
        nest: list[list] = []  # 配列・辞書の入れ子（["[", 中身] / ["<<", 中身]）
        mc: list = []  # (ActualText, 文字の数)
        saved: list = []  # q / Q で保存・復元する書体
        egs = doc.resolve(res.get("ExtGState")) or {}
        for kind, val in tokens(data):
            if kind in ("[", "<<"):
                nest.append([kind, []])
                continue
            if kind in ("]", ">>"):
                if not nest:
                    continue
                k0, items = nest.pop()
                if k0 == "<<":
                    obj = {items[i].name: items[i + 1] for i in range(0, len(items) - 1, 2) if isinstance(items[i], Name)}
                else:
                    obj = items
                (nest[-1][1] if nest else stack).append(obj)
                continue
            if kind != "o":
                (nest[-1][1] if nest else stack).append(val)
                continue
            op = val
            strings = []
            if op == "q":
                saved.append(font)
            elif op == "Q":
                if saved:
                    font = saved.pop()
            elif op == "Tf" and len(stack) >= 2 and isinstance(stack[-2], Name):
                ref = fonts.get(stack[-2].name) if isinstance(fonts, dict) else None
                font = load_font(doc, ref, stack[-2].name) if ref is not None else None
            elif op == "gs" and stack and isinstance(stack[-1], Name) and isinstance(egs, dict):
                g = doc.resolve(egs.get(stack[-1].name))
                fnt = doc.resolve(g.get("Font")) if isinstance(g, dict) else None
                if isinstance(fnt, list) and len(fnt) == 2:
                    font = load_font(doc, fnt[0], "gs")
            elif op in ("Tj", "'") and stack and isinstance(stack[-1], (bytes, PdfString)):
                strings = [stack[-1]]
            elif op == '"' and stack and isinstance(stack[-1], (bytes, PdfString)):
                strings = [stack[-1]]
            elif op == "TJ" and stack and isinstance(stack[-1], list):
                strings = [x for x in stack[-1] if isinstance(x, (bytes, PdfString))]
            elif op == "Do" and stack and isinstance(stack[-1], Name):
                xo = doc.resolve(xobjs.get(stack[-1].name)) if isinstance(xobjs, dict) else None
                if isinstance(xo, Stream) and getattr(xo.dict.get("Subtype"), "name", "") == "Form":
                    try:
                        run(doc.decode(xo), doc.resolve(xo.dict.get("Resources")) or res, depth + 1)
                    except Exception:  # noqa: BLE001
                        pass
            elif op in ("BDC", "BMC"):
                p = stack[-1] if op == "BDC" and stack else None
                if isinstance(p, Name):
                    p = doc.resolve(props.get(p.name)) if isinstance(props, dict) else None
                at = p.get("ActualText") if isinstance(p, dict) else None
                mc.append((decode_text(at) if isinstance(at, (bytes, PdfString)) else None, len(out)))
            elif op == "EMC" and mc:
                at, start = mc.pop()
                if at is not None and len(out) > start:
                    del out[start:]
                    out.append(at)
            if font is not None:
                for s in strings:
                    for _c, t, *_ in font.decode(bytes(s)):
                        if t:
                            out.append(t)
            stack.clear()

    run(doc.page_content(page), doc.resolve(page.get("Resources")) or {}, 0)
    return out


def _count(texts) -> Counter:
    c = Counter()
    for t in texts:
        for ch in t:
            if not ch.isspace() and ch not in BOXES:
                c[ch] += 1
    return c


def check_reading(doc: PdfDocument, win: dict) -> list[Check]:
    missing, extra, notes = [], [], []
    dup_total = 0
    for pg, page in zip(win["pages"], doc.pages()):
        raw = _count(raw_chars(doc, page))
        shown = _count(pg["_plain"])
        dropped = _count(t for _k, t in pg["_dropped"])
        dup_total += sum(1 for k, _ in pg["_dropped"] if k == "dup")
        want = raw - dropped
        for ch, n in (want - shown).items():
            missing.append(f"ページ {pg['number']}: 「{ch}」が {n} 字、窓に出ていません")
        for ch, n in (shown - raw).items():
            extra.append(f"ページ {pg['number']}: 「{ch}」が {n} 字、元にないのに窓に出ています")
    total = sum(p["chars"] for p in win["pages"])
    out = [Check("読み落とし（ページの文字）", not missing and not extra,
                 f"ページの文字 {total:,} 字を、窓とは別の読み方（文字の命令を順に）で拾い、すべて窓に出ていることを確かめました"
                 + (f"（同じ字を同じ位置に重ねた太字などの {dup_total:,} 字は1字として読んでいます）" if dup_total else ""),
                 missing + extra)]
    unmapped = sum(p["unmapped"] for p in win["pages"])
    out.append(Check("文字に直せない字", unmapped == 0,
                     "すべての字を文字（Unicode）に直せました" if not unmapped else
                     f"{unmapped} 字は、フォントに文字の対応がなく、文字に直せませんでした（窓では �）",
                     [f"ページ {p['number']}: {p['unmapped']} 字" for p in win["pages"] if p["unmapped"]]))
    img = [p["number"] for p in win["pages"] if p.get("image_only")]
    out.append(Check("画像だけのページ", not img,
                     "すべてのページに文字の情報があります" if not img else
                     f"{len(img)} ページは画像だけで、文字の情報がありません（画像を見て読む必要があります）",
                     [f"ページ {n}" for n in img]))
    return out


def _sample(f: dict):
    k = f["kind"]
    if k == "checkbox":
        return not bool(f.get("value"))
    if k in ("circle", "radio"):
        opts = f.get("options") or []
        cur = f.get("value")
        for o in reversed(opts):
            if o != cur:
                return o
        return opts[0] if opts else None
    if k in ("combo", "list"):
        opts = f.get("options") or []
        cur = f.get("value")
        for o in opts:
            if o not in (cur, "", " "):
                return o
        return opts[0] if opts else None
    if k == "comb":
        return ("1234567890" * 4)[: f["boxes"]]
    fmt = f.get("format") or ""
    if fmt.startswith("日付"):
        return "2026-04-01"
    if fmt.startswith(("数", "百分率")):
        return "1"
    if fmt.startswith("郵便番号（5"):
        return "12345"
    if fmt.startswith("郵便番号（9"):
        return "123456789"
    if fmt.startswith("電話"):
        return "0312345678"
    if fmt.startswith("社会保障"):
        return "123456789"
    ml = f.get("maxlen")
    if f.get("comb") and ml:
        return ("1234567890" * 4)[:ml]
    if ml and ml < 3:
        return "1" * ml
    # 欄の大きさに入る値（小さな欄は1字。書く値が入りきらないときは、書き込みが「入りきりません」と断る）
    x0, y0, x1, y1 = f.get("rect", (0, 0, 100, 20))
    w = x1 - x0
    from .fill import FillError, layout_text

    for cand in (("12", "1") if k == "inline" else ("テスト", "12", "1")):
        try:
            layout_text(cand, w, y1 - y0, min(float(f.get("size") or 10.5), 10.5), multiline=(y1 - y0) > 25)
            return cand
        except FillError:
            continue
    return "1"


def check_writing(doc: PdfDocument, win: dict) -> list[Check]:
    targets: dict[str, object] = {}
    skipped, not_writable = [], []
    for pg in win["pages"]:
        for f in pg["fields"]:
            if f["id"] in targets:
                continue
            if f.get("readonly"):
                not_writable.append(f"{f['id']}（読み取り専用）")
                continue
            if f["kind"] in ("signature", "button"):
                not_writable.append(f"{f['id']}（{'電子署名' if f['kind'] == 'signature' else '押すだけのボタン'}）")
                continue
            v = _sample(f)
            if v is None:
                skipped.append(f"{f['id']}: 試しの値を決められませんでした")
                continue
            targets[f["id"]] = v
    n_fields = sum(len(p["fields"]) for p in win["pages"])
    try:
        res = fill(doc, targets)
    except FillError as e:
        return [Check("全入力欄への書き込み", False, f"{len(targets)} 欄に書き込めませんでした", str(e).splitlines()[:30])]
    pr = prove(doc.blob, res.blob, set(targets))
    na = f"（読み取り専用・電子署名・ボタンの {len(not_writable)} 欄は、値を書く欄ではないので除外）" if not_writable else ""
    out = [Check("全入力欄への書き込み", pr.ok and not skipped,
                 f"入力欄 {n_fields:,} か所のうち {len(targets):,} に試しの値を書き込み{na}。{pr.headline()}",
                 [*pr.unexpected, *pr.lint_errors, *skipped])]
    after = PdfDocument(res.blob, doc.filename)
    win2 = build(after)
    got = {}
    for pg in win2["pages"]:
        for f in pg["fields"]:
            got[f["id"]] = f.get("value")
    wrong = []
    written = {p_.id: p_.value for p_ in res.patches}
    formatted = 0
    for k, v in targets.items():
        g = got.get(k)
        w = written.get(k, v)
        if not isinstance(v, bool) and str(w) != str(v):
            formatted += 1  # 欄の書式（日付・数など）を通した値
        v = w if not isinstance(v, bool) else v
        if isinstance(v, bool):
            ok = bool(g) == v
        else:
            ok = str(g) == str(v)
        if not ok:
            wrong.append(f"{k}: 書いた値 {v!r} → 窓では {g!r}")
    note = f"（うち {formatted} 欄は、欄の書式どおりの形にして書いた値）" if formatted else ""
    out.append(Check("読み戻し", not wrong, f"書いた {len(targets):,} 欄のうち {len(targets) - len(wrong):,} 欄が、書き込み後の窓にそのまま出ている{note}", wrong))
    return out


def run(blob: bytes, filename: str = "") -> Audit:
    doc = PdfDocument(blob, filename)
    win = build(doc)
    au = Audit(filename)
    au.checks += check_reading(doc, win)
    parts = markdown_parts(win)
    biggest = max(len(t) for _, t in parts)
    total = sum(len(t) for _, t in parts)
    au.checks.append(Check("窓の大きさ", biggest <= 150_000,
                           f"Markdown {total:,} 文字" + (f"を {len(parts)} 部分に分け、いちばん大きい部分も {biggest:,} 文字"
                                                        if len(parts) > 1 else "") + "（AIが1回で読める目安 15万文字以内）"))
    au.checks += check_writing(PdfDocument(blob, filename), win)
    fm = win["form"]
    scripts = _scripts(doc)
    au.checks.append(Check("フォームの仕掛け", not fm["xfa"] and not scripts,
                           "XFA（Adobe の独自のフォーム）も、独自のスクリプトもありません（Acrobat の標準の書式・計算は、同じ規則で行います）"
                           if not fm["xfa"] and not scripts else
                           ("XFA を含みます（Adobe Acrobat 以外では、PDF の欄の値が使われます）。" if fm["xfa"] else "")
                           + (f"独自のスクリプトが {len(scripts)} 個あります（実行しません。その欄は、書いた後に確かめてください）"
                              if scripts else ""), scripts))
    return au


def _scripts(doc: PdfDocument) -> list[str]:
    """Acrobat の標準の関数（日付・数・計算など）でない、独自のスクリプト。"""
    from . import acrojs

    out = []
    for f in read_fields(doc):
        node = doc.resolve(f.ref) if f.ref is not None else None
        for k, js in acrojs.actions(doc, node).items():
            if not acrojs.known(js):
                out.append(f"{f.name}（{ {'C': '計算', 'F': '表示の書式', 'K': '入力の制限', 'V': '値の確認'}[k] }: {js[:60]}）")
    return out
