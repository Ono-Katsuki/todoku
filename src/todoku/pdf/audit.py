"""PDF：エージェントがこの書類を「全て正確に読み書きできるか」を、決定的な処理で確かめる。

読む: ページの描画の命令から、窓とは別の読み方（並べずに、文字の命令だけを順に）で全部の字を拾い、窓に1字残らず出ているかを確かめる。
      文字に直せない字・画像だけのページ・注釈の文字も確かめる。
書く: 窓が示したすべての入力欄に試しの値を一度に書き込み、元のバイト列がそのまま残っていること、
      変わったのが書いた欄とその見た目だけであること、修復なしに開けること、書いた値が窓に読み戻せることを確かめる。
"""

from __future__ import annotations

import re
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
    flags = [(p["number"], fl) for p in win["pages"] for fl in p.get("read_flags") or []]
    hard = [f"{n} ページ：{fl['title']}（{'、'.join(fl['where'][:3])}）" for n, fl in flags if fl["grade"] in "CDEF"]
    soft = [f"{n} ページ：{fl['title']}（{'、'.join(fl['where'][:3])}）" for n, fl in flags if fl["grade"] not in "CDEF"]
    out.append(Check("窓の字を信用できない所", not hard,
                     "窓の字をそのまま信用できない所はありません" if not flags else
                     f"窓の字をそのまま信用できない所が {len(flags)} 種類あります（窓のページの頭に、位置と確かめ方を書いています）",
                     hard + soft, agent=bool(soft) and not hard))
    img = [p["number"] for p in win["pages"] if p.get("image_only")]
    out.append(Check("画像だけのページ", not img,
                     "すべてのページに文字の情報があります" if not img else
                     f"{len(img)} ページは画像だけで、文字の情報がありません（画像を見て読む必要があります）",
                     [f"ページ {n}" for n in img]))
    return out


def _sample(f: dict):
    if f.get("source") == "xfa":
        return f.get("_sample")  # XFA の欄：欄の決まり（書式・桁・選択肢）に合う値
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


def _all_fields(win: dict) -> list[dict]:
    """ページの入力欄と、ページに置かれていない XFA の入力欄（動的なフォームの欄など）。"""
    out = [f for pg in win["pages"] for f in pg["fields"]]
    seen = {f["id"] for f in out}
    out += [f for f in (win.get("xfa") or {}).get("fields", []) if f["id"] not in seen]
    return out


def check_writing(doc: PdfDocument, win: dict, password: str | None = None) -> list[Check]:
    targets: dict[str, object] = {}
    skipped, not_writable = [], []
    for f in _all_fields(win):
        if f["id"] in targets:
            continue
        if f.get("readonly"):
            not_writable.append(f"{f['id']}（{f.get('why_not') or '読み取り専用'}）")
            continue
        if f["kind"] in ("signature", "button"):
            not_writable.append(f"{f['id']}（{'電子署名' if f['kind'] == 'signature' else '押すだけのボタン'}）")
            continue
        if any(o in targets for o in f.get("shared_with") or []):
            continue  # 同じデータを共有する欄（XFA の global）は、1つだけ書く
        v = _sample(f)
        if v is None:
            skipped.append(f"{f['id']}: 試しの値を決められませんでした")
            continue
        targets[f["id"]] = v
    n_fields = len(_all_fields(win))
    try:
        res = fill(doc, targets)
    except FillError as e:
        return [Check("全入力欄への書き込み", False, f"{len(targets)} 欄に書き込めませんでした", str(e).splitlines()[:30])]
    pr = prove(doc.blob, res.blob, set(targets), password=password)
    na = (f"（読み取り専用・電子署名・ボタン・データに結びつかない欄など {len(not_writable)} 欄は、書けない欄なので除外）"
          if not_writable else "")
    out = [Check("全入力欄への書き込み", pr.ok and not skipped,
                 f"入力欄 {n_fields:,} か所のうち {len(targets):,} に試しの値を書き込み{na}。{pr.headline()}",
                 [*pr.unexpected, *pr.lint_errors, *skipped])]
    after = PdfDocument(res.blob, doc.filename, password=password)
    win2 = build(after)
    got = {f["id"]: f.get("value") for f in _all_fields(win2)}
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


def anchored_check(blob: bytes, filename: str, password: str | None = None) -> Check:
    """見出しの字を目印にした書き込み（窓に欄として示していない所に、エージェントが書く）：ページの中に1つしかない字の右の
    空いている所に試しの字を書き、証明（目印から決め直した位置と、足した欄の位置が同じ）と読み戻しを確かめる。"""
    import re

    from .anchor import AnchorError, _lines, resolve_write
    from .fill import FillError, fill
    from .form import read_fields
    from .proof import prove

    doc = PdfDocument(blob, filename, password=password)
    names = {f.name for f in read_fields(doc)}
    layouts: dict = {}
    tried = 0
    for n in range(1, len(doc.pages()) + 1):
        try:
            resolve_write(doc, {"after": "\x00", "value": "", "page": n}, layouts=layouts)
        except AnchorError:
            pass  # ページの配置を読んでおく（見つからないのは当然）
        lay = layouts.get(n)
        if lay is None:
            continue
        texts = [ln.text for _c, ln in _lines(lay)]
        for _cell, ln in _lines(lay):
            core = ln.text.rstrip(" \u3000")
            m = re.search(r"\S{3,}$", core)
            if not m or sum(t.count(m[0]) for t in texts) != 1:
                continue
            spec = {"after": m[0], "value": "（試し）", "page": n}
            try:
                sd = resolve_write(doc, spec, layouts=layouts, taken_names=names)
            except AnchorError:
                continue
            r = sd["rect"]
            if r[2] - r[0] < 5 * sd["size"] or r[3] - r[1] < 0.9 * sd["size"]:
                continue  # 試しの字「（試し）」（4字）が、ふつうの大きさで入る空きを選ぶ
            tried += 1
            try:
                res = fill(PdfDocument(blob, filename, password=password), {"@1": spec})
            except FillError as e:
                return Check("字を目印にした書き込み", False, f"{n} ページの「{m[0]}」の右に書けませんでした", [str(e)])
            pr = prove(blob, res.blob, {sd["id"]}, {sd["id"]: spec}, password=password)
            got = {f["id"]: f.get("value") for pg in build(PdfDocument(res.blob, filename, password=password))["pages"]
                   for f in pg["fields"]}
            ok = pr.ok and got.get(sd["id"]) == "（試し）"
            return Check("字を目印にした書き込み", ok,
                         f"窓に欄として示していない所への書き込み（{n} ページの「{m[0]}」の右に「（試し）」）を {sd['id']} として書き、"
                         f"{pr.headline()}。読み戻し：{'窓に出ている' if got.get(sd['id']) == '（試し）' else '窓に出ていない'}",
                         [*pr.unexpected, *pr.lint_errors])
            if tried >= 5:
                break
    return Check("字を目印にした書き込み", True, "右に空きのある、ページの中に1つしかない字がないので、試していません")


def run(blob: bytes, filename: str = "", password: str | None = None) -> Audit:
    """password：開くのにパスワードが要る PDF のとき（本人に聞いたもの。結果には出さない）。"""
    doc = PdfDocument(blob, filename, password=password)
    win = build(doc)
    au = Audit(filename)
    au.checks += check_reading(doc, win)
    parts = markdown_parts(win)
    biggest = max(len(t) for _, t in parts)
    total = sum(len(t) for _, t in parts)
    au.checks.append(Check("窓の大きさ", biggest <= 150_000,
                           f"Markdown {total:,} 文字" + (f"を {len(parts)} 部分に分け、いちばん大きい部分も {biggest:,} 文字"
                                                        if len(parts) > 1 else "") + "（AIが1回で読める目安 15万文字以内）"))
    au.checks += check_writing(PdfDocument(blob, filename, password=password), win, password)
    from ..xlsx.audit import agent_places_check

    ap = agent_places_check([f for p_ in win["pages"] for f in p_["fields"]])
    if ap is not None:
        au.checks.append(ap)
    if (win.get("xfa") or {}).get("readable"):
        au.checks.append(Check("字を目印にした書き込み", True, "XFA のフォームでは、字を目印にした書き込みはしません"
                               "（XFA の欄でない所に書いても Adobe では表示されないことがあるので、XFA の欄の id を指定して書きます）"))
    elif not any(p_.get("image_only") for p_ in win["pages"]) or any(p_.get("fields") for p_ in win["pages"]):
        au.checks.append(anchored_check(blob, filename, password))
    fm = win["form"]
    scripts = _scripts(doc)
    xw = win.get("xfa")
    if xw and xw.get("readable"):
        au.checks += check_xfa(doc, win)
        auto = [s_ for s_ in xw["scripts"] if s_["automatic"]]
        au.checks.append(Check("フォームの仕掛け", True,
                               f"XFA のフォームを読みました（{'動的' if xw['kind'] == 'dynamic' else '静的'}なフォーム）。"
                               + (f"XFA のスクリプト {len(xw['scripts'])} 種類（自動で動くもの {len(auto)}）は実行しません。"
                                  "ソースを窓に出したので、値や欄の出し入れを変えるものかを、エージェントが読んで判断します"
                                  if xw["scripts"] else "スクリプトはありません"),
                               [" / ".join(s_["where"][:2]) for s_ in auto[:20]], agent=bool(auto)))
        return au
    au.checks.append(Check("フォームの仕掛け", not fm["xfa"] and not scripts,
                           "XFA（Adobe の独自のフォーム）も、独自のスクリプトもありません（Acrobat の標準の書式・計算は、同じ規則で行います）"
                           if not fm["xfa"] and not scripts else
                           ("XFA を含みます（Adobe Acrobat 以外では、PDF の欄の値が使われます）。" if fm["xfa"] else "")
                           + (f"独自のスクリプトが {len(scripts)} 個あります（実行しません。その欄は、書いた後に確かめてください）"
                              if scripts else ""), scripts))
    return au


def _norm_ws(t: str) -> str:
    return " ".join(t.split())


def check_xfa(doc: PdfDocument, win: dict) -> list[Check]:
    """XFA：template と datasets の字（文字の節・属性の値）を、窓とは別の読み手（xml.etree）で拾い、すべて窓に出ているかを確かめる。
    静的なフォームでは、PDF の欄の名前が、すべて XFA の欄の名前（id）になっているかも確かめる。"""
    import xml.etree.ElementTree as ET

    from .xfa import Xdp, repair_xml

    xw = win["xfa"]
    xdp = Xdp(doc)
    fields = xw["fields"]
    shown = list(xw["_plain"])
    for f in fields:
        for k in ("label", "tooltip", "value", "display"):
            if isinstance(f.get(k), str):
                shown.append(f[k])
        if isinstance(f.get("default"), str):
            shown.append(f["default"])
        for k in ("options", "option_labels", "export_values", "rules", "value", "assist"):
            if isinstance(f.get(k), list):
                shown += [str(v) for v in f[k]]
        for k in ("on_value", "off_value"):
            if f.get(k):
                shown.append(f[k])
    window_text = _norm_ws("\n".join(shown))
    atoms = {_norm_ws(t) for t in shown}
    scripts = {_norm_ws(s_) for s_ in xw["_scripts"]}
    lists_text = _norm_ws("\n".join(xw["lists"]))
    missing: list[str] = []
    images = 0
    total = 0

    def ok_text(t: str, where: str) -> bool:
        return t in atoms or t in window_text or (where == "lists" and t in lists_text)

    def walk(e, in_script: bool, kind: str):
        nonlocal images, total
        tag = e.tag.rsplit("}", 1)[-1] if isinstance(e.tag, str) else ""
        if tag == "image":
            images += 1
            return
        script = in_script or tag == "script"
        texts = [e.text or ""] + [c.tail or "" for c in e]
        if kind == "data" or kind == "lists":
            texts += [v for k_, v in e.attrib.items() if "xmlns" not in k_ and "xfa-data" not in k_ and "XMLSchema" not in k_]
        for t in texts:
            t = _norm_ws(t)
            if not t:
                continue
            total += 1
            if script and kind == "template":
                if _norm_ws(e.text or "") not in scripts:
                    missing.append(f"スクリプト（{tag}）：{t[:60]}")
                break
            if not ok_text(t, kind):
                missing.append(f"{'template' if kind == 'template' else 'datasets'} の {tag}：「{t[:60]}」")
        for c in e:
            if isinstance(c.tag, str):
                walk(c, script, kind)

    def et(el):
        data = xdp.combined[el.s:el.e]
        fixed, _ = repair_xml(data)
        try:
            return ET.fromstring(fixed)
        except ET.ParseError:
            return ET.fromstring(b'<w xmlns:xfa="http://www.xfa.org/schema/xfa-data/1.0/">' + fixed + b"</w>")[0]

    t_el = xdp.parts.get("template")
    if t_el is not None:
        walk(et(t_el), False, "template")
    ds = xdp.parts.get("datasets")
    if ds is not None:
        for c in et(ds):
            if not isinstance(c.tag, str) or "data-description" in c.tag:
                continue
            walk(c, False, "data" if c.tag.endswith("}data") else "lists")
    out = [Check("読み落とし（XFA の字）", not missing,
                 f"template と datasets の字 {total:,} 個（文字の節と、データの属性の値）を、窓とは別の読み手で拾い、"
                 + ("すべて窓に出ていることを確かめました" if not missing else f"{len(missing)} 個が窓に出ていません")
                 + f"（画像 {images} 個は字ではないので除く）", missing[:30])]
    if xw["kind"] == "static":
        from .form import read_fields

        ids = {f["id"] for f in fields}
        names = [f.name for f in read_fields(doc)]
        # ページの共通部分（pageArea）は、ページごとに同じものが PDF の欄になる（Page1[1] など）：0 番目に置き換えて照らす
        base = [re.sub(r"(#pageSet\[\d+\]\.[^.]+)\[\d+\]", r"\1[0]", n) for n in names]
        bad = [n for n, b in zip(names, base) if n not in ids and b not in ids]
        out.append(Check("XFA の名前と PDF の欄", not bad,
                         f"PDF の欄 {len(names)} 個の名前が、すべて XFA の欄の id（名前の式）と同じです" if not bad else
                         f"PDF の欄 {len(bad)} 個の名前が、XFA の欄の id と合いません", bad[:30]))
    return out


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
