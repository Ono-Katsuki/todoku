"""Word：エージェントがこの書類を「全て正確に読み書きできるか」を、決定的な処理で確かめる。

読む: 本文・表・ヘッダー/フッター・脚注・コメント・テキストボックスの字を、窓とは別の読み方（XML の文字の要素を直接）で
      全部拾い、窓に1字残らず出ているかを確かめる。
書く: 窓が示したすべての入力欄（文書の保護で書けない欄を除く）に試しの値を一度に書き込み、
      変わったのが書いた欄の段落だけであること、XML として正しいこと、書いた値が窓に読み戻せることを確かめる。
"""

from __future__ import annotations

import re
from collections import Counter

from ..xlsx.audit import Audit, Check
from .document import WINGDINGS, DocxDocument
from .fill import FillError, fill
from .proof import prove
from .window import build, markdown_parts

BOXES = set("□☐☑☒■")


def raw_texts(doc: DocxDocument) -> list[str]:
    """文字の要素（w:t）と記号（w:sym）を、XML から直接拾う。AlternateContent の代わりの中身（Fallback）は除く。"""
    out = []
    for _kind, part, _blocks in doc.all_blocks():
        xml = doc.xml(part)
        xml = re.sub(rb"<mc:Fallback\b.*?</mc:Fallback>", b"", xml, flags=re.S)
        xml = re.sub(rb"<w:instrText\b[^>]*>.*?</w:instrText>", b"", xml, flags=re.S)
        for m in re.finditer(rb"<(?:w|m):t(?:\s[^>]*)?>(.*?)</(?:w|m):t>|<w:sym\b([^>]*)/>", xml, re.S):
            if m[1] is not None:
                t = m[1].decode("utf-8", "replace")
                out.append(t.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&"))
            else:
                a = m[2].decode("latin-1")
                cm = re.search(r'char="([0-9A-Fa-f]+)"', a)
                fm = re.search(r'font="([^"]*)"', a)
                if cm:
                    code = int(cm[1], 16)
                    code2 = code + 0xF000 if code < 0xF000 else code
                    if fm and "wingding" in fm[1].lower():
                        out.append(WINGDINGS.get(code2, "◆"))
    return out


def _count(texts) -> Counter:
    c = Counter()
    for t in texts:
        for ch in t:
            if not ch.isspace() and ch not in BOXES:
                c[ch] += 1
    return c


def check_reading(doc: DocxDocument, win: dict) -> list[Check]:
    raw = _count(raw_texts(doc))
    shown = Counter()
    for st in win["stories"]:
        shown += _count(st["_plain"])
    missing = [f"「{ch}」が {n} 字、窓に出ていません" for ch, n in (raw - shown).items()]
    extra = [f"「{ch}」が {n} 字、元にないのに窓に出ています" for ch, n in (shown - raw).items()]
    total = sum(raw.values())
    return [Check("読み落とし（文書の文字）", not missing and not extra,
                  f"本文・表・ヘッダー/フッター・脚注・コメント・テキストボックスの字 {total:,} 字を、窓とは別の読み方（XML の文字の要素を直接）で拾い、"
                  "すべて窓に出ていることを確かめました", missing + extra)]


def _sample(f: dict):
    k = f["kind"]
    if k == "checkbox":
        return not bool(f.get("value"))
    if f.get("options"):
        cur = f.get("value")
        for o in reversed(f["options"]):
            if o != cur:
                return o
        return f["options"][0]
    if k == "date":
        return "2026-04-01"
    if k == "inline":
        return "12"
    ml = f.get("maxlen")
    if ml and ml < 3:
        return "1" * ml
    return "テスト"


def check_writing(doc: DocxDocument, win: dict) -> list[Check]:
    targets = {}
    protected, skipped = [], []
    for f in win["fields"]:
        if f.get("readonly"):
            protected.append(f["id"])
            continue
        if f["kind"] == "picture":
            skipped.append(f"{f['id']}（画像の欄）")
            continue
        targets[f["id"]] = _sample(f)
    try:
        res = fill(doc, targets)
    except FillError as e:
        return [Check("全入力欄への書き込み", False, f"{len(targets)} 欄に書き込めませんでした", str(e).splitlines()[:30])]
    from .fill import written_of

    pr = prove(doc.blob, res.blob, written_of(res.patches))
    note = f"（文書の保護で書けない・変えられない欄 {len(protected)} か所は除外）" if protected else ""
    out = [Check("全入力欄への書き込み", pr.ok and not skipped,
                 f"入力欄 {len(win['fields']):,} か所のうち {len(targets):,} に試しの値を書き込み{note}。{pr.headline()}",
                 [*pr.unexpected, *pr.lint_errors, *skipped])]
    after = build(DocxDocument(res.blob, doc.filename))
    got = {f["id"]: f.get("value") for f in after["fields"]}
    written = {p.id: p.value for p in res.patches}
    wrong = []
    for k, v in targets.items():
        g = got.get(k)
        w = written.get(k, v)
        if isinstance(v, bool):
            ok = bool(g) == v
        else:
            ok = str(g) == str(w)
        if not ok:
            wrong.append(f"{k}: 書いた値 {w!r} → 窓では {g!r}")
    out.append(Check("読み戻し", not wrong, f"書いた {len(targets):,} 欄のうち {len(targets) - len(wrong):,} 欄が、書き込み後の窓にそのまま出ている", wrong))
    return out


def anchored_check(blob: bytes, filename: str, opener, fill_fn, prove_fn, main_end: int | None = None) -> Check:
    """字を目印にした書き込み（規則で示していない所に、エージェントが書く）：文書の中に1つしかない4字の後ろに試しの字を書き、
    証明（目印から同じ所に決まり、その段落だけが変わった）と読み戻し（［P段落.w番号＝値］）を確かめる。
    main_end：.doc の本文の終わりの位置（しおりを置けるのは本文だけなので、本文の段落を目印にする）。"""
    from .document import iter_paras
    from .fill import written_of
    from .window import Spots, _in_field

    doc = opener(blob, filename)
    sp = Spots(doc)
    paras = [(sp.para_index[id(q)], q) for _k, _p, bl in doc.all_blocks() for q in iter_paras(bl)]
    full = "\n".join(q.text for _n, q in paras)
    tried = 0
    last_err = ""
    for n, q in paras:
        if main_end is not None and getattr(q, "cp0", 0) >= main_end:
            continue
        t = q.text
        pos = q.positions()
        for i in range(0, len(t) - 3):
            frag = t[i:i + 4]
            if frag != frag.strip(" \u3000\t") or full.count(frag) != 1 or any(_in_field(q, j) for j in range(i, i + 5)):
                continue
            if any(q.segs[pos[k][0]].kind != "t" for k in range(i, min(i + 5, len(pos)))):
                continue  # 本文の字（字の要素）の中だけを目印にする（図の説明などは目印にならない）
            tried += 1
            spec = {"old": frag, "new": frag + "（試し）"}
            try:
                res = fill_fn(opener(blob, filename), {"@1": spec})
            except FillError as e:
                last_err = str(e)
                if tried >= 5:
                    break
                continue
            pr = prove_fn(blob, res.blob, written_of(res.patches))
            fid = res.patches[0].id
            got = {f["id"]: f.get("value") for f in build(opener(res.blob, filename))["fields"]}
            ok = pr.ok and got.get(fid) == "（試し）"
            return Check("字を目印にした書き込み", ok,
                         f"規則で示していない所への書き込み（段落 {n} の「{frag}」の後ろに「（試し）」）を {fid} として書き、"
                         f"{pr.headline()}。読み戻し：{'窓に［' + fid + '＝（試し）］と出ている' if got.get(fid) == '（試し）' else '窓に出ていない'}",
                         [*pr.unexpected, *pr.lint_errors])
        if tried >= 5:
            break
    if not tried:
        return Check("字を目印にした書き込み", True, "文書の中に1つしかない字がないので、試していません（同じ字の所は near・occurrence で指す）")
    return Check("字を目印にした書き込み", False, f"字を目印にした書き込みを {tried} か所で試し、書けませんでした", [last_err])


def run(blob: bytes, filename: str = "") -> Audit:
    doc = DocxDocument(blob, filename)
    win = build(doc)
    au = Audit(filename)
    au.checks += check_reading(doc, win)
    parts = markdown_parts(win)
    biggest = max(len(t) for _, t in parts)
    total = sum(len(t) for _, t in parts)
    au.checks.append(Check("窓の大きさ", biggest <= 150_000,
                           f"Markdown {total:,} 文字" + (f"を {len(parts)} 部分に分け、いちばん大きい部分も {biggest:,} 文字"
                                                        if len(parts) > 1 else "") + "（AIが1回で読める目安 15万文字以内）"))
    au.checks += check_writing(DocxDocument(blob, filename), win)
    from ..xlsx.audit import agent_places_check

    ap = agent_places_check(win["fields"])
    if ap is not None:
        au.checks.append(ap)
    if not (doc.protection_enforced and doc.protection):
        au.checks.append(anchored_check(blob, filename, DocxDocument, fill, prove))
    au.checks.append(macro_check(doc, win))
    return au


def macro_check(doc, win) -> Check:
    """マクロ：ソースをすべて窓に出したか。動作は規則では再現しないので、判断はエージェント。"""
    mi = win.get("macros")
    if not doc.has_macros:
        return Check("マクロ", True, "マクロはありません")
    if not isinstance(mi, dict) or not mi.get("readable"):
        return Check("マクロ", False, "マクロを含みますが、マクロの入れ物を読めませんでした（何をするマクロか分かりません）")
    n_proc = sum(len(m["procedures"]) for m in mi["modules"])
    auto = [f"{e['module']}.{e['procedure']}（{e['when']}）" for e in mi["events"]]
    return Check("マクロ", True, f"マクロのソース（モジュール {len(mi['modules'])} 個・手続き {n_proc} 個）をすべて窓に出しました。"
                 "マクロの動作は規則では再現しないので、自動で動く処理が要るかは、エージェントがソースを読んで判断します",
                 auto, agent=True)
