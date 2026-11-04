"""HWPX：エージェントが「全て正確に読み書きできるか」を、決まった処理で確かめる。

読む: 本文の節（とバタンページ）の hp:t の字を、窓とは別の読み方（Python の標準の XML の読み手 expat で、hp:t の中の字を
      そのまま拾う）で数え、窓に1字残らず出ているかを確かめる（空白と □☐☑☒■ は数えない）。フォームの部品の値・選択肢も、
      窓の入力欄に出ているかを確かめる。
書く: 窓が示したすべての入力欄（保護で書けない欄を除く）に試しの値を一度に書き込み、証明と読み戻しを確かめる。
"""

from __future__ import annotations

import xml.parsers.expat as expat
from collections import Counter

from ..docx.audit import _sample, already_written, anchored_check
from ..docx.fill import FillError
from ..xlsx.audit import AUDIT_BOXES, Audit, Check, agent_places_check, count_chars
from .document import HwpxDocument
from .fill import fill, written_of
from .package import HNS
from .proof import prove
from .window import build, markdown_parts

BOXES = AUDIT_BOXES


def raw_chars(xml: bytes) -> Counter:
    """expat で、한글の hp:t の中の字を数える（空白・□☐☑☒■ は数えない）。"""
    p = expat.ParserCreate(namespace_separator="}")
    p.buffer_text = True
    c = Counter()
    depth = [0]

    def start(tag, attrs):
        if tag.startswith(HNS) and tag.endswith("}t"):
            depth[0] += 1

    def end(tag):
        if tag.startswith(HNS) and tag.endswith("}t"):
            depth[0] -= 1

    def chars(data):
        if depth[0] > 0:
            for ch in data:
                if not ch.isspace() and ch not in BOXES:
                    c[ch] += 1

    p.StartElementHandler = start
    p.EndElementHandler = end
    p.CharacterDataHandler = chars
    p.Parse(xml, True)
    return c


def _count(texts) -> Counter:
    return count_chars(texts, BOXES)


def check_reading(doc: HwpxDocument, win: dict) -> list[Check]:
    raw = Counter()
    parts = list(doc.parts) + [n for n in doc.pkg.names if n.startswith("Contents/masterpage") and n.endswith(".xml")]
    for part in parts:
        raw += raw_chars(doc.pkg.xml(part))
    shown = Counter()
    for st in win["stories"]:
        shown += _count(st["_plain"])
    missing = [f"「{ch}」が {n} 字、窓に出ていません" for ch, n in (raw - shown).items()]
    extra = [f"「{ch}」が {n} 字、元にないのに窓に出ています" for ch, n in (shown - raw).items()]
    out = [Check("読み落とし（文書の文字）", not missing and not extra,
                 f"本文・表・図形の中・ヘッダー/フッター・脚注・メモの字（hp:t）{sum(raw.values()):,} 字を、窓とは別の読み方"
                 "（Python の標準の XML の読み手で、hp:t の字をそのまま）で数え、すべて窓に出ていることを確かめました", missing + extra)]
    by = {f["id"]: f for f in win["fields"]}
    from ..docx.window import Spots

    sp = Spots(doc)
    ctl_ids = {id(it["control"]): k for k, it in sp.items.items() if it.get("control") is not None}
    bad = []
    n = 0
    for c in doc.controls:
        n += 1
        f = by.get(ctl_ids.get(id(c), ""))
        if f is None:
            bad.append(f"欄「{c.name or c.source}」が入力欄にありません")
            continue
        if c.source == "hwpx-clickhere":
            want = c.value
        else:
            want = c.checked if c.kind == "checkbox" else c.value
        if c.kind != "checkbox" and want in (None, ""):
            want = None
        got = f.get("value")
        if (c.kind == "checkbox" and bool(got) != bool(want)) or (c.kind != "checkbox" and (got or None) != want):
            bad.append(f"欄「{f['id']}」の値 {want!r} が、窓では {got!r}")
        if c.options and f.get("options") != c.options:
            bad.append(f"欄「{f['id']}」の選択肢が、窓と違います")
    if n:
        out.append(Check("누름틀・フォームの部品", not bad, f"누름틀・フォームの部品 {n} 個の値・選択肢が、窓の入力欄に出ていることを確かめました", bad))
    return out


def check_writing(blob: bytes, filename: str, win: dict) -> list[Check]:
    doc = HwpxDocument(blob, filename)
    targets = {}
    protected, skipped = [], []
    done, bad_done = already_written(doc, win)
    for f in win["fields"]:
        if f["id"] in done:
            continue
        if f.get("readonly"):
            protected.append(f["id"])
            continue
        if f["kind"] in ("picture", "label", "other"):
            skipped.append(f"{f['id']}（{f['kind']}）")
            continue
        targets[f["id"]] = "12" if f.get("format") == "数字だけ" else _sample(f)
    try:
        res = fill(doc, targets)
    except FillError as e:
        return [Check("全入力欄への書き込み", False, f"{len(targets)} 欄に書き込めませんでした", str(e).splitlines()[:30])]
    pr = prove(blob, res.blob, written_of(res.patches))
    note = f"（保護で書けない欄 {len(protected)} か所は除外）" if protected else ""
    out = [Check("全入力欄への書き込み", pr.ok and not bad_done,
                 f"入力欄 {len(win['fields']):,} か所のうち {len(targets):,} に試しの値を書き込み{note}。{pr.headline()}",
                 [*pr.unexpected, *pr.lint_errors, *bad_done])]
    after = build(HwpxDocument(res.blob, filename))
    got = {f["id"]: f.get("value") for f in after["fields"]}
    written = {p.id: p.value for p in res.patches}
    wrong = []
    for k, v in targets.items():
        g = got.get(k)
        w = written.get(k, v)
        ok = bool(g) == v if isinstance(v, bool) else str(g) == str(w)
        if not ok:
            wrong.append(f"{k}: 書いた値 {w!r} → 窓では {g!r}")
    out.append(Check("読み戻し", not wrong, f"書いた {len(targets):,} 欄のうち {len(targets) - len(wrong):,} 欄が、書き込み後の窓にそのまま出ている", wrong))
    return out


def _locked(filename: str, e: Exception) -> Audit:
    au = Audit(filename)
    au.checks.append(Check("開けるか", False, f"文書を読めません：{e}"))
    return au


def run(blob: bytes, filename: str = "") -> Audit:
    from .package import HwpxEncrypted

    try:
        doc = HwpxDocument(blob, filename)
    except HwpxEncrypted as e:
        return _locked(filename, e)
    win = build(doc)
    au = Audit(filename)
    au.checks += check_reading(doc, win)
    parts = markdown_parts(win)
    biggest = max(len(t) for _, t in parts)
    au.checks.append(Check("窓の大きさ", biggest <= 150_000,
                           f"Markdown {sum(len(t) for _, t in parts):,} 文字" + (f"を {len(parts)} 部分に分け、いちばん大きい部分も {biggest:,} 文字"
                                                                             if len(parts) > 1 else "") + "（AIが1回で読める目安 15万文字以内）"))
    au.checks += check_writing(blob, filename, win)
    ap = agent_places_check(win["fields"])
    if ap is not None:
        au.checks.append(ap)
    au.checks.append(anchored_check(blob, filename, HwpxDocument, fill, prove))
    if doc.scripts:
        au.checks.append(Check("マクロ", True, f"한글のマクロ（JScript）{len(doc.scripts)} 個のソースを窓に出しました。動作は規則では再現しないので、"
                               "自動で動く処理が要るかは、エージェントがソースを読んで判断します", [n for n, _s in doc.scripts], agent=True))
    else:
        au.checks.append(Check("マクロ", True, "マクロはありません"))
    return au


# ====================================================================== HWP 5（読むだけ）
def run5(blob: bytes, filename: str = "") -> Audit:
    """HWP 5：読み落とし（すべての段落の PARA_TEXT の字を、窓とは別に数える）と窓の大きさ。書き込みは HWP 5 ではしない。"""
    from .hwp5 import Hwp5Document, Hwp5Locked
    from .window import build5

    try:
        doc = Hwp5Document(blob, filename)
    except Hwp5Locked as e:
        return _locked(filename, e)
    win = build5(doc)
    au = Audit(filename)
    raw = _count(doc.raw_texts)
    shown = Counter()
    for st in win["stories"]:
        shown += _count(st["_plain"])
    missing = [f"「{ch}」が {n} 字、窓に出ていません" for ch, n in (raw - shown).items()]
    extra = [f"「{ch}」が {n} 字、元にないのに窓に出ています" for ch, n in (shown - raw).items()]
    au.checks.append(Check("読み落とし（文書の文字）", not missing and not extra,
                           f"本文・表・글상자・ヘッダー/フッター・脚注の段落の字（PARA_TEXT）{sum(raw.values()):,} 字を、レコードから"
                           "そのまま数え、すべて窓に出ていることを確かめました", missing + extra))
    parts = markdown_parts(win)
    biggest = max(len(t) for _, t in parts)
    au.checks.append(Check("窓の大きさ", biggest <= 150_000, f"Markdown {sum(len(t) for _, t in parts):,} 文字（AIが1回で読める目安 15万文字以内）"))
    au.checks.append(Check("書き込み", True, "HWP 5（.hwp）には書けません。HWPX で保存し直せば書ける（한글の「다른 이름으로 저장」で .hwpx を選ぶ）",
                           agent=True))
    return au
