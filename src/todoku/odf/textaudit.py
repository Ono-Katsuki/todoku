"""OpenDocument の文書：エージェントが「全て正確に読み書きできるか」を、決まった処理で確かめる。

読む: 本文・表・枠・ヘッダー/フッター・脚注・コメント・フィールドの字を、窓とは別の読み方（Python の標準の XML の読み手 expat で、
      content.xml・styles.xml の文字をそのまま拾う）で数え、窓に1字残らず出ているかを確かめる（空白は数えない）。
      フォームの部品の値・見出しも、窓の入力欄に出ているかを確かめる。
書く: 窓が示したすべての入力欄（保護で書けない欄を除く）に試しの値を一度に書き込み、証明と読み戻しを確かめる。
"""

from __future__ import annotations

import xml.parsers.expat as expat
from collections import Counter

from ..docx.audit import _sample, anchored_check, macro_check
from ..docx.fill import FillError
from ..xlsx.audit import Audit, Check, agent_places_check
from .text import OdtDocument
from .textfill import fill, written_of
from .textproof import prove
from .textwindow import build, markdown_parts
from .xml import NS

BOXES = set("□☐☑☒■")
_SKIP = {NS["text"] + "}" + x for x in ("tracked-changes", "note-citation", "variable-decls", "sequence-decls", "user-field-decls",
                                        "dde-connection-decls", "table-of-content-source", "illustration-index-source",
                                        "table-index-source", "object-index-source", "user-index-source",
                                        "alphabetical-index-source", "bibliography-source")} | {
    NS["office"] + "}forms", NS["office"] + "}event-listeners", NS["svg"] + "}title", NS["svg"] + "}desc",
    NS["dc"] + "}creator", NS["dc"] + "}date", NS["meta"] + "}date-string"}


def raw_chars(xml: bytes, area: str) -> Counter:
    """expat で、area（office:text・office:master-styles）の中の文字を、そのまま数える（空白・□☐☑☒■ は数えない）。"""
    p = expat.ParserCreate(namespace_separator="}")
    p.buffer_text = True
    c = Counter()
    stack: list[bool] = []  # 数える所か
    inside = [0]
    target = NS["office"] + "}" + area

    def start(tag, attrs):
        skip = tag in _SKIP or (tag == NS["text"] + "}hidden-text" and attrs.get(NS["text"] + "}is-hidden") == "true")
        on = (stack[-1] if stack else False) or tag == target
        stack.append(on and not skip)

    def end(tag):
        stack.pop()

    def chars(data):
        if stack and stack[-1]:
            for ch in data:
                if not ch.isspace() and ch not in BOXES:
                    c[ch] += 1

    p.StartElementHandler = start
    p.EndElementHandler = end
    p.CharacterDataHandler = chars
    p.Parse(xml, True)
    del inside
    return c


def _count(texts) -> Counter:
    c = Counter()
    for t in texts:
        for ch in t:
            if not ch.isspace() and ch not in BOXES:
                c[ch] += 1
    return c


def check_reading(doc: OdtDocument, win: dict) -> list[Check]:
    raw = raw_chars(doc.pkg.xml(doc.part), "text")
    if doc.pkg.flat or doc.pkg.has("styles.xml"):
        raw += raw_chars(doc.pkg.xml(doc.pkg.part_of("styles.xml")), "master-styles") if not doc.pkg.flat else Counter()
        if doc.pkg.flat:
            raw += raw_chars(doc.pkg.xml(doc.part), "master-styles")
    shown = Counter()
    for st in win["stories"]:
        shown += _count(st["_plain"])
    missing = [f"「{ch}」が {n} 字、窓に出ていません" for ch, n in (raw - shown).items()]
    extra = [f"「{ch}」が {n} 字、元にないのに窓に出ています" for ch, n in (shown - raw).items()]
    out = [Check("読み落とし（文書の文字）", not missing and not extra,
                 f"本文・表・枠・ヘッダー/フッター・脚注・コメント・フィールドの字 {sum(raw.values()):,} 字を、窓とは別の読み方"
                 "（Python の標準の XML の読み手で、文字をそのまま）で数え、すべて窓に出ていることを確かめました", missing + extra)]
    # フォームの部品：値と見出しが、窓の入力欄に出ているか
    by = {f["id"]: f for f in win["fields"]}
    bad = []
    n = 0
    for c in doc.controls:
        if c.source not in ("odf-form", "odf-radio"):
            continue
        n += 1
        f = by.get(c.name)
        if f is None:
            bad.append(f"部品「{c.name}」が入力欄にありません")
            continue
        want = c.checked if c.kind == "checkbox" else c.value
        if c.kind != "checkbox" and want in (None, ""):
            want = None
        got = f.get("value")
        if (c.kind == "checkbox" and bool(got) != bool(want)) or (c.kind != "checkbox" and (got or None) != want):
            bad.append(f"部品「{c.name}」の値 {want!r} が、窓では {got!r}")
        if c.options and f.get("options") != c.options:
            bad.append(f"部品「{c.name}」の選択肢が、窓と違います")
    if n:
        out.append(Check("フォームの部品", not bad, f"フォームの部品 {n} 個の値・選択肢が、窓の入力欄に出ていることを確かめました", bad))
    return out


def check_writing(blob: bytes, filename: str, win: dict) -> list[Check]:
    targets = {}
    protected, skipped = [], []
    for f in win["fields"]:
        if f.get("readonly"):
            protected.append(f["id"])
            continue
        if f["kind"] in ("picture", "label", "other"):
            skipped.append(f"{f['id']}（{f['kind']}）")
            continue
        targets[f["id"]] = _sample(f)
    try:
        res = fill(OdtDocument(blob, filename), targets)
    except FillError as e:
        return [Check("全入力欄への書き込み", False, f"{len(targets)} 欄に書き込めませんでした", str(e).splitlines()[:30])]
    pr = prove(blob, res.blob, written_of(res.patches))
    note = f"（保護で書けない欄 {len(protected)} か所は除外）" if protected else ""
    out = [Check("全入力欄への書き込み", pr.ok, f"入力欄 {len(win['fields']):,} か所のうち {len(targets):,} に試しの値を書き込み{note}。{pr.headline()}",
                 [*pr.unexpected, *pr.lint_errors])]
    after = build(OdtDocument(res.blob, filename))
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


def run(blob: bytes, filename: str = "") -> Audit:
    doc = OdtDocument(blob, filename)
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
    if not doc.protection_enforced:
        au.checks.append(anchored_check(blob, filename, OdtDocument, fill, prove))
    au.checks.append(macro_check(doc, win))
    return au
