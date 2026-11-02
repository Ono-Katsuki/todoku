"""OpenDocument の表計算：エージェントが「全て正確に読み書きできるか」を、決まった処理で確かめる。

読む: セルの字・コメント・図形の字・入力規則の説明・ヘッダー/フッターを、窓とは別の読み方（Python の標準の XML の読み手
      expat で、content.xml の文字をそのまま。繰り返しの行・セルは繰り返しの数だけ）で数え、窓に1字残らず出ているかを確かめる。
書く: 窓が示したすべての入力欄に、入力規則に合う試しの値を一度に書き込み、証明（変わったのは書いたセルと、その値で
      変わる数式の結果だけ）と読み戻しを確かめる。数式は、保存された計算結果と照らす。
"""

from __future__ import annotations

import xml.parsers.expat as expat
from collections import Counter

from ..formula.parser import addr, parse_cell
from ..formula.values import to_text
from ..xlsx.audit import Audit, Check, _sample
from ..xlsx.calc import Calc, _run_deep, values_equal, verify
from ..xlsx.edit import EditError, interpret_input, resolve_target
from .sheet import OdsWorkbook
from .sheetfill import fill, prove
from .sheetwindow import build_wb, markdown_parts
from .xml import NS

_TB = NS["table"] + "}"
_OF = NS["office"] + "}"


def raw_counts(xml: bytes) -> dict[str, Counter]:
    """expat で、表のセル・コメント・図形・入力規則の説明の字を数える（空白は数えない）。"""
    p = expat.ParserCreate(namespace_separator="}")
    p.buffer_text = True
    out = {"セルの字": Counter(), "コメント": Counter(), "図形の字": Counter(), "入力規則の説明": Counter()}
    stack: list[tuple[str | None, int]] = []  # (いまの区分, 繰り返しの数)
    meta = {NS["dc"] + "}creator", NS["dc"] + "}date", NS["meta"] + "}date-string"}

    def start(tag, attrs):
        cat, rep = stack[-1] if stack else (None, 1)
        if tag == _TB + "table-row":
            rep = rep * int(attrs.get(_TB + "number-rows-repeated", "1") or 1)
        elif tag in (_TB + "table-cell", _TB + "covered-table-cell"):
            rep = rep * int(attrs.get(_TB + "number-columns-repeated", "1") or 1)
            cat = "セルの字"
        elif tag == _OF + "annotation":
            cat = "コメント"
        elif tag.startswith(NS["draw"] + "}") and cat in (None, "セルの字"):
            cat = "図形の字"
        elif tag in (_TB + "help-message", _TB + "error-message"):
            cat = "入力規則の説明"
        elif tag in meta or tag in (_OF + "forms", _TB + "named-expressions", _OF + "event-listeners"):
            cat = "skip"
        stack.append((cat, rep))

    def end(tag):
        stack.pop()

    def chars(data):
        if not stack:
            return
        cat, rep = stack[-1]
        if cat in out:
            for ch in data:
                if not ch.isspace():
                    out[cat][ch] += rep

    p.StartElementHandler = start
    p.EndElementHandler = end
    p.CharacterDataHandler = chars
    p.Parse(xml, True)
    return out


def _count(texts) -> Counter:
    c = Counter()
    for t in texts:
        for ch in str(t):
            if not ch.isspace():
                c[ch] += 1
    return c


def check_reading(wb: OdsWorkbook, doc: dict) -> list[Check]:
    raw = raw_counts(wb.data)
    shown = {"セルの字": Counter(), "コメント": Counter(), "図形の字": Counter(), "入力規則の説明": Counter()}
    for sh in doc["sheets"]:
        for c in sh.get("cells", []):
            if "display" in c:
                shown["セルの字"] += _count([c["display"]])
            elif c.get("value") is not None:
                v = c["value"]
                shown["セルの字"] += _count([("TRUE" if v else "FALSE") if isinstance(v, bool) else to_text(v)])
        shown["コメント"] += _count(n["text"] for n in sh.get("notes", []))
        shown["図形の字"] += _count(s["text"] + s["alt"] for s in sh.get("shapes", []))
        for v in sh.get("validations", []):
            shown["入力規則の説明"] += _count([v.get("prompt") or "", v.get("prompt_title") or "", v.get("error") or "",
                                        v.get("error_title") or ""])
    out = []
    for kind, cnt in raw.items():
        if kind == "入力規則の説明":
            # 入力規則の説明：ファイルには、どのセルにも使われていない規則の説明もある（窓は、セルに使われている規則だけ）
            extra = [f"「{ch}」が {n} 字、元にないのに窓に出ています" for ch, n in (shown[kind] - cnt).items()]
            if not cnt:
                continue
            out.append(Check(f"読み落とし（{kind}）", not extra, f"セルに使われている入力規則の説明 {sum(shown[kind].values()):,} 字が、"
                             f"元の説明（すべての規則で {sum(cnt.values()):,} 字）のとおりに窓に出ています", extra))
            continue
        missing = [f"「{ch}」が {n} 字、窓に出ていません" for ch, n in (cnt - shown[kind]).items()]
        extra = [f"「{ch}」が {n} 字、元にないのに窓に出ています" for ch, n in (shown[kind] - cnt).items()]
        if not cnt and not shown[kind]:
            continue
        out.append(Check(f"読み落とし（{kind}）", not missing and not extra,
                         f"{kind} {sum(cnt.values()):,} 字を、窓とは別の読み方（Python の標準の XML の読み手で、繰り返しの数だけ）で数え、"
                         "すべて窓に出ていることを確かめました" if not missing and not extra else f"{kind}：窓と、別の読み方で数えた字が合いません",
                         (missing + extra)[:30]))
    if not out:
        out.append(Check("読み落とし（セルの字）", True, "字のあるセルはありません"))
    return out


def check_writing(blob: bytes, filename: str, doc: dict) -> list[Check]:
    wb = OdsWorkbook(blob, filename)
    fields = [(sh["name"], f) for sh in doc["sheets"] for f in sh.get("fields", [])]
    targets: dict[str, object] = {}
    skipped: list[str] = []
    later = []
    for sheet, f in fields:
        if f.get("readonly"):
            continue
        if f["kind"] == "choice":
            later.append((sheet, f))
            continue
        v = _sample(f, wb, sheet, None)
        if v is None:
            skipped.append(f"{f['id']}: 試しの値を決められませんでした")
            continue
        targets[f["id"]] = v
    na = []
    if later:
        over = {}
        for k, v in targets.items():
            try:
                s, r, c = resolve_target(wb, k)
                over[(s, r, c)] = interpret_input(wb, s, r, c, v)[0]
            except EditError:
                pass

        def pick():
            calc = Calc(wb, over)
            return [(f, _sample(f, wb, sheet, calc)) for sheet, f in later]

        for f, v in _run_deep(pick):
            if v is None:
                na.append(f["id"])
            else:
                targets[f["id"]] = v
    try:
        res = fill(wb, targets)
    except EditError as e:
        return [Check("全入力欄への書き込み", False, f"{len(targets)} 欄に書き込めませんでした", str(e).splitlines()[:30])]
    written = {(p.sheet, *parse_cell(p.address)): p.after for p in res.patches if p.kind == "input"}
    recalc_ = {(p.sheet, *parse_cell(p.address)): p.after for p in res.patches if p.kind == "formula"}
    pr = prove(blob, res.blob, written, recalc_)
    note = f"（選択肢が空で選べない欄 {len(na)} か所は除外）" if na else ""
    stale = f"。計算し直せない数式 {len(res.stale)} 個は、結果が古いままと知らせました" if res.stale else ""
    out = [Check("全入力欄への書き込み", pr.ok and not skipped,
                 f"入力欄 {len(fields):,} か所のうち {len(targets):,} に試しの値を書き込み{note}。{pr.headline()}"
                 f"（計算し直した数式 {len(recalc_)} 個{stale}）", [*pr.unexpected, *pr.lint_errors, *skipped])]
    after = OdsWorkbook(res.blob, filename)
    wrong = []
    for (s, r, c), v in written.items():
        cell = after.sheet(s).cells.get((r, c))
        g = cell.value if cell is not None else None
        if not (values_equal(v, g) or (v in (None, "") and g in (None, ""))):
            wrong.append(f"{s}!{addr(r, c)}: 書いた値 {v!r} → 読み直すと {g!r}")
    out.append(Check("読み戻し", not wrong, f"書いた {len(written):,} セルのうち {len(written) - len(wrong):,} セルが、書き込み後のファイルを読み直して、書いた値のまま",
                     wrong))
    return out


def run(blob: bytes, filename: str = "") -> Audit:
    wb = OdsWorkbook(blob, filename)
    doc = build_wb(wb)
    au = Audit(filename)
    au.checks += check_reading(wb, doc)
    parts = markdown_parts(doc)
    biggest = max(len(t) for _, t in parts)
    au.checks.append(Check("窓の大きさ", biggest <= 150_000,
                           f"Markdown {sum(len(t) for _, t in parts):,} 文字" + (f"を {len(parts)} 部分に分け、いちばん大きい部分も {biggest:,} 文字"
                                                                             if len(parts) > 1 else "") + "（AIが1回で読める目安 15万文字以内）"))
    rep = verify(wb)
    s = rep.summary()
    bad = [f"{r.sheet}!{r.address}: {r.status}（{r.reason or ''}）{r.formula[:40]}" for r in rep.results
           if r.status in ("mismatch", "unsupported", "cycle")]
    au.checks.append(Check("数式", not bad, f"数式 {s['total']:,} 個を、OpenFormula を Excel の書き方にして計算し直し、"
                           f"保存された計算結果と照らしました（一致 {s['match']:,}・毎回変わる関数 {s['volatile']}・"
                           f"一致しない {s['mismatch']}・計算できない {s['unsupported']}）", bad[:30], agent=bool(bad)))
    au.checks += check_writing(blob, filename, doc)
    if doc.get("unread"):
        au.checks.append(Check("読んでいない部分", False, "同じ中身を繰り返す、とても大きな範囲は、最初の部分だけ読みました", doc["unread"]))
    mi = doc.get("macros")
    if not mi:
        au.checks.append(Check("マクロ", True, "マクロはありません"))
    else:
        n_proc = sum(len(m["procedures"]) for m in mi["modules"])
        au.checks.append(Check("マクロ", bool(mi.get("readable")),
                               f"マクロのソース（モジュール {len(mi['modules'])} 個・手続き {n_proc} 個）をすべて窓に出しました。"
                               "マクロの動作は規則では再現しないので、要るかどうかはエージェントがソースを読んで判断します",
                               [f"{e['module']}.{e['procedure']}（{e['when']}）" for e in mi["events"]], agent=True))
    return au
