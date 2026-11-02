"""CSV・TSV・JSON・JSON Lines：エージェントが「全て正確に読み書きできるか」を、決まった処理で確かめる。

読む: 窓とは別の読み方（Python の標準の csv・json の読み手）でファイルを読み、すべてのマス・値が、同じ値で窓に出ているか
      （1字残らず）を確かめる。
書く: すべてのマス・値（と足す行・足す名前）に試しの値を一度に書き、変わったのが書いた所だけであること（証明）と、
      書いた値を読み直せることを確かめる。
"""

from __future__ import annotations

import csv
import io
import json
from collections import Counter

from ..xlsx.audit import Audit, Check
from .csvfile import CsvError, CsvFile
from .csvfile import fill as csv_fill
from .csvfile import prove as csv_prove
from .jsonfile import JsonError, JsonFile, esc_token
from .jsonfile import fill as json_fill
from .jsonfile import prove as json_prove
from .window import build_csv, build_json, cell_text, parts_csv, parts_json, row_line


def _chars(texts) -> Counter:
    c = Counter()
    for t in texts:
        for ch in str(t):
            if not ch.isspace():
                c[ch] += 1
    return c


# ====================================================================== CSV・TSV
def csv_reading(f: CsvFile, doc: dict, md: str) -> Check:
    """Python の csv モジュール（RFC 4180 の別の実装）で読み、マスごとに値を照らし、窓の表の行がそのまま出ているかを見る。"""
    other = list(csv.reader(io.StringIO(f.text, newline=""), delimiter=f.delim, quotechar='"', strict=False))
    mine = f.rows
    problems = []
    if len(other) != len(mine):
        problems.append(f"行の数が違います（窓 {len(mine)}・別の読み方 {len(other)}）")
    for r, (x, y) in enumerate(zip(mine, other), 1):
        if x != y:
            k = next((c for c in range(max(len(x), len(y))) if (x[c] if c < len(x) else None) != (y[c] if c < len(y) else None)), 0)
            problems.append(f"R{r}.C{k + 1}：窓「{(x[k] if k < len(x) else '（なし）')[:30]}」・別の読み方「{(y[k] if k < len(y) else '（なし）')[:30]}」")
            if len(problems) > 20:
                break
    width = max(doc["width"], 1)
    missing = [f"R{r}" for r, vals in enumerate(other, 1) if row_line(f"R{r}" + ("（見出し）" if r == 1 else ""), vals, width) not in md]
    if missing:
        problems.append(f"窓に、そのままの形で出ていない行 {len(missing)} 行：" + "、".join(missing[:15]))
    total = sum(_chars(v for row in other for v in row).values())
    return Check("読み落とし（表の字）", not problems,
                 f"マス {sum(len(r) for r in other):,} か所・字 {total:,} 字を、窓とは別の読み方（Python の標準の csv の読み手）で読み、"
                 "すべてのマスが同じ値で窓の表に出ていることを確かめました", problems)


def _csv_sample(f: CsvFile, k: int) -> str:
    samples = ["テスト", 'カンマ,と"引用符"', "改行\nあり", "12", "  前後の空白  "]
    s = samples[k % len(samples)]
    try:
        s.encode(f.dec.encoding)
    except UnicodeEncodeError:
        s = ["test", 'comma,and"quote"', "line\nbreak", "12", "  spaces  "][k % 5]
    if f.kind == "tsv":
        s = s.replace("\t", " ")
    return s


def csv_writing(blob: bytes, filename: str, kind: str, doc: dict) -> list[Check]:
    f = CsvFile(blob, filename, kind)
    targets = {}
    for k, fd in enumerate(doc["fields"]):
        targets[fd["id"]] = _csv_sample(f, k)
    try:
        res = csv_fill(f, targets)
    except CsvError as e:
        return [Check("全マスへの書き込み", False, f"{len(targets)} マスに書き込めませんでした", str(e).splitlines()[:20])]
    written = {(p.row, p.col): p.value for p in res.patches}
    pr = csv_prove(blob, res.blob, written, filename, kind)
    note = "（マスが多いので、空のマスと足す行だけ）" if doc.get("fields_truncated") else ""
    out = [Check("全マスへの書き込み", pr.ok, f"マス {len(targets):,} か所に試しの値（引用符・カンマ・改行を含む）を書き込み{note}。{pr.headline()}",
                 [*pr.unexpected, *pr.lint_errors])]
    g = CsvFile(res.blob, filename, kind)
    wrong = [f"R{r}.C{c}: 書いた値 {v!r} → 読み直すと {g.value(r, c)!r}" for (r, c), v in written.items() if g.value(r, c) != v]
    other = list(csv.reader(io.StringIO(g.text, newline=""), delimiter=g.delim, quotechar='"', strict=False))
    wrong += [f"R{r}.C{c}: 別の読み方では {(other[r - 1][c - 1] if r <= len(other) and c <= len(other[r - 1]) else None)!r}"
              for (r, c), v in written.items()
              if not (r <= len(other) and c <= len(other[r - 1]) and other[r - 1][c - 1] == v)]
    out.append(Check("読み戻し", not wrong, f"書いた {len(written):,} マスのうち {len(written) - len(set(w.split(':')[0] for w in wrong)):,} マスが、"
                     "書き込み後のファイルを（窓の読み方と、別の読み方の両方で）読み直して、書いた値のまま", wrong[:30]))
    return out


# ====================================================================== JSON・JSON Lines
def _independent_leaves(f: JsonFile) -> list[tuple[str, object]]:
    """Python の json モジュールで読んだ値の、すべての葉（id, 値）。"""
    docs = []
    if f.kind == "jsonl":
        for k, line in enumerate(f.text.split("\n"), 1):
            if line.strip():
                docs.append((f"L{k}", json.loads(line)))
    else:
        docs.append(("", json.loads(f.text)))
    out = []

    def walk(pre: str, v):
        if isinstance(v, dict):
            for key, x in v.items():
                walk(pre + "/" + esc_token(key), x)
        elif isinstance(v, list):
            for i, x in enumerate(v):
                walk(f"{pre}/{i}", x)
        else:
            out.append((pre or "/", v))

    for pre, v in docs:
        walk(pre, v)
    return out


def _keys(f: JsonFile) -> list[str]:
    out = []
    if f.kind == "jsonl":
        vals = [json.loads(x) for x in f.text.split("\n") if x.strip()]
    else:
        vals = [json.loads(f.text)]
    stack = list(vals)
    while stack:
        v = stack.pop()
        if isinstance(v, dict):
            out += list(v.keys())
            stack += list(v.values())
        elif isinstance(v, list):
            stack += v
    return out


def json_reading(f: JsonFile, doc: dict, md: str) -> Check:
    problems = []
    try:
        other = _independent_leaves(f)
    except ValueError as e:
        return Check("読み落とし（値）", False, f"別の読み方（Python の json）で読めません：{e}")
    if f.duplicates:
        problems.append("同じ名前が2つ以上あるオブジェクトがあります（別の読み方では後のものだけが残る）：" + "、".join(f.duplicates[:10]))
    mine = {v["id"]: v for v in doc["values"]}
    for i, v in other:
        w = mine.get(i)
        if w is None:
            problems.append(f"{i}：窓にありません")
            continue
        try:
            got = json.loads(w["json"])
        except ValueError:
            got = object()
        same = (got == v and type(got) is type(v)) or (isinstance(v, (int, float)) and not isinstance(v, bool)
                                                      and isinstance(got, (int, float)) and not isinstance(got, bool) and got == v)
        if not same:
            problems.append(f"{i}：窓 {w['json'][:40]}・別の読み方 {json.dumps(v, ensure_ascii=False)[:40]}")
        elif w["json"] not in md:
            problems.append(f"{i}：値 {w['json'][:40]} が窓の Markdown に出ていません")
        if len(problems) > 30:
            break
    if len(other) != len(mine) and not f.duplicates:
        problems.append(f"値の数が違います（窓 {len(mine)}・別の読み方 {len(other)}）")
    names = _keys(f)
    missing_names = [k for k in set(names) if esc_token(k) not in md and k not in md]
    if missing_names:
        problems.append("窓に出ていない名前：" + "、".join(f"「{k}」" for k in missing_names[:15]))
    total = sum(_chars([str(v) for _i, v in other if v is not None] + names).values())
    return Check("読み落とし（値）", not problems,
                 f"値 {len(other):,} 個・名前 {len(names):,} 個（字 {total:,} 字）を、窓とは別の読み方（Python の標準の json の読み手）で読み、"
                 "すべての値が同じ値で、すべての名前とともに窓に出ていることを確かめました", problems)


def _json_sample(kind: str, k: int):
    if kind in ("true", "false"):
        return kind == "false"
    if kind == "number":
        return 12
    return ["テスト", '引用符"と\\逆斜線', "改行\nあり"][k % 3]


def json_writing(blob: bytes, filename: str, kind: str, doc: dict) -> list[Check]:
    f = JsonFile(blob, filename, kind)
    targets = {}
    for k, fd in enumerate(doc["fields"]):
        targets[fd["id"]] = _json_sample(fd.get("json_type", "string"), k)
    root = f.docs[0].root if f.docs else None
    if kind == "json" and root is not None and root.kind == "object":
        targets["/_todoku_試し"] = "足した名前"  # ない名前に書く（オブジェクトの終わりに足す）
    try:
        res = json_fill(f, targets)
    except JsonError as e:
        return [Check("全値への書き込み", False, f"{len(targets)} か所に書き込めませんでした", str(e).splitlines()[:20])]
    written = {p.id: p.value for p in res.patches}
    pr = json_prove(blob, res.blob, written, filename, kind)
    note = "（値が多いので、空の値だけ）" if doc.get("fields_truncated") else ""
    out = [Check("全値への書き込み", pr.ok, f"値 {len(targets):,} か所に試しの値を書き込み（ない名前に足すのも含む）{note}。{pr.headline()}",
                 [*pr.unexpected, *pr.lint_errors])]
    g = JsonFile(res.blob, filename, kind)
    got = {i: n.value for i, n, _d in g.leaves()}
    wrong = [f"{i}: 書いた値 {v!r} → 読み直すと {got.get(i)!r}" for i, v in written.items() if got.get(i) != v or type(got.get(i)) is not type(v)]
    try:
        other = dict(_independent_leaves(g))
        wrong += [f"{i}: 別の読み方では {other.get(i)!r}" for i, v in written.items() if other.get(i) != v]
    except ValueError as e:
        wrong.append(f"書き込み後のファイルを、別の読み方（Python の json）で読めません：{e}")
    out.append(Check("読み戻し", not wrong, f"書いた {len(written):,} 個の値が、書き込み後のファイルを（窓の読み方と、別の読み方の両方で）読み直して、書いた値のまま"
                     if not wrong else f"書いた {len(written):,} 個のうち、読み直して違う値があります", wrong[:30]))
    return out


# ====================================================================== 入口
def run(blob: bytes, filename: str = "", kind: str = "csv") -> Audit:
    au = Audit(filename)
    if kind in ("csv", "tsv"):
        f = CsvFile(blob, filename, kind)
        doc = build_csv(f)
        parts = parts_csv(doc)
        md = "\n".join(t for _n, t in parts)
        au.checks.append(csv_reading(f, doc, md))
        if f.problems:
            au.checks.append(Check("書き方", False, "RFC 4180 に合わない書き方があります（窓に、どう読んだかを書きました）", f.problems[:20]))
    else:
        f = JsonFile(blob, filename, kind)
        doc = build_json(f)
        parts = parts_json(doc)
        md = "\n".join(t for _n, t in parts)
        au.checks.append(json_reading(f, doc, md))
    biggest = max(len(t) for _, t in parts)
    au.checks.append(Check("窓の大きさ", biggest <= 150_000,
                           f"Markdown {sum(len(t) for _, t in parts):,} 文字" + (f"を {len(parts)} 部分に分け、いちばん大きい部分も {biggest:,} 文字"
                                                                             if len(parts) > 1 else "") + "（AIが1回で読める目安 15万文字以内）"))
    au.checks += csv_writing(blob, filename, kind, doc) if kind in ("csv", "tsv") else json_writing(blob, filename, kind, doc)
    return au


__all__ = ["run", "cell_text"]
