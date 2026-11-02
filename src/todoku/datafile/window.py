"""CSV・TSV・JSON・JSON Lines の窓：データをそのまま、表（CSV は行と列、JSON は JSON Pointer）として出す。

データそのものなので、書く所を推定しない。どのマス・どの値にも id で書ける（CSV：R行.C列、JSON：/名前/番号）。
入力欄の一覧には、すべてのマス・値を出す（多いときは、空のマスと足す行だけ）。新しい行（CSV の R{行数+1}）は、
足すかどうかをエージェントが決める所として示す。
"""

from __future__ import annotations

import json

from .csvfile import CsvFile
from .jsonfile import JsonFile, JNode, unesc_token

BUDGET = 120_000
MAX_FIELDS = 3000
NEW_ROW_WHY = "新しい行（足すかどうかは、エージェントが決める。書けばファイルの終わりに行を足す）"


def cell_text(v: str | None) -> str:
    """表のマスの見せ方：| は \\|、マスの中の改行は ↵、タブは ⇥（字そのものは変えない）。"""
    if v is None:
        return "（マスなし）"
    if v == "":
        return "［］"
    return v.replace("|", "\\|").replace("\r\n", "↵").replace("\n", "↵").replace("\r", "↵").replace("\t", "⇥")


def row_line(label: str, values: list, width: int) -> str:
    cells = [cell_text(values[c]) if c < len(values) else "（マスなし）" for c in range(width)]
    return "| " + label + " | " + " | ".join(cells) + " |"


# ====================================================================== CSV・TSV
def build_csv(f: CsvFile) -> dict:
    rows = f.rows
    n = len(rows)
    width = f.width
    head = f.header
    fields = []
    cells = sum(len(r) for r in rows[1:])
    every = cells + width <= MAX_FIELDS
    for r in range(2, n + 1):
        vals = rows[r - 1]
        for c in range(1, width + 1):
            v = vals[c - 1] if c <= len(vals) else None
            if not every and v not in (None, ""):
                continue
            name = head[c - 1].strip() if c <= len(head) and head[c - 1].strip() else f"C{c}"
            fields.append({"id": f"R{r}.C{c}", "kind": "text", "label": f"{name}（{r}行目）", "value": v,
                           "row": r, "col": c, "column": name, "provenance": "deterministic",
                           "context": row_line(f"R{r}", vals, width)})
    for c in range(1, max(width, 1) + 1):
        name = head[c - 1].strip() if c <= len(head) and head[c - 1].strip() else f"C{c}"
        fields.append({"id": f"R{n + 1}.C{c}", "kind": "text", "label": f"{name}（足す行 {n + 1}行目）", "value": None,
                       "row": n + 1, "col": c, "column": name, "agent": True, "why": NEW_ROW_WHY, "provenance": "deterministic"})
    truncated = len(fields) > MAX_FIELDS
    return {
        "todoku": "0.2", "kind": f.kind,
        "source": {"file": f.filename, "sha256": f.sha256(), "bytes": len(f.blob),
                   "note": "元のファイルが正本。書き込みは、書いたマスのバイトだけを置き換える。"},
        "format": f.format_info(),
        "notation": {"R3.C2": "3行目・2列目のマス（1行目も R1。列は見出しの字でも指せる：R3.氏名・行3.列「氏名」）",
                     "［］": "空のマス", "↵": "マスの中の改行", "⇥": "マスの中のタブ", "\\|": "マスの中の「|」",
                     "（マスなし）": "その行は、この列までマスがない（書けば区切りを足す）"},
        "header": head, "width": width, "rows": rows, "row_count": n,
        "fields": fields[:MAX_FIELDS], "fields_truncated": truncated or not every,
        "problems": list(f.problems),
    }


def md_csv_head(doc: dict) -> list[str]:
    fm = doc["format"]
    n = doc["row_count"]
    kind = "TSV（タブ区切り）" if doc["kind"] == "tsv" else "CSV"
    out = [f"# {doc['source']['file'] or '表'}", "",
           f"- 元ファイルの SHA-256: `{doc['source']['sha256']}`（{doc['source']['bytes']:,} バイト）、{kind}",
           f"- 書き方：文字コード {fm['encoding_ja']}／改行 {fm['newline']}"
           + ("（行によって改行が違います）" if fm["mixed_newlines"] else "")
           + f"／区切り {fm['delimiter_ja']}／引用符 {fm['quoting']}"
           + ("" if fm["ends_with_newline"] else "／最後の行に改行なし"),
           "- この窓は読むためのものです。書き込み（`todoku fill`）は、書いたマスのバイトだけを変えます"
           "（ほかのマスの字・引用符・改行・文字コードはそのまま）。",
           f"- マスの id は「R行.C列」（1行目も R1）。列は見出しの字でも指せます（R2.氏名・行2.列「氏名」）。"
           f"新しい行は R{n + 1}.C1 のように、今ある {n} 行の次から足せます。",
           "- 記号: " + "、".join(f"{k} は{v}" for k, v in doc["notation"].items() if k != "R3.C2"),
           "- 1行目を列の見出しとして、入力欄の見出しに使いました（1行目が見出しでないときは、列の番号で指してください）。"]
    if fm.get("ascii_only"):
        out.append("- このファイルは ASCII の字だけなので、文字コードは決まりません（日本語を書くと UTF-8 で書きます）。")
    for p in doc.get("problems") or []:
        out.append(f"- **書き方の問題**：{p}")
    out.append("")
    return out


def md_csv_table(doc: dict, start: int = 1, stop: int | None = None) -> list[str]:
    width = max(doc["width"], 1)
    rows = doc["rows"]
    stop = len(rows) if stop is None else stop
    out = ["| 行 | " + " | ".join(f"C{c}" for c in range(1, width + 1)) + " |", "|---" * (width + 1) + "|"]
    if start > 1 and rows:
        out.append(row_line("R1（見出し）", rows[0], width))
    for r in range(start, stop + 1):
        out.append(row_line(f"R{r}" + ("（見出し）" if r == 1 else ""), rows[r - 1], width))
    if stop == len(rows):
        n = len(rows)
        out.append("| R" + str(n + 1) + "（足す行） | " + " | ".join(f"［R{n + 1}.C{c}？］" for c in range(1, width + 1)) + " |")
    return out


def to_markdown_csv(doc: dict) -> str:
    out = md_csv_head(doc)
    out += [f"## 表（{doc['row_count']:,} 行・{doc['width']} 列）", ""]
    out += md_csv_table(doc)
    return "\n".join(out) + "\n"


def parts_csv(doc: dict, budget: int = BUDGET) -> list[tuple[str, str]]:
    whole = to_markdown_csv(doc)
    if len(whole) <= budget:
        return [("全体", whole)]
    head = md_csv_head(doc)
    width = max(doc["width"], 1)
    parts: list[tuple[int, int]] = []
    start, size = 1, 0
    for r, vals in enumerate(doc["rows"], 1):
        ln = len(row_line(f"R{r}", vals, width)) + 1
        if size + ln > budget - 4000 and r > start:
            parts.append((start, r - 1))
            start, size = r, 0
        size += ln
    parts.append((start, len(doc["rows"])))
    index = head + ["## この窓の目次", "", f"表が大きいので、行の範囲で {len(parts)} の部分に分けました（どの部分にも1行目を付けています）。", ""]
    index += [f"- 部分 {i}: R{a}〜R{b}" for i, (a, b) in enumerate(parts, 1)]
    out = [("目次", "\n".join(index))]
    for i, (a, b) in enumerate(parts, 1):
        out.append((f"R{a}〜R{b}", "\n".join(head[:1] + ["", f"## 表 R{a}〜R{b}（全 {doc['row_count']:,} 行）", ""]
                                              + md_csv_table(doc, a, b))))
    return out


# ====================================================================== JSON・JSON Lines
def literal_text(n: JNode, text: str) -> str:
    """値の見せ方：JSON の書き方のまま（文字列は "…"）。数は元の字のまま。"""
    if n.kind == "string":
        return json.dumps(n.value, ensure_ascii=False)
    if n.kind == "number":
        return n.raw
    return text[n.a:n.b]


def _records(nodes: list[tuple[str, JNode]]) -> list[str] | None:
    """同じ名前の並びのオブジェクト（中は値だけ）が続くなら、その名前の並び（表にする）。"""
    if len(nodes) < 2:
        return None
    keys = None
    for _i, n in nodes:
        if n.kind != "object" or not n.members or not all(m[3].is_leaf for m in n.members):
            return None
        ks = [m[0] for m in n.members]
        if len(set(ks)) != len(ks):
            return None
        if keys is None:
            keys = ks
        elif ks != keys:
            return None
    return keys


def build_json(f: JsonFile) -> dict:
    values = []
    fields = []
    for i, n, _d in f.leaves():
        lit = literal_text(n, f.text)
        label = unesc_token(i.rsplit("/", 1)[-1]) if "/" in i else i
        values.append({"id": i, "kind": n.kind, "json": lit})
        kind = {"true": "checkbox", "false": "checkbox", "number": "number"}.get(n.kind, "text")
        fields.append({"id": i, "kind": kind, "label": label or i, "value": n.value, "json_type": n.kind,
                       "provenance": "deterministic", "context": f"{i} ＝ {lit}"})
    truncated = len(fields) > MAX_FIELDS
    if truncated:
        fields = [x for x in fields if x["value"] in (None, "")][:MAX_FIELDS]
    blocks = _blocks(f)
    return {
        "todoku": "0.2", "kind": f.kind,
        "source": {"file": f.filename, "sha256": f.sha256(), "bytes": len(f.blob),
                   "note": "元のファイルが正本。書き込みは、書いた値の字だけを置き換える。"},
        "format": {"encoding": f.dec.encoding, "encoding_ja": f.dec.label, "bom": bool(f.dec.bom),
                   "escape_nonascii": f.escape_nonascii, "documents": len(f.docs)},
        "notation": {"/a/b": "JSON Pointer（RFC 6901）。配列は /items/0（0から）。名前の中の / は ~1、~ は ~0",
                     "L3/a": "JSON Lines の3行目の値の中（L3 は3行目の全体）"},
        "values": values, "blocks": blocks, "fields": fields, "fields_truncated": truncated,
        "duplicates": list(f.duplicates),
    }


def _blocks(f: JsonFile) -> list[dict]:
    """窓に出す順のかたまり：同じ形のオブジェクトの並びは表に、ほかは値の一覧に。"""
    out: list[dict] = []
    if f.kind == "jsonl":
        rows = [(f"L{d.line}", d.root) for d in f.docs]
        keys = _records(rows)
        if keys is not None:
            out.append({"kind": "table", "id": "", "title": f"各行のオブジェクト（{len(rows)} 行）", "keys": keys,
                        "rows": [(rid, [literal_text(m[3], f.text) for m in n.members]) for rid, n in rows]})
            return out
        for rid, n in rows:
            _walk(f, n, rid, out)
        return out
    _walk(f, f.docs[0].root, "", out)
    return out


def _walk(f: JsonFile, n: JNode, pre: str, out: list[dict]) -> None:
    stack = [n]
    pending: list[tuple[str, str]] = []

    def flush():
        if pending:
            out.append({"kind": "list", "items": list(pending)})
            pending.clear()

    while stack:
        x = stack.pop()
        i = (pre + x.path) or "/"
        if x.is_leaf:
            pending.append((i, literal_text(x, f.text)))
            continue
        kids = [(pre + m[3].path, m[3]) for m in x.members] if x.kind == "object" else [(pre + k.path, k) for k in x.items]
        keys = _records(kids) if x.kind == "array" else None
        if keys is not None:
            flush()
            out.append({"kind": "table", "id": i, "title": f"{i}（配列・{len(kids)} 個。列は名前、行は番号）", "keys": keys,
                        "rows": [(rid, [literal_text(m[3], f.text) for m in k.members]) for rid, k in kids]})
            continue
        if not kids:
            pending.append((i, "{}" if x.kind == "object" else "[]"))
            continue
        stack.extend(reversed([k for _rid, k in kids]))
    flush()


def to_markdown_json(doc: dict) -> str:
    return "\n".join(md_json_head(doc) + md_json_body(doc)) + "\n"


def md_json_head(doc: dict) -> list[str]:
    fm = doc["format"]
    kind = "JSON Lines（1行に1つの JSON）" if doc["kind"] == "jsonl" else "JSON"
    out = [f"# {doc['source']['file'] or 'データ'}", "",
           f"- 元ファイルの SHA-256: `{doc['source']['sha256']}`（{doc['source']['bytes']:,} バイト）、{kind}・{fm['encoding_ja']}",
           "- この窓は読むためのものです。書き込み（`todoku fill`）は、書いた値の字だけを変えます（ほかの値・空白・改行・名前の順はそのまま）。",
           "- 値の id は JSON Pointer（RFC 6901）です：/applicant/name、配列は /items/0（0から）。"
           "ない名前に書くと、そのオブジェクトの終わりに足します。配列の終わりに足すときは /items/- と書きます。",
           "- 値は JSON の書き方のまま出しています（文字列は \"…\"、値のないものは null）。"]
    if doc["kind"] == "jsonl":
        out.append("- JSON Lines の id は、行の番号を付けて L3/name のように書きます（L3 は3行目の値の全体）。")
    for p in doc.get("duplicates") or []:
        out.append(f"- **同じ名前が2つ以上あります**：{p}（JSON Pointer では1つに決まらないので、ここには書けません）")
    out.append("")
    return out


def md_json_body(doc: dict) -> list[str]:
    out = [f"## 値（{len(doc['values']):,} 個）", ""]
    for b in doc["blocks"]:
        if b["kind"] == "list":
            out += [f"- `{i}` ＝ {v}" for i, v in b["items"]]
        else:
            out += ["", f"### {b['title']}", "", "| 位置 | " + " | ".join(_md(k) for k in b["keys"]) + " |",
                    "|---" * (len(b["keys"]) + 1) + "|"]
            out += ["| " + rid + " | " + " | ".join(_md(v) for v in vals) + " |" for rid, vals in b["rows"]]
            out.append("")
    return out


def _md(s: str) -> str:
    return str(s).replace("|", "\\|").replace("\n", "↵")


def parts_json(doc: dict, budget: int = BUDGET) -> list[tuple[str, str]]:
    whole = to_markdown_json(doc)
    if len(whole) <= budget:
        return [("全体", whole)]
    head = md_json_head(doc)
    lines = md_json_body(doc)
    parts, cur, size = [], [], 0
    for ln in lines:
        if cur and size + len(ln) + 1 > budget - 4000:
            parts.append(cur)
            cur, size = [], 0
        cur.append(ln)
        size += len(ln) + 1
    if cur:
        parts.append(cur)
    index = head + ["## この窓の目次", "", f"値が多いので、{len(parts)} の部分に分けました。", ""]
    index += [f"- 部分 {i}: {sum(len(x) + 1 for x in p):,} 文字" for i, p in enumerate(parts, 1)]
    return [("目次", "\n".join(index))] + [(f"部分 {i}", "\n".join(head[:1] + [""] + p)) for i, p in enumerate(parts, 1)]


# ====================================================================== 入口
def build(blob: bytes, filename: str, kind: str) -> dict:
    if kind in ("csv", "tsv"):
        return build_csv(CsvFile(blob, filename, kind))
    return build_json(JsonFile(blob, filename, kind))


def markdown_parts(doc: dict) -> list[tuple[str, str]]:
    return parts_csv(doc) if doc["kind"] in ("csv", "tsv") else parts_json(doc)
