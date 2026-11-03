"""Excel のテーブル（リストオブジェクト）の見出し・集計の行のセル。

テーブルの列の名前（xl/tables/tableN.xml の tableColumn/@name）は、見出しのセルの字と同じでなければならない（違うと、Excel は
開くときに修復を求め、構造化参照 T1[金額] の指す列が変わる）。集計の行の見出し（totalsRowLabel）も、そのセルの字と同じ。

- 見出し・集計の行の見出しのセルへの書き込みは、ふつうは書かない（edit.plan_inputs が断る）
- force で書くと決めたときは、Excel で見出しを書き換えたときと同じく、列の名前（tableColumn/@name）・集計の行の見出し
  （totalsRowLabel）と、その列を指す構造化参照（セルの数式・calculatedColumnFormula・totalsRowFormula）を合わせて書き換える
- 証明（proof.py）は、書いた後のブックで、見出しのセルの字と列の名前が合っていること、数式の構造化参照がすべて、ある列を
  指していること、書き換わった数式・テーブルの部品が、列の名前の書き換えのほかに変わっていないことを確かめる
"""

from __future__ import annotations

import re
from dataclasses import replace
from xml.sax.saxutils import escape, unescape

from ..formula.parser import StructRef, _read_bracket, _unesc

HEADER_REASON = "表の見出し：書き換えると表の定義とずれる"
TOTALS_REASON = "表の集計の行の見出し：書き換えると表の定義とずれる"


def _esc_col(name: str) -> str:
    return re.sub(r"([\[\]#'])", r"'\1", name)


def _plain_name(name: str) -> bool:
    """括弧で囲まずに T1[名前] と書ける列の名前か（字・数字・_ だけ）。"""
    return bool(name) and all(ch.isalnum() or ch == "_" for ch in name)


def _rename_item(item: str, cols: dict[str, str]) -> str:
    """[…] の中身（列の名前・#Totals など）を、名前が cols の鍵なら新しい名前にする（' のエスケープを付け直す）。"""
    raw = item
    lead = ""
    if raw.startswith("@"):
        lead, raw = "@", raw[1:]
    name = _unesc(raw)
    new = cols.get(name.lower()) if not name.startswith("#") else None
    if new is None and name.strip() != name:
        new = cols.get(name.strip().lower())
    if new is None:
        return item
    return lead + _esc_col(new)


def _rename_content(content: str, cols: dict[str, str]) -> str:
    """T1[…] の […] の中身の、列の名前を書き換える。"""
    c = content
    if c.strip() == "":
        return c
    if c.startswith("@"):
        rest = c[1:]
        if rest.lstrip().startswith("["):
            k = rest.index("[")
            inner, j = _read_bracket(rest, k)
            return "@" + rest[:k] + "[" + _rename_item(inner, cols) + "]" + rest[j:]
        new = _rename_item(rest, cols)
        if new == rest:
            return c
        return "@" + (new if _plain_name(_unesc(new)) else "[" + new + "]")
    if not c.lstrip().startswith("["):
        new = _rename_item(c, cols)
        if new == c:
            return c
        return new if _plain_name(_unesc(new)) else "[" + new + "]"
    out, i = [], 0
    while i < len(c):
        if c[i] == "[":
            inner, j = _read_bracket(c, i)
            out.append("[" + _rename_item(inner, cols) + "]")
            i = j
        else:
            out.append(c[i])
            i += 1
    return "".join(out)


_WORD = re.compile(r"[A-Za-z_\\\u0080-￿][\w.\u0080-￿]*")


def rename_in_formula(text: str, renames: dict[str, dict[str, str]], here: str | None) -> str:
    """数式の字の、構造化参照の列の名前を書き換える。renames：テーブルの名前（小文字）→ {元の列の名前（小文字）: 新しい名前}。
    here：数式のあるセルが入っているテーブルの名前（テーブルの名前のない参照 [列] の、テーブル）。字の中（"…"）は変えない。"""
    if not renames or "[" not in text:
        return text
    out, i, n = [], 0, len(text)
    while i < n:
        ch = text[i]
        if ch == '"':
            j = i + 1
            while j < n:
                if text[j] == '"':
                    if j + 1 < n and text[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            out.append(text[i:j + 1])
            i = j + 1
            continue
        if ch == "'":
            # 'シート名'!A1：シート名の中は変えない
            j = text.find("'", i + 1)
            while j != -1 and j + 1 < n and text[j + 1] == "'":
                j = text.find("'", j + 2)
            j = n - 1 if j == -1 else j
            out.append(text[i:j + 1])
            i = j + 1
            continue
        if ch == "[":
            if re.match(r"\[\d+\]", text[i:]):
                out.append(ch)  # 外部ブックの番号 [1]Sheet1!A1
                i += 1
                continue
            content, j = _read_bracket(text, i)
            cols = renames.get((here or "").lower())
            out.append("[" + (_rename_content(content, cols) if cols else content) + "]")
            i = j
            continue
        m = _WORD.match(text, i)
        if m and (i == 0 or not (text[i - 1].isalnum() or text[i - 1] in "_.")):
            word = m[0]
            j = m.end()
            if j < n and text[j] == "[":
                content, j2 = _read_bracket(text, j)
                cols = renames.get(word.lower())
                out.append(word + "[" + (_rename_content(content, cols) if cols else content) + "]")
                i = j2
                continue
            out.append(word)
            i = j
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def rename_ast(node, renames: dict[str, dict[str, str]], here: str | None):
    """構文木の構造化参照の列の名前を書き換えた構文木（証明で、書いた後の数式と意味で比べる）。"""
    from dataclasses import fields, is_dataclass

    if isinstance(node, StructRef):
        cols = renames.get((node.table or here or "").lower())
        if not cols:
            return node
        return replace(node, col1=cols.get(node.col1.lower(), node.col1) if node.col1 else node.col1,
                       col2=cols.get(node.col2.lower(), node.col2) if node.col2 else node.col2)
    if is_dataclass(node):
        kw = {}
        for f in fields(node):
            v = getattr(node, f.name)
            if isinstance(v, list):
                kw[f.name] = [rename_ast(x, renames, here) for x in v]
            elif is_dataclass(v):
                kw[f.name] = rename_ast(v, renames, here)
        return replace(node, **kw) if kw else node
    return node


def struct_refs(node):
    """構文木の構造化参照をすべて。"""
    from dataclasses import fields, is_dataclass

    if isinstance(node, StructRef):
        yield node
        return
    if is_dataclass(node):
        for f in fields(node):
            v = getattr(node, f.name)
            for x in (v if isinstance(v, list) else [v]):
                if is_dataclass(x):
                    yield from struct_refs(x)


def table_at(wb, sheet: str, r: int, c: int):
    """そのセルが入っているテーブル（なければ None）。"""
    for t in getattr(wb, "tables", {}).values():
        if t.sheet == sheet and t.r1 <= r <= t.r2 and t.c1 <= c <= t.c2:
            return t
    return None


def label_cell(wb, sheet: str, r: int, c: int):
    """テーブルの見出しか集計の行のセルなら (テーブル, "header"|"totals", 列の番号)、ほかは None。"""
    t = table_at(wb, sheet, r, c)
    if t is None:
        return None
    if t.header_rows and r == t.r1:
        return t, "header", c - t.c1
    if t.totals_rows and r == t.r2:
        return t, "totals", c - t.c1
    return None


# ---------------------------------------------------------------- テーブルの部品（xl/tables/tableN.xml）
_COL_TAG = re.compile(rb"<(?:[\w.-]+:)?tableColumn\b[^>]*?/?>")
_FORMULA_EL = re.compile(rb"(<(?:[\w.-]+:)?(calculatedColumnFormula|totalsRowFormula)\b[^>]*?>)(.*?)(</(?:[\w.-]+:)?\2>)", re.S)


def _set_attr(tag: bytes, name: bytes, value: str | None) -> bytes:
    pat = re.compile(rb"(\s" + name + rb'=")[^"]*(")')
    if value is None:
        return re.sub(rb"\s" + name + rb'="[^"]*"', b"", tag, count=1)
    v = escape(value, {'"': "&quot;"}).encode("utf-8")
    if pat.search(tag):
        return pat.sub(lambda m: m[1] + v + m[2], tag, count=1)
    end = -2 if tag.endswith(b"/>") else -1
    return tag[:end].rstrip() + b" " + name + b'="' + v + b'"' + tag[end:]


def table_xml(xml: bytes, table_name: str, col_changes: dict[int, dict[str, str]],
              renames: dict[str, dict[str, str]]) -> bytes:
    """テーブルの部品の XML：列 i の属性を col_changes[i]（name・totalsRowLabel）にし、数式の要素（calculatedColumnFormula・
    totalsRowFormula）の構造化参照の列の名前を renames で書き換える。ほかのバイトは変えない。"""
    tags = list(_COL_TAG.finditer(xml))
    out, last = bytearray(), 0
    for i, m in enumerate(tags):
        if i not in col_changes:
            continue
        tag = m[0]
        for k, v in col_changes[i].items():
            tag = _set_attr(tag, k.encode(), v)
        out += xml[last:m.start()] + tag
        last = m.end()
    out += xml[last:]
    xml = bytes(out)

    def sub(m: re.Match) -> bytes:
        text = unescape(m[3].decode("utf-8"))
        new = rename_in_formula(text, renames, table_name)
        if new == text:
            return m[0]
        return m[1] + escape(new).encode("utf-8") + m[4]

    return _FORMULA_EL.sub(sub, xml)


def column_attrs(xml: bytes) -> list[dict[str, str]]:
    """テーブルの部品の列ごとの属性（name・totalsRowLabel・totalsRowFunction）。"""
    out = []
    for m in _COL_TAG.finditer(xml):
        d = {}
        for k, v in re.findall(rb'\s([\w:]+)="([^"]*)"', m[0]):
            d[k.decode().split(":")[-1]] = unescape(v.decode("utf-8"), {"&quot;": '"', "&apos;": "'"})
        out.append(d)
    return out
