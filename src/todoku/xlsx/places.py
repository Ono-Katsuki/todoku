"""範囲の所：1 マスずつ欄にすると多すぎる宣言（入力規則・数式の参照・テーブル・配列数式）を、区間のまま持つ。

入力規則 C5:C40000 のような宣言は、1 マスずつの欄にしなくても、同じ事実を一定の大きさで言える：
「S!C5:C40000（空のセル 39,996 個・入力規則・数式が読む）」。読めない所は作らない（unread にしない）。
- どのマスも番地で書ける・読める（広げない）
- list_fields・read_window に rows を渡すと、頼んだ行の分だけ 1 マスずつの欄として返す（expand_fields）
- 重なる宣言は、重ならない区間に分ける（どの区間も、中のマスの事実が同じ）

ファイルに要素のあるマス・結合の左上は、いままでどおり 1 マスずつの欄になる。ここで数えるのは、ファイルに要素のない、空のマスだけ。
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .. import limits
from ..formula.parser import addr, num_to_col
from ..formula.values import ExcelError

MAX_DECOMPOSE = 2048  # 区間に分ける宣言の数（これを超えると、宣言ごとの区間のまま。重なりは overlap で伝える）
READERS_MAX = 20


def fits_window(r1: int, c1: int, r2: int, c2: int) -> bool:
    """宣言を 1 マスずつの欄にして、窓に入るか（limits.CONTEXT_BUDGET）。窓に書く字そのもので数える：行ごとの「- 77000: 」と、
    マスごとの「［B77000？］ ｜ 」（どれも範囲の右下の番地の長さで）。入らない宣言は範囲の所にする（頼まれた行だけ、1 マスずつの欄にする）。"""
    a = addr(r2, c2)
    return (r2 - r1 + 1) * (len(f"- {r2}: ") + (c2 - c1 + 1) * len(f"［{a}？］ ｜ ")) <= limits.CONTEXT_BUDGET


@dataclass
class Declared:
    """1 つの広い宣言：何の宣言か（what）と、その範囲。"""

    r1: int
    c1: int
    r2: int
    c2: int
    what: str  # validation / unlocked / formula_ref / table / array
    rule: object = None  # 入力規則（Validation）
    table: str = ""
    master: tuple | None = None


@dataclass
class Place:
    r1: int
    c1: int
    r2: int
    c2: int
    sources: list = field(default_factory=list)  # Declared の並び
    count: int = 0
    overlap: bool = False


def decompose(decls: list[Declared]) -> list[Place]:
    """宣言の範囲を、重ならない区間に分ける（同じ宣言の組の区間は、つながっていれば 1 つにする）。"""
    if not decls:
        return []
    if len(decls) > MAX_DECOMPOSE:
        return [Place(d.r1, d.c1, d.r2, d.c2, [d], overlap=True) for d in decls]
    cuts = sorted({d.c1 for d in decls} | {d.c2 + 1 for d in decls})
    bands: list[tuple[int, int, list]] = []
    for ca, cb in zip(cuts, cuts[1:]):
        active = [d for d in decls if d.c1 <= ca and d.c2 >= cb - 1]
        if not active:
            continue
        rows = sorted({d.r1 for d in active} | {d.r2 + 1 for d in active})
        pieces: list[list] = []
        for ra, rb in zip(rows, rows[1:]):
            here = [d for d in active if d.r1 <= ra and d.r2 >= rb - 1]
            if not here:
                continue
            key = tuple(id(d) for d in here)
            if pieces and pieces[-1][3] == key and pieces[-1][1] == ra - 1:
                pieces[-1][1] = rb - 1
            else:
                pieces.append([ra, rb - 1, here, key])
        if bands and bands[-1][1] == ca - 1 and [p[:2] + p[3:] for p in bands[-1][2]] == [p[:2] + p[3:] for p in pieces]:
            bands[-1] = (bands[-1][0], cb - 1, pieces)
        else:
            bands.append((ca, cb - 1, pieces))
    out = []
    for ca, cb, pieces in bands:
        for ra, rb, here, _k in pieces:
            out.append(Place(ra, ca, rb, cb, here))
    out.sort(key=lambda p: (p.r1, p.c1))
    return out


class Eligible:
    """シートのマスのうち、欄になれる（見えていて、保護でロックされていない）マスを数える。宣言の広さによらず、行・列の宣言の数で数える。"""

    def __init__(self, wb, sh):
        self.wb = wb
        self.sh = sh
        self.protected = sh.protected
        self.hidden_rows = sorted(sh.hidden_rows)
        self.hidden_cols = sorted(sh.hidden_cols)

    def _hidden_between(self, seq: list[int], a: int, b: int) -> int:
        import bisect

        return bisect.bisect_right(seq, b) - bisect.bisect_left(seq, a)

    def count(self, r1: int, c1: int, r2: int, c2: int) -> int:
        rows = (r2 - r1 + 1) - self._hidden_between(self.hidden_rows, r1, r2)
        cols = (c2 - c1 + 1) - self._hidden_between(self.hidden_cols, c1, c2)
        if rows <= 0 or cols <= 0:
            return 0
        if not self.protected:
            return rows * cols
        sh, wb = self.sh, self.wb
        ok_cols = sum(1 for c in range(c1, c2 + 1) if c not in sh.hidden_cols and not wb.style(sh.col_styles.get(c, 0)).locked)
        styled = [r for r in sh.row_styles if r1 <= r <= r2 and r not in sh.hidden_rows]
        out = (rows - len(styled)) * ok_cols
        out += sum(cols for r in styled if not wb.style(sh.row_styles[r]).locked)
        return out

    def one(self, r: int, c: int) -> bool:
        sh = self.sh
        if sh.is_hidden(r, c):
            return False
        if not self.protected:
            return True
        style = sh.row_styles[r] if r in sh.row_styles else sh.col_styles.get(c, 0)
        return not self.wb.style(style).locked


def free_count(wb, sh, el: Eligible, p: Place, links: dict) -> int:
    """区間の、ファイルに要素のない空のマスのうち、欄になれるマスの数。"""
    n = el.count(p.r1, p.c1, p.r2, p.c2)
    seen_in_merge: set = set()
    for rg in sh.__dict__.get("repeat_ranges") or ():  # OpenDocument の繰り返しのマスは、空ではない
        if rg.content:
            ra, rb = max(rg.r1, p.r1), min(rg.r2, p.r2)
            ca, cb = max(rg.c1, p.c1), min(rg.c2, p.c2)
            if ra <= rb and ca <= cb:
                n -= el.count(ra, ca, rb, cb)
                seen_in_merge.add((ra, ca, rb, cb))
    for m in sh.merges:
        ra, rb = max(m[0], p.r1), min(m[2], p.r2)
        ca, cb = max(m[1], p.c1), min(m[3], p.c2)
        if ra <= rb and ca <= cb:
            n -= el.count(ra, ca, rb, cb)
            seen_in_merge.add((ra, ca, rb, cb))
    for r, c in sh.existing_in(p.r1, p.c1, p.r2, p.c2):
        if el.one(r, c) and not any(a <= r <= b and x <= c <= y for a, x, b, y in seen_in_merge):
            n -= 1
    for r, c in links:
        if p.r1 <= r <= p.r2 and p.c1 <= c <= p.c2 and (r, c) not in sh.cells and el.one(r, c) \
                and not any(a <= r <= b and x <= c <= y for a, x, b, y in seen_in_merge):
            n -= 1
    return max(n, 0)


def _near(sh, p: Place) -> dict:
    """区間の上の見出しと、左の字（そのまま）。"""
    out = {}

    def text_at(r: int, c: int) -> str | None:
        cell = sh.cells.get((r, c))
        if cell is None or not isinstance(cell.value, str) or not cell.value.strip():
            return None
        return cell.value.strip()

    above = [x for x in sh.existing_in(max(1, p.r1 - 64), p.c1, p.r1 - 1, p.c2) if text_at(*x)] if p.r1 > 1 else []
    if above:
        r0 = max(r for r, _c in above)
        heads = [text_at(r0, c) for r, c in above if r == r0][:8]
        out["above"] = "・".join(h for h in heads if h)[:120]
    left = [x for x in sh.existing_in(p.r1, max(1, p.c1 - 16), min(p.r2, p.r1 + 8), p.c1 - 1) if text_at(*x)] if p.c1 > 1 else []
    if left:
        out["left"] = text_at(*left[0])[:60]
    return out


def _readers_summary(readers, p: Place) -> dict | None:
    """区間のどこかを読む数式：数と、最初の数個の番地。"""
    orders = []
    seen = set()
    for r1, c1, r2, c2, order in readers.boxes:
        if r1 <= p.r2 and r2 >= p.r1 and c1 <= p.c2 and c2 >= p.c1 and order not in seen:
            seen.add(order)
            orders.append(order)
    if not orders:
        return None
    orders.sort()
    return {"n": len(orders), "cells": [readers._name(o) for o in orders[:READERS_MAX]]}


WHY_JA = {"validation": "入力規則", "unlocked": "保護の中でロックを外した", "formula_ref": "数式が読む", "table": "テーブルの中"}


def describe(p: dict) -> str:
    """窓の 1 行：S!C5:C40000（空のセル 39,996 個・数式が読む・入力規則 …）。"""
    if p["why_key"] == ["array"]:
        return (f"{p['id']}（配列数式 {p['master']} の結果 {p['count']:,} マス。マスの値は try_calculation の cells で計算する。"
                "配列数式の範囲なので書けない）")
    if p["why_key"] == ["repeat"]:
        return f"{p['id']}（同じ中身「{p.get('repeats_value', '')}」を繰り返す {p['count']:,} マス。どのマスも番地で読めて、書ける）"
    if p["why_key"] == ["repeat_blank"]:
        return f"{p['id']}（枠・塗りなどの書式だけの空のマス {p['count']:,} 個。どのマスも番地で書ける）"
    bits = [f"空のセル {p['count']:,} 個"] + [WHY_JA.get(k, k) for k in p["why_key"]]
    for r in (p.get("rules") or [])[:2]:
        bits.append(_rule_text(r))
    near = p.get("near") or {}
    if near.get("above"):
        bits.append(f"上の見出し「{near['above']}」")
    if near.get("left"):
        bits.append(f"左の字「{near['left']}」")
    rb = p.get("referenced_by")
    if rb:
        bits.append(f"読む数式 {rb['n']:,} 個（{'、'.join(rb['cells'][:3])}{'…' if rb['n'] > 3 else ''}）")
    if p.get("table"):
        bits.append(f"テーブル {p['table']}")
    return f"{p['id']}（{'・'.join(bits)}）"


def _rule_text(r: dict) -> str:
    t = r.get("type") or ""
    return f"入力規則の種類 {t}" + (f"（{r['message']}）" if r.get("message") else "")


def place_docs(wb, sh, lay, opts) -> list[dict]:
    """シートの範囲の所：窓の JSON。"""
    from .window import AGENT, LAYOUT, _rule_dict, _declared_room

    decls: list[Declared] = list(lay.wide_inputs)
    decls += [Declared(r1, c1, r2, c2, "formula_ref") for r1, c1, r2, c2 in lay.readers.wide] if lay.shown_sheet else []
    for t in (getattr(wb, "tables", None) or {}).values():
        if t.sheet != sh.name or not lay.shown_sheet:
            continue
        first, last = t.r1 + t.header_rows, t.r2 - t.totals_rows
        if last >= first and (last - first + 1) * (t.c2 - t.c1 + 1) > _declared_room(wb):
            decls.append(Declared(first, t.c1, last, t.c2, "table", table=t.name))
    out: list[dict] = []
    if decls:
        el = Eligible(wb, sh)
        for p in decompose(decls):
            p.count = free_count(wb, sh, el, p, lay.links)
            if p.count <= 0:
                continue
            kinds = list(dict.fromkeys(d.what for d in p.sources))
            cell = f"{addr(p.r1, p.c1)}:{addr(p.r2, p.c2)}"
            is_input = "validation" in kinds or "unlocked" in kinds
            d = {"id": f"{sh.name}!{cell}", "cell": cell, "kind": "range", "count": p.count, "value": None, "why_key": kinds,
                 "rows": f"{p.r1}-{p.r2}", "cols": f"{num_to_col(p.c1)}-{num_to_col(p.c2)}",
                 "write_as": "どのマスも番地（シート!" + addr(p.r1, p.c1) + " など）で書けます。list_fields に sheet と rows を渡すと、その行の分を 1 マスずつの欄で返します",
                 "provenance": LAYOUT if is_input else AGENT}
            if not is_input:
                d["agent"] = True
            near = _near(sh, p)
            if near:
                d["near"] = near
            rules = [x.rule for x in p.sources if x.rule is not None]
            if rules:
                d["rules"] = [_rule_dict(v) for v in rules]
                for v in rules:
                    if v.type == "list":
                        o, dep = opts.get(sh.name, v, p.r1, p.c1)
                        if o is not None:
                            d["options"] = o[:500]
                            if len(o) > 500:
                                d["options_total"] = len(o)
                        if dep:
                            d["options_depend_on"] = dep
            rb = _readers_summary(lay.readers, p)
            if rb is not None:
                d["referenced_by"] = rb
                if "formula_ref" not in kinds:
                    d["why_key"] = kinds + ["formula_ref"]
            tables = [x.table for x in p.sources if x.table]
            if tables:
                d["table"] = tables[0]
            if p.overlap:
                d["overlap"] = True  # ほかの範囲の所と重なることがある（数は、この範囲の中の数）
            if not any(sh.in_print_area(r, c) for r, c in ((p.r1, p.c1), (p.r2, p.c2))):
                d["outside_print_area"] = True
            out.append(d)
    for rg in sh.__dict__.get("repeat_ranges") or ():  # OpenDocument の同じ中身の繰り返し
        cell = f"{addr(rg.r1, rg.c1)}:{addr(rg.r2, rg.c2)}"
        d = {"id": f"{sh.name}!{cell}", "cell": cell, "kind": "range", "count": (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1), "value": None,
             "why_key": ["repeat" if rg.content else "repeat_blank"], "rows": f"{rg.r1}-{rg.r2}", "cols": f"{num_to_col(rg.c1)}-{num_to_col(rg.c2)}",
             "write_as": "どのマスも番地（シート!" + addr(rg.r1, rg.c1) + " など）で書けます", "provenance": AGENT}
        if rg.content:
            first = sh.virtual_cell(rg.r1, rg.c1)
            if first is not None:
                d["repeats_value"] = first.value.code if isinstance(first.value, ExcelError) else first.value
        else:
            d["agent"] = True
        out.append(d)
    for r1, c1, r2, c2, m in getattr(sh, "array_spans", ()) or ():
        cell = f"{addr(r1, c1)}:{addr(r2, c2)}"
        out.append({"id": f"{sh.name}!{cell}", "cell": cell, "kind": "range", "count": (r2 - r1 + 1) * (c2 - c1 + 1), "value": None,
                    "why_key": ["array"], "master": addr(*m), "readonly": True, "rows": f"{r1}-{r2}", "cols": f"{num_to_col(c1)}-{num_to_col(c2)}",
                    "provenance": LAYOUT})
    return out


def markdown(places: list[dict]) -> list[str]:
    """窓の節。"""
    if not places:
        return []
    out = [f"### 範囲の所（{len(places)} か所。空のマスをまとめて示します）", "",
           "どのマスも番地で読めて、書けます。list_fields に sheet と rows（例 \"5-40\"）を渡すと、その行の分を 1 マスずつの欄で返します。", ""]
    return out + [f"- {describe(p)}" for p in places] + [""]


def expand_fields(wb, sh, lay, opts, places: list[dict], a: int, b: int, limit: int, only_empty: bool = True) -> tuple[list[dict], int | None]:
    """範囲の所のうち、行 a〜b の分を、1 マスずつの欄にする。(欄, 続きの行) を返す（limit で切ったとき）。"""
    from ..formula.parser import parse_cell
    from .window import AGENT_WHY, AGENT, LAYOUT, Box, _field_kind, _rule_dict, evidence_text, reason_codes

    out: list[dict] = []
    el = Eligible(wb, sh)
    rects = []
    for p in places:
        if p["why_key"] not in (["array"], ["repeat"]):
            (r1, c1), (r2, c2) = (parse_cell(x) for x in p["cell"].split(":"))
            rects.append((r1, c1, r2, c2, p))
    if not rects:
        return out, None
    a = max(a, min(x[0] for x in rects))
    b = min(b, max(x[2] for x in rects))
    for r in range(a, b + 1):
        for r1, c1, r2, c2, p in rects:
            if not r1 <= r <= r2:
                continue
            for c in range(c1, c2 + 1):
                if (r, c) in sh.cells or (r, c) in lay.links or sh.merged_at(r, c) is not None or not el.one(r, c):
                    continue
                if len(out) >= limit:
                    return out, r
                read = lay.readers.get((r, c))
                rules = [v for v in sh.validations if v.covers(r, c)]
                box = Box(r, c, r, c, "agent", None, "", outside=not sh.in_print_area(r, c))
                d = lay.describe_agent(box)
                is_input = bool(rules) or "unlocked" in p["why_key"]
                f = {"id": f"{sh.name}!{addr(r, c)}", "cell": addr(r, c), "kind": _field_kind(wb, None, rules), "value": None, **d,
                     "provenance": LAYOUT if is_input else AGENT, "range_place": p["id"]}
                if rules:
                    f["rules"] = [_rule_dict(v) for v in rules]
                    f["reason"] = ["入力規則があるセル"]
                    for v in rules:
                        if v.type == "list":
                            o, _dep = opts.get(sh.name, v, r, c)
                            if o is not None:
                                f["options"] = o[:500]
                elif is_input:
                    f["reason"] = ["シート保護の中で編集できるセル"]
                else:
                    whys = ["formula_ref"] if read is not None else []
                    ev = {"referenced_by": read} if read is not None else {}
                    f.update({"agent": True, "why": "・".join(AGENT_WHY[k] for k in whys), "why_key": whys, "evidence": ev,
                              "evidence_text": evidence_text(ev, lay.side("左")), "reason": [AGENT_WHY[k] for k in whys]})
                if read is not None:
                    f["referenced_by"] = read
                if f.get("reason"):
                    f["reason_codes"] = reason_codes(f["reason"])
                out.append(f)
    return out, None
