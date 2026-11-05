"""規則の説明の中の食い違いを確かめる：text・a1・formula・r1c1・reads・range が、同じ 1 つの数式を指しているか。

説明は同じ形（R1C1）の規則で構文木を分け合うので、はじめて見たセルの番地がほかの規則の説明に出ると、字だけが別の所を指す。
ここでは説明の字（JSON）を、数式の字とは別に解析し直して比べる：
- a1 の {row}・{col} を規則の左上に当てはめて解析した形が、ファイルの数式（formula）の形・r1c1 と同じか
- text が「範囲 = a1」か
- 名前を通さない reads の pattern が a1 の中にあり、当てはめた番地が reads の range に入るか

確かめられなかった規則（読めない数式・データテーブルなど）は unchecked に数と理由を残す：「確かめた規則の数」が 0 のまま
通ることがないよう、数をそのまま結果に出す。"""

from __future__ import annotations

import re

from ..formula.parser import MAX_COL, MAX_ROW, RefNode, addr, iter_refs, num_to_col, parse
from .shape import r1c1, render

_SLOT = re.compile(r"\{(row|col)([+-]\d+)?\}")
_CELL = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)")
LISTED = 50


def _fill(text: str, r: int, c: int) -> str:
    """{row}・{col+2} を、(r, c) のセルでの番地の字にする。"""

    def sub(m: re.Match) -> str:
        d = int(m[2] or 0)
        return str(r + d) if m[1] == "row" else num_to_col(c + d)

    return _SLOT.sub(sub, text)


def _canon(formula: str, r: int, c: int) -> str:
    return render(parse(formula.lstrip("=")), lambda n: r1c1(n, r, c))


def _rects(rng) -> list[tuple[int, int, int, int]]:
    from ..formula.parser import col_to_num

    out = []
    for x in rng if isinstance(rng, list) else [rng]:
        a, _, b = str(x).partition(":")
        m, n = _CELL.fullmatch(a), _CELL.fullmatch(b or a)
        if m and n:
            out.append((int(m[2]), col_to_num(m[1]), int(n[2]), col_to_num(n[1])))
    return out


def _inside(rect: tuple[int, int, int, int], rects: list[tuple[int, int, int, int]]) -> bool:
    return any(x[0] <= rect[0] and x[1] <= rect[1] and rect[2] <= x[2] and rect[3] <= x[3] for x in rects)


def check_rule(d: dict, origin: tuple[int, int], listed_reads: int = 10 ** 9) -> list[str]:
    """規則 1 つの説明（rule_json の辞書）の食い違い。ないなら []。"""
    r, c = origin
    bad: list[str] = []
    a1 = d["a1"]
    filled = _fill(a1, r, c)
    try:
        got = _canon(filled, r, c)
    except Exception as e:  # noqa: BLE001  説明の字が解析できないことも食い違い
        return [f"a1「{a1}」を {addr(r, c)} に当てはめた式を解析できません（{type(e).__name__}）"]
    if d.get("r1c1") and got != d["r1c1"].lstrip("="):
        bad.append(f"a1「{a1}」の形が r1c1「{d['r1c1']}」と違います")
    if d.get("formula"):
        try:
            real = _canon(d["formula"].lstrip("{").rstrip("}"), r, c)
        except Exception as e:  # noqa: BLE001
            bad.append(f"formula「{d['formula']}」を解析できません（{type(e).__name__}）")
        else:
            if got != real:
                bad.append(f"a1「{a1}」を {addr(r, c)} に当てはめた「{filled}」が、ファイルの数式「{d['formula']}」と違います")
    body = a1[1:] if a1.startswith("=") else a1
    if d["kind"] in ("array", "dynamic_array", "array_cell"):
        body = "{=" + body + "}"
    if d.get("text") != f"{d['range']} = {body}":
        bad.append(f"text「{d.get('text')}」が「{d['range']} = {body}」と違います")
    for rd in d.get("reads") or []:
        pat = rd.get("pattern") or ""
        if not pat or rd.get("name") or rd.get("external") or rd.get("missing_sheet") or "[" in pat:
            continue
        here = _fill(pat, r, c)
        try:
            nodes = [n for n in iter_refs(parse(here)) if isinstance(n, RefNode)]
        except Exception:  # noqa: BLE001
            nodes = []
        if len(nodes) != 1:
            bad.append(f"reads の pattern「{pat}」を番地として解析できません")
            continue
        n = nodes[0]
        if pat not in a1:
            # 字の決まった INDIRECT は a1 に番地として出ないので、数式の中の字で比べる
            lit = rd.get("pattern", "").replace("$", "")
            if lit not in (d.get("formula") or "").replace("$", ""):
                bad.append(f"reads の pattern「{pat}」が a1「{a1}」にありません")
        rects = _rects(rd.get("range")) if rd.get("range") else []
        if rects and n.kind in ("range", "cell"):
            rect = (min(n.r1, n.r2), min(n.c1, n.c2), max(n.r1, n.r2), max(n.c1, n.c2))
            if MAX_ROW >= rect[2] and MAX_COL >= rect[3] and not _inside(rect, rects) and len(rects) < listed_reads:
                bad.append(f"reads の pattern「{pat}」が {addr(r, c)} で読む {here} が、reads の range「{rd.get('range')}」に入りません")
    return bad


def check(rule_dicts: list[dict], origins: dict[str, tuple[int, int]], listed_reads: int) -> dict:
    """{rules, checked, unchecked, unchecked_why, failures, failure_count}。"""
    failures: list[dict] = []
    checked = unchecked = n_failed = 0
    why: dict[str, int] = {}
    for d in rule_dicts:
        if not d.get("a1"):
            unchecked += 1
            k = "読めない数式" if d.get("unreadable") else d.get("kind", "?")
            why[k] = why.get(k, 0) + 1
            continue
        checked += 1
        bad = check_rule(d, origins[d["id"]], listed_reads)
        if bad:
            n_failed += 1
            if len(failures) < LISTED:
                failures.append({"rule": d["id"], "sheet": d.get("sheet"), "range": d.get("range"), "what": bad})
    return {"rules": len(rule_dicts), "checked": checked, "unchecked": unchecked, "unchecked_why": why,
            "failed": n_failed, "failures": failures, "ok": n_failed == 0 and checked + unchecked == len(rule_dicts)}
