"""書き出したモジュールを、別のプロセス（とどくを読み込めない状態）で動かし、保存された値・とどくの計算と 1 セルずつ比べる。"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

from ..formula.parser import addr
from ..formula.values import ExcelError
from .. import limits
from .pygen import Generated
from .rules import Rule
from .structure import plain

MISMATCH_LISTED = 200
CELL_BYTES = 420  # 計算したセル 1 つが別のプロセスで使うメモリの見積もり（行と列の鍵・値・控え）


def _decode(x):
    if isinstance(x, dict):
        if "e" in x:
            return ExcelError(x["e"])
        return x
    return x


class _RuleIndex:
    def __init__(self, rules: list[Rule]):
        from .rules import RectIndex

        self.ix = RectIndex()
        self.by_col: dict[str, dict[int, list]] = {}
        for u in rules:
            cols = self.by_col.setdefault(u.sheet, {})
            for m in u.members:
                for c in range(m.c1, m.c2 + 1):
                    cols.setdefault(c, []).append((m.r1, m.r2, u))
        for cols in self.by_col.values():
            for c in cols:
                cols[c].sort(key=lambda t: t[0])

    def at(self, sheet: str, r: int, c: int) -> Rule | None:
        col = self.by_col.get(sheet, {}).get(c)
        if not col:
            return None
        lo, hi = 0, len(col)
        while lo < hi:
            m = (lo + hi) // 2
            if col[m][1] < r:
                lo = m + 1
            else:
                hi = m
        if lo < len(col) and col[lo][0] <= r <= col[lo][1]:
            return col[lo][2]
        return None


def _evaluator_book(src, blob: bytes, name: str, k: str, cells: int):
    """とどくの計算に渡すブック。行の多いブックは、メモリに入るときだけ丸ごと読む。"""
    wb = getattr(src, "wb", None)
    if wb is not None:
        return wb, None
    total = sum(d["index"].cells for d in src.lz.sheet_meta if d["index"] is not None)
    if total * 700 > limits.MEMORY_LIMIT // 2:
        return None, f"セルが多い（{total:,} 個）ので、とどくの計算との比較はしていません（保存された値とだけ比べました）"
    from ..book import open_book

    return open_book(blob, name), None


def run_and_compare(gen: Generated, target: Path, src, internal: dict, res: dict, blob: bytes, name: str, k: str) -> dict:
    from ..xlsx.calc import Calc, _run_deep, values_equal

    rules: list[Rule] = internal["rules"]
    sheets = internal["sheets"]
    t0 = time.monotonic()
    target.parent.mkdir(parents=True, exist_ok=True)
    budget = max(limits.TIME_LIMIT * 0.5, 2.0)  # 別のプロセスの計算は、呼び出しの時間の半分まで（残りで比べて返す）
    max_cells = max(limits.MEMORY_LIMIT // 2 // CELL_BYTES, 1024)
    with tempfile.TemporaryDirectory(prefix="todoku-verify-") as td:
        mod = Path(td) / "module_under_test.py"
        mod.write_text(gen.source, encoding="utf-8")
        dump = Path(td) / "dump.jsonl"
        env = {k2: v for k2, v in os.environ.items() if k2 in ("PATH", "TMPDIR", "SYSTEMROOT", "HOME")}
        try:
            p = subprocess.run([sys.executable, "-I", "-S", str(mod), "--dump", str(dump), str(budget), str(max_cells)],
                               cwd=td, env=env, capture_output=True, text=True, timeout=budget + 30)
            run_err = None if p.returncode == 0 else (p.stderr or "")[-2000:]
        except subprocess.TimeoutExpired:
            run_err = f"{budget + 30:.0f} 秒で終わりませんでした"
        report = _compare(dump if dump.exists() else None, run_err, src, internal, res, blob, name, k, Calc, _run_deep, values_equal)
    report["seconds"] = round(time.monotonic() - t0, 2)
    full = None
    if report["mismatch_count"] > len(report["mismatches"]):
        full = target.with_name(target.name + ".mismatches.json")
        full.write_text(json.dumps(report.pop("_all_mismatches"), ensure_ascii=False, indent=0), encoding="utf-8")
        report["mismatches_file"] = str(full)
    report.pop("_all_mismatches", None)
    # モジュールに、確かめた結果を書いて置く
    summary = {k2: report[k2] for k2 in ("ran", "stored", "evaluator", "rules_proven", "rules_not_proven", "mismatch_count", "notes")}
    summary["not_proven"] = {uid: s for uid, s in report["rule_status"].items() if s["status"] != "proven"}
    final = gen.source.replace(gen.placeholder, "\nVERIFICATION = " + repr(summary) + "\n", 1)
    tmp = target.with_name(target.name + ".tmp")
    tmp.write_text(final, encoding="utf-8")
    os.replace(tmp, target)
    return report


def _compare(dump, run_err, src, internal, res, blob, name, k, Calc, _run_deep, values_equal) -> dict:
    rules: list[Rule] = internal["rules"]
    sheets = internal["sheets"]
    ri = _RuleIndex(rules)
    notes: list[str] = []
    cells = sum(s.formula_cells for s in sheets)
    wb, why = _evaluator_book(src, blob, name, k, cells)
    if why:
        notes.append(why)
    stats = {u.id: {"cells": u.count, "stored_match": 0, "stored_mismatch": 0, "no_stored": 0, "unknown": 0, "not_reached": 0,
                    "evaluator_match": 0, "evaluator_mismatch": 0, "evaluator_na": 0, "volatile": 0} for u in rules}
    mism: list[dict] = []
    stored_total = {"compared": 0, "match": 0}
    eval_total = {"compared": 0, "match": 0}
    ran = run_err is None and dump is not None
    stopped = None

    def lines():
        nonlocal stopped
        if dump is None:
            return
        with open(dump, encoding="utf-8") as f:
            for ln in f:
                x = json.loads(ln)
                if isinstance(x, dict):
                    stopped = x.get("stopped")
                    return
                yield x

    def body():
        calc = Calc(wb) if wb is not None else None
        late = time.monotonic() + limits.TIME_LIMIT * 0.15  # ここまで（全体の時間の上限の手前）で、とどくの計算との比較をやめる
        it = lines()
        cur = next(it, None)
        for si, sh in enumerate(sheets):
            sheet = sh.name
            for r, c, stored, is_str in _stored_with_gaps(src, si, sh, ri):
                u = ri.at(sheet, r, c)
                while cur is not None and (cur[0], cur[1], cur[2]) < (si, r, c):
                    cur = next(it, None)
                module_v, have = None, False
                if cur is not None and (cur[0], cur[1], cur[2]) == (si, r, c):
                    module_v, have = _decode(cur[3]), True
                    cur = next(it, None)
                if u is None:
                    continue
                st = stats[u.id]
                volatile = bool(u.__dict__.get("volatile") or u.__dict__.get("depends_on_volatile"))
                if not have:
                    st["not_reached"] += 1
                    continue
                if isinstance(module_v, dict) and "u" in module_v:
                    st["unknown"] += 1
                    st.setdefault("why", module_v["u"])
                    continue
                if volatile:
                    st["volatile"] += 1
                # 保存された値
                if stored is None and not is_str:
                    st["no_stored"] += 1
                elif not volatile:
                    stored_total["compared"] += 1
                    if values_equal(stored, module_v):
                        st["stored_match"] += 1
                        stored_total["match"] += 1
                    else:
                        st["stored_mismatch"] += 1
                        mism.append({"cell": f"{sheet}!{addr(r, c)}", "rule": u.id, "kind": "stored", "stored": plain(stored), "module": plain(module_v)})
                # とどくの計算
                if calc is not None and time.monotonic() > late:
                    calc = None
                    notes.append("時間の上限が近いので、とどくの計算との比較を途中でやめました（残りは保存された値とだけ比べました）")
                if calc is None:
                    st["evaluator_na"] += 1
                else:
                    actual = calc.cell_value(sheet, r, c)
                    status = calc.status.get((sheet, r, c), ("ok", ""))[0]
                    if status == "ok" and not volatile:
                        eval_total["compared"] += 1
                        if values_equal(actual, module_v) or _both_empty(actual, module_v):
                            st["evaluator_match"] += 1
                            eval_total["match"] += 1
                        else:
                            st["evaluator_mismatch"] += 1
                            mism.append({"cell": f"{sheet}!{addr(r, c)}", "rule": u.id, "kind": "evaluator", "evaluator": plain(actual), "module": plain(module_v)})
                    else:
                        st["evaluator_na"] += 1

    _run_deep(body)
    if stopped:
        notes.append(f"別のプロセスの計算を途中で止めました（{stopped}）。止まった先の規則は not_reached です。TODOKU_TIME_LIMIT で時間を延ばせます")
    if run_err:
        notes.append("モジュールが動きませんでした：" + run_err)
    status = {}
    proven = 0
    for u in rules:
        st = stats[u.id]
        status[u.id] = _rule_status(u, st, ran)
        if status[u.id]["status"] == "proven":
            proven += 1
    stored_rate = _rate(stored_total)
    eval_rate = _rate(eval_total)
    ordered = sorted(mism, key=lambda m: (m["rule"], m["cell"]))
    return {"ran": ran, "stored": {**stored_total, "rate": stored_rate}, "evaluator": {**eval_total, "rate": eval_rate} if wb is not None else None,
            "rules_proven": proven, "rules_not_proven": len(rules) - proven, "mismatch_count": len(mism),
            "mismatches": ordered[:MISMATCH_LISTED], "_all_mismatches": ordered, "notes": notes, "rule_status": status,
            "standalone": "別のプロセスで、とどくを読み込めない状態（python -I -S）で動かしました"}


def _both_empty(a, b) -> bool:
    return a in (None, "", 0.0) and b in (None, "", 0.0)


def _rate(t: dict) -> str:
    if not t["compared"]:
        return "比べたセルがありません"
    if t["match"] == t["compared"]:
        return "100%"
    return f"{100 * t['match'] / t['compared']:.4f}%"


def _rule_status(u: Rule, st: dict, ran: bool) -> dict:
    counts = {k: v for k, v in st.items() if v and k not in ("why",)}
    why = []
    if not ran:
        return {"status": "not_proven", "why": ["モジュールが動かなかったので、比べていない"], "counts": counts}
    if st["unknown"]:
        why.append("再現できない：" + st.get("why", ""))
    if st["stored_mismatch"]:
        why.append(f"保存された値と違うセルが {st['stored_mismatch']} 個")
    if st["evaluator_mismatch"]:
        why.append(f"とどくの計算と違うセルが {st['evaluator_mismatch']} 個")
    if st["not_reached"]:
        why.append(f"時間かメモリの上限で、比べていないセルが {st['not_reached']} 個")
    if st["volatile"]:
        why.append("計算のたびに値が変わる関数に頼るので、保存された値とは比べられない")
    if st["no_stored"]:
        why.append(f"保存された値がないセルが {st['no_stored']} 個")
    if u.kind == "data_table":
        why.append("データテーブルは計算しない")
    if not why and not st["stored_match"]:
        why.append("比べたセルがない")
    if why:
        return {"status": "mismatch" if st["stored_mismatch"] or st["evaluator_mismatch"] else "not_proven", "why": why, "counts": counts}
    return {"status": "proven", "counts": counts}


def _stored_with_gaps(src, si, sh, ri):
    """保存された値を、シートの順に。規則のセルなのにファイルに要素がないマス（配列数式のこぼれた所など）も出す。"""
    cols = ri.by_col.get(sh.name, {})
    want = []
    for c, items in cols.items():
        for a, b, u in items:
            if u.kind in ("array", "dynamic_array", "data_table"):
                want += [(r, c) for r in range(a, b + 1)]
    want = sorted(want)
    i = 0
    for r, c, v, is_str in src.stored(si):
        while i < len(want) and want[i] < (r, c):
            yield want[i][0], want[i][1], None, False
            i += 1
        if i < len(want) and want[i] == (r, c):
            i += 1
        yield r, c, v, is_str
    while i < len(want):
        yield want[i][0], want[i][1], None, False
        i += 1
