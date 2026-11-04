#!/usr/bin/env python3
"""処理の取り出しの評価（Excel の書類）：extract_processing が書き出した Python が、保存された計算結果と合うか。
- 数式の値が保存された値と合う割合（とどくの verification.stored）。保存された値は、Excel が最後に計算した値（とどくが作った値ではない）
- 規則の数と proven の割合。確かめられなかった規則は、その理由ごとに数える（黙って通らない）
- 計算できない関数を言うか（unknown_functions）。XML から数えた関数と数式のセルの数に合わせて、言わずに落としたものがないかを見る
- 比べる道具でできること：openpyxl は数式の字を読むだけ（値の再計算も規則のまとまりもない）
使い方: processing_eval.py --files 一覧 --out 出力 --baseline-python PY --todoku-python PY [--label dev]"""

from __future__ import annotations

import argparse
import collections
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from common import ROOT, atomic_json, doc_id, guard_path, kind_of, last_json, load_json, md_table, pct, read_list, run_watched  # noqa: E402


def summarize(recs: list[dict]) -> dict:
    ok = [r for r in recs if not r.get("todoku", {}).get("error")]
    with_f = [r for r in ok if (r["todoku"].get("formula_cells") or 0) > 0]
    stored_c = sum(r["todoku"].get("stored_compared") or 0 for r in with_f)
    stored_m = sum(r["todoku"].get("stored_match") or 0 for r in with_f)
    rules = sum(r["todoku"].get("rules") or 0 for r in with_f)
    proven = sum(r["todoku"].get("rules_proven") or 0 for r in with_f)
    unknown = collections.Counter(fn for r in with_f for fn in r["todoku"].get("unknown_functions") or [])
    status = collections.Counter()
    why = collections.Counter()
    for r in with_f:
        status.update(r["todoku"].get("rule_status") or {})
        why.update(r["todoku"].get("not_proven_why") or [])
    return {"docs": len(recs), "failed": [{"path": r["path"], "why": r["todoku"].get("error")} for r in recs if r.get("todoku", {}).get("error")],
            "docs_with_formulas": len(with_f), "formula_cells": sum(r["todoku"].get("formula_cells") or 0 for r in with_f),
            "xml_formula_cells": sum(r.get("baseline", {}).get("xml_formula_cells") or 0 for r in with_f),
            "stored_compared": stored_c, "stored_match": stored_m, "rules": rules, "rules_proven": proven, "rule_status": dict(status),
            "unknown_functions": dict(unknown), "docs_reporting_unknown": sum(1 for r in with_f if r["todoku"].get("unknown_functions")),
            "not_proven_why": why.most_common(10), "explain_ok": sum(1 for r in with_f if r["todoku"].get("explain_ok")),
            "baseline_formula_text_cells": sum(r.get("baseline", {}).get("openpyxl_formula_text_cells") or 0 for r in with_f)}


def markdown(s: dict, label: str) -> str:
    rows = [["数式のある書類", f"{s['docs_with_formulas']} / {s['docs']}", "—"],
            ["数式のセル（とどく・XML から数えた数）", f"{s['formula_cells']:,}・{s['xml_formula_cells']:,}", "字だけ " + f"{s['baseline_formula_text_cells']:,}（openpyxl）"],
            ["計算して保存された値と合う", pct(s["stored_match"], s["stored_compared"], 2), "計算しない"],
            ["規則（同じ形をまとめたもの）", f"{s['rules']:,}", "まとめない"],
            ["確かめられた規則（proven）", pct(s["rules_proven"], s["rules"]), "—"],
            ["計算できない関数を言った書類", f"{s['docs_reporting_unknown']}（{', '.join(s['unknown_functions']) or 'なし'}）", "言わない（計算しないので関係しない）"],
            ["explain_calculation が動いた", f"{s['explain_ok']} / {s['docs_with_formulas']}", "—"]]
    L = [f"# 処理の取り出し（{label}）", "", md_table(["", "とどく", "比べる道具"], rows), "", f"規則の結果：{s['rule_status']}", ""]
    if s["not_proven_why"]:
        L += ["確かめられなかった規則の理由（多い順）："] + [f"- {w}（{n}）" for w, n in s["not_proven_why"]] + [""]
    if s["failed"]:
        L += ["処理の取り出しが失敗した書類："] + [f"- {Path(x['path']).name}：{x['why']}" for x in s["failed"]] + [""]
    return "\n".join(L)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--files", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--baseline-python", required=True)
    ap.add_argument("--todoku-python", required=True)
    ap.add_argument("--label", default="")
    a = ap.parse_args()
    base = Path(a.out)
    (base / "scores").mkdir(parents=True, exist_ok=True)
    files = [p for p in read_list(a.files) if kind_of(p) in ("xlsx", "xls")]
    recs = []
    for i, p in enumerate(files, 1):
        guard_path(p)
        sp = base / "scores" / f"{doc_id(p)}.json"
        rec = load_json(sp)
        if rec is None:
            out_py = base / "work" / f"{doc_id(p)}.py"
            out_py.parent.mkdir(parents=True, exist_ok=True)
            res = run_watched([a.todoku_python, "-I", HERE / "todoku_process.py", ROOT / "src", p, out_py], limit_s=600, mem_mb=3000)
            rec = {"path": p, "todoku": last_json(res) if not res.get("reason") else {"error": res["reason"]},
                   "baseline": last_json(run_watched([a.baseline_python, "-I", HERE / "baseline_process.py", p], limit_s=300))}
            atomic_json(sp, rec)
            out_py.unlink(missing_ok=True)
            t = rec["todoku"]
            print(f"[{i}/{len(files)}] {Path(p).name} 数式 {t.get('formula_cells')} 合う {t.get('stored_match')}/{t.get('stored_compared')} 規則 {t.get('rules_proven')}/{t.get('rules')} {t.get('error') or ''}", flush=True)
        recs.append(rec)
    s = summarize(recs)
    atomic_json(base / "summary.json", s)
    (base / "summary.md").write_text(markdown(s, a.label), encoding="utf-8")
    print(markdown(s, a.label))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
