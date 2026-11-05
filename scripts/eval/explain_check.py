#!/usr/bin/env python3
"""計算の説明が正しいかを、別のエージェントが確かめる（計画：several forms）。数式の多い書類から種で選び、とどくの説明（explain_calculation と同じ calc_model・規則）と、
openpyxl で読んだ数式の字を、作業フォルダに置いて比べさせる。エージェントはファイルを読むだけで、todoku の道具は持たない。
使い方: explain_check.py --processing 処理の結果のフォルダ --out 出力 --baseline-python PY --todoku-python PY [--n 5] [--model opus]"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from common import ROOT, atomic_json, doc_id, guard_path, load_json, md_table, run_watched  # noqa: E402
from fill_runner import blind_check, claude_cmd, run_agent  # noqa: E402

DUMP = r"""
import json, sys
sys.path.insert(0, sys.argv[1])
from todoku import api
m = api.calc_model(sys.argv[2])
r = api.extract_processing(sys.argv[2])
keep = {k: r.get(k) for k in ("summary", "rules", "inputs", "parameters", "order", "cycles", "unknown_functions", "external_references") if k in r}
json.dump({"calc_model": m, "processing": keep}, open(sys.argv[3], "w"), ensure_ascii=False, indent=0, default=str)
"""
FORMULAS = r"""
import json, sys, openpyxl
wf = openpyxl.load_workbook(sys.argv[1])
wv = openpyxl.load_workbook(sys.argv[1], data_only=True)
out = []
for ws in wf.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                out.append({"sheet": ws.title, "cell": c.coordinate, "formula": c.value, "cached": wv[ws.title][c.coordinate].value})
json.dump(out[:3000], open(sys.argv[2], "w"), ensure_ascii=False, indent=0, default=str)
"""
PROMPT = """あなたは計算の説明の確かめ役です。作業フォルダ {D} の 2 つのファイルだけを Read で読みます。
- explain.json：道具（とどく）がこの Excel の書類の計算を説明したもの（規則・入力・順番など）
- formulas.json：別の道具で読んだ、数式の字と保存された値（正解の基準。3000 セルまで）

規則を 10 個まで選び（種類が違うものを優先）、説明が数式と合っているかを確かめます。範囲・参照先・計算の中身・入力と出力の別・計算できないと言っている所が正しいか。
甘くしないでください。Write で {R} に、次の JSON を書いて終えます。
{{"rules_checked": 0, "correct": 0, "wrong": [{{"rule": "規則の id", "what": "何が違うか", "formula": "数式の字"}}], "misleading": ["誤解を招く説明"], "notes": "短く"}}
"""


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--processing", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--baseline-python", required=True)
    ap.add_argument("--todoku-python", required=True)
    ap.add_argument("--n", type=int, default=5)
    ap.add_argument("--model", default="opus")
    ap.add_argument("--seed", type=int, default=20261105)
    a = ap.parse_args()
    base = Path(a.out)
    base.mkdir(parents=True, exist_ok=True)
    recs = [load_json(p) for p in sorted((Path(a.processing) / "scores").glob("*.json"))]
    cands = [r for r in recs if (r.get("todoku", {}).get("formula_cells") or 0) >= 5 and r["path"].lower().endswith((".xlsx", ".xlsm"))]
    cands.sort(key=lambda r: hashlib.sha256(f"{a.seed}:{r['path']}".encode()).hexdigest())
    out = []
    for r in cands[: a.n]:
        guard_path(r["path"])
        d = base / doc_id(r["path"])
        d.mkdir(exist_ok=True)
        res_path = d / "check.json"
        if not res_path.exists():
            run_watched([a.todoku_python, "-I", "-c", DUMP, ROOT / "src", r["path"], d / "explain.json"], limit_s=600, mem_mb=3000)
            run_watched([a.baseline_python, "-I", "-c", FORMULAS, r["path"], d / "formulas.json"], limit_s=300)
            if (d / "explain.json").exists() and (d / "formulas.json").exists():
                allowed = [f"Read(//{str(d).lstrip('/')}/**)", f"Edit(//{str(d).lstrip('/')}/**)"]
                cmd = claude_cmd(PROMPT.format(D=d, R=res_path), d, a.model, None, None, "Read,Write", allowed)
                run_agent(cmd, d, d / "transcript.jsonl", 20 * 60)
                atomic_json(d / "blind.json", blind_check(d / "transcript.jsonl", d, {"Read", "Write"}))
        got = load_json(res_path) or {}
        out.append({"file": Path(r["path"]).name, "formula_cells": r["todoku"].get("formula_cells"), **got})
        print(Path(r["path"]).name, got.get("correct"), "/", got.get("rules_checked"), flush=True)
    s = {"docs": len(out), "rules_checked": sum(x.get("rules_checked") or 0 for x in out), "correct": sum(x.get("correct") or 0 for x in out), "items": out}
    atomic_json(base / "summary.json", s)
    rows = [[x["file"], x["formula_cells"], f"{x.get('correct')}/{x.get('rules_checked')}", "；".join(str(w.get("what"))[:80] for w in x.get("wrong") or [])] for x in out]
    (base / "summary.md").write_text("# 計算の説明の確かめ（エージェントが数式の字と比べた）\n\n" + md_table(["書類", "数式のセル", "正しい規則", "違い"], rows) + "\n", encoding="utf-8")
    print(json.dumps({k: s[k] for k in ("docs", "rules_checked", "correct")}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
