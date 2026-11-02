"""集めた様式すべてに todoku audit をかけ、どの検査で何件落ちたかをまとめる（汎用性の検査）。

1件ずつ別のプロセスで動かし、時間がかかりすぎるもの・途中で落ちるものも結果として記録する。
落ちたものは「その様式のせい」にせず、とどくの側の一般的な原因を探して直すための材料にする。

使い方: uv run python scripts/audit_corpus.py [corpus/wide] [--jobs 4] [--timeout 900]
"""

from __future__ import annotations

import argparse
import json
import multiprocessing as mp
import sys
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def one(path: str) -> dict:
    from todoku.xlsx.audit import run

    if path.lower().endswith(".pdf"):
        from todoku.pdf.audit import run
    elif path.lower().endswith((".docx", ".docm")):
        from todoku.docx.audit import run
    elif path.lower().endswith(".doc"):
        from todoku.doc.audit import run
    t = time.time()
    try:
        au = run(Path(path).read_bytes(), Path(path).name)
        d = au.to_dict()
    except Exception as e:  # noqa: BLE001
        d = {"file": Path(path).name, "ok": False, "crash": f"{type(e).__name__}: {e}",
             "traceback": traceback.format_exc()[-3000:], "checks": []}
    d["path"] = path
    d["seconds"] = round(time.time() - t, 1)
    return d


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("dir", nargs="?", default=str(ROOT / "corpus" / "wide"))
    ap.add_argument("--jobs", type=int, default=4)
    ap.add_argument("--timeout", type=int, default=900)
    a = ap.parse_args(argv)
    base = Path(a.dir)
    files = sorted(str(p) for p in base.rglob("*") if p.suffix.lower() in (".xlsx", ".xlsm", ".xls", ".pdf", ".docx", ".docm", ".doc") and p.is_file())
    print(f"{len(files)} 件の様式を検査します", flush=True)
    results = []
    ctx = mp.get_context("spawn")
    # 小さいファイルから順に（大きなファイルの結果を待たずに、終わったものから表示する）
    files.sort(key=lambda f: Path(f).stat().st_size)
    with ctx.Pool(a.jobs, maxtasksperchild=1) as pool:
        it = pool.imap_unordered(one, files)
        while True:
            try:
                r = it.next(timeout=a.timeout)
            except StopIteration:
                break
            except mp.TimeoutError:
                print(f"[待ち] {a.timeout} 秒のあいだ、どのファイルも終わりませんでした。打ち切ります", flush=True)
                break
            results.append(r)
            mark = ("OK・判断はエージェント" if r.get("needs_agent") else "OK") if r["ok"] else ("落ちた" if r.get("crash") else "要確認")
            bad = [c["name"] for c in r.get("checks", []) if not c["ok"]]
            print(f"[{mark}] {Path(r['path']).relative_to(base)} {r.get('seconds', '')}s {r.get('crash', '')} {bad}", flush=True)
    done = {r["path"] for r in results}
    for f in files:
        if f not in done:
            results.append({"file": Path(f).name, "path": f, "ok": False, "crash": "時間内に終わらなかった", "checks": []})
    (base / "audit.json").write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding="utf-8")
    n = len(results)
    ok = sum(r["ok"] for r in results)
    crash = sum(1 for r in results if r.get("crash"))
    by_check: dict[str, list[int]] = {}
    for r in results:
        for c in r.get("checks", []):
            s = by_check.setdefault(c["name"], [0, 0])
            s[0] += c["ok"]
            s[1] += 1
    agent = sum(1 for r in results if r["ok"] and r.get("needs_agent"))
    print(f"\n全体: {ok}/{n} 件で「全て正確に読み書きできる」（うち {agent} 件は、規則では決められない所を窓に示し、"
          f"エージェントの判断に任せた）。処理が落ちた・終わらなかった {crash} 件")
    for name, (good, total) in by_check.items():
        print(f"  {name}: {good}/{total}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
