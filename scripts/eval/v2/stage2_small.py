#!/usr/bin/env python3
"""第 2 段の新しい標本がない 24 か国を測るための書類を選ぶ（docs/再評価の計画.md 15 節）。結果を見ずに決まる。

対象：stage2_plan.json に層がない国の、world/sample.jsonl の書類（この国の層は候補が 50 件以下で、新しい書類がない）。
選び方：調査で全体を取得できた書類のうち、固定版の subset.json に入っていない書類を先に、種と digest のハッシュの順で、各層 2 件。
足りない層だけ subset.json の書類から足す。使い方: stage2_small.py <固定版の subset.json> <書き出す subset.json>"""

import collections
import hashlib
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
PLAN = json.loads((Path(__file__).parent / "stage2_plan.json").read_text(encoding="utf-8"))
SEED = PLAN["seed"]
PER_STRATUM = 2
FIELDS = ("id", "country", "family", "rank", "domain_source", "pi", "crawl", "url", "timestamp", "mime", "mime_detected", "digest", "filename", "offset", "length", "block", "block_stratum", "sha256")  # 調査の段階と理由は渡さない


def main() -> int:
    measured = {d["id"] for d in json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))["docs"]}
    used = {a["country"] for a in PLAN["allocation"]}
    strata = collections.defaultdict(list)
    for line in (ROOT / "world/results.jsonl").open(encoding="utf-8"):
        r = json.loads(line)
        if r["country"] not in used and r.get("reach") == "ok":
            strata[(r["country"], r["family"], r["domain_source"])].append(r)
    docs = []
    for key in sorted(strata):
        order = sorted(strata[key], key=lambda r: (r["id"] in measured, hashlib.sha256(f"{SEED}:{r['digest']}".encode()).hexdigest()))
        for r in order[:PER_STRATUM]:
            docs.append({k: r[k] for k in FIELDS if k in r})
    out = {"seed": SEED, "per_stratum": PER_STRATUM, "strata": len(strata), "docs": docs}
    Path(sys.argv[2]).write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
    print(f"国 {len({k[0] for k in strata})}・層 {len(strata)}・書類 {len(docs)}（固定版の subset と重なる {sum(d['id'] in measured for d in docs)}）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
