"""第 2 段が時間の上限で止まったとき：採点の済んでいない書類を、数の対象から外す（順は結果によらないので、残りは無作為な一部）。
docs_map.json の、その書類を {"skip": "時間の上限で測らなかった"} にする。使い方: stage2_trim.py <出力のフォルダ>"""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import atomic_json, doc_id, kind_of  # noqa: E402

out = Path(sys.argv[1])
mp = json.loads((out / "docs_map.json").read_text(encoding="utf-8"))
n = 0
for i, m in mp.items():
    p = m.get("path")
    if not p or not kind_of(p):
        continue
    rd = (out / "read" / "scores" / f"{doc_id(p)}.json").exists()
    wr = (out / "write" / "scores" / f"{doc_id(p)}.json").exists() or kind_of(p) not in {"xlsx", "docx", "pdf", "xls", "doc"}
    if not (rd and wr):
        mp[i] = {"skip": "時間の上限で測らなかった", "path_was": p}
        n += 1
atomic_json(out / "docs_map.json", mp)
print(f"測らなかった書類 {n} 件を外した")
