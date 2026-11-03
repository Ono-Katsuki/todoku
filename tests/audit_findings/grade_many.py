"""PDF をいくつも開いて、読む段階・理由を1行ずつ出す。使い方: PYTHONPATH=src python grade_many.py DIR [OUT.json]"""

import json
import sys
from pathlib import Path

from todoku import api

out = {}
root = Path(sys.argv[1])
for p in sorted(root.rglob("*.pdf")):
    key = str(p.relative_to(root))
    try:
        g = api.grade(str(p))
        out[key] = {"read": g.get("grade"), "write": (g.get("write") or {}).get("grade"),
                    "reasons": [(r["grade"], r["title"], r["where"][:4]) for r in g.get("reasons", [])]}
    except Exception as e:  # noqa: BLE001
        out[key] = {"error": repr(e)[:200]}
    r = out[key]
    print(key, r.get("read"), r.get("write"), [x[:2] for x in r.get("reasons", [])], r.get("error", ""))
if len(sys.argv) > 2:
    Path(sys.argv[2]).write_text(json.dumps(out, ensure_ascii=False, indent=1))
