"""PDF の窓の読み順の行だけを出す。使い方: PYTHONPATH=src python read_lines.py FILE.pdf"""
import sys
from todoku import api

k, doc, parts = api.window(sys.argv[1])
md = "\n".join(m for _t, m in parts)
for ln in md.splitlines():
    if ln.startswith("- ") and "SHA-256" not in ln and "記号" not in ln and "窓です" not in ln and "目印" not in ln:
        print(ln)
g = api.grade(sys.argv[1])
print("grade:", g.get("grade"), [r["title"] for r in g.get("reasons", [])], [a["title"] for a in g.get("alerts", [])])
