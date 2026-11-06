"""とどくで読む：MCP の read_window が返すのと同じ Markdown の窓の全部。段階の理由（grade）と計算の説明の字も別に残す（古い値の印を確かめるため）。
使い方（python -I）: todoku_read.py <src のフォルダ> <ファイル> <出力>   最後の行の JSON に結果"""

import json
import sys
import time

sys.dont_write_bytecode = True  # 道具の版のフォルダを変えない
sys.path.insert(0, sys.argv[1])
from todoku import api  # noqa: E402

p, out = sys.argv[2], sys.argv[3]
o = {"tool": "todoku"}
t = time.time()
try:
    k, _doc, parts = api.window(p)
    txt = "\n".join(s for _n, s in parts)
    with open(out, "w", encoding="utf-8") as f:
        f.write(txt)
    o["chars"] = len(txt)
    o["doc_kind"] = k
    o["seconds"] = round(time.time() - t, 2)
    try:
        g = api.grade(p)
        o["read_grade"] = g.get("grade")
        o["write_grade"] = (g.get("write") or {}).get("grade")
        o["grade_reasons"] = [str(x.get("title"))[:200] for x in g.get("reasons") or []][:20]
    except Exception as e:  # noqa: BLE001
        o["grade_error"] = f"{type(e).__name__}: {str(e)[:120]}"
except Exception as e:  # noqa: BLE001
    o["error"] = f"{type(e).__name__}: {str(e)[:200]}"
    o["seconds"] = round(time.time() - t, 2)
print(json.dumps(o, ensure_ascii=False, default=str))
