"""とどくで書く（比べのため）：pick は書く欄を 1 つ選ぶ（比べる道具にも同じ欄を渡す：道具が欄を見つける力は、ここでは問わない）。fill は書く。
使い方（python -I）: todoku_write.py <src> <ファイル> pick | fill <出力> <欄の id> <値>"""

import json
import sys
import time

sys.path.insert(0, sys.argv[1])
from todoku import api  # noqa: E402

p, mode = sys.argv[2], sys.argv[3]
o = {}
t = time.time()
try:
    if mode == "pick":
        k, doc, _parts = api.window(p)
        fl = api.fields(p)
        names = {}
        if k == "pdf":
            for pg in doc.get("pages", []):
                for f in pg.get("fields", []):
                    names[f["id"]] = {"name": f.get("name"), "source": f.get("source"), "rect": f.get("rect"), "page": pg["number"]}
        cands = [f for f in fl if not f.get("readonly") and f.get("kind") in ("text", "number", "date", None, "") and f.get("value") in (None, "")]
        pref = [f for f in cands if not f.get("agent")] or cands
        acro = [f for f in fl if names.get(f["id"], {}).get("source") == "acroform" and not f.get("readonly") and f.get("kind") in ("text", "number", "date", None, "")]
        if acro:
            pref = acro
        # 字の入る幅のある欄を選ぶ（1 字ずつのマス・チェックボックスの小さな欄では、値の長さで比べにならない）。幅だけで決める
        def width(f):
            r = names.get(f["id"], {}).get("rect")
            return abs(r[2] - r[0]) if r else 0

        if any(width(f) for f in pref):
            wide = [f for f in pref if width(f) >= 80]
            pref = wide or sorted(pref, key=width, reverse=True)
        if pref:
            f = pref[0]
            o = {"id": f["id"], "kind": f.get("kind"), "label": (f.get("label") or "")[:60], "agent": bool(f.get("agent")), **names.get(f["id"], {})}
        else:
            o = {"id": None}
        o["kind_doc"] = k
    else:
        out, fid, val = sys.argv[4], sys.argv[5], sys.argv[6]
        r = api.fill(p, {fid: val}, out, overwrite=True)
        o = {"proof_ok": r.get("proof_ok"), "read_back_ok": r.get("read_back_ok"), "warnings": [str(w)[:160] for w in (r.get("warnings") or [])][:5]}
except Exception as e:  # noqa: BLE001
    o["error"] = f"{type(e).__name__}: {str(e)[:300]}"
o["seconds"] = round(time.time() - t, 2)
print(json.dumps(o, ensure_ascii=False, default=str))
