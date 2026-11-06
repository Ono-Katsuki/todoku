"""とどくで書く（比べのため）：pick は書く欄を 1 つ選ぶ（比べる道具にも同じ欄を渡す：道具が欄を見つける力は、ここでは問わない）。fill は書く。
使い方（python -I）: todoku_write.py <src> <ファイル> pick | fill <出力> <欄の id> <値>"""

import json
import re
import sys
import time

sys.dont_write_bytecode = True  # 道具の版のフォルダを変えない
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
        skip = set(json.loads(sys.argv[4])) if len(sys.argv) > 4 else set()

        def width(f):
            r = names.get(f["id"], {}).get("rect")
            return abs(r[2] - r[0]) if r else 0

        def dark(f):
            """黒い字が読めない塗りのマス（事実の fill。#RRGGBB の明るさが半分未満）。"""
            m = re.fullmatch(r"#([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})", str((f.get("facts") or {}).get("fill") or ""))
            return bool(m) and 0.299 * int(m[1], 16) + 0.587 * int(m[2], 16) + 0.114 * int(m[3], 16) < 128

        usable = [f for f in fl if not f.get("readonly") and f["id"] not in skip and not dark(f)]
        plain = ("text", "number", "date", None, "")
        cands = [f for f in usable if f.get("kind") in plain and f.get("value") in (None, "")]
        pref = [f for f in cands if not f.get("agent")] or cands
        acro = [f for f in usable if names.get(f["id"], {}).get("source") == "acroform" and f.get("kind") in plain]
        if acro:
            pref = acro
        # 字の入る幅のある欄を選ぶ（1 字ずつのマス・チェックボックスの小さな欄では、値の長さで比べにならない）。幅だけで決める
        if any(width(f) for f in pref):
            wide = [f for f in pref if width(f) >= 80]
            pref = wide or sorted(pref, key=width, reverse=True)
        # v2：文字の中の空欄（inline）も書ける所。前の定義の所がないときだけ選ぶ。規則で決めた所を先に、広い順に
        inl = sorted((f for f in usable if f.get("kind") == "inline"), key=lambda f: (bool(f.get("agent")), width(f) < 80, -width(f)))
        order = pref + inl
        if order:
            f = order[0]
            o = {"id": f["id"], "kind": f.get("kind"), "label": (f.get("label") or "")[:60], "agent": bool(f.get("agent")), "inline": f.get("kind") == "inline", **names.get(f["id"], {})}
        else:
            o = {"id": None}
        o["kind_doc"] = k
    else:
        out, fid, val = sys.argv[4], sys.argv[5], sys.argv[6]
        # 文字の中の空欄の並び（マスの全体を読み戻す Excel など）：空欄ごとの値のリストで書く（道具の説明のとおり）。最初の空欄に書く
        f = next((x for x in api.fields(p) if x["id"] == fid), {})
        if f.get("kind") == "inline" and "!" in fid:
            n = max(1, (f.get("label") or "").count("［"))
            val = [val] + [None] * (n - 1)
        r = api.fill(p, {fid: val}, out, overwrite=True)
        o = {"proof_ok": r.get("proof_ok"), "read_back_ok": r.get("read_back_ok"), "warnings": [str(w)[:160] for w in (r.get("warnings") or [])][:5]}
except Exception as e:  # noqa: BLE001
    o["error"] = f"{type(e).__name__}: {str(e)[:300]}"
o["seconds"] = round(time.time() - t, 2)
print(json.dumps(o, ensure_ascii=False, default=str))
