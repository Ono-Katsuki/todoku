"""とどくの処理の取り出し（別の処理で 1 ブックずつ）。使い方（python -I）: todoku_process.py <src> <ファイル> <出力 .py>"""

import json
import sys
import time

sys.path.insert(0, sys.argv[1])
from todoku import api  # noqa: E402

p, out_py = sys.argv[2], sys.argv[3]
o = {}
t = time.time()
try:
    r = api.extract_processing(p, format="python", out_path=out_py, overwrite=True)
    v = r.get("verification") or {}
    st = v.get("stored") or {}
    ev = v.get("evaluator") or {}
    rules = r.get("rules") or []
    statuses = {}
    for d in rules:
        s = (d.get("verification") or {}).get("status", "none")
        statuses[s] = statuses.get(s, 0) + 1
    o = {"formula_cells": (r.get("summary") or {}).get("formula_cells"), "rules": (r.get("summary") or {}).get("rules"),
         "stored_compared": st.get("compared"), "stored_match": st.get("match"), "stored_rate": st.get("rate"), "evaluator_compared": ev.get("compared"),
         "evaluator_match": ev.get("match"), "rules_proven": v.get("rules_proven"), "rules_not_proven": v.get("rules_not_proven"),
         "explanation_check": {k: (r.get("explanation_check") or {}).get(k) for k in ("rules", "checked", "unchecked", "failed", "ok")},
         "mismatch_count": v.get("mismatch_count"), "rule_status": statuses, "unknown_functions": r.get("unknown_functions") or [],
         "external_references": len(r.get("external_references") or []), "not_expressible": len(r.get("not_expressible_by_cells") or []),
         "macros": bool((r.get("macros") or {}).get("procedures")) if isinstance(r.get("macros"), dict) else bool(r.get("macros")),
         "notes": [str(x)[:200] for x in (v.get("notes") or [])][:5],
         "not_proven_why": [str(x.get("why"))[:200] for x in rules if (x.get("verification") or {}).get("status") not in (None, "proven")][:8],
         "mismatch_examples": [{k: str(m.get(k))[:80] for k in ("cell", "kind", "stored", "module", "evaluator")} for m in (v.get("mismatches") or [])[:5]]}
    # エージェントが使う入口も、同じブックで動くこと（explain_calculation）
    try:
        m = api.calc_model(p)
        o["explain_ok"] = isinstance(m, dict)
    except Exception as e:  # noqa: BLE001
        o["explain_ok"] = False
        o["explain_error"] = f"{type(e).__name__}: {str(e)[:160]}"
except Exception as e:  # noqa: BLE001
    o["error"] = f"{type(e).__name__}: {str(e)[:300]}"
o["seconds"] = round(time.time() - t, 2)
print(json.dumps(o, ensure_ascii=False, default=str))
