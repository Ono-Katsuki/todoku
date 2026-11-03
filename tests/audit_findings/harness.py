"""監査用：書類を開き、欄・書き込み・証明・読み戻し・検査を一度に見る（todoku の api だけを使う）。

使い方（リポジトリの根で）：
  PYTHONPATH=src .venv/bin/python tests/audit_findings/harness.py 書類 ['{"欄": 値}' 出力] [--audit] [--nowin]
"""
import json
import sys
import time

from todoku import api


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    path = args[0]
    values = (json.loads(open(args[1][1:]).read()) if args[1].startswith("@") else json.loads(args[1])) if len(args) > 1 else None
    out = args[2] if len(args) > 2 else None
    s = api.summary(path)
    print("SUMMARY", json.dumps({k: s.get(k) for k in ("kind", "kind_ja", "fields", "writable_fields", "window_parts")}, ensure_ascii=False))
    if "--nowin" not in sys.argv:
        _k, _d, parts = api.window(path)
        print("WINDOW\n" + "\n".join(p[1] for p in parts)[:8000])
    for f in api.fields(path):
        print("FIELD", json.dumps({k: f.get(k) for k in ("id", "label", "kind", "options", "value", "readonly")}, ensure_ascii=False))
    if values is not None and out:
        try:
            r = api.fill(path, values, out)
        except Exception as e:  # noqa: BLE001
            print("FILL ERROR", type(e).__name__, e)
            r = None
        if r:
            print("FILL", json.dumps({k: r.get(k) for k in ("proof_ok", "proof", "read_back_ok", "read_back_differences", "warnings",
                                                         "unexpected", "structure_errors", "recalculated", "written_as")},
                                     ensure_ascii=False, default=str))
            print("SAY\n" + str(r.get("say_to_user")))
            for f in api.fields(out):
                print("AFTER", json.dumps({k: f.get(k) for k in ("id", "kind", "value")}, ensure_ascii=False))
    if "--audit" in sys.argv:
        t = time.time()
        au = api.audit(path)
        print("AUDIT", au["ok"], round(time.time() - t, 2))
        for c in au["checks"]:
            print("  ", c["ok"], c["name"], str(c["detail"])[:400])
        g = au.get("grade")
        if g:
            print("GRADE", g["headline"])


main()
