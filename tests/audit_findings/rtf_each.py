"""監査用：.rtf（LibreOffice が Word の様式から書き出したもの）の欄に1つずつ書く。

  PYTHONPATH=src .venv/bin/python tests/audit_findings/rtf_each.py フォルダ/word_all.rtf
"""
import sys

from todoku import api

src = sys.argv[1]
_k, _d, parts = api.window(src)
print("\n".join(ln for ln in parts[0][1].splitlines() if ln.startswith("- ")))
for f in api.fields(src):
    if f["kind"] not in ("text", "checkbox", "dropdown", "inline"):
        continue
    v = True if f["kind"] == "checkbox" else ("女" if f["kind"] == "dropdown" else "12")
    try:
        r = api.fill(src, {f["id"]: v}, src + f".{f['id']}.rtf")
        print(f"[{f['id']} {f['kind']}] proof={r['proof_ok']} readback={r['read_back_ok']}")
    except Exception as e:  # noqa: BLE001
        print(f"[{f['id']} {f['kind']}] 断った: {e}")
