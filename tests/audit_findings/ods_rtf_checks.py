"""監査用：.ods の入力規則（整数 1〜10）・ありもしないグラフの行、.rtf の書き込み。

  PYTHONPATH=src .venv/bin/python tests/audit_findings/ods_rtf_checks.py フォルダ
  （フォルダに xl_all.ods・plain.ods・word_all.rtf：make_xlsx.py / make_docx.py の結果を LibreOffice で変換したもの）
"""
import sys
from pathlib import Path

from todoku import api

d = Path(sys.argv[1])
for v in (5, 11, 0):
    try:
        r = api.fill(str(d / "xl_all.ods"), {"申請!B4": v}, str(d / f"ods_b4_{v}.ods"))
        print(f"[ods 整数 1〜10 に {v}] 書いた proof={r['proof_ok']}")
    except Exception as e:  # noqa: BLE001
        print(f"[ods 整数 1〜10 に {v}] 断った: {e}")
for name in ("plain.ods", "xl_all.ods"):
    _k, _d, parts = api.window(str(d / name))
    t = "\n".join(p[1] for p in parts)
    print(f"[{name}] グラフの行:", [ln for ln in t.splitlines() if "グラフ" in ln][:2])
fl = api.fields(str(d / "word_all.rtf"))
print("[rtf] 欄:", [(f["id"], f.get("label"), f["kind"]) for f in fl])
vals = {f["id"]: ("12" if f["kind"] == "text" else True) for f in fl if f["kind"] in ("text", "checkbox")}
r = api.fill(str(d / "word_all.rtf"), vals, str(d / "rtf_filled.rtf"))
print("[rtf] proof", r["proof_ok"], "readback", r["read_back_ok"], r["read_back_differences"][:3])
print(r["say_to_user"])
