#!/usr/bin/env python3
"""何で測ったかを残す（コミット・タグ・道具の版・評価のスクリプトのハッシュ）。使い方: provenance.py --out 出力 --tag TAG --commit SHA --baseline-python PY --todoku-python PY [--pkr P] [--dry-run]
同じ出力に続きから動かすときは、コミットが同じことを確かめる（違えば止める）。"""

import argparse
import datetime as dt
import hashlib
import json
import platform
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent


def ver(cmd: list, lines: int = 1) -> str:
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        return "\n".join((r.stdout + r.stderr).strip().splitlines()[:lines])
    except (OSError, subprocess.SubprocessError) as e:
        return f"なし（{type(e).__name__}）"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("--tag", required=True)
    ap.add_argument("--commit", required=True)
    ap.add_argument("--baseline-python", required=True)
    ap.add_argument("--todoku-python", required=True)
    ap.add_argument("--pkr", default="")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    p = Path(a.out) / "provenance.json"
    if p.exists():
        old = json.loads(p.read_text(encoding="utf-8"))
        if old.get("commit") != a.commit:
            print(f"この出力は別のコミット（{old.get('commit')}）で始めたものです。別の出力を指定してください", file=sys.stderr)
            return 2
        return 0
    scripts = {f.name: hashlib.sha256(f.read_bytes()).hexdigest()[:16] for f in sorted(HERE.iterdir()) if f.is_file()}
    freeze = subprocess.run(["uv", "pip", "freeze", "--python", a.baseline_python], capture_output=True, text=True).stdout.splitlines()
    want = ("markitdown", "openpyxl", "xlrd", "pdfplumber", "pypdf", "python-docx", "pdfminer", "mammoth", "pillow", "reportlab")
    d = {"tag": a.tag, "commit": a.commit, "dry_run": a.dry_run, "started": dt.datetime.now().astimezone().isoformat(timespec="seconds"),
         "platform": platform.platform(), "todoku_python": ver([a.todoku_python, "--version"]), "baselines": [x for x in freeze if x.lower().startswith(want)],
         "poppler": ver(["pdftotext", "-v"]), "libreoffice": ver(["soffice", "--version"]),
         "excel": ver(["defaults", "read", "/Applications/Microsoft Excel.app/Contents/Info", "CFBundleShortVersionString"]),
         "claude": ver(["claude", "--version"]), "pdfkit_tool": a.pkr or None, "eval_scripts_sha256": scripts}
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps({k: d[k] for k in ("tag", "commit", "poppler", "libreoffice", "excel", "claude")}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
