"""tests/forms.json に記録した公開様式を corpus/forms/ に取得し、SHA-256を確かめる。

書類はリポジトリに含めない（再配布できるか分からないため）。robots.txt を守り、間隔を空けて取得する。
使い方: uv run python scripts/fetch_forms.py
"""

from __future__ import annotations

import hashlib
import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from fetch_sample import allowed, get  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent


def main() -> int:
    manifest = json.loads((ROOT / "tests" / "forms.json").read_text(encoding="utf-8"))
    out = ROOT / "corpus" / "forms"
    out.mkdir(parents=True, exist_ok=True)
    bad = 0
    for f in manifest["forms"]:
        dst = out / f"{f['id']}.xlsx"
        if dst.exists() and hashlib.sha256(dst.read_bytes()).hexdigest() == f["sha256"]:
            print("取得済み", f["id"])
            continue
        if not allowed(f["url"]):
            print("robots.txt で禁止", f["url"])
            bad += 1
            continue
        code, body = get(f["url"], timeout=60)
        if code != 200:
            print(f"取得できません（HTTP {code}）", f["url"])
            bad += 1
            continue
        h = hashlib.sha256(body).hexdigest()
        dst.write_bytes(body)
        if h != f["sha256"]:
            print(f"注意: {f['id']} の内容が記録時と違います（様式が差し替えられた可能性）")
            bad += 1
        else:
            print("取得", f["id"])
        time.sleep(2)
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
