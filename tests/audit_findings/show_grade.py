"""段階（読む・書く）と、検査（audit）で通らなかった所・注記を出す。

使い方: PYTHONPATH=src python show_grade.py FILE
"""

from __future__ import annotations

import json
import sys

from todoku import api


def main(path):
    g = api.grade(path)
    print("grade:", json.dumps(g, ensure_ascii=False)[:1500])
    a = api.audit(path)
    for c in a.get("checks", a if isinstance(a, list) else []):
        if isinstance(c, dict):
            print(("OK  " if c.get("ok") else "NG  ") + str(c.get("name") or c.get("title")), "-", str(c.get("detail"))[:300])
            for w in (c.get("where") or c.get("items") or [])[:5]:
                print("      ", w)
    if not isinstance(a, list) and "checks" not in a:
        print(json.dumps(a, ensure_ascii=False)[:3000])


if __name__ == "__main__":
    main(sys.argv[1])
