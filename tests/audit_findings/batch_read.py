"""書類ごとに、窓の本文の行（符号つき）・読む段階・検査で通らなかった所を出す。

使い方: PYTHONPATH=src python batch_read.py FILE [FILE ...]
"""

from __future__ import annotations

import sys

from todoku import api

SKIP = ("- 元ファイル", "- この文書", "- 欄として", "- 記号", "- フォームの欄はありません", "# ")


def body(parts) -> list[str]:
    out = []
    for _t, md in parts:
        for ln in md.splitlines():
            if ln.strip() and not ln.startswith(SKIP):
                out.append(ln)
    return out


def main(paths):
    for p in paths:
        print("=" * 20, p.rsplit("/", 1)[-1])
        try:
            k, doc, parts = api.window(p)
        except Exception as e:  # noqa: BLE001
            print("  window error:", type(e).__name__, e)
            continue
        for ln in body(parts)[:40]:
            print("  ", ln)
            if any(ord(c) > 0x7F for c in ln):
                print("     ", " ".join(f"{ord(c):04X}" for c in ln if ord(c) > 0x20)[:400])
        try:
            g = api.grade(p)
            print("  READ", g.get("grade"), "WRITE", (g.get("write") or {}).get("grade"),
                  [r.get("title") for r in g.get("reasons", [])], "alerts", [a.get("title") for a in g.get("alerts") or []])
            a = api.audit(p)
            for c in a.get("checks", []):
                if not c.get("ok"):
                    print("  AUDIT NG", c.get("name") or c.get("title"), str(c.get("detail"))[:200], (c.get("where") or [])[:3])
        except Exception as e:  # noqa: BLE001
            print("  grade/audit error:", type(e).__name__, e)


if __name__ == "__main__":
    main(sys.argv[1:])
