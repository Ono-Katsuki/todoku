"""窓（Markdown）をそのまま出す。字の符号も見たいときは --codes（行ごとに U+XXXX を並べる）。

使い方: PYTHONPATH=src python dump_window.py FILE [--codes] [--grep 字]
"""

from __future__ import annotations

import sys

from todoku import api


def main(argv):
    path = argv[0]
    codes = "--codes" in argv
    pat = argv[argv.index("--grep") + 1] if "--grep" in argv else None
    k, doc, parts = api.window(path)
    for _t, md in parts:
        for ln in md.splitlines():
            if pat and pat not in ln:
                continue
            print(ln)
            if codes:
                print("   ", " ".join(f"{ord(c):04X}" for c in ln))


if __name__ == "__main__":
    main(sys.argv[1:])
