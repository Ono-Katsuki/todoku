"""見本の文を入れた書類を todoku の窓で読み、文ごとに、窓の字が論理順の正しい字と同じかを比べる。

使い方: PYTHONPATH=src python check_read.py FILE [FILE ...]
各行: 名前 / 結果（OK・NFC だけ違う・違う・見つからない）/ 窓の字 / 期待の字。最後に段階と、読めない字の数。
"""

from __future__ import annotations

import sys
import unicodedata
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from scripts_samples import SAMPLES  # noqa: E402

from todoku import api  # noqa: E402


def window_text(path: str) -> tuple[str, dict]:
    k, doc, parts = api.window(path)
    return "\n".join(md for _t, md in parts), doc


def find_line(text: str, key: str) -> str | None:
    for ln in text.splitlines():
        if f"{key} |" in ln or f"| {key}" in ln:
            return ln
    return None


def main(paths):
    for path in paths:
        print("=" * 30, path)
        text, doc = window_text(path)
        for k, (exp, _d) in SAMPLES.items():
            ln = find_line(text, k)
            if ln is None:
                print(f"{k:14} MISSING")
                continue
            got = ln
            if exp in got:
                st = "OK"
            elif unicodedata.normalize("NFC", exp) in unicodedata.normalize("NFC", got):
                st = "NFC-ONLY"
            elif unicodedata.normalize("NFKC", exp) in unicodedata.normalize("NFKC", got):
                st = "NFKC-ONLY"
            else:
                st = "WRONG"
            print(f"{k:14} {st:9} got={got!r}")
            if st == "WRONG":
                print(f"{'':14} {'':9} exp={exp!r}")
        try:
            g = api.grade(path)
            print("grade:", {kk: g[kk] for kk in g if kk in ("read", "write", "grade", "level")} if isinstance(g, dict) else g)
        except Exception as e:  # noqa: BLE001
            print("grade error", e)
        pages = doc.get("pages") or []
        print("unmapped:", [p.get("unmapped") for p in pages])


if __name__ == "__main__":
    main(sys.argv[1:])
