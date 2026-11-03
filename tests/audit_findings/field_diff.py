"""PDF の窓の入力欄（id・見出し・矩形）を JSON に出す。2つの版の出力を比べるのに使う。
使い方: PYTHONPATH=src python field_diff.py DIR OUT.json   /   python field_diff.py --diff A.json B.json"""

import json
import sys
from pathlib import Path

if sys.argv[1] == "--diff":
    a, b = (json.loads(Path(x).read_text()) for x in sys.argv[2:4])
    for k in sorted(set(a) | set(b)):
        fa = {(f[1], tuple(f[2])) for f in a.get(k, [])}
        fb = {(f[1], tuple(f[2])) for f in b.get(k, [])}
        gone, new = fa - fb, fb - fa
        if gone or new:
            print(f"{k}: {len(a.get(k, []))} → {len(b.get(k, []))}")
            for x in sorted(gone)[:8]:
                print("   -", x)
            for x in sorted(new)[:8]:
                print("   +", x)
    sys.exit(0)

from todoku.pdf.doc import PdfDocument  # noqa: E402
from todoku.pdf.window import build  # noqa: E402

out = {}
root = Path(sys.argv[1])
for p in sorted(root.rglob("*.pdf")):
    w = build(PdfDocument(p.read_bytes(), p.name))
    out[str(p.relative_to(root))] = [(f["id"], f["label"][:30], f["rect"]) for pg in w["pages"] for f in pg["fields"]]
Path(sys.argv[2]).write_text(json.dumps(out, ensure_ascii=False))
