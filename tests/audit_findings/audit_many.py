"""PDF をいくつも audit して、NG の検査だけを出す。使い方: PYTHONPATH=src python audit_many.py FILE_OR_DIR ..."""

import sys
from pathlib import Path

from todoku.pdf.audit import run

paths = []
for a in sys.argv[1:]:
    p = Path(a)
    paths += sorted(p.rglob("*.pdf")) if p.is_dir() else [p]
for p in paths:
    try:
        au = run(p.read_bytes(), p.name)
    except Exception as e:  # noqa: BLE001
        print(p.name, "ERROR", repr(e)[:200])
        continue
    bad = [(c.name, c.items[:2]) for c in au.checks if not c.ok]
    print(p.parent.name + "/" + p.name, "OK" if not bad else bad)
