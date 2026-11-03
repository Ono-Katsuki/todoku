"""PDF のページの字（グリフ）を、描いた順に、位置・字・符号と並べて出す。

使い方: PYTHONPATH=src python dump_glyphs.py FILE.pdf [ページ番号] [--grep 字]
"""

from __future__ import annotations

import sys
from pathlib import Path

from todoku.pdf.content import page_content
from todoku.pdf.doc import PdfDocument


def main(argv):
    path = argv[0]
    n = int(argv[1]) if len(argv) > 1 and argv[1].isdigit() else 1
    doc = PdfDocument(Path(path).read_bytes(), Path(path).name)
    page = doc.pages()[n - 1]
    pc = page_content(doc, page)
    for g in pc.glyphs:
        print(f"x0={g.x0:7.2f} x1={g.x1:7.2f} y={g.oy:7.2f} size={g.size:5.1f} text={g.text!r} "
              f"{' '.join(f'U+{ord(c):04X}' for c in g.text)}")


if __name__ == "__main__":
    main(sys.argv[1:])
