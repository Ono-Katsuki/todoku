"""PDF の /ToUnicode を、同じ長さの別の名前（/XoUnicode）に変えて、ToUnicode のない PDF を作る（xref はそのまま正しい）。

フォントの辞書が圧縮されていない PDF（Chrome・Quartz・LibreOffice の出力）で使える。
使い方: python strip_tounicode.py IN.pdf OUT.pdf [フォント名の一部 ...]  （名前を渡すと、そのフォントだけ外す）
"""

from __future__ import annotations

import re
import sys
from pathlib import Path


def strip(blob: bytes, only: list[str]) -> tuple[bytes, int]:
    if not only:
        n = blob.count(b"/ToUnicode")
        return blob.replace(b"/ToUnicode", b"/XoUnicode"), n
    out = bytearray(blob)
    n = 0
    # obj ... endobj ごとに、BaseFont が名前に合うものだけ
    for m in re.finditer(rb"\d+ 0 obj(.*?)endobj", blob, re.S):
        body = m.group(1)
        bf = re.search(rb"/BaseFont\s*/([^\s/<>\[\]]+)", body)
        if bf and any(o.encode() in bf.group(1) for o in only) and b"/ToUnicode" in body:
            i = m.start(1) + body.index(b"/ToUnicode")
            out[i:i + 10] = b"/XoUnicode"
            n += 1
    return bytes(out), n


if __name__ == "__main__":
    src, dst, *only = sys.argv[1:]
    data, n = strip(Path(src).read_bytes(), only)
    Path(dst).write_bytes(data)
    print(f"{n} ToUnicode removed -> {dst}")
