"""CID（文字の番号）から Unicode への対応表を、Adobe の公開資料（BSD-3-Clause）から作る。

PDFのフォントには、文字の番号と Unicode の対応（ToUnicode）を持たないものがある。そのとき、
フォントの文字集合（Adobe-Japan1 など）の公式の対応表で文字を決める。
入手先: https://github.com/adobe-type-tools/mapping-resources-pdf （pdf2unicode）
出力: src/todoku/pdf/data/cid2unicode.json.gz と LICENSE-adobe-mapping-resources.txt
"""

from __future__ import annotations

import gzip
import json
import re
import sys
import time
import urllib.request
from pathlib import Path

BASE = "https://raw.githubusercontent.com/adobe-type-tools/mapping-resources-pdf/master/"
NAMES = ["Adobe-Japan1-UCS2", "Adobe-GB1-UCS2", "Adobe-CNS1-UCS2", "Adobe-Korea1-UCS2", "Adobe-KR-UCS2"]
OUT = Path(__file__).resolve().parent.parent / "src" / "todoku" / "pdf" / "data"
UA = "todoku-research/0.1 (accessibility research on public forms; polite crawler)"


def fetch(path: str) -> bytes:
    req = urllib.request.Request(BASE + path, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read()


def hexs(s: str) -> str:
    b = bytes.fromhex(s)
    return b.decode("utf-16-be", errors="surrogatepass")


def parse(text: str) -> dict[int, str]:
    out: dict[int, str] = {}
    for block in re.findall(r"beginbfchar(.*?)endbfchar", text, re.S):
        for src, dst in re.findall(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", block):
            out[int(src, 16)] = hexs(dst)
    for block in re.findall(r"beginbfrange(.*?)endbfrange", text, re.S):
        for m in re.finditer(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(<[0-9A-Fa-f]+>|\[[^\]]*\])", block):
            a, b, dst = int(m[1], 16), int(m[2], 16), m[3]
            if dst.startswith("["):
                for i, d in enumerate(re.findall(r"<([0-9A-Fa-f]+)>", dst)):
                    out[a + i] = hexs(d)
            else:
                raw = bytes.fromhex(dst[1:-1])
                base = int.from_bytes(raw, "big")
                for i in range(b - a + 1):
                    out[a + i] = (base + i).to_bytes(len(raw), "big").decode("utf-16-be", errors="surrogatepass")
    return out


def runs(m: dict[int, str]) -> list[list]:
    """連続する番号が連続する1文字に対応するところを、まとめて [始まりのCID, 個数, 始まりの文字コード] にする。
    複数の文字に対応するものは [CID, 0, 文字列]。"""
    out: list[list] = []
    for cid in sorted(m):
        s = m[cid]
        if len(s) == 1:
            u = ord(s)
            if out and out[-1][1] and out[-1][0] + out[-1][1] == cid and out[-1][2] + out[-1][1] == u:
                out[-1][1] += 1
                continue
            out.append([cid, 1, u])
        else:
            out.append([cid, 0, s])
    return out


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    data = {}
    for n in NAMES:
        raw = fetch("pdf2unicode/" + n).decode("latin-1")
        m = parse(raw)
        ros = n.replace("-UCS2", "")
        data[ros] = runs(m)
        print(f"{n}: {len(m):,} 文字（{len(data[ros]):,} まとまり）")
        time.sleep(2)
    (OUT / "cid2unicode.json.gz").write_bytes(gzip.compress(json.dumps(data, ensure_ascii=False, separators=(",", ":")).encode("utf-8"), 9, mtime=0))
    (OUT / "LICENSE-adobe-mapping-resources.txt").write_bytes(fetch("LICENSE.txt"))
    print("書き出し:", OUT / "cid2unicode.json.gz", (OUT / "cid2unicode.json.gz").stat().st_size, "バイト")
    return 0


if __name__ == "__main__":
    sys.exit(main())
