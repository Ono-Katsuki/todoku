"""書き込みで使う CMap（Unicode → CID）の対応表を、Adobe の公開資料（cmap-resources、BSD-3-Clause）から作る。

とどくは、日本語・中国語・韓国語の字を、埋め込まない CID フォント（HeiseiKakuGo-W5 など）と、定義済みの CMap
（UniJIS-UCS2-HW-H など）で書く。書く字が、その CMap でどの CID になるかが分かれば、宣言した字形の集まりの補遺
（Supplement）の範囲に入るか・幅はいくつかを、書く前と証明で確かめられる（範囲の外の CID は、ビューアーの字形にない
ことがあり、白く抜ける）。

入手先: https://github.com/adobe-type-tools/cmap-resources （<集まり>/CMap/<名前>）。手元の写し（TeX Live の
fonts/cmap/adobemapping/cmap-resources など）のフォルダを引数に渡すと、それを読む。
ライセンス: adobe-type-tools/cmap-resources の BSD 3-Clause（src/todoku/pdf/data/LICENSE-adobe-cmap-resources.txt）
出力: src/todoku/pdf/data/uni2cid.json.gz（{CMap の名前: [[始まりの Unicode, 個数, 始まりの CID], …]}）
"""

from __future__ import annotations

import gzip
import json
import re
import sys
import urllib.request
from pathlib import Path

BASE = "https://raw.githubusercontent.com/adobe-type-tools/cmap-resources/master/"
CMAPS = {"UniJIS-UCS2-HW-H": "Adobe-Japan1-7", "UniJIS-UCS2-H": "Adobe-Japan1-7", "UniKS-UCS2-H": "Adobe-Korea1-2",
         "UniGB-UCS2-H": "Adobe-GB1-5", "UniCNS-UCS2-H": "Adobe-CNS1-7"}
OUT = Path(__file__).resolve().parent.parent / "src" / "todoku" / "pdf" / "data"
UA = "todoku-research/0.1 (accessibility research on public forms; polite crawler)"


def read(name: str, local: Path | None) -> str:
    rel = f"{CMAPS[name]}/CMap/{name}"
    if local is not None:
        return (local / rel).read_text(encoding="latin-1")
    req = urllib.request.Request(BASE + rel, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read().decode("latin-1")


def parse(name: str, local: Path | None) -> dict[int, int]:
    text = read(name, local)
    out: dict[int, int] = {}
    use = re.search(r"/(\S+)\s+usecmap", text)
    if use:
        out.update(parse(use[1], local))
    for blk in re.findall(r"begincidrange(.*?)endcidrange", text, re.S):
        for a, b, c in re.findall(r"<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*(\d+)", blk):
            for k in range(int(a, 16), int(b, 16) + 1):
                out[k] = int(c) + k - int(a, 16)
    for blk in re.findall(r"begincidchar(.*?)endcidchar", text, re.S):
        for a, c in re.findall(r"<([0-9a-fA-F]+)>\s*(\d+)", blk):
            out[int(a, 16)] = int(c)
    return out


def runs(m: dict[int, int]) -> list[list[int]]:
    out: list[list[int]] = []
    for u in sorted(m):
        c = m[u]
        if out and out[-1][0] + out[-1][1] == u and out[-1][2] + out[-1][1] == c:
            out[-1][1] += 1
        else:
            out.append([u, 1, c])
    return out


def main(argv: list[str]) -> int:
    local = Path(argv[0]) if argv else None
    data = {}
    for name in CMAPS:
        if name == "UniJIS-UCS2-H":
            continue  # HW-H が使う（usecmap）ので、HW-H に含めて書き出す
        m = parse(name, local)
        data[name] = runs(m)
        print(f"{name}: {len(m):,} 字（{len(data[name]):,} まとまり）")
    p = OUT / "uni2cid.json.gz"
    p.write_bytes(gzip.compress(json.dumps(data, separators=(",", ":")).encode(), 9, mtime=0))
    print("書き出し:", p, p.stat().st_size, "バイト")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
