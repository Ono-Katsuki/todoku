"""字の組み立て（シェーピング）に使う Unicode の文字データを、Unicode の公開データ（UCD）から作る。

作るもの：src/todoku/pdf/data/shaping.json.gz
- joining：アラビア文字の結合の型（ArabicShaping.txt の明示した値。書いていない字は、一般カテゴリが Mn・Me・Cf なら T、
  ほかは U：ArabicShaping.txt の冒頭の決まり）
- indic：インド系の文字の音節の種類と位置（IndicSyllabicCategory.txt・IndicPositionalCategory.txt）
- mirror：右から左の並びで鏡に映す字（BidiMirroring.txt）
- brackets：対になる括弧（BidiBrackets.txt、UAX #9 の N0）
- scripts：文字の種類（Scripts.txt。組み立てる文字の範囲だけ）

版は、Python の unicodedata（unicodedata.unidata_version）にそろえる（双方向の種類・一般カテゴリと食い違わないように）。

使い方：python scripts/make_shaping_data.py [UCD のファイルのあるフォルダ]
フォルダを渡さなければ https://www.unicode.org/Public/<版>/ucd/ から取る。
"""

from __future__ import annotations

import gzip
import json
import re
import sys
import unicodedata
import urllib.request
from pathlib import Path

OUT = Path(__file__).resolve().parent.parent / "src" / "todoku" / "pdf" / "data" / "shaping.json.gz"
FILES = ("ArabicShaping", "IndicSyllabicCategory", "IndicPositionalCategory", "BidiMirroring", "BidiBrackets", "Scripts")
# 組み立てる文字（Scripts.txt の名前）
SCRIPTS = ("Arabic", "Hebrew", "Devanagari", "Bengali", "Gurmukhi", "Gujarati", "Oriya", "Tamil", "Telugu", "Kannada",
           "Malayalam", "Sinhala", "Thai", "Lao", "Khmer", "Myanmar", "Syriac", "Thaana", "Nko", "Inherited", "Common")


def _read(src: str | None, name: str) -> str:
    if src:
        return (Path(src) / f"{name}.txt").read_text(encoding="utf-8")
    url = f"https://www.unicode.org/Public/{unicodedata.unidata_version}/ucd/{name}.txt"
    with urllib.request.urlopen(url, timeout=60) as r:  # noqa: S310（決まった URL）
        return r.read().decode("utf-8")


def _rows(text: str):
    for line in text.splitlines():
        line = line.split("#", 1)[0].strip()
        if line:
            yield [x.strip() for x in line.split(";")]


def _range(s: str) -> tuple[int, int]:
    if ".." in s:
        a, b = s.split("..")
        return int(a, 16), int(b, 16)
    return int(s, 16), int(s, 16)


def _ranges(pairs: dict[int, str]) -> list[list]:
    """{字: 値} → [[始まり, 終わり, 値], ...]（続く同じ値をまとめる）。"""
    out: list[list] = []
    for cp in sorted(pairs):
        v = pairs[cp]
        if out and out[-1][1] == cp - 1 and out[-1][2] == v:
            out[-1][1] = cp
        else:
            out.append([cp, cp, v])
    return out


def build(src: str | None) -> dict:
    head = _read(src, "ArabicShaping").splitlines()[0]
    joining: dict[int, str] = {}
    for r in _rows(_read(src, "ArabicShaping")):
        joining[int(r[0], 16)] = r[2]
    indic_cat: dict[int, str] = {}
    for r in _rows(_read(src, "IndicSyllabicCategory")):
        a, b = _range(r[0])
        for cp in range(a, b + 1):
            indic_cat[cp] = r[1]
    indic_pos: dict[int, str] = {}
    for r in _rows(_read(src, "IndicPositionalCategory")):
        a, b = _range(r[0])
        for cp in range(a, b + 1):
            indic_pos[cp] = r[1]
    mirror = {int(r[0], 16): int(r[1], 16) for r in _rows(_read(src, "BidiMirroring"))}
    brackets = {int(r[0], 16): [int(r[1], 16), r[2]] for r in _rows(_read(src, "BidiBrackets"))}
    scripts: dict[int, str] = {}
    for r in _rows(_read(src, "Scripts")):
        if r[1] in SCRIPTS:
            a, b = _range(r[0])
            for cp in range(a, b + 1):
                scripts[cp] = r[1]
    return {
        "version": unicodedata.unidata_version,
        "source": f"Unicode Character Database {unicodedata.unidata_version} ({head.strip('# ').strip()} ほか)",
        "joining": _ranges(joining),
        "indic_cat": _ranges(indic_cat),
        "indic_pos": _ranges(indic_pos),
        "mirror": sorted([a, b] for a, b in mirror.items()),
        "brackets": sorted([a, b, t] for a, (b, t) in brackets.items()),
        "scripts": _ranges(scripts),
    }


def main() -> None:
    src = sys.argv[1] if len(sys.argv) > 1 else None
    data = build(src)
    if not re.match(r"\d+\.\d+\.\d+$", data["version"]):
        raise SystemExit("版が読めません")
    OUT.write_bytes(gzip.compress(json.dumps(data, ensure_ascii=False, separators=(",", ":")).encode("utf-8"), mtime=0))
    print(OUT, OUT.stat().st_size, "bytes", data["source"])


if __name__ == "__main__":
    main()
