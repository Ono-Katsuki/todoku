"""src/todoku/chars.py の、Python の unicodedata にない Unicode の資料の表（かっこの組・文の終わりの字・字の文字の種類）を
作り直す。

unicodedata にある性質（一般カテゴリ・東アジアの字の幅 UAX #11・双方向の種類・字の名前）は、chars.py が動いている Python の
unicodedata から読む（表を持たない）。表に持つのは、unicodedata にない資料だけ：
- https://www.unicode.org/Public/<版>/ucd/BidiBrackets.txt  （Bidi_Paired_Bracket・Bidi_Paired_Bracket_Type）
- https://www.unicode.org/Public/<版>/ucd/PropList.txt      （Sentence_Terminal・Terminal_Punctuation）
- https://www.unicode.org/Public/<版>/ucd/Scripts.txt       （Script：字の文字の種類。UAX #24）
表の版（TABLES_UCD_VERSION）は、Python の版と結びつけない。新しい版の表は、古い版の Python でも使える（その Python が
まだ知らない字にも、資料の値を返すだけ）。

  python3 scripts/make_chars.py                    # 版 DEFAULT_VERSION の資料を取ってきて書き直す
  python3 scripts/make_chars.py --version 17.0.0   # 版を指定する
  python3 scripts/make_chars.py --check            # chars.py の表が、その版の資料と同じかを確かめるだけ
  python3 scripts/make_chars.py --dir <フォルダ>    # 取ってこずに、フォルダの中の資料を読む
"""

from __future__ import annotations

import argparse
import re
import sys
import urllib.request
from pathlib import Path

DEFAULT_VERSION = "17.0.0"
OUT = Path(__file__).resolve().parent.parent / "src" / "todoku" / "chars.py"
UA = "todoku-research/0.1 (accessibility research on public forms)"
BEGIN = "# ---- ここから scripts/make_chars.py が作る（手で直さない）"
END = "# ---- ここまで scripts/make_chars.py が作る"


def fetch(version: str, name: str, folder: str | None) -> str:
    if folder:
        return (Path(folder) / name).read_text(encoding="utf-8")
    req = urllib.request.Request(f"https://www.unicode.org/Public/{version}/ucd/{name}", headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read().decode("utf-8")


def _rows(text: str):
    for line in text.splitlines():
        line = line.split("#", 1)[0].strip()
        if line:
            yield [x.strip() for x in line.split(";")]


def brackets(text: str) -> list[tuple[int, int]]:
    """BidiBrackets.txt の、開きかっこ（o）と対になる閉じかっこ。"""
    return [(int(cp, 16), int(pair, 16)) for cp, pair, typ in _rows(text) if typ == "o"]


def prop_ranges(text: str, prop: str) -> list[tuple[int, int]]:
    out = []
    for cps, name in _rows(text):
        if name == prop:
            a, _, b = cps.partition("..")
            out.append((int(a, 16), int(b or a, 16)))
    return _join(sorted(out))


def _join(rr: list[tuple[int, int]]) -> list[tuple[int, int]]:
    out: list[list[int]] = []
    for a, b in rr:
        if out and out[-1][1] + 1 == a:
            out[-1][1] = b
        else:
            out.append([a, b])
    return [(a, b) for a, b in out]


def script_ranges(text: str) -> tuple[list[str], list[tuple[int, int, int]]]:
    """Scripts.txt の範囲を、隣り合う同じ種類どうしでつないだもの：(種類の名前の一覧, [(始め, 終わり, 名前の番号)])。
    載っていない符号位置は Unknown（表に入れない）。"""
    rr = []
    for cps, name in _rows(text):
        a, _, b = cps.partition("..")
        rr.append((int(a, 16), int(b or a, 16), name))
    rr.sort()
    merged: list[list] = []
    for a, b, n in rr:
        if merged and merged[-1][2] == n and merged[-1][1] + 1 == a:
            merged[-1][1] = b
        else:
            merged.append([a, b, n])
    names = sorted({n for _a, _b, n in merged})
    return names, [(a, b, names.index(n)) for a, b, n in merged]


def _packed(lines: list[str], name: str, items: list[str], comment: str | None) -> None:
    if comment:
        lines.append(comment)
    lines.append(f"{name}: str = (")
    row = ""
    for piece in items:
        piece += " "
        if len(row) + len(piece) > 112:
            lines.append(f'    "{row}"')
            row = ""
        row += piece
    lines.append(f'    "{row.rstrip()}"')
    lines.append(")")


def render(version, pairs, terms, tpunct, scripts) -> str:
    names, table = scripts
    lines = [BEGIN, f'TABLES_UCD_VERSION = "{version}"  # 下の表を作った Unicode 文字データベースの版', ""]
    _packed(lines, "BRACKET_TABLE", [f"{o:X}.{c:X}" for o, c in pairs],
            "# BidiBrackets.txt：開きかっこ.閉じかっこ（Bidi_Paired_Bracket_Type = Open。16 進）")
    lines.append("")
    _packed(lines, "SENTENCE_TERMINAL_TABLE", [f"{a:X}.{b:X}" for a, b in terms],
            "# PropList.txt：Sentence_Terminal の範囲（始め.終わり。両端を含む。16 進）")
    lines.append("")
    _packed(lines, "TERMINAL_PUNCTUATION_TABLE", [f"{a:X}.{b:X}" for a, b in tpunct],
            "# PropList.txt：Terminal_Punctuation の範囲（文・句の区切りの字。読点・コロン・セミコロンも含む）")
    lines.append("")
    lines.append("# Scripts.txt：字の文字の種類（Script）。SCRIPT_NAMES は種類の名前、SCRIPT_TABLE は「始め.終わり.名前の番号」"
                 "（16 進。隣り合う同じ種類はつないだ。表にない符号位置は Unknown）")
    lines.append("SCRIPT_NAMES: tuple[str, ...] = (")
    row = "   "
    for n in names:
        piece = f' "{n}",'
        if len(row) + len(piece) > 118:
            lines.append(row)
            row = "   "
        row += piece
    lines.append(row)
    lines.append(")")
    _packed(lines, "SCRIPT_TABLE", [f"{a:X}.{b:X}.{i:X}" for a, b, i in table], None)
    lines.append(END)
    return "\n".join(lines)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--version", default=DEFAULT_VERSION)
    ap.add_argument("--dir")
    ap.add_argument("--check", action="store_true")
    a = ap.parse_args()
    props = fetch(a.version, "PropList.txt", a.dir)
    block = render(a.version, brackets(fetch(a.version, "BidiBrackets.txt", a.dir)),
                   prop_ranges(props, "Sentence_Terminal"), prop_ranges(props, "Terminal_Punctuation"),
                   script_ranges(fetch(a.version, "Scripts.txt", a.dir)))
    src = OUT.read_text(encoding="utf-8")
    new = re.sub(re.escape(BEGIN) + r".*?" + re.escape(END), lambda _m: block, src, flags=re.S)
    if a.check:
        print("同じ" if new == src else "違う（python3 scripts/make_chars.py で作り直す）")
        sys.exit(0 if new == src else 1)
    OUT.write_text(new, encoding="utf-8")
    print(f"書き直しました（UCD {a.version}）")


if __name__ == "__main__":
    main()
