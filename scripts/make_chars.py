"""src/todoku/chars.py の、Unicode の資料から作る表（かっこの組・文の終わりの字・コロン）を作り直す。

入手先（Unicode 文字データベース。版は、この Python の unicodedata と同じにする）:
- https://www.unicode.org/Public/<版>/ucd/BidiBrackets.txt  （Bidi_Paired_Bracket・Bidi_Paired_Bracket_Type）
- https://www.unicode.org/Public/<版>/ucd/PropList.txt      （Sentence_Terminal・Terminal_Punctuation）
- https://www.unicode.org/Public/<版>/ucd/Scripts.txt       （Script：字の文字の種類。UAX #24）
- コロン：この Python の unicodedata で、一般カテゴリが Po（そのほかの句読点）で、名前に語「COLON」を含む字
  （「SEMICOLON」は含めない。U+061B ؛ ARABIC SEMICOLON はセミコロン、U+0589 ։ ARMENIAN FULL STOP は文の終わりの字で、
  どちらも見出しのコロンではない）

  python3 scripts/make_chars.py            # 取ってきて書き直す
  python3 scripts/make_chars.py --check    # chars.py の表が、いまの資料と同じかを確かめるだけ
"""

from __future__ import annotations

import re
import sys
import unicodedata
import urllib.request
from pathlib import Path

VERSION = unicodedata.unidata_version
BASE = f"https://www.unicode.org/Public/{VERSION}/ucd/"
OUT = Path(__file__).resolve().parent.parent / "src" / "todoku" / "chars.py"
UA = "todoku-research/0.1 (accessibility research on public forms)"
BEGIN = "# ---- ここから scripts/make_chars.py が作る（手で直さない）"
END = "# ---- ここまで scripts/make_chars.py が作る"


def fetch(name: str) -> str:
    req = urllib.request.Request(BASE + name, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read().decode("utf-8")


def brackets(text: str) -> list[tuple[int, int]]:
    """BidiBrackets.txt の、開きかっこ（o）と対になる閉じかっこ。"""
    out = []
    for line in text.splitlines():
        line = line.split("#", 1)[0].strip()
        if not line:
            continue
        cp, pair, typ = (x.strip() for x in line.split(";"))
        if typ == "o":
            out.append((int(cp, 16), int(pair, 16)))
    return out


def prop_ranges(text: str, prop: str) -> list[tuple[int, int]]:
    out = []
    for line in text.splitlines():
        line = line.split("#", 1)[0].strip()
        if not line:
            continue
        cps, name = (x.strip() for x in line.split(";"))
        if name != prop:
            continue
        a, _, b = cps.partition("..")
        out.append((int(a, 16), int(b or a, 16)))
    return sorted(out)


def script_ranges(text: str) -> tuple[list[str], list[tuple[int, int, int]]]:
    """Scripts.txt の範囲を、隣り合う同じ種類どうしでつないだもの：(種類の名前の一覧, [(始め, 終わり, 名前の番号)])。
    載っていない符号位置は Unknown（表に入れない）。"""
    rr = []
    for line in text.splitlines():
        line = line.split("#", 1)[0].strip()
        if not line:
            continue
        cps, name = (x.strip() for x in line.split(";"))
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


def colons() -> list[int]:
    return [c for c in range(0x110000)
            if unicodedata.category(chr(c)) == "Po" and re.search(r"\bCOLON\b", unicodedata.name(chr(c), ""))]


def _ranges(lines: list[str], name: str, rr) -> None:
    lines.append(f"{name}: tuple[tuple[int, int], ...] = (")
    for a, b in rr:
        lines.append(f"    (0x{a:04X}, 0x{b:04X}),")
    lines.append(")")


def _scripts(lines: list[str], names: list[str], rr) -> None:
    lines.append("# Scripts.txt：字の文字の種類（Script）。SCRIPT_NAMES は種類の名前、SCRIPT_TABLE は「始め.終わり.名前の番号」"
                 "（16 進）を空白で区切って並べたもの（隣り合う同じ種類はつないだ。表にない符号位置は Unknown）")
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
    lines.append("SCRIPT_TABLE: str = (")
    row = ""
    for a, b, i in rr:
        piece = f"{a:X}.{b:X}.{i:X} "
        if len(row) + len(piece) > 112:
            lines.append(f'    "{row}"')
            row = ""
        row += piece
    lines.append(f'    "{row.rstrip()}"')
    lines.append(")")


def render(pairs, terms, tpunct, cols, scripts) -> str:
    lines = [BEGIN, f'UCD_VERSION = "{VERSION}"  # unicodedata.unidata_version と同じ版で作った', "",
             "# BidiBrackets.txt：開きかっこ → 閉じかっこ（Bidi_Paired_Bracket_Type = Open）",
             "BRACKET_PAIRS: dict[str, str] = {"]
    for o, c in pairs:
        lines.append(f"    \"\\u{o:04x}\": \"\\u{c:04x}\",  # {unicodedata.name(chr(o))} / {unicodedata.name(chr(c))}")
    lines.append("}")
    lines.append("")
    lines.append("# PropList.txt：Sentence_Terminal の範囲（両端を含む）")
    _ranges(lines, "SENTENCE_TERMINAL_RANGES", terms)
    lines.append("")
    lines.append("# PropList.txt：Terminal_Punctuation の範囲（文・句の区切りの字。読点・コロン・セミコロンも含む）")
    _ranges(lines, "TERMINAL_PUNCTUATION_RANGES", tpunct)
    lines.append("")
    lines.append("# 一般カテゴリ Po で、名前に語「COLON」を含む字（SEMICOLON は含めない）")
    lines.append("LABEL_COLONS = frozenset(")
    lines.append("    \"" + "".join(f"\\u{c:04x}" if c <= 0xFFFF else f"\\U{c:08x}" for c in cols) + "\"")
    lines.append(")")
    lines.append("")
    _scripts(lines, *scripts)
    lines.append(END)
    return "\n".join(lines)


def main() -> None:
    props = fetch("PropList.txt")
    block = render(brackets(fetch("BidiBrackets.txt")), prop_ranges(props, "Sentence_Terminal"),
                   prop_ranges(props, "Terminal_Punctuation"), colons(), script_ranges(fetch("Scripts.txt")))
    src = OUT.read_text(encoding="utf-8")
    new = re.sub(re.escape(BEGIN) + r".*?" + re.escape(END), lambda _m: block, src, flags=re.S)
    if "--check" in sys.argv:
        print("同じ" if new == src else "違う（python3 scripts/make_chars.py で作り直す）")
        sys.exit(0 if new == src else 1)
    OUT.write_text(new, encoding="utf-8")
    print(f"書き直しました（UCD {VERSION}）")


if __name__ == "__main__":
    main()
