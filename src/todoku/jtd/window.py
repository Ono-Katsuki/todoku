"""一太郎の窓：Word の窓と同じ形（本文・表・レイアウト枠・柱・脚注）。

書き込みは字数を変えない上書きなので、欄ごとに上書きできる字数（「N 字まで」）を示し、字のない所（空のマスなど）は
「変更できない欄」として示す（読むことはできる。書くなら、空白・伏せ字のある所を字を目印にして上書きする）。
"""

from __future__ import annotations

import struct

from .book import JtdFile
from .fill import capacity


def sheet_names(doc: JtdFile) -> list[str]:
    """シート（一太郎の文書の見出しのタブ）の名前。ObjectSheets/SheetInfo：シートの数・…・名前の字数・名前（UTF-16LE）。"""
    if not doc.cfb.has("ObjectSheets/SheetInfo"):
        return []
    d = doc.cfb.read("ObjectSheets/SheetInfo")
    try:
        n = struct.unpack_from("<I", d, 0)[0]
        if n != 1 or len(d) < 0x2E:
            return []
        cch = struct.unpack_from("<I", d, 0x2A)[0]
        return [d[0x2E:0x2E + 2 * cch].decode("utf-16-le")]
    except (struct.error, UnicodeDecodeError):
        return []


def build(doc: JtdFile) -> dict:
    from ..docx.window import Spots
    from ..docx.window import build as wbuild

    win = wbuild(doc)
    win["kind"] = "jtd"
    sp = Spots(doc)
    for f in win["fields"]:
        it = sp.items.get(f["id"])
        room = capacity(doc, it) if it is not None else 0
        if f["kind"] == "checkbox":
            if not room:
                f["readonly"] = True
            continue
        if room:
            f["maxlen"] = room
        else:
            f["readonly"] = True
    win["jtd"] = {"compressed": doc.compressed, "sheets": sheet_names(doc)}
    return win


def to_markdown(win: dict) -> str:
    from ..docx.window import to_markdown as wmd

    out = []
    for ln in wmd(win).split("\n"):
        if ln.startswith("- 元ファイルの SHA-256:"):
            ln = ln.replace("、Word", "、一太郎（.jtd）" + ("。圧縮して保存された文書（中の文書を解いて読みました）"
                                                       if win["jtd"]["compressed"] else ""), 1)
            out.append(ln)
            if win["jtd"]["sheets"]:
                out.append("- シート：" + "、".join(win["jtd"]["sheets"]))
            continue
        if ln.startswith("- この文書は読むための窓です。"):
            ln = ("- この文書は読むための窓です。書き込みは `todoku fill` で、元の字を同じ字数の字で上書きします（字数・位置・書式と、"
                  "ほかのバイトは変えません）。一太郎の形式は公開された仕様がなく、字数を変えると位置の表がずれるおそれがあるため、"
                  "空白・伏せ字（○○）の字数までしか書けません（欄の「N 字まで」）。字のない所（空のマスなど）は「変更できない欄」です")
        elif ln.startswith("- 欄として示していない所に書くときは"):
            ln = ("- 欄として示していない所は、字を目印にして上書きできます：`{\"@1\": {\"old\": \"元の字\", \"new\": \"同じ字数の新しい字\"}}`"
                  "（old と new は同じ字数にする。足りない分は空白「　」で埋める。old がいくつもあるときは、`near`（近くの欄の id）か"
                  " `occurrence`（何番目か）で1つに決める）")
        out.append(ln)
    return "\n".join(out)


def markdown_parts(win: dict, budget: int = 120_000) -> list[tuple[str, str]]:
    """Word の窓と同じ分け方：大きいときは、行の切れ目で部分に分け、目次をつける。"""
    whole = to_markdown(win)
    if len(whole) <= budget:
        return [("全体", whole)]
    lines = whole.split("\n")
    parts: list[list[str]] = []
    cur: list[str] = []
    for line in lines:
        if cur and sum(len(x) + 1 for x in cur) + len(line) > budget - 2000:
            parts.append(cur)
            cur = []
        cur.append(line)
    if cur:
        parts.append(cur)
    head = lines[:6]
    index = head + ["", "## この窓の目次", "", f"窓が大きいので、{len(parts)} の部分に分けました。", ""]
    index += [f"- 部分 {i}: {sum(len(x) + 1 for x in p):,} 文字" for i, p in enumerate(parts, 1)]
    return [("目次", "\n".join(index))] + [(f"部分 {i}", "\n".join(head[:1] + [""] + p)) for i, p in enumerate(parts, 1)]
