"""一太郎の窓：Word の窓と同じ形（本文・表・レイアウト枠・柱・脚注）。

本文の欄は、字数を変える書き込み（空のマスに足す・空白や伏せ字より長い値に替える）ができるかを文書の構造で確かめ、
できる欄は字数の制限なしで示す。できない欄（枠・柱の中、構造の意味を確かめられない文書など）は、上書きできる字数
（「N 字まで」）を示し、上書きできる字もない所は「変更できない欄」として示す（f["why_fixed"] に理由）。
"""

from __future__ import annotations

import struct

from .book import JtdFile
from .fill import capacity, growable


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
        why = growable(doc, it) if it is not None else "欄が見つかりません"
        if why is None:
            continue  # 字数を変える書き込みができる（字数の制限なし）
        f["why_fixed"] = why
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
            ln = ("- この文書は読むための窓です。書き込みは `todoku fill` で行います。値が空白・伏せ字（○○）の字数に収まるときは、"
                  "元の字を同じ字数で上書きします（ほかのバイトは変えません）。本文の欄は、空のマスに字を足すことも、空白・伏せ字より"
                  "長い値に替えることもできます（文字の書式の区切り・表の位置の表などを合わせて直し、行・ページの割り付けの控えを"
                  "取り除きます。割り付けは一太郎が開くときに作り直します）。「N 字まで」とある欄は、その字数までの上書きだけができる欄、"
                  "「変更できない欄」は書けない欄です（理由は欄の why_fixed）")
        elif ln.startswith("- 欄として示していない所に書くときは"):
            ln = ("- 欄として示していない所は、字を目印にして書けます：`{\"@1\": {\"old\": \"元の字\", \"new\": \"新しい字\"}}`"
                  "（本文なら old と new の字数が違ってもよい。枠・柱の中などは同じ字数にする（足りない分は空白「　」で埋める）。"
                  "old がいくつもあるときは、`near`（近くの欄の id）か `occurrence`（何番目か）で1つに決める）")
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
