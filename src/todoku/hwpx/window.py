"""HWPX の窓：Word の窓と同じ書き方（本文を上から、表は行ごとに ｜ で区切り、入力欄は［］）。

入力欄の推定（表の空のマス・段落の中の空欄・規則では決められない所）は Word と同じ処理。そのうえで、
- 누름틀（クリックして書く欄）：名前（fieldBegin の name）を id にし、案内の字（Direction）を placeholder に添える
- 保護したマス（hp:tc の protect="1"）の中の欄は「書けない欄」にする
- 한글のマクロ（Scripts の JScript）は、ソースを窓に出す
"""

from __future__ import annotations

from ..docx.window import STORY_JA, Spots
from ..docx.window import build as word_build
from ..docx.window import markdown_parts as word_parts
from .document import HwpxDocument

KIND_JA = "한글の文書（.hwpx）"
STORY_JA.setdefault("master", "바탕쪽（マスターページ）")


def build(doc: HwpxDocument) -> dict:
    win = word_build(doc)
    win["kind"] = "hwpx"
    win["kind_ja"] = KIND_JA
    win["source"]["note"] = ("元のファイルが正本。書き込みは、書いた欄のある本文の節（Contents/sectionN.xml）の、その段落だけを変える差分として"
                             "適用する（ほかの部品はバイトのまま）。字を変えた段落は、한글の行の配置の控え（hp:linesegarray）を除き、"
                             "한글が開くときに配置し直す")
    win["notation"]["［P12.c5］"] += "（段落は、表のマス・図形の中の段落も数える）"
    win["notation"]["［名前］"] = "누름틀（クリックして書く欄）・フォームの部品。値がなければ、案内の字を placeholder に示す"
    sp = Spots(doc)
    for f in win["fields"]:
        it = sp.items.get(f["id"]) or {}
        c = it.get("control")
        p = it.get("para")
        if p is None and c is not None:
            p = c.para
        if c is not None and c.source == "hwpx-clickhere":
            if c.direction:
                f["placeholder"] = c.direction  # 値のないとき、한글が欄に表示する案内の字（印刷しない）
                if not f.get("label"):
                    f["label"] = c.direction
            if getattr(c, "help_text", ""):
                f["help"] = c.help_text
            f["kind"] = "text"
        if c is not None and getattr(c, "num_only", False):
            f["format"] = "数字だけ"  # hp:edit の numOnly="1"
        cell = p.cell if p is not None else None
        if (cell is not None and getattr(cell, "protected", False)) or (p is not None and getattr(p, "protected", False)):
            f["readonly"] = True
            f["protected"] = True  # 保護したマスの中（その中の図形の字も）
        if c is not None and c.source == "hwpx-clickhere" and c.end_ctrl is None:
            f["readonly"] = True  # 段落をまたぐ 누름틀：値の字が後ろの段落に続くので、欄の id では書かない
            f["why"] = ("この 누름틀は段落をまたいでいます（終わりの印が後ろの段落にある）。欄の id では書けないので、"
                        "中の字を目印にした書き込み（{\"@1\": {\"old\": …, \"new\": …}}）で書き換えてください")
    # 한글のマクロ（JScript）：ソースを窓に出す（動作は規則では再現しない）
    win["scripts"] = [{"name": n, "source": src} for n, src in doc.scripts]
    win["unread"] = list(doc.unread)
    return win


def build5(doc) -> dict:
    """HWP 5（読むだけ）の窓。"""
    win = word_build(doc)
    win["kind"] = "hwp"
    win["kind_ja"] = "한글の文書（.hwp・HWP 5。読むだけ）"
    win["source"]["note"] = ("元のファイルが正本。HWP 5 には書けません。HWPX で保存し直せば書ける"
                             "（한글の「다른 이름으로 저장」で .hwpx を選ぶ）")
    win["notation"]["［名前］"] = "누름틀（クリックして書く欄）。値がなければ、案内の字を placeholder に示す"
    sp = Spots(doc)
    for f in win["fields"]:
        f["readonly"] = True  # HWP 5 には書けない
        it = sp.items.get(f["id"]) or {}
        c = it.get("control")
        if c is not None and c.source == "hwp5-clickhere":
            f["kind"] = "text"
            if c.direction:
                f["placeholder"] = c.direction
                if not f.get("label"):
                    f["label"] = c.direction
    win["unread"] = list(dict.fromkeys(doc.unread))
    win["numbered_paragraphs"] = doc.numbered
    win["scripts"] = []
    return win


def markdown_parts(win: dict) -> list[tuple[str, str]]:
    parts = word_parts(win)
    if win.get("scripts"):
        lines = ["", "## 한글のマクロ（JScript）", "",
                 "- 文書を開く・閉じるときなどに한글が動かす処理です。動作は規則では再現しないので、自動で動く処理が要るかを読んで判断してください。", ""]
        for sc in win["scripts"]:
            lines += [f"### {sc['name']}", "", "```javascript", sc["source"], "```", ""]
        name, text = parts[-1]
        parts[-1] = (name, text + "\n".join(lines))
    return parts
