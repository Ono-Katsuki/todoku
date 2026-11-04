"""PowerPoint の窓：Word の窓と同じ書き方（スライドごとの節に、字を上から。表は行ごとに ｜ で区切り、入力欄は［］）。

入力欄の推定も Word と同じ（docx/window.py の Spots：表の空のマス T1.R2.C3・段落の中の空欄 P12.c5・□・「？」の所）。
段落の番号は、スライドの順に（スライド・そのノート）数える。スライドの型（レイアウト・マスター）の、プレースホルダーでない
図形の字は、すべてのスライドに出る字として別の節に出す（欄にはしない）。
"""

from __future__ import annotations

from ..docx.window import Spots, _render_blocks
from ..docx.window import build as word_build
from ..docx.window import markdown_parts as word_parts
from .document import PptxDocument

TEMPLATE = "スライドの型の字（レイアウト・マスター。そのスライドにも出る）"


def build(doc: PptxDocument) -> dict:
    win = word_build(doc)
    win["kind"] = "pptx"
    win["kind_ja"] = "PowerPoint（.pptx）"
    win["source"]["note"] = ("元のファイルが正本。書き込みは、書いた段落のあるスライドの部品だけを、書いた字の差分として変える"
                             "（ほかの部品は、圧縮したバイトのまま）。")
    win["notation"]["［P12.c5］"] += "（段落は、スライドの順に数える）"
    win["slides"] = [{"n": s["n"], "part": s["part"], "hidden": s["hidden"], "notes": s["notes"]} for s in doc.slides]
    sp = Spots(doc)
    for part, blocks in doc.templates:
        lines: list[str] = []
        plain: list[str] = []
        _render_blocks(blocks, sp, {}, lines, plain)
        if lines:
            win["stories"].append({"kind": TEMPLATE, "part": part, "reading": lines, "_plain": plain, "template": True})
    win["pictures"] = [{"id": f"図{p['n']}", "alt": p["alt"] or None, "decorative": p["decorative"],
                        "where": f"スライド {p['slide']}" + (f"の「{p['name']}」" if p["name"] else "")} for p in doc.pictures]
    win["unread"] = list(doc.unread)
    return win


def markdown_parts(win: dict) -> list[tuple[str, str]]:
    return word_parts(win)
