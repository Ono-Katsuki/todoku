"""RTF の窓：Word の窓と同じ書き方（本文を上から、表は行ごとに ｜ で区切り、入力欄は［］）。入力欄の推定も Word と同じ。"""

from __future__ import annotations

from ..docx.window import build as word_build
from ..docx.window import markdown_parts as word_parts
from .document import RtfDocument


def build(doc: RtfDocument) -> dict:
    win = word_build(doc)
    win["kind"] = "rtf"
    win["kind_ja"] = "RTF（リッチテキスト）"
    win["source"]["note"] = "元のファイルが正本。書き込みは、書いた欄の字だけを変える差分として適用する（ほかのバイトはそのまま）。"
    win["encoding"] = {"ansicpg": doc.ansicpg}
    win["unread"] = list(doc.unread)
    return win


def markdown_parts(win: dict) -> list[tuple[str, str]]:
    return word_parts(win)
