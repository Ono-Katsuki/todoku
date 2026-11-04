"""読むだけの形式（.ppt・.xps・.epub・.eml・.dbf など）の共通の窓・段階・検査。

各形式は、節の並び（見出しと行）を作るだけ。書き込みはしない（書く D。理由を添える）。
読み落としの検査：形式ごとに、窓を作る処理とは別の読み方で字を数え（raw）、窓の字と字の種類ごとに照らす。
"""

from __future__ import annotations

import hashlib
import re
from collections import Counter

PART_LIMIT = 120_000  # Markdown の 1 部分の字数の目安


def make(kind: str, kind_ja: str, name: str, blob: bytes, sections: list[tuple[str, list[str]]], note: str,
         unread: list[str] | None = None, **extra) -> dict:
    return {"kind": kind, "kind_ja": kind_ja, "file": name,
            "source": {"file": name, "sha256": hashlib.sha256(blob).hexdigest(), "bytes": len(blob), "note": note},
            "sections": [{"heading": h, "lines": list(ls)} for h, ls in sections], "fields": [], "unread": list(unread or []), **extra}


def markdown_parts(win: dict) -> list[tuple[str, str]]:
    head = [f"# {win['file']}（{win['kind_ja']}）", "", f"> {win['source']['note']}", ""]
    head += [f"- **読んでいない部分**：{u}" for u in win.get("unread") or []]
    if win.get("unread"):
        head.append("")
    parts: list[tuple[str, str]] = []
    cur: list[str] = list(head)
    size = sum(len(x) for x in cur)
    label = "本文"
    for sec in win["sections"]:
        block = [f"## {sec['heading']}", ""] + (sec["lines"] or ["（字はありません）"]) + [""]
        n = sum(len(x) + 1 for x in block)
        if cur and size + n > PART_LIMIT and size > 0:
            parts.append((label, "\n".join(cur) + "\n"))
            cur, size, label = [], 0, sec["heading"]
        cur += block
        size += n
    from .links import markdown as _links

    cur += _links(win.get("links") or [])
    parts.append((label, "\n".join(cur) + "\n"))
    return parts


def window_text(win: dict) -> str:
    return "".join(ln for s in win["sections"] for ln in s["lines"]) + "".join(s["heading"] for s in win["sections"])


def chars(text) -> Counter:
    """字の種類ごとの数（空白と、窓の区切りの字は数えない）。"""
    s = text if isinstance(text, str) else "".join(text)
    return Counter(re.sub(r"[\s　｜／]", "", s))


def grade(win: dict, read_grade: str, title: str, detail: str, write_title: str, write_detail: str, more=None):
    """読む段階（理由つき）と、書く D。unread は C（読んでいない部分を窓に書いた）。"""
    from .grade import Reason, _final

    reasons = [Reason(read_grade, title, detail)]
    reasons += [Reason("C", "読んでいない部分", u) for u in win.get("unread") or []]
    reasons += list(more or [])
    return _final(reasons, [Reason("D", write_title, write_detail)])


def audit(win: dict, name: str, raw, what: str, write_note: str, extra_checks=None):
    """raw：窓を作るのとは別の読み方で取り出した字（文字列か、文字列の並び）。窓の字が raw の字をすべて含むかを照らす。"""
    from .xlsx.audit import Audit, Check

    au = Audit(name)
    a, b = chars(raw), chars(window_text(win))
    missing = a - b
    items = [f"「{c}」が {n} 字、窓に出ていません" for c, n in list(missing.items())[:20]]
    au.checks.append(Check("読み落とし（字）", not missing, f"{what}の字 {sum(a.values()):,} 字を、窓を作るのとは別に数え、すべて窓に出ていることを確かめました"
                           if not missing else f"{what}の字のうち {sum(missing.values()):,} 字が窓に出ていません", items))
    parts = markdown_parts(win)
    biggest = max(len(t) for _, t in parts)
    au.checks.append(Check("窓の大きさ", biggest <= 150_000, f"Markdown {sum(len(t) for _, t in parts):,} 文字を {len(parts)} 部分に分けました（AIが1回で読める目安 15万文字以内）"))
    if win.get("unread"):
        au.checks.append(Check("読んでいない部分", True, "読んでいない部分を窓と段階に書きました", list(win["unread"]), agent=True))
    au.checks += list(extra_checks or [])
    au.checks.append(Check("書き込み", True, write_note, agent=True))
    return au
