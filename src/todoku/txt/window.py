"""文字だけのファイル（.txt）の窓：すべての行を「行番号｜字」の形で、上から順に（空の行も、行の頭の空白・タブもそのまま）。

入力欄の推定は Word と同じ（docx/window.py の Spots：「令和　年　月　日」などの空欄・□・「？」の所）。1行を1段落として
数えるので、欄の id の「P12.1」は 12 行目の1つ目の空欄。空欄の字は［P12.1］に置き換えて示す（Word の窓と同じ）。
"""

from __future__ import annotations

import re

from ..docx.window import AGENT, KIND_JA, Spots
from ..docx.window import build as word_build
from .document import TxtDocument

BUDGET = 120_000


def _shown(text: str, spans: list[tuple[int, int, str, bool]]) -> str:
    """行の字の、空欄を［id］（規則では決められない所は［id？］、□ は ☐［id］）に置き換えたもの。spans：(始め, 終わり, id, 種類)。"""
    out, at = [], 0
    for a, b, fid, kind in sorted(spans):
        if a < at:
            continue
        out.append(text[at:a])
        out.append(mark(fid, kind, text[a:b]))
        at = b
    out.append(text[at:])
    return "".join(out)


def mark(fid: str, kind: str, covered: str) -> str:
    """空欄の印。kind：inline / checkbox / agent / placeholder（伏せ字の字は、＝のあとにそのまま出す）。"""
    if kind == "checkbox":
        return f"☐［{fid}］"
    if kind == "placeholder":
        return f"［{fid}？＝{covered}］"
    if kind == "agent":
        return f"［{fid}？］"
    return f"［{fid}］"


def build(doc: TxtDocument) -> dict:
    win = word_build(doc)  # 入力欄（Spots）・見出し・その行の文脈は、Word と同じ決め方
    sp = Spots(doc)
    by_line: dict[int, list] = {}
    for it in sp.items.values():
        p = it.get("para")
        if p is None or "span" not in it:
            continue
        a, b = it["span"]
        kind = it["kind"] if it["kind"] == "checkbox" else (
            ("placeholder" if it.get("placeholder") else "agent") if it.get("provenance") == AGENT else "inline")
        by_line.setdefault(p.number, []).append((a, b, it["id"], kind))
    lines = []
    for ln in doc.lines:
        spans = sorted(by_line.get(ln.number, []))
        lines.append({"n": ln.number, "text": ln.text, "terminator": ln.terminator,
                      "spans": [list(s) for s in spans], "shown": _shown(ln.text, spans)})
    win["kind"] = "txt"
    win["kind_ja"] = "テキスト（.txt）"
    win["source"]["note"] = ("元のファイルが正本。書き込みは、書いた所の字のバイトだけを、そのファイルの文字コードで作ったバイト列に"
                             "置き換える（ほかのバイト・改行・BOM はそのまま）。")
    win["notation"] = {"12｜": "12行目", "［P12.1］": "12行目の中の空欄（1つ目）",
                       "［P12.a1？］": "規則では決められない、書き込む所かもしれない所（エージェントが読んで判断する。書くなら id を指定する。"
                                      "書かなければ何も変わらない。［P12.a1？＝○○］は、伏せ字の○○を書き換える所）",
                       "☐［P12.1］": "チェック（□）"}
    win["format"] = doc.format_info()
    win["lines"] = lines
    for st in win["stories"]:
        st["reading"] = [x["shown"] for x in lines]  # 空の行・行の頭の空白も、そのまま
    win["unread"] = list(doc.unread)
    return win


def _md(s) -> str:
    return str(s).replace("|", "\\|").replace("\n", " ")


def _head(win: dict) -> list[str]:
    fm = win["format"]
    out = [f"# {win['source']['file'] or 'テキスト'}", "",
           f"- 元ファイルの SHA-256: `{win['source']['sha256']}`（{win['source']['bytes']:,} バイト）、{win['kind_ja']}",
           f"- 書き方：文字コード {fm['encoding_ja']}／改行 {fm['newline']}"
           + ("（行によって改行が違います）" if fm["mixed_newlines"] else "")
           + ("" if fm["ends_with_newline"] or not fm["lines"] else "／最後の行に改行なし") + f"／{fm['lines']:,} 行",
           "- この窓は読むためのものです。書き込み（`todoku fill`）は、書いた所の字のバイトだけを変えます"
           "（ほかの字・改行・文字コードはそのまま）。",
           "- 欄として示していない所に書くときは、字を目印にして書けます："
           "`{\"@1\": {\"old\": \"行の中の元の字\", \"new\": \"書き換えた後の字\"}}`（old がいくつもあるときは、"
           "`near`（近くの欄の id）か `occurrence`（何番目か）で1つに決める）。変わるのは old と new の違う所だけです。"
           "改行は書けません（1つの行の中の字だけ）。",
           "- 行は「行番号｜字」の形で、すべての行を出しています（空の行・行の頭の空白・タブもそのまま）。",
           "- 記号: " + "、".join(f"{k} は{v}" for k, v in win["notation"].items())]
    if fm.get("ascii_only"):
        out.append("- このファイルは ASCII の字だけなので、文字コードは決まりません（日本語を書くと UTF-8 で書きます）。")
    if fm["encoding"] == "latin-1":
        out.append("- **文字コードを決められませんでした**：UTF-8・Shift_JIS・EUC-JP のどれとしても読めないので、Latin-1 として読みました"
                   "（字が正しいとは限りません）。")
    out.append("")
    if win["fields"]:
        out += [f"## 入力欄（{len(win['fields'])} か所）", "", "| 欄 | 見出し | 種類 | いまの値 |", "|---|---|---|---|"]
        for f in win["fields"]:
            k = KIND_JA.get(f["kind"], f["kind"])
            if f.get("agent"):
                k = f"規則では決められない：{f['why']}"
            v = f.get("value")
            vs = ("☑ あり" if v else "☐ なし") if f["kind"] == "checkbox" else ("" if v in (None, "") else str(v))
            out.append(f"| {_md(f['id'])} | {_md(f['label'])} | {_md(k)} | {_md(vs)} |")
        out.append("")
    return out


def line_text(n: int, shown: str) -> str:
    """窓の1行：「行番号｜字」。"""
    return f"{n}｜{shown}"


def markdown_parts(win: dict, budget: int = BUDGET) -> list[tuple[str, str]]:
    head = _head(win)
    body = [line_text(x["n"], x["shown"]) for x in win["lines"]]
    title = f"## 本文（{len(body):,} 行）"
    lines = head + [title, ""] + body
    whole = "\n".join(lines) + "\n"
    if len(whole) <= budget:
        return [("全体", whole)]
    # 大きい窓：入力欄の一覧も本文も、行の切れ目で分ける（1つの行は分けない）
    parts: list[list[str]] = []
    cur: list[str] = []
    size = 0
    for line in lines:
        if cur and size + len(line) + 1 > budget - 2000:
            parts.append(cur)
            cur, size = [], 0
        cur.append(line)
        size += len(line) + 1
    if cur:
        parts.append(cur)
    num = re.compile(r"(\d+)｜")
    index = head[:9] + ["## この窓の目次", "", f"窓が大きいので、{len(parts)} の部分に分けました（入力欄の一覧と、本文の行）。", ""]
    for i, p in enumerate(parts, 1):
        ns = [m[1] for x in p for m in [num.match(x)] if m]
        index.append(f"- 部分 {i}: " + (f"本文の {ns[0]}〜{ns[-1]} 行目" if ns else "入力欄の一覧など"))
    return [("目次", "\n".join(index))] + [(f"部分 {i}", "\n".join(head[:1] + [""] + p)) for i, p in enumerate(parts, 1)]
