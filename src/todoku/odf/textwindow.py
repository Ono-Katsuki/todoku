"""OpenDocument の文書の窓：Word の窓と同じ書き方（本文を上から、表は行ごとに ｜ で区切り、入力欄は［］）。

入力欄の推定（表の空のマス・段落の中の空欄・規則では決められない所）は Word と同じ処理。そのうえで、
- 保護された節・表の中の欄は「書けない欄」にする
- 同じ中身を繰り返す表のマス（table:number-…-repeated）は、1つだけには書けないので「書けない欄」にする
- マクロ（LibreOffice Basic）は、ソースとそのきっかけを窓に出す
"""

from __future__ import annotations

from ..docx.window import Spots
from ..docx.window import build as word_build
from ..docx.window import markdown_parts as word_parts
from .basic import macro_info
from .text import OdtDocument
from .xml import top


def build(doc: OdtDocument) -> dict:
    win = word_build(doc)
    win["kind"] = "fodt" if doc.pkg.flat else "odt"
    if doc.unread:
        win["unread"] = list(dict.fromkeys(doc.unread))
    if doc.repeats:
        win["repeats"] = list(dict.fromkeys(doc.repeats))
    win["kind_ja"] = "OpenDocument の文書（.fodt・1つの XML）" if doc.pkg.flat else "OpenDocument の文書（.odt）"
    win["source"]["note"] = "元のファイルが正本。書き込みは、書いた欄の部分だけを変える差分として適用する（ほかの部品はバイトのまま）。"
    win["notation"]["［P12.c5］"] += "（段落は、表のマスの段落も数える）"
    sp = Spots(doc)
    for f in win["fields"]:
        it = sp.items.get(f["id"]) or {}
        p = it.get("para")
        c_ = it.get("control")
        if p is None and c_ is not None:
            p = c_.para
        if c_ is not None and c_.source == "odf-form" and not c_.help and p is not None:
            near = _beside(p, c_)
            if near:
                f["label"] = near  # 見出しのない部品：段落の中で、部品のすぐ後ろ（チェック）か前（字の欄）の字
                f["name"] = c_.name
        if p is not None and getattr(p, "protected", False):
            f["readonly"] = True
            f["protected"] = True  # 保護された節・表の中
        c = p.cell if p is not None else None
        if c is not None and getattr(c, "repeated", False) and it.get("control") is None:
            f["readonly"] = True
            f["repeated_cell"] = True
    info = macro_info(doc.pkg, top(doc.pkg.tree(doc.part)))
    win["macros"] = info
    return win


def markdown_parts(win: dict) -> list[tuple[str, str]]:
    return word_parts(win)


def _beside(p, c) -> str:
    """部品の見出し：チェックは、段落の中のすぐ後ろの字（次の部品・タブの前まで）、字の欄は、すぐ前の字（前の部品・タブの後ろから）。"""
    from ..xlsx.window import label_text

    segs = [s for s in p.segs if not s.hidden]
    idx = next((k for k, s in enumerate(segs) if s.field is c), None)
    if idx is None:
        return ""
    if c.kind == "checkbox":
        run = []
        for s in segs[idx + 1:]:
            if s.field is not None or s.kind in ("tab", "br") or (s.node is not None and s.node.kind == "ctl"):
                break
            run.append(s.text)
        text = "".join(run)
    else:
        run = []
        for s in reversed(segs[:idx]):
            if s.field is not None or s.kind in ("tab", "br") or (s.node is not None and s.node.kind == "ctl"):
                break
            run.insert(0, s.text)
        text = "".join(run)
    # 前後の空白（一般カテゴリ Zs）と、見出しのあとのコロン（chars.LABEL_COLONS）を除く。字の意味は見ない
    from .. import chars

    trim = "".join(ch for ch in set(text) if chars.is_space(ch) or chars.is_label_colon(ch))
    return label_text(text.strip(trim)[:40])
