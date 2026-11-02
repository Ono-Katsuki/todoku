"""Word（.docx・.doc）：記入例と白紙の様式を照らして、入力欄の見落としがないかを確かめる。

白紙と記入例の段落を順に対応づけ（文字の並びが違っても、表の同じマス・似た段落どうし）、
記入例で字が足された・書き換えられた所が、白紙で見つけた入力欄（空のマス・文字の中の空欄・□・フォームの欄）に
かかっているかを調べる。かかっていない所は「見落とした入力欄」の候補として返す。
"""

from __future__ import annotations

import difflib
import re
import unicodedata
from dataclasses import dataclass, field

from .document import Para, Table, iter_paras
from .window import Spots


@dataclass
class WordMissed:
    para: int
    text: str
    context: str


@dataclass
class WordExampleCheck:
    added: int = 0
    covered: int = 0  # 規則で見つけた入力欄、または規則では決められない所（エージェントの判断）にかかっている
    covered_agent: int = 0  # そのうち、規則では決められない所（［P12.a1？］）にかかっている
    template: int = 0
    notes: int = 0
    drawn: int = 0  # 記入例で描き足された図形（○で囲む印・吹き出し）
    missed: list[WordMissed] = field(default_factory=list)
    unpaired: int = 0
    kept: float = 1.0  # 白紙の字のうち、記入例にそのまま残っている割合（低いと、様式の版が違う）

    @property
    def ok(self) -> bool:
        return not self.missed

    @property
    def same_version(self) -> bool:
        return self.kept >= 0.65

    def headline(self) -> str:
        s = (f"記入例で書き込まれた所 {self.added} か所のうち {self.covered - self.covered_agent} か所が、規則で見つけた入力欄に、"
             f"{self.covered_agent} か所が、規則では決められない所（エージェントの判断）にかかっています")
        if self.template or self.notes or self.drawn:
            s += (f"（様式の版の違い {self.template} か所・記入例の説明 {self.notes} か所・"
                  f"描き足された図形 {self.drawn} か所は除外）")
        if self.missed:
            s += f"。入力欄の外：{len(self.missed)} か所（見落としの候補）"
        if self.unpaired:
            s += f"。対応づけられなかった段落 {self.unpaired}"
        return s


def _open(blob: bytes):
    from ..book import kind_of

    if kind_of(blob) == "doc":
        from ..doc.book import DocFile

        return DocFile(blob, "x.doc")
    from .document import DocxDocument

    return DocxDocument(blob, "x.docx")


def _paras(doc) -> list[Para]:
    out = []
    for _k, _p, blocks in doc.all_blocks():
        out += list(iter_paras(blocks))
    return out


def _skeleton(t: str) -> str:
    """段落の型（空白・数字・○などの書き込みやすい字を除いた並び）。対応づけに使う。"""
    return re.sub(r"[\s　\d０-９○〇□■☑☐☒✓レ\-－‐―・.,、。]", "", t)


def _pair(pa: list[Para], pb: list[Para]) -> list[tuple[int, int]]:
    if len(pa) == len(pb):
        return list(zip(range(len(pa)), range(len(pb))))
    sa = [_skeleton(p.text) for p in pa]
    sb = [_skeleton(p.text) for p in pb]
    out = []
    sm = difflib.SequenceMatcher(None, sa, sb, autojunk=False)
    for tag, a1, a2, b1, b2 in sm.get_opcodes():
        if tag == "equal" or (tag == "replace" and a2 - a1 == b2 - b1):
            out += list(zip(range(a1, a2), range(b1, b2)))
    return out


def _is_note(t: str) -> bool:
    return any(w in t for w in ("記入例", "記載例", "作成例", "見本")) or ("ください" in t and len(t) >= 10)


def check(blank: bytes, example: bytes) -> WordExampleCheck:
    a = _open(blank)
    b = _open(example)
    sp = Spots(a)
    pa, pb = _paras(a), _paras(b)
    by_para: dict[int, list[dict]] = {}
    for it in sp.items.values():
        p = it.get("para")
        if p is None and it.get("control") is not None:
            p = it["control"].para
        if it.get("cell") is not None:
            for q in it["cell"].paras():  # 空のマスの欄は、マスのどの段落に書いても、その欄
                by_para.setdefault(id(q), []).append(it)
        elif p is not None:
            by_para.setdefault(id(p), []).append(it)
    out = WordExampleCheck()
    pairs = _pair(pa, pb)
    out.unpaired = len(pa) - len(pairs)
    # 白紙の字（空白を除く）のうち、対応づけた記入例の段落にそのまま残っている字の割合
    total = sum(len(re.sub(r"\s", "", q.text)) for q in pa)
    kept = 0
    for i, j in pairs:
        sm = difflib.SequenceMatcher(None, pa[i].text, pb[j].text, autojunk=False)
        kept += sum(len(re.sub(r"\s", "", pa[i].text[m.a:m.a + m.size])) for m in sm.get_matching_blocks())
    out.kept = round(kept / total, 2) if total else 1.0
    for i, j in pairs:
        x, y = pa[i], pb[j]
        if x.text == y.text:
            continue
        items = by_para.get(id(x), [])
        sm = difflib.SequenceMatcher(None, x.text, y.text, autojunk=False)
        for tag, a1, a2, b1, b2 in sm.get_opcodes():
            if tag == "equal":
                continue
            new = y.text[b1:b2]
            old = x.text[a1:a2]
            blank_chars = " \u3000\u2002\t"
            if not new.strip(blank_chars) and not old.strip(blank_chars):
                continue  # 空白だけの違い
            if not new.strip(blank_chars) and tag != "delete":
                continue  # 空白を足しただけ
            if unicodedata.normalize("NFKC", new).strip() == unicodedata.normalize("NFKC", old).strip():
                out.template += 1  # 全角・半角だけの違い（様式の版の違い）
                continue
            near = "".join(q.text for q in pa[max(0, i - 2):i + 3])
            if len(new.strip()) >= 3 and new.strip(blank_chars) in near:
                out.template += 1  # となりの段落にある字（見出しの分け方の違い）
                continue
            if tag == "delete" and old.strip(" 　"):
                # 白紙の字が消えた：様式の版の違いか、□ を ■ にしたもの
                if old.strip() in ("□", "☐"):
                    continue
                out.template += 1
                continue
            if _is_note(new):
                out.notes += 1
                continue
            if new.replace("〔図〕", "").strip(blank_chars) == "" and "〔図〕" in new:
                out.drawn += 1  # 図形だけ（○で囲む印・吹き出し）：字の書き込みではない
                continue
            out.added += 1
            by = _covered(x, a1, a2, items)
            if by:
                out.covered += 1
                if by == "agent":
                    out.covered_agent += 1
            elif old.strip(" 　 ") and difflib.SequenceMatcher(None, old, new).ratio() >= 0.5 and len(old) >= 4:
                out.added -= 1
                out.template += 1  # 似た文への書き換え：様式の版の違い
            else:
                out.missed.append(WordMissed(i + 1, new.strip()[:40], x.text.strip()[:60]))
    return out


def _covered(p: Para, a1: int, a2: int, items: list[dict]) -> str | None:
    """記入例の書き込み [a1, a2)（白紙の段落の位置）がかかる欄：規則で見つけた欄なら "rules"、規則では決められない所だけなら "agent"。"""
    agent = False
    for it in items:
        if it.get("cell") is not None and it["kind"] == "cell":
            return "rules"  # 空のマス：段落ぜんぶが入力欄
        if "span" in it:
            a, b = it["span"]
            if a - 1 <= a1 <= b + 1 or a - 1 <= a2 <= b + 1 or (a1 <= a and b <= a2):
                if it.get("provenance") == "agent":
                    agent = True
                    continue
                return "rules"
        c = it.get("control")
        if c is not None:
            return "rules"
    return "agent" if agent else None
