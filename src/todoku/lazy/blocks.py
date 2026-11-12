"""段落と表でできた書類（Word・一太郎・한글・PowerPoint・OpenDocument・RTF）を、まとまり（chunk）ごとに読む共通の部品。

どの形式も、読むと Word と同じ形（段落 Para・表 Table）になり、窓は docx/window の Spots と _render_blocks で作る。そこで：
- BlocksLazy：まとまりごとに小さな書類を読み、段落の範囲で窓を返す・欄を返す。番号（段落・表・欄の名前）は、書類全体を読んだときと同じ
- SubDoc：読んだ書類のブロックの一部だけを見せる窓の入口（all_blocks を絞り、番号の前までの数を添える）
- ViewLazy：全体を読んでから（窓は作らずに）分ける形式（一太郎・Word 6.0/95・RTF）
- SeqLazy：単位（unit）ごとに小さく読める形式（.doc・HWP・OpenDocument・HWPX）。ひととおり読むときは 1 単位ずつ読んで手放す（メモリは 1 単位の分）。
  ヘッダー・脚注などは本文の後ろの番号になるので、全体と同じ順（種類の順・単位の順）に並べ直す
"""

from __future__ import annotations

import bisect
import re
from collections import Counter

from .. import limits
from ..docx.fontfacts import count_colors, most_common_color
from .base import LazyDoc, Page, RangeError, comma, parse_range

_P_ID = re.compile(r"P(\d+)\.")
STORY_JA = {"header": "ヘッダー", "footer": "フッター", "footnotes": "脚注", "endnotes": "文末脚注", "comments": "コメント",
            "textbox": "テキストボックス", "header-textbox": "ヘッダーのテキストボックス", "master": "マスターページ"}
VIEW_CHUNK_PARAS = 400  # 全体を読んである書類の、1 まとまりの段落の数の目安


def _para_of(f: dict) -> int | None:
    m = _P_ID.match(str(f.get("id")))
    return int(m[1]) if m else None


class Chunk:
    """1 つのまとまりを読んだもの。"""

    def __init__(self, doc, sp, by_para, units):
        self.doc, self.sp, self.by_para, self.units = doc, sp, by_para, units


class SubDoc:
    """読んだ書類の一部を、書類全体のように見せる。blob は空（窓の頭のハッシュを、まとまりごとに計算しない）。"""

    def __init__(self, doc, blocks: list, para_base: int = 0, table_base: int = 0, control_ids: dict | None = None, objects: bool = False, orphans: bool = False):
        self._doc = doc
        self._blocks = blocks  # [(種類, 部品, ブロック)]
        self.para_base = para_base
        self.table_base = table_base
        self.control_ids = control_ids or {}
        self.blob = b""
        from ..docx.document import iter_paras

        seen = {id(p) for _k, _part, bl in blocks for p in iter_paras(bl)}
        # 段落に結び付かない欄（置き場所のないフォームの部品）は、はじめのまとまりにだけ見せる
        self.controls = [c for c in getattr(doc, "controls", []) if (c.para is not None and id(c.para) in seen) or (c.para is None and orphans)]
        if not objects:
            self.embeds, self.pictures, self.charts, self.diagrams = [], [], [], []
            self.unread = []
            self.scripts = []
            self.has_macros = False
            self.vba = None

    def __getattr__(self, name):
        return getattr(self._doc, name)

    def all_blocks(self):
        yield from self._blocks

    @property
    def body(self):
        return next((b for k, _p, b in self._blocks if k == "body"), [])

    @property
    def stories(self):
        return [t for t in self._blocks if t[0] != "body"]

    def marks(self):
        real = getattr(self._doc, "marks", None)
        if real is None:
            from ..docx.window import _bookmarks

            return _bookmarks(self)
        from ..docx.document import iter_paras

        seen = {id(p) for _k, _part, bl in self._blocks for p in iter_paras(bl)}
        return {k: v for k, v in real().items() if id(v[0]) in seen}


class BlocksLazy(LazyDoc):
    """chunks：[{"para": 前までの段落の数, "kind": 本文でなければ種類, "part": 部品}]。派生は chunk_doc(i)・build_window(doc) を作る。"""

    streams = False
    range_keys = ("paragraphs",)
    chunks: list = []
    paragraphs = 0
    tables = 0
    name = ""

    def chunk_doc(self, i: int, real: tuple = ()):
        raise NotImplementedError

    def told_doc(self, i: int):
        """窓・欄にするまとまりの文書。字の色の「ふつう」は、まとまりごとには決めない（まとまりで印が変わる）：
        段落の位置を調べるときに全体で数えた色（normal_font_color）。数えていない形式は、黒をふつうとし、黒でない色にはみな印を付ける。"""
        from ..xlsx.fontfacts import AUTO

        d = self.chunk_doc(i)
        d.__dict__.setdefault("_normal_font_color", getattr(self, "normal_font_color", AUTO))
        return d

    def build_window(self, doc) -> dict:
        from ..docx.window import build

        return build(doc)

    def chunk(self, i: int) -> Chunk:
        cache = self.__dict__.setdefault("_cache", {})
        got = cache.get(i)
        if got is not None:
            return got
        from ..docx.document import iter_paras
        from ..docx.window import Spots

        d = self.told_doc(i)
        sp = Spots(d)
        by_para: dict[int, list[dict]] = {}
        for it in sp.items.values():
            q = it.get("para")
            if q is not None:
                by_para.setdefault(id(q), []).append(it)
        units = []
        for _kind, _part, blocks in d.all_blocks():
            for b in blocks:
                nums = [sp.para_index[id(q)] for q in iter_paras([b]) if id(q) in sp.para_index]
                if nums:
                    units.append((min(nums), max(nums), b))
        units.sort(key=lambda u: u[0])
        got = Chunk(d, sp, by_para, units)
        cache[i] = got
        order = self.__dict__.setdefault("_cache_order", [])
        order.append(i)
        while len(order) > 2:
            cache.pop(order.pop(0), None)
        return got

    def _chunk_at(self, para: int) -> int:
        firsts = self.__dict__.get("_firsts")
        if firsts is None:
            firsts = self._firsts = [c["para"] for c in self.chunks]
        return max(0, bisect.bisect_left(firsts, para) - 1)

    # ------------------------------------------------------------------ 読む
    def read(self, rng: dict, max_chars: int) -> Page:
        from ..docx.window import _render_blocks

        cur = rng.get("cursor")
        if cur and str(cur).startswith("extras:"):
            return self._read_extras(int(str(cur)[7:]), max_chars)
        total = self.paragraphs
        if cur:
            m = re.fullmatch(r"paragraphs:(\d+):(\d+)(:w)?", str(cur))
            if not m:
                raise RangeError(f"cursor「{cur}」はこの書類の続きの印ではありません（前の結果の next をそのまま渡してください）")
            a, b, whole = int(m[1]), int(m[2]), bool(m[3])
        else:
            a, b = parse_range(rng.get("paragraphs"), "段落")
            whole = b is None
            a = 1 if a is None else a
            b = total if b is None else min(b, total)
        if a > max(total, 1):
            raise RangeError(f"段落 {a} はありません（全 {comma(total)} 段落）")
        body: list[str] = []
        size = 0
        stop = None
        i = self._chunk_at(a)
        said = None
        while i < len(self.chunks) and stop is None:
            ch = self.chunks[i]
            if ch["para"] >= b:
                break
            limits.progress("段落を読んでいます", ch["para"], total)
            c = self.chunk(i)
            if ch.get("kind") and said != i and c.units and c.units[-1][1] >= a:
                said = i
                body.append(f"〔{STORY_JA.get(ch['kind'], ch['kind'])}（{ch.get('part', '')}）〕")
            for lo, hi, block in c.units:
                if hi < a or lo > b:
                    continue
                lines: list[str] = []
                plain: list[str] = []
                _render_blocks([block], c.sp, c.by_para, lines, plain)
                text = "\n".join(f"〔段落 {lo}" + (f"〜{hi}" if hi != lo else "") + f"〕{ln}" for ln in lines)
                if body and size + len(text) > max_chars:
                    stop = lo
                    break
                if text:
                    body.append(text)
                    size += len(text) + 1
            i += 1
        cov = f"段落 {a}〜{(stop - 1) if stop else b}（全 {comma(total)} 段落）"
        if stop is None:
            if whole and self.has_extras():
                return Page("\n".join(body), cov, {"cursor": "extras:0"}, "埋め込み・グラフ・マクロ・画像の一覧の節", body)
            return Page("\n".join(body), cov, None, "", body)
        return Page("\n".join(body), cov, {"cursor": f"paragraphs:{stop}:{b}" + (":w" if whole else "")}, f"{comma(b - stop + 1)} 段落", body)

    def has_extras(self) -> bool:
        return bool(self.extras_lines())

    def extras_lines(self) -> list[str]:
        got = self.__dict__.get("_extras")
        if got is None:
            got = self._extras = self._make_extras()
        return got

    def _make_extras(self) -> list[str]:
        return []

    def _read_extras(self, at: int, max_chars: int) -> Page:
        text = "\n".join(self.extras_lines())
        end = min(at + max_chars, len(text))
        if end < len(text):
            nl = text.rfind("\n", at, end)
            if nl > at:
                end = nl + 1
        piece = text[at:end]
        cov = f"埋め込み・グラフ・マクロ・画像の一覧の節 {at + 1:,}〜{end:,} 字（全 {len(text):,} 字）"
        if end >= len(text):
            return Page(piece, cov, None, "", [piece])
        return Page(piece, cov, {"cursor": f"extras:{end}"}, f"{comma(len(text) - end)} 字", [piece])

    # ------------------------------------------------------------------ 欄
    def _chunk_fields(self, i: int) -> list[dict]:
        return self.build_window(self.told_doc(i))["fields"]

    def fields(self, rng: dict, only_empty: bool, limit: int) -> dict:
        from ..api import _field

        total = self.paragraphs
        cur = rng.get("cursor")
        if cur:
            m = re.fullmatch(r"fields:(\d+):(\d+)", str(cur))
            if not m:
                raise RangeError(f"cursor「{cur}」はこの書類の続きの印ではありません")
            start_chunk, b = int(m[1]), int(m[2])
            a = 1
        else:
            a, b = parse_range(rng.get("paragraphs"), "段落")
            a = 1 if a is None else a
            b = total if b is None else min(b, total)
            start_chunk = self._chunk_at(a)
        out: list[dict] = []
        i = start_chunk
        while i < len(self.chunks) and self.chunks[i]["para"] < b:
            limits.progress("欄を探しています", self.chunks[i]["para"], total)
            for f in self._chunk_fields(i):
                pn = _para_of(f)
                if pn is not None and not a <= pn <= b:
                    continue
                d = _field(f)
                if only_empty and d.get("value") not in (None, "", False):
                    continue
                out.append(d)
            i += 1
            if len(out) >= limit and i < len(self.chunks) and self.chunks[i]["para"] < b:
                return {"fields": out, "total_paragraphs": total, "looked_paragraphs": f"{a}〜{self.chunks[i]['para']}",
                        "next": {"cursor": f"fields:{i}:{b}"}, "remaining": f"{comma(b - self.chunks[i]['para'])} 段落"}
        return {"fields": out, "total_paragraphs": total, "looked_paragraphs": f"{a}〜{b}", "next": None}

    # ------------------------------------------------------------------ 概要
    def overview(self) -> dict:
        out = {"bytes": getattr(self, "size", None), "window_chars": self.est_chars, "paragraphs": self.paragraphs, "tables": self.tables,
               "ranges": "paragraphs（段落の範囲）か cursor"}
        if self._normal_note():
            out["normal_font_color"] = self.normal_font_color
        return out

    def _normal_note(self) -> str:
        """黒でない色がふつうの文書：印の付かない字が黒とは限らないので、ふつうの色そのものを伝える（ふつうの道の窓と同じ）。"""
        from ..xlsx.fontfacts import AUTO

        c = getattr(self, "normal_font_color", AUTO)
        return "" if c == AUTO else f"- この文書のふつうの字の色は #{c} です（黒ではない。この色の字には印を付けません）\n"

    def overview_text(self) -> str:
        from ..api import KIND_JA

        return (f"# {self.name}\n\n- {KIND_JA.get(self.kind, self.kind)}（{getattr(self, 'size', 0) or 0:,} バイト）・段落 {self.paragraphs:,}・表 {self.tables:,}\n"
                f"- 大きいので、ここでは本文を読んでいません（窓の全体は見積もりで {self.est_chars:,} 字）。read_window に paragraphs: \"1-200\" のように段落の範囲を渡して読みます"
                "（cursor をたどれば、すべてを読めます。max_chars で 1 回の大きさを変えられます。窓を作るのが上限で止まるときは raw: true で字のまま読めます）。"
                "list_fields に paragraphs を渡すと、その範囲の欄を出します。\n" + self._normal_note())

    # ------------------------------------------------------------------ 段階・検査
    def grade(self):
        from ..grade import Reason, _final

        write = [Reason("B", "欄を範囲ごとに見つける", "大きな文書なので、欄は頼まれた段落の範囲ごとに見つけます")]
        read = []
        if getattr(self, "tracked", 0):  # ふつうの道（grade.word_reasons）と同じ理由
            read.append(Reason("B", "変更履歴", f"変更履歴で削除された字が {self.tracked} 段落に残っています。窓では本文に入れず"
                               "〔変更履歴で削除された字〕と示しています。印刷される字は、ソフトの変更履歴の表示の設定で変わります"))
        g = _final(read, write)
        g.deferred = ["未確認：全段落の字の読み落とし（audit_document で全部のまとまりを確かめます）"]
        return g

    def audit_begin(self) -> dict:
        return {"next_chunk": 0, "chars": 0, "fields": 0, "problems": [], "done": False, "checks": []}

    def audit_step(self, state: dict) -> dict:
        """いままでの検査で全体を確かめられる大きさなら、それを動かす。大きすぎるときは、まとまりごとに窓を作り、段落の字が窓に出ていることを確かめる。"""
        import os

        if state.get("next_chunk", 0) == 0 and not state.get("skip_whole") and os.path.getsize(self.path) <= limits.MAX_FILE // 8:
            from .. import api

            limits.progress("書類の全体を検査しています")
            au = api.audit_whole(self.path, getattr(self, "password", None), None)
            state["checks"] = au.get("checks", [])
            state["full"] = au
            state["done"] = True
            return state
        from collections import Counter

        from ..docx.audit import _count
        from ..docx.document import iter_paras

        i = state["next_chunk"]
        n = len(self.chunks)
        if i < n:
            d = self.told_doc(i)
            win = self.build_window(d)
            raw = Counter()
            for _k, _part, blocks in d.all_blocks():
                raw += _count(p.raw_text() for p in iter_paras(blocks))
            shown = Counter()
            for st in win["stories"]:
                shown += _count(st["_plain"])
            state["chars"] += sum(raw.values())
            state["fields"] += len(win["fields"])
            for ch, k in list((raw - shown).items())[:5]:
                if len(state["problems"]) < 30:
                    state["problems"].append(f"段落 {self.chunks[i]['para'] + 1}〜：「{ch}」が {k} 字、窓に出ていません")
            state["next_chunk"] = i + 1
            limits.progress("まとまりごとに確かめています", i + 1, n)
            return state
        state["checks"] = [{"name": "読み落とし（文書の文字）", "ok": not state["problems"], "agent": False, "items": list(state["problems"]),
                            "detail": f"全 {self.paragraphs:,} 段落・字 {state['chars']:,} 字が、まとまり {n} 個ごとの窓に出ていることを確かめました。見つけた欄 {state['fields']:,} か所"},
                           {"name": "窓の大きさ", "ok": True, "agent": False, "items": [], "detail": f"窓の全体は見積もりで {self.est_chars:,} 字。段落の範囲で分けて読めます"}]
        state["still_deferred"] = ["未確認：書き込みの試し（大きすぎるので検査では試していません。fill_document のときに証明します）"]
        state["done"] = True
        return state


def _window_extras(win: dict, md_parts) -> list[str]:
    parts = md_parts(win)
    text = "\n".join(t for _n, t in (parts[1:] if len(parts) > 1 else parts))
    k = text.find("\n## ")
    return text[k + 1:].split("\n") if k >= 0 else []


class ViewLazy(BlocksLazy):
    """全体を読んで（窓は作らずに）、ブロックを分けて窓にする。"""

    def __init__(self, path: str, name: str, kind: str, doc, build, md_parts, size: int, password: str | None = None):
        from ..docx.document import iter_paras

        self.path, self.name, self.kind, self.doc, self.build_fn, self.md_parts, self.size = path, name, kind, doc, build, md_parts, size
        self.password = password
        self.chunks = []
        self.control_ids = _control_ids(doc)
        para = table = chars = tracked = 0
        colors: Counter = Counter()
        for k, part, blocks in doc.all_blocks():
            start = 0
            first = True
            for j, b in enumerate(blocks):
                n = sum(1 for _ in iter_paras([b]))
                if j > start and (self._chunk_paras(blocks, start, j) + n > VIEW_CHUNK_PARAS):
                    self._add(k, part, blocks[start:j], para, table, first)
                    para, table = self._counts(blocks[start:j], para, table)
                    first = False
                    start = j
            if blocks[start:] or k == "body":
                self._add(k, part, blocks[start:], para, table, first)
                para, table = self._counts(blocks[start:], para, table)
            chars += sum(len(p.text) for p in iter_paras(blocks))
            tracked += sum(1 for p in iter_paras(blocks) if hasattr(p, "deleted_text") and p.deleted_text().strip())
            count_colors(iter_paras(blocks), colors)
        self.tracked = tracked  # 変更履歴で削除された字の残る段落の数
        self.normal_font_color = most_common_color(colors)  # 文書のふつうの字の色（ふつうの道と同じ数え方）
        self.paragraphs, self.tables = para, table
        self.est_chars = int(chars * 1.25) + para * 12 + len(self.control_ids) * 60

    @staticmethod
    def _chunk_paras(blocks, a, b) -> int:
        from ..docx.document import iter_paras

        return sum(1 for _ in iter_paras(blocks[a:b]))

    @staticmethod
    def _counts(blocks, para, table):
        from ..docx.document import iter_paras
        from ..docx.window import _tables

        return para + sum(1 for _ in iter_paras(blocks)), table + sum(1 for _ in _tables(blocks))

    def _add(self, k, part, blocks, para, table, first) -> None:
        self.chunks.append({"para": para, "table": table, "kind": None if k == "body" else k, "part": part, "k": k, "blocks": blocks})

    def chunk_doc(self, i: int, real: tuple = ()):
        ch = self.chunks[i]
        return SubDoc(self.doc, [(ch["k"], ch["part"], ch["blocks"])], ch["para"], ch["table"], self.control_ids)

    def build_window(self, doc) -> dict:
        return self.build_fn(doc)

    def _make_extras(self) -> list[str]:
        from ..imaging.docimages import list_images, markdown

        d = SubDoc(self.doc, [], 0, 0, self.control_ids, objects=True)
        d.controls = []
        win = self.build_fn(d)
        lines = _window_extras(win, self.md_parts)
        try:
            images = list_images(self.kind, self.doc.blob, win)
        except Exception:  # noqa: BLE001  一覧を作れないだけ（中身は読める）
            images = []
        if images:
            lines = lines + markdown(images).split("\n")
        return lines


def _control_ids(doc) -> dict:
    """欄の名前を、書類全体の順で決める（Spots の決め方と同じ）。id(欄) → 名前。"""
    out: dict[int, str] = {}
    used: set[str] = set()
    for k, c in enumerate(getattr(doc, "controls", []), 1):
        name = c.name or f"ctl{k}"
        base = name
        j = 2
        while name in used:
            name = f"{base}#{j}"
            j += 1
        used.add(name)
        out[id(c)] = name
    return out


class SeqLazy(BlocksLazy):
    """単位ごとに読む。派生は units（数）・parse_unit(u, state)・state_after(doc, state)・window_kind を作る。"""

    window_fn = None
    md_fn = None

    def index_units(self) -> None:
        """1 単位ずつ読み、段落・表の数と、続きの状態を控える。読んだ単位は手放す。"""
        from ..docx.document import iter_paras
        from ..docx.window import _tables

        state = self.first_state()
        self.states = []
        self.story_rank: dict[str, int] = {}
        bodies = []
        stories = []
        chars = 0
        self.names: list[list[str]] = []
        used: set[str] = set()
        total_controls = 0
        n = self.unit_count()
        self.object_units: list[int] = []
        self.tracked = 0  # 変更履歴で削除された字の残る段落の数
        colors: Counter = Counter()

        def gone(blocks) -> int:
            count_colors(iter_paras(blocks), colors)
            return sum(1 for p in iter_paras(blocks) if hasattr(p, "deleted_text") and p.deleted_text().strip())

        for u in range(n):
            limits.progress("段落の位置を調べています", u, n)
            self.states.append(state)
            d = self.parse_unit(u, state)
            pa = sum(1 for _ in iter_paras(d.body))
            ta = sum(1 for _ in _tables(d.body))
            bodies.append((u, pa, ta))
            chars += sum(len(p.text) for p in iter_paras(d.body))
            self.tracked += gone(d.body)
            for j, (k, part, blocks) in enumerate(d.stories):
                self.story_rank.setdefault(k, len(self.story_rank))
                ps = sum(1 for _ in iter_paras(blocks))
                ts = sum(1 for _ in _tables(blocks))
                stories.append((k, u, j, part, ps, ts))
                chars += sum(len(p.text) for p in iter_paras(blocks))
                self.tracked += gone(blocks)
            names = []
            for c in self._unit_controls(d, u):
                total_controls += 1
                name = c.name or f"ctl{total_controls}"
                base = name
                jj = 2
                while name in used:
                    name = f"{base}#{jj}"
                    jj += 1
                used.add(name)
                names.append(name)
            self.names.append(names)
            if self.unit_has_objects(d):
                self.object_units.append(u)
            state = self.state_after(d, state)
            del d
        self.chunks = []
        para = table = 0
        for u, pa, ta in bodies:
            self.chunks.append({"unit": u, "which": "body", "para": para, "table": table, "kind": None})
            para += pa
            table += ta
        self.body_paragraphs = para
        for k, u, j, part, ps, ts in sorted(stories, key=lambda s: (self.story_key(s[0]), s[1], s[2])):
            self.chunks.append({"unit": u, "which": j, "para": para, "table": table, "kind": k, "part": part})
            para += ps
            table += ts
        self.paragraphs, self.tables = para, table
        self.normal_font_color = most_common_color(colors)  # 文書のふつうの字の色（ふつうの道と同じ数え方）
        self.est_chars = int(chars * 1.25) + para * 12 + len(used) * 60

    def first_state(self) -> dict:
        return {}

    def story_key(self, kind: str) -> int:
        """ヘッダー・脚注などの、全体を読んだときの順（種類の順）。"""
        return self.story_rank[kind]

    def unit_count(self) -> int:
        raise NotImplementedError

    def parse_unit(self, u: int, state: dict):
        raise NotImplementedError

    def state_after(self, doc, state: dict) -> dict:
        return state

    def unit_has_objects(self, doc) -> bool:
        return bool(getattr(doc, "embeds", None) or getattr(doc, "charts", None))

    def chunk_doc(self, i: int, real: tuple = ()):
        ch = self.chunks[i]
        d = self.parse_unit(ch["unit"], self.states[ch["unit"]])
        if ch["which"] == "body":
            blocks = [("body", getattr(d, "part", "body"), d.body)]
        else:
            blocks = [d.stories[ch["which"]]]
        ids = {id(c): nm for c, nm in zip(self._unit_controls(d, ch["unit"]), self.names[ch["unit"]])}
        return SubDoc(d, blocks, ch["para"], ch["table"], ids, orphans=i == 0)

    @staticmethod
    def _unit_controls(d, u: int) -> list:
        """この単位の欄。段落に結び付かない欄は、どの単位にも出るので、はじめの単位にだけ数える。"""
        return [c for c in getattr(d, "controls", []) if c.para is not None or u == 0]

    def build_window(self, doc) -> dict:
        return self.window_fn(doc)

    def extras_doc(self):
        """窓の終わりの節の材料：埋め込み・グラフのある単位を読み、一覧をつなぐ。"""
        last = None
        embeds, pictures, charts, unread = [], [], [], []
        for u in self.object_units or [0]:
            d = self.parse_unit(u, self.states[u])
            last = d
            embeds += [e for e in getattr(d, "embeds", []) if not e.get("stub")]
            charts += [e for e in getattr(d, "charts", []) if not e.get("stub")]
            pictures += [p for p in getattr(d, "pictures", []) if not p.get("stub")]
            unread += list(getattr(d, "unread", []))
        sub = SubDoc(last, [], 0, 0, {}, objects=True)
        sub.controls = []
        sub.embeds, sub.charts, sub.pictures = embeds, charts, []
        sub.unread = list(dict.fromkeys(unread))
        return sub

    def _make_extras(self) -> list[str]:
        win = self.window_fn(self.extras_doc())
        lines = _window_extras(win, self.md_fn)
        images = self.list_images()
        if images:
            from ..imaging.docimages import markdown

            lines = lines + markdown(images).split("\n")
        return lines

    def list_images(self) -> list[dict]:
        return []
