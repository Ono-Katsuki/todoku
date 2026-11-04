"""窓を作ってから分ける書類（Word・一太郎・RTF・OpenDocument・PowerPoint・한글・HTML・.xls・JSON・ZIP）。

窓の全体を作り（api.window の控えを使う）、測った字の数で判断する。読む範囲は、窓の Markdown の字の位置で決まる（part と同じ中身を、
どの大きさでも分けて読める）。長い 1 行も、字の位置で分けて読める。書き込みは、いままでの api.fill のまま。
Word（.docx・.doc）は、段落の範囲（paragraphs）でも読める。
"""

from __future__ import annotations

import re

from .. import limits
from .base import LazyDoc, Page, RangeError, comma, parse_range


class BuiltDoc(LazyDoc):
    streams = False

    def __init__(self, path: str, password: str | None, encoding: str | None):
        from .. import api

        self.path = path
        self.password = password
        self.encoding = encoding
        self.k, self.doc, self.parts = api.window(path, password, encoding)
        self.kind = self.k
        self.est_chars = sum(len(t) for _n, t in self.parts)
        self._word = None

    def whole(self) -> str:
        """窓の中身：目次のある窓は、目次のほかの部分をつないだもの。"""
        if len(self.parts) == 1:
            return self.parts[0][1]
        return "\n".join(t for _n, t in self.parts[1:])

    def overview(self) -> dict:
        out = {"bytes": (self.doc.get("source") or {}).get("bytes"), "window_chars": self.est_chars, "window_parts": len(self.parts),
               "ranges": "paragraphs（段落の範囲）か cursor" if self.k in ("docx", "doc") else "cursor（字の位置）"}
        if self.k in ("docx", "doc"):
            out["fields_total"] = len(self.doc.get("fields") or [])
            out["paragraphs"] = self._paragraph_count()
        return out

    def overview_text(self) -> str:
        head = self.parts[0][1]
        if len(self.parts) == 1:
            lines = head.split("\n")[:12]
        else:
            lines = head.split("\n")
        how = (f"段落の範囲（paragraphs: \"1-200\"。全 {self._paragraph_count():,} 段落）か、" if self.k in ("docx", "doc") else "")
        lines = lines + ["", f"- 窓の全体は {self.est_chars:,} 字あります。read_window に {how}part（部分の番号）か cursor を渡して読みます"
                         "（cursor をたどれば、すべてを読めます。max_chars で 1 回の大きさを変えられます）。"
                         "list_fields に paragraphs を渡すと、その範囲の欄を出します。"]
        return "\n".join(lines)

    # ------------------------------------------------------------------ 読む
    def read(self, rng: dict, max_chars: int) -> Page:
        if rng.get("paragraphs") is not None or str(rng.get("cursor") or "").startswith("paragraphs:"):
            return self._read_paragraphs(rng, max_chars)
        if rng.get("sheet") is not None or rng.get("rows") is not None:
            return self._read_sheet(rng, max_chars)
        if rng.get("part") is not None:
            n = int(rng["part"])
            if not 1 <= n <= len(self.parts):
                raise RangeError(f"part は 1〜{len(self.parts)} です")
            title, text = self.parts[n - 1]
            return Page(text, f"部分 {n}/{len(self.parts)}：{title}", {"part": n + 1} if n < len(self.parts) else None,
                        f"{len(self.parts) - n} の部分", text.split("\n"))
        cur = rng.get("cursor")
        at = 0
        if cur:
            m = re.fullmatch(r"chars:(\d+)", str(cur))
            if not m:
                raise RangeError(f"cursor「{cur}」はこの書類の続きの印ではありません（前の結果の next をそのまま渡してください）")
            at = int(m[1])
        text = self.whole()
        end = min(at + max_chars, len(text))
        if end < len(text):
            nl = text.rfind("\n", at, end)
            if nl > at:
                end = nl + 1
        piece = text[at:end]
        cov = f"{at + 1:,}〜{end:,} 字目（全 {len(text):,} 字）"
        if end >= len(text):
            return Page(piece, cov, None, "", [piece])
        return Page(piece, cov, {"cursor": f"chars:{end}"}, f"{comma(len(text) - end)} 字", [piece])

    # ------------------------------------------------------------------ シート（表だけの形式）
    def _sheet_text(self, name: str, rows) -> str:
        """窓の中の 1 つのシートの節。読んでいない部分の注のうち、そのシートの名前のあるものを前に添える。"""
        text = self.whole()
        heads = list(re.finditer(r"^## シート「(.*)」(?:（非表示[^）]*）)?$", text, re.M))  # 非表示のシートも、名前で読める
        if not heads:
            raise RangeError("この書類はシートの範囲では読めません（part か cursor で読んでください）")
        at = next((i for i, m in enumerate(heads) if m[1] == name), None)
        if at is None:
            raise RangeError(f"シート「{name}」はありません（ある名前：{'・'.join(m[1] for m in heads[:30])}）")
        body = text[heads[at].start():heads[at + 1].start() if at + 1 < len(heads) else len(text)]
        if rows is not None:
            a, b = parse_range(rows, "行")
            kept = []
            for line in body.split("\n"):
                m = re.match(r"- (\d+): ", line)
                if m is None or a <= int(m[1]) <= b:
                    kept.append(line)
            body = "\n".join(kept)
        notes = [ln for ln in text[:heads[0].start()].split("\n") if ln.startswith("- **読んでいない部分**") and f"「{name}」" in ln]
        return "\n".join(notes + [body]) if notes else body

    def _read_sheet(self, rng: dict, max_chars: int) -> Page:
        name = rng.get("sheet")
        if name is None:
            raise RangeError("rows は sheet（シートの名前）といっしょに渡してください")
        text = self._sheet_text(str(name), rng.get("rows"))
        at = 0
        cur = rng.get("cursor")
        if cur:
            m = re.fullmatch(r"chars:(\d+)", str(cur))
            if not m:
                raise RangeError(f"cursor「{cur}」はこの書類の続きの印ではありません（前の結果の next をそのまま渡してください）")
            at = int(m[1])
        end = min(at + max_chars, len(text))
        if end < len(text):
            nl = text.rfind("\n", at, end)
            if nl > at:
                end = nl + 1
        piece = text[at:end]
        cov = f"シート「{name}」の {at + 1:,}〜{end:,} 字目（全 {len(text):,} 字）"
        if end >= len(text):
            return Page(piece, cov, None, "", [piece])
        nxt = {"sheet": name, "cursor": f"chars:{end}"}
        if rng.get("rows") is not None:
            nxt["rows"] = rng["rows"]
        return Page(piece, cov, nxt, f"{comma(len(text) - end)} 字", [piece])

    # ------------------------------------------------------------------ Word の段落
    def _word_units(self):
        if self._word is not None:
            return self._word
        from .. import api
        from ..docx.document import Para, Table
        from ..docx.window import Spots, iter_paras

        blob, name, k, _lock = api._read(self.path, self.password)
        d = api._word(blob, name, k)
        sp = Spots(d)
        by_para: dict[int, list[dict]] = {}
        for it in sp.items.values():
            p = it.get("para")
            if p is not None:
                by_para.setdefault(id(p), []).append(it)
        units = []
        for kind, _part, blocks in d.all_blocks():
            for b in blocks:
                ps = list(iter_paras([b]))
                nums = [sp.para_index[id(q)] for q in ps if id(q) in sp.para_index]
                if nums:
                    units.append((min(nums), max(nums), kind, b))
        units.sort(key=lambda u: u[0])
        self._word = (sp, by_para, units, d, (Para, Table))
        return self._word

    def _paragraph_count(self) -> int:
        try:
            sp = self._word_units()[0]
        except Exception:  # noqa: BLE001  Word 6.0/95 など：段落の数が分からなくても、ほかの読み方で全部読める
            return 0
        return len(sp.para_index)

    def _read_paragraphs(self, rng: dict, max_chars: int) -> Page:
        from ..docx.window import _render_blocks

        sp, by_para, units, _d, _t = self._word_units()
        total = len(sp.para_index)
        cur = rng.get("cursor")
        if cur:
            m = re.fullmatch(r"paragraphs:(\d+):(\d+)", str(cur))
            if not m:
                raise RangeError(f"cursor「{cur}」はこの書類の続きの印ではありません")
            a, b = int(m[1]), int(m[2])
        else:
            a, b = parse_range(rng.get("paragraphs"), "段落")
            a = 1 if a is None else a
            b = total if b is None else min(b, total)
        body: list[str] = []
        size = 0
        stop = None
        for lo, hi, _kind, block in units:
            if hi < a or lo > b:
                continue
            lines: list[str] = []
            plain: list[str] = []
            _render_blocks([block], sp, by_para, lines, plain)
            text = "\n".join(f"〔段落 {lo}" + (f"〜{hi}" if hi != lo else "") + f"〕{ln}" for ln in lines)
            if body and size + len(text) > max_chars:
                stop = lo
                break
            if text:
                body.append(text)
                size += len(text) + 1
        cov = f"段落 {a}〜{(stop - 1) if stop else b}（全 {comma(total)} 段落）"
        if stop is None:
            return Page("\n".join(body), cov, None, "", body)
        return Page("\n".join(body), cov, {"cursor": f"paragraphs:{stop}:{b}"}, f"{comma(b - stop + 1)} 段落", body)

    # ------------------------------------------------------------------ 欄
    def fields(self, rng: dict, only_empty: bool, limit: int) -> dict:
        from .. import api

        fl = api.fields(self.path, self.password, self.encoding)
        if rng.get("paragraphs") is not None:
            a, b = parse_range(rng["paragraphs"], "段落")
            a = a or 1
            b = b or 10 ** 12

            def inside(f: dict) -> bool:
                m = re.match(r"P(\d+)\.", str(f.get("id")))
                return bool(m) and a <= int(m[1]) <= b
            fl = [f for f in fl if inside(f)]
        if only_empty:
            fl = [f for f in fl if f.get("value") in (None, "", False)]
        start = int(str(rng.get("cursor") or "fields:0").split(":")[-1])
        got = fl[start:start + limit]
        res = {"fields": got, "total_fields": len(fl), "next": None}
        if start + limit < len(fl):
            res["next"] = {"cursor": f"fields:{start + limit}"}
            res["remaining"] = f"{comma(len(fl) - start - limit)} か所"
        return res

    # ------------------------------------------------------------------ 検査：いままでの検査を、1 区切りで動かす
    def audit_step(self, state: dict) -> dict:
        from .. import api

        limits.progress("書類の全体を検査しています")
        au = api.audit_whole(self.path, self.password, self.encoding)
        state["checks"] = au.get("checks", [])
        state["full"] = au
        state["done"] = True
        return state
