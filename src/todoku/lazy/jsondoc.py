"""大きな JSON（1 つの文書）を分けて読む：字句を流し読みして、値の位置をたどる。メモリは 1 回に読む値の分だけ。

- 窓：値の一覧・表は、datafile の窓と同じ書き方（`/a/b` ＝ 値。同じ形のオブジェクトの配列は表）。
- 読む：path（JSON Pointer）でその値を読む。続きの印（cursor）は、いま読んでいる入れ物の位置の並び
- とても大きな入れ物の子は、1 段ずつ降りて読む。とても長い字の値は、字の範囲で分けて印を付ける
- 字句の確かめ：開くときに全体を 1 度たどる（かっこの対応）。読む値は、窓と同じ読み手（datafile/jsonfile）で厳密に読む
- 書く：いままでの api.fill（全体を読む）。文字コードは UTF-8 だけ（ほかは全体を読む道へ）
"""

from __future__ import annotations

import json
import re

from .. import limits
from ..datafile.jsonfile import JsonError, _Parser, esc_token, unesc_token
from ..datafile.window import _md, _records, _walk, literal_text
from .base import LazyDoc, Page, RangeError, comma, cut_note
from .textstream import Source, detect

STEP = 1 << 20
_STRUCT = re.compile(rb'["{}\[\]]')
_STR_END = re.compile(rb'["\\]')
_WS = b" \t\r\n"


class _Win:
    """読む窓（STEP ずつ控える）。"""

    def __init__(self, src: Source):
        self.src = src
        self.a = 0
        self.b = b""

    def load(self, pos: int) -> None:
        self.a = pos
        self.b = self.src.read(pos, pos + STEP)

    def byte(self, pos: int) -> int:
        if not self.a <= pos < self.a + len(self.b):
            self.load(pos)
            if not self.b:
                return -1
        return self.b[pos - self.a]


class JsonLazy(LazyDoc):
    kind = "json"
    streams = False
    range_keys = ("path",)

    def __init__(self, path: str, name: str, encoding: str | None = None, source=None):
        self.path, self.name = path, name
        self.src = source if source is not None else Source(path)  # source：書庫の中のファイルの、取り出した中身
        self.size = self.src.size
        self.dec = detect(self.src, encoding)
        if self.dec.encoding.lower().replace("-", "") not in ("utf8", "utf_8", "ascii"):
            raise ValueError("UTF-8 でない JSON は、全体を読む道で読みます")
        self.start = len(self.dec.bom)
        self.win = _Win(self.src)
        self.values = 0
        self.root_at, self.end_at = self._check_all()
        self.est_chars = int(self.size * 1.2) + self.values * 8

    def close(self) -> None:
        self.src.close()

    # ------------------------------------------------------------------ 字句
    def _skip_ws(self, pos: int) -> int:
        w = self.win
        while True:
            c = w.byte(pos)
            if c < 0 or c not in _WS:
                return pos
            pos += 1

    def _check_all(self) -> tuple[int, int]:
        """全体を 1 度たどる：かっこの対応と、値の数の目安。"""
        pos = self.start
        depth = 0
        in_str = False
        total = self.size
        first = None
        last = None
        count = 0
        stack: list[int] = []
        while pos < total:
            c = self.src.read(pos, pos + STEP)
            if not c:
                break
            i = 0
            n = len(c)
            while i < n:
                if in_str:
                    m = _STR_END.search(c, i)
                    if not m:
                        i = n
                        break
                    if c[m.start()] == 0x5C:
                        i = m.start() + 2
                        continue
                    in_str = False
                    i = m.end()
                    continue
                m = _STRUCT.search(c, i)
                if not m:
                    count += c.count(b",", i)
                    break
                count += c.count(b",", i, m.start())
                ch = c[m.start()]
                i = m.end()
                if first is None:
                    first = pos + m.start()
                if ch == 0x22:
                    in_str = True
                elif ch in b"{[":
                    stack.append(ch)
                else:
                    if not stack or (stack.pop() == 0x7B) != (ch == 0x7D):
                        raise ValueError("かっこの対応が合いません")
                    last = pos + m.start()
            pos += n
            limits.progress("JSON の字句を調べています", pos, total)
        if stack or in_str:
            raise ValueError("JSON が途中で終わっています")
        self.values = count + 1
        root = self._skip_ws(self.start)
        del depth, first, last
        end = total
        return root, end

    def value_end(self, pos: int) -> int:
        """pos で始まる値の終わりの次。"""
        w = self.win
        c = w.byte(pos)
        if c == 0x22:
            return self._string_end(pos)
        if c in (0x7B, 0x5B):
            depth = 0
            in_str = False
            p = pos
            while True:
                if not w.a <= p < w.a + len(w.b):
                    w.load(p)
                    if not w.b:
                        raise ValueError("JSON が途中で終わっています")
                b = w.b
                i = p - w.a
                n = len(b)
                while i < n:
                    if in_str:
                        m = _STR_END.search(b, i)
                        if not m:
                            i = n
                            break
                        if b[m.start()] == 0x5C:
                            i = m.start() + 2
                            continue
                        in_str = False
                        i = m.end()
                        continue
                    m = _STRUCT.search(b, i)
                    if not m:
                        i = n
                        break
                    ch = b[m.start()]
                    i = m.end()
                    if ch == 0x22:
                        in_str = True
                    elif ch in b"{[":
                        depth += 1
                    else:
                        depth -= 1
                        if depth == 0:
                            return w.a + i
                p = w.a + i
        # 数・true・false・null
        p = pos
        while True:
            c = w.byte(p)
            if c < 0 or c in b",]} \t\r\n":
                return p
            p += 1

    def _string_end(self, pos: int) -> int:
        w = self.win
        p = pos + 1
        while True:
            if not w.a <= p < w.a + len(w.b):
                w.load(p)
                if not w.b:
                    raise ValueError("JSON の字が途中で終わっています")
            m = _STR_END.search(w.b, p - w.a)
            if not m:
                p = w.a + len(w.b)
                continue
            if w.b[m.start()] == 0x5C:
                p = w.a + m.start() + 2
                continue
            return w.a + m.end()

    def children(self, at: int, kind: int, from_pos: int | None = None, from_no: int = 0):
        """入れ物の子を順に：(番号, 名前か None, 値の始まり, 値の終わり)。from_pos：途中の子の始まり（名前の始まり）から。"""
        pos = from_pos if from_pos is not None else self._skip_ws(at + 1)
        no = from_no
        while True:
            pos = self._skip_ws(pos)
            c = self.win.byte(pos)
            if c in (0x7D, 0x5D):
                return
            if c < 0:
                raise ValueError("JSON が途中で終わっています")
            name = None
            start = pos
            if kind == 0x7B:
                e = self._string_end(pos)
                name = json.loads(self.src.read(pos, e).decode("utf-8"))
                pos = self._skip_ws(e)
                if self.win.byte(pos) != 0x3A:
                    raise ValueError("名前の後ろに : がありません")
                pos = self._skip_ws(pos + 1)
            e = self.value_end(pos)
            yield no, name, start, pos, e
            no += 1
            pos = self._skip_ws(e)
            c = self.win.byte(pos)
            if c == 0x2C:
                pos += 1
            elif c not in (0x7D, 0x5D):
                raise ValueError("値の後ろに , がありません")

    def locate(self, pointer: str) -> tuple[int, int, str]:
        """JSON Pointer → (値の始まり, 終わり, 正しい書き方のパス)。"""
        at = self.root_at
        path = ""
        p = str(pointer or "")
        if p == "/":
            p = ""
        if p and not p.startswith("/"):
            raise RangeError(f"path「{pointer}」は JSON Pointer ではありません（/ で始めます。例 /applicant/name）")
        tokens = [unesc_token(x) for x in p.split("/")[1:]] if p else []
        end = None
        for t in tokens:
            c = self.win.byte(at)
            if c not in (0x7B, 0x5B):
                raise RangeError(f"path「{pointer}」：{path or '/'} は入れ物ではありません")
            for no, name, _s, a, b in self.children(at, c):
                limits.progress(f"{path or '/'} の中を探しています", no, 0)
                if (c == 0x7B and name == t) or (c == 0x5B and t.isdigit() and no == int(t)):
                    at, end = a, b
                    path += "/" + esc_token(t)
                    break
            else:
                raise RangeError(f"path「{pointer}」：{path or '/'} の中に「{t}」がありません")
        if end is None:
            end = self.value_end(at)
        return at, end, path

    # ------------------------------------------------------------------ 値を窓の字に
    def _parse(self, a: int, b: int, path: str):
        text = self.src.read(a, b).decode("utf-8")
        p = _Parser(text)
        try:
            node = p.value(path)
        except JsonError as e:
            raise ValueError(f"{path or '/'}：{e}") from None
        return node, text

    def _render(self, items: list[tuple[str, int, int]], array_parent: bool) -> list[str]:
        """子の値たちの窓の行。同じ形のオブジェクトの配列は表にする。"""
        if not items:
            return []
        parsed = [self._parse(a, b, path) for path, a, b in items]
        nodes = [(n.path, n) for n, _t in parsed]
        keys = _records(nodes) if array_parent else None
        if keys is not None:
            rows = []
            for n, t in parsed:
                rows.append((n.path, [literal_text(m[3], t) for m in n.members]))
            return ["| 位置 | " + " | ".join(_md(k) for k in keys) + " |", "|---" * (len(keys) + 1) + "|"] + \
                ["| " + rid + " | " + " | ".join(_md(v) for v in vals) + " |" for rid, vals in rows]
        out: list[str] = []
        for n, t in parsed:
            blocks: list[dict] = []
            f = _Text(t)
            _walk(f, n, "", blocks)
            for bl in blocks:
                if bl["kind"] == "list":
                    out += [f"- `{i}` ＝ {v}" for i, v in bl["items"]]
                else:
                    out += [f"### {bl['title']}", "| 位置 | " + " | ".join(_md(k) for k in bl["keys"]) + " |", "|---" * (len(bl["keys"]) + 1) + "|"]
                    out += ["| " + rid + " | " + " | ".join(_md(v) for v in vals) + " |" for rid, vals in bl["rows"]]
        return out

    # ------------------------------------------------------------------ 読む
    def read(self, rng: dict, max_chars: int) -> Page:
        cur = rng.get("cursor")
        stop_end = None
        frames: list[dict]
        if cur:
            try:
                st = json.loads(str(cur)[5:]) if str(cur).startswith("json:") else None
            except ValueError:
                st = None
            if not st:
                raise RangeError(f"cursor「{cur}」はこの書類の続きの印ではありません（前の結果の next をそのまま渡してください）")
            frames = st["f"]
            stop_end = st.get("e")
        else:
            a, b, path = self.locate(rng.get("path") or "")
            if rng.get("path"):
                stop_end = b
            frames = [{"at": a, "path": path, "pos": None, "no": 0, "s": 0, "b": b}]
        body: list[str] = []
        size = 0
        limit_bytes = max_chars  # 1 つの値を読む大きさの目安（字はバイトを超えない）
        while frames:
            fr = frames[-1]
            at, path = fr["at"], fr["path"]
            kind = self.win.byte(at)
            if kind not in (0x7B, 0x5B) or fr.get("leaf"):
                # 入れ物でない値（とても長い字もここ）
                done, lines, nxt = self._leaf(fr, max_chars - size)
                body += lines
                size += sum(len(x) + 1 for x in lines)
                if not done:
                    fr["s"] = nxt
                    return self._page(body, frames, stop_end, "この値の残り")
                frames.pop()
                continue
            if fr["b"] - at <= limit_bytes and fr["no"] == 0 and fr["pos"] is None:
                # 入れ物の全体が 1 回に入る
                lines = self._render([(path, at, fr["b"])], False)
                n = sum(len(x) + 1 for x in lines)
                if body and size + n > max_chars:
                    return self._page(body, frames, stop_end, "残りの値")
                body += lines
                size += n
                frames.pop()
                continue
            # 子を順に。1 回に入る分はまとめて読む
            batch: list[tuple[str, int, int]] = []
            batch_bytes = 0
            pushed = False
            finished = True
            for no, name, start, a, b in self.children(at, kind, fr["pos"], fr["no"]):
                sub = path + "/" + (esc_token(name) if name is not None else str(no))
                span = b - a
                if span > limit_bytes:
                    # 大きな子：ここまでを出して、子へ降りる
                    lines = self._render(batch, kind == 0x5B)
                    body += lines
                    size += sum(len(x) + 1 for x in lines)
                    batch = []
                    fr["pos"], fr["no"] = b, no + 1
                    fr["pos"] = self._after(b)
                    frames.append({"at": a, "path": sub, "pos": None, "no": 0, "s": 0, "b": b, "leaf": self.win.byte(a) not in (0x7B, 0x5B)})
                    pushed = True
                    finished = False
                    break
                if batch and size + batch_bytes + span > max_chars:
                    fr["pos"], fr["no"] = start, no
                    finished = False
                    break
                batch.append((sub, a, b))
                batch_bytes += span
            lines = self._render(batch, kind == 0x5B)
            body += lines
            size += sum(len(x) + 1 for x in lines)
            if pushed:
                if size >= max_chars:
                    return self._page(body, frames, stop_end, "残りの値")
                continue
            if not finished:
                return self._page(body, frames, stop_end, "残りの値")
            frames.pop()
            if size >= max_chars and frames:
                return self._page(body, frames, stop_end, "残りの値")
        what = f"{rng['path']} の値" if rng.get("path") else "JSON の値（全 " + comma(self.size) + " バイト）"
        return Page("\n".join(body), f"{what}を終わりまで", None, "", body)

    def _after(self, b: int) -> int:
        pos = self._skip_ws(b)
        if self.win.byte(pos) == 0x2C:
            pos += 1
        return pos

    def _page(self, body: list[str], frames: list[dict], stop_end, what: str) -> Page:
        st = {"f": frames}
        if stop_end is not None:
            st["e"] = stop_end
        fr = frames[-1]
        at = fr.get("pos") or fr["at"]
        return Page("\n".join(body), f"JSON の値（{at:,} バイト目の手前まで／全 {comma(self.size)} バイト）",
                    {"cursor": "json:" + json.dumps(st, ensure_ascii=False, separators=(",", ":"))}, f"{what}（{comma(max(self.size - at, 0))} バイト）", body)

    def _leaf(self, fr: dict, room: int) -> tuple[bool, list[str], int]:
        """入れ物でない値。字の値が大きいときは、字の範囲で分ける。"""
        a, b = fr["at"], fr["b"]
        size = b - a
        if size <= max(room, 1 << 16) and not fr.get("s"):
            return True, self._render([(fr["path"], a, b)], False), 0
        node, _t = self._parse(a, b, fr["path"])
        text = literal_text(node, _t)
        s = fr.get("s", 0)
        step = max(room, 2000)
        e = min(s + step, len(text))
        head = f"- `{fr['path'] or '/'}` ＝ " if s == 0 else ""
        line = head + text[s:e] + ("" if e >= len(text) else "\n" + cut_note(s, e, len(text)))
        return e >= len(text), [line], e

    # ------------------------------------------------------------------ 欄
    def fields(self, rng: dict, only_empty: bool, limit: int) -> dict:
        from ..api import _field

        cur = rng.get("cursor")
        if cur:
            m = re.fullmatch(r"fieldsat:(\d+):(\d+):(.*)", str(cur), re.S)
            if not m:
                raise RangeError(f"cursor「{cur}」はこの書類の続きの印ではありません")
            pos, no, base = int(m[1]), int(m[2]), m[3]
        else:
            pos, no, base = None, 0, rng.get("path") or ""
        a, b, path = self.locate(base)
        kind = self.win.byte(a)
        out: list[dict] = []
        nxt = None

        def add(node, text, doc_id=None):
            stack = [node]
            while stack:
                n = stack.pop()
                if n.is_leaf:
                    lit = literal_text(n, text)
                    i = n.path or "/"
                    k = {"true": "checkbox", "false": "checkbox", "number": "number"}.get(n.kind, "text")
                    d = _field({"id": i, "kind": k, "label": (unesc_token(i.rsplit("/", 1)[-1]) if "/" in i else i) or i, "value": n.value,
                                "json_type": n.kind, "provenance": "deterministic", "context": f"{i} ＝ {lit}"})
                    if not only_empty or d.get("value") in (None, "", False):
                        out.append(d)
                elif n.kind == "object":
                    stack.extend(reversed([m_[3] for m_ in n.members]))
                else:
                    stack.extend(reversed(n.items))

        if kind in (0x7B, 0x5B):
            for k_no, name, start, va, vb in self.children(a, kind, pos, no):
                if len(out) >= limit:
                    nxt = {"cursor": f"fieldsat:{start}:{k_no}:{base}"}
                    break
                sub = path + "/" + (esc_token(name) if name is not None else str(k_no))
                node, text = self._parse(va, vb, sub)
                add(node, text)
        else:
            node, text = self._parse(a, b, path)
            add(node, text)
        return {"fields": out, "next": nxt, "note": "大きな入れ物は path（JSON Pointer）で中を指せます"}

    # ------------------------------------------------------------------ 概要
    def overview(self) -> dict:
        return {"bytes": self.size, "window_chars": self.est_chars, "values_about": self.values, "top_level": self._top(), "ranges": "path（JSON Pointer）か cursor"}

    def _top(self) -> list[dict]:
        out = []
        kind = self.win.byte(self.root_at)
        if kind not in (0x7B, 0x5B):
            return out
        for no, name, _s, a, b in self.children(self.root_at, kind):
            if no >= 60:
                out.append({"more": True})
                break
            out.append({"path": "/" + (esc_token(name) if name is not None else str(no)), "bytes": b - a})
        return out

    def overview_text(self) -> str:
        top = self._top()
        lines = [f"# {self.name}", "", f"- JSON（{self.size:,} バイト。値はおよそ {self.values:,} 個）。全体を一度には出しません。read_window に path（JSON Pointer。例 /items/0）を渡して、その値を読みます"
                 "（path なしで cursor をたどれば、すべての値を読めます。max_chars で 1 回の大きさを変えられます。窓の読み手が上限で止まるときは raw: true で字のまま読めます）。"
                 "list_fields も path で、その中の値を出します。", "", "## いちばん上の値", ""]
        lines += [f"- `{d['path']}`（{d['bytes']:,} バイト）" if "path" in d else "- ほか" for d in top]
        return "\n".join(lines) + "\n"

    # ------------------------------------------------------------------ 段階・検査
    def grade(self):
        from ..grade import Reason, _final

        g = _final([Reason("A", "大きな JSON", "かっこの対応は全体で確かめました")], [Reason("B", "値を範囲ごとに読む", "大きいので、値は path か cursor で読みます")])
        g.deferred = ["未確認：全部の値の字の並びの厳密な確かめ（audit_document で全部を読みます）"]
        return g

    def audit_begin(self) -> dict:
        return {"cursor": None, "pages": 0, "problems": [], "done": False, "checks": []}

    def audit_step(self, state: dict) -> dict:
        rng = {"cursor": state["cursor"]} if state["cursor"] else {}
        pg = self.read(rng, limits.PAGE_CHARS)
        state["pages"] += 1
        limits.progress("値を順に読んでいます", state["pages"], 0)
        if pg.next is None:
            state["checks"] = [{"name": "読み落とし（値）", "ok": True, "agent": False, "items": [],
                                "detail": f"全部の値を窓の読み手で読みました（{state['pages']:,} 回に分けて）"}]
            state["done"] = True
        else:
            state["cursor"] = pg.next["cursor"]
        return state


class _Text:
    """_walk が使う、JsonFile の字の部分だけ。"""

    def __init__(self, text: str):
        self.text = text
