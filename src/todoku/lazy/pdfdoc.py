"""大きな PDF を分けて読み書きする：ページの木と xref だけを読み、頼まれたページだけを配置する。

- 窓の字の数の見積もり：等間隔の数ページを実際に配置して測り、ページの数を掛ける（5 ページ以下は全部を作っても同じなので測らない）
- 読む：pages の範囲。ページの窓は xlsx でなく pdf/window.page_window そのもの。XFA のフォームの節は、ページのあとに続けて読める
- 書く：書く所のページだけの窓で、いままでの api._fill を通す（pdf.window.page_scope）。増分更新なので、元のバイトは 1 バイトも変わらない
- 段階：費用の小さな確かめ（ページの木・フォーム・XFA・暗号・フォントの字の対応）は全部のページで行う。字の読み落としは全部のページを配置しないと
  分からないので「未確認」と示し、検査（audit_document の長い仕事）で確かめる
"""

from __future__ import annotations

import re
from pathlib import Path

from .. import limits
from .base import LazyDoc, Page, RangeError, comma, parse_range

SAMPLE_PAGES = 5


class _OnePage:
    """1 ページだけの書類に見せる（検査の部品を、ページごとに使うため）。"""

    def __init__(self, doc, page):
        self._doc = doc
        self._page = page

    def __getattr__(self, name):
        return getattr(self._doc, name)

    def pages(self):
        return [self._page]


class PdfLazy(LazyDoc):
    kind = "pdf"
    streams = False

    def __init__(self, path: str, name: str, password: str | None):
        from .. import api
        from ..pdf.form import read_fields
        from ..pdf.window import _xfa_window, build, has_xfa

        self.path = path
        self.name = name
        self.password = password
        blob, _n, k, lock = api._read(path, password)
        self.blob = blob
        self.lock = lock
        self.doc = api._pdf(blob, name, lock)
        self.page_list = self.doc.pages()
        self.fields_all = read_fields(self.doc)
        self.xw = _xfa_window(self.doc) if has_xfa(self.doc) else None
        self.head_doc = build(self.doc, pages=())
        self._lay = {}
        self.est_chars = self._estimate()

    @property
    def layout_on(self) -> bool:
        return self.xw is None or not self.xw.get("readable")

    def _estimate(self) -> int:
        n = len(self.page_list)
        if n <= SAMPLE_PAGES:
            return 0  # 測っても、全部を作るのと同じ（いままでどおり）
        picks = sorted({round(i * (n - 1) / (SAMPLE_PAGES - 1)) + 1 for i in range(SAMPLE_PAGES)})
        from ..pdf.window import _md_page

        total = 0
        for p in picks:
            total += sum(len(x) + 1 for x in _md_page(self.page(p)))
        return total * n // len(picks)

    def page(self, number: int) -> dict:
        from ..pdf.window import page_window, with_label_keys

        got = self._lay.get(number)
        if got is not None:
            return got
        pg = page_window(self.doc, self.page_list[number - 1], number, self.fields_all, layout=self.layout_on)
        with_label_keys(pg["fields"])
        if len(self._lay) >= 16:
            self._lay.pop(next(iter(self._lay)))
        self._lay[number] = pg
        return pg

    def overview(self) -> dict:
        widgets: dict[int, int] = {}
        for f in self.fields_all:
            for w in f.widgets:
                widgets[w.page + 1] = widgets.get(w.page + 1, 0) + 1
        return {"pages": len(self.page_list), "bytes": len(self.blob), "form_fields": len(self.fields_all), "xfa": bool(self.xw),
                "pages_with_form_fields": sorted(widgets)[:200], "outline": self._outline(), "ranges": "pages（ページの範囲）"}

    def _outline(self) -> list[dict]:
        """しおり（目次）：どこに何があるか。"""
        doc = self.doc
        out: list[dict] = []
        try:
            root = doc.root()
            ol = doc.resolve(root.get("Outlines"))
            if not isinstance(ol, dict):
                return out
            pages = {id(p): i + 1 for i, p in enumerate(self.page_list)}
            item = doc.resolve(ol.get("First"))
            seen = 0
            while isinstance(item, dict) and seen < 200:
                seen += 1
                title = item.get("Title")
                t = title.decode("utf-16", "replace") if isinstance(title, bytes) and title[:2] in (b"\xfe\xff", b"\xff\xfe") else (
                    title.decode("latin-1") if isinstance(title, bytes) else str(title or ""))
                d = {"title": t}
                dest = doc.resolve(item.get("Dest"))
                if isinstance(dest, list) and dest:
                    p = pages.get(id(doc.resolve(dest[0])))
                    if p:
                        d["page"] = p
                out.append(d)
                item = doc.resolve(item.get("Next"))
        except Exception:  # noqa: BLE001  しおりは案内なので、読めなくても書類は読める
            return out
        return out

    def overview_text(self) -> str:
        from ..pdf.window import _md_head

        ov = self.overview()
        doc = dict(self.head_doc)
        doc["pages"] = self.page_list  # 頭の字にはページの数だけを使う
        out = _md_head(doc)
        out += [f"- 大きいので、ここでは全部のページを配置していません（窓の全体は見積もりで {self.est_chars:,} 字）。read_window に pages: \"1-20\" のようにページの範囲を渡して読みます"
                "（cursor をたどれば、すべてのページを読めます）。list_fields も pages で、そのページの欄を出します。"]
        if ov["outline"]:
            out += ["", "## しおり", ""] + [f"- {d['title']}" + (f"（{d['page']} ページ）" if d.get("page") else "") for d in ov["outline"][:60]]
        if ov["pages_with_form_fields"]:
            out += ["", f"フォームの欄のあるページ：{ov['pages_with_form_fields'][:60]}" + ("ほか" if len(ov["pages_with_form_fields"]) > 60 else "")]
        return "\n".join(out) + "\n"

    # ------------------------------------------------------------------ 読む
    def _xfa_lines(self) -> list[str]:
        from ..pdf.window import _xfa_units

        if not self.xw:
            return []
        doc = dict(self.head_doc)
        doc["xfa"] = self.xw
        return [ln for _t, lines in _xfa_units(doc) for ln in lines]

    def read(self, rng: dict, max_chars: int) -> Page:
        from ..pdf.window import _md_page

        n = len(self.page_list)
        cur = rng.get("cursor")
        skip = 0
        if cur and str(cur).startswith("xfa:"):
            return self._read_xfa(int(str(cur)[4:]), max_chars)
        if cur:
            m = re.fullmatch(r"pages:(\d+)(?::(\d+))?(?::(\d+))?", str(cur))
            if not m:
                raise RangeError(f"cursor「{cur}」はこの書類の続きの印ではありません（前の結果の next をそのまま渡してください）")
            a, skip = int(m[1]), int(m[2] or 0)
            b = int(m[3]) if m[3] else n
            whole = not m[3]
        else:
            a, b = parse_range(rng.get("pages"), "ページ")
            whole = b is None
            a = 1 if a is None else a
            b = n if b is None else min(b, n)
        if a > max(n, 1):
            raise RangeError(f"{a} ページ目はありません（全 {n:,} ページ）")
        body: list[str] = []
        size = 0
        p = a
        stop = None
        to = "" if whole else f":{b}"
        while p <= b:
            limits.progress("ページを配置しています", p, n)
            text = "\n".join(_md_page(self.page(p)))
            if skip:
                text = text[skip:]
            if len(text) > max_chars:
                if body:
                    stop = (p, 0)
                    break
                cut = text[:max_chars]
                k = cut.rfind("\n")
                if k > max_chars // 2:
                    cut = cut[:k + 1]
                body.append(cut + f"\n〔{p} ページ目はここで分けました〕")
                stop = (p, skip + len(cut))
                break
            if body and size + len(text) > max_chars:
                stop = (p, 0)
                break
            body.append(text)
            size += len(text) + 1
            skip = 0
            p += 1
        if stop is None:
            cov = f"{a}〜{b} ページ（全 {comma(n)} ページ）"
            if whole and self.xw:
                return Page("\n\n".join(body), cov, {"cursor": "xfa:0"}, "XFA のフォームの節", body)
            return Page("\n\n".join(body), cov, None, "", body)
        cov = f"{a}〜{stop[0] - (0 if stop[1] else 1)} ページ" + (f"（{stop[0]} ページ目は途中まで）" if stop[1] else "") + f"（全 {comma(n)} ページ）"
        nxt = {"cursor": f"pages:{stop[0]}" + (f":{stop[1]}" if stop[1] or to else "") + to}
        return Page("\n\n".join(body), cov, nxt, f"{comma(b - stop[0] + 1)} ページ", body)

    def _read_xfa(self, at: int, max_chars: int) -> Page:
        lines = self._xfa_lines()
        text = "\n".join(lines)
        end = min(at + max_chars, len(text))
        piece = text[at:end]
        cov = f"XFA のフォームの節 {at + 1:,}〜{end:,} 字（全 {len(text):,} 字）"
        if end >= len(text):
            return Page(piece, cov, None, "", [piece])
        return Page(piece, cov, {"cursor": f"xfa:{end}"}, f"{comma(len(text) - end)} 字", [piece])

    # ------------------------------------------------------------------ 欄
    def fields(self, rng: dict, only_empty: bool, limit: int) -> dict:
        from ..api import _field

        n = len(self.page_list)
        cur = rng.get("cursor")
        if cur:
            m = re.fullmatch(r"fields:(\d+):(\d+)", str(cur))
            if not m:
                raise RangeError(f"cursor「{cur}」はこの書類の続きの印ではありません")
            a, b = int(m[1]), int(m[2])
        else:
            a, b = parse_range(rng.get("pages"), "ページ")
            a = a or 1
            b = min(b or n, n)
        out: list[dict] = []
        p = a
        stop = None
        while p <= b:
            limits.progress("ページを配置しています", p, n)
            for f in self.page(p)["fields"]:
                d = _field(f, page=p)
                if only_empty and d.get("value") not in (None, "", False):
                    continue
                out.append(d)
            p += 1
            if len(out) >= limit and p <= b:
                stop = p
                break
        res = {"fields": out, "total_pages": n, "looked_pages": f"{a}〜{(stop - 1) if stop else b}", "next": None}
        if stop:
            res["next"] = {"cursor": f"fields:{stop}:{b}"}
            res["remaining"] = f"{comma(b - stop + 1)} ページ"
        return res

    # ------------------------------------------------------------------ 書く
    def pages_for(self, values: dict) -> set[int]:
        """書く所のページ。ページが決まらない目印の書き込みは、頼み方を伝える。"""
        from .. import api

        pages: set[int] = set()
        by_name = {f.name: f for f in self.fields_all}
        for key, v in values.items():
            k = str(key)
            m = re.match(r"p(\d+)\.", k)
            if k.startswith("@"):
                if not isinstance(v, dict) or v.get("page") is None:
                    raise api.ApiError(f"{k}: 大きな PDF では、字を目印にした書き込みに page（ページの番号）を添えてください"
                                       "（read_window に pages を渡して、書く所のページを確かめられます）")
                pages.add(int(v["page"]))
            elif k in by_name:
                pages.update(w.page + 1 for w in by_name[k].widgets)
            elif m:
                pages.add(int(m[1]))
        return pages or {1}

    # ------------------------------------------------------------------ 段階：費用の小さな確かめは全部のページで
    def _font_gaps(self) -> list[str]:
        """字の対応（ToUnicode）のない埋め込みのフォント：ページの資源を、全部のページでたどる（フォントごとに 1 度だけ調べる）。"""
        doc = self.doc
        memo: dict[int, bool] = {}
        pages_bad: list[str] = []

        def bad_font(fd) -> bool:
            sub = fd.get("Subtype")
            sub = getattr(sub, "name", sub)
            if fd.get("ToUnicode") is not None:
                return False
            return str(sub).lstrip("/") == "Type0" or (fd.get("Encoding") is None and fd.get("FontDescriptor") is not None and str(sub).lstrip("/") != "Type1")

        for i, pg in enumerate(self.page_list, 1):
            try:
                res = doc.resolve(pg.get("Resources")) or {}
                fonts = doc.resolve(res.get("Font")) or {}
                for v in fonts.values():
                    fd = doc.resolve(v)
                    if not isinstance(fd, dict):
                        continue
                    key = id(fd)
                    if key not in memo:
                        memo[key] = bad_font(fd)
                    if memo[key]:
                        pages_bad.append(f"{i}ページ")
                        break
            except Exception:  # noqa: BLE001
                pages_bad.append(f"{i}ページ（フォントを読めません）")
        return pages_bad

    def grade(self):
        from ..grade import Reason, _final, _locked

        if not self.page_list:
            return _locked("ページがない", "ページの木をたどっても、ページが見つかりません", "ページがないので、書けません")
        reasons: list = []
        write = [Reason("B", "欄をページごとに見つける", "大きな PDF なので、配置から見つける欄は、頼まれたページごとに見つけます。"
                        f"フォームの欄は {len(self.fields_all):,} か所あり、全部を読みました")]
        gaps = self._font_gaps()
        if gaps:
            reasons.append(Reason("C", "字の対応がないかもしれないフォント",
                                  f"{len(gaps)} ページに、文字の対応（ToUnicode）のない埋め込みのフォントがあります。字に直せない字があるかは、配置して確かめます", gaps))
        if self.xw:
            reasons.append(Reason("B", "XFA のフォーム", "XFA のフォームがあります。中身は窓の「XFA のフォーム」の節にあります"))
        g = _final(reasons, write)
        g.deferred = ["未確認：全ページの字の読み落とし", "未確認：画像だけのページ・字に直せない字（audit_document で全部のページを確かめます）"]
        return g

    # ------------------------------------------------------------------ 検査
    def audit_begin(self) -> dict:
        return {"next_page": 1, "chars": 0, "problems": [], "image_pages": [], "unmapped": 0, "fields": 0, "done": False, "checks": []}

    def audit_step(self, state: dict) -> dict:
        from ..pdf.audit import _check_unmapped_and_images, check_glyphs
        from ..pdf.window import page_window

        n = len(self.page_list)
        p = state["next_page"]
        if p <= n:
            pg = page_window(self.doc, self.page_list[p - 1], p, self.fields_all, layout=self.layout_on)
            view = _OnePage(self.doc, self.page_list[p - 1])
            win = {"pages": [pg]}
            c1 = check_glyphs(view, win)
            state["chars"] += pg.get("chars", 0)
            state["fields"] += len(pg["fields"])
            if pg.get("image_only"):
                state["image_pages"].append(p)
            if not c1.ok and len(state["problems"]) < 30:
                state["problems"] += list(c1.items)[:5] or [f"{p} ページ：{c1.detail}"]
            for c in _check_unmapped_and_images(view, win):
                if not c.ok and len(state["problems"]) < 30:
                    state["problems"] += list(c.items)[:5] or [f"{p} ページ：{c.detail}"]
            state["next_page"] = p + 1
            limits.progress("ページを確かめています", p, n)
            return state
        state["checks"] = [{"name": "読み落とし（ページの文字）", "ok": not state["problems"], "agent": False, "items": list(state["problems"]),
                            "detail": f"全 {n:,} ページ・文字 {state['chars']:,} 字を、窓とは別の読み方（文字の命令を順に）で拾い、窓に出ていることを確かめました。"
                                      f"見つけた欄 {state['fields']:,} か所" + (f"。画像だけのページ {len(state['image_pages'])}" if state["image_pages"] else "")}]
        state["checks"].append(self._write_trial())
        state["done"] = True
        return state

    def _write_trial(self) -> dict:
        import tempfile

        from .. import api

        wr = None
        for p in range(1, min(len(self.page_list), 50) + 1):
            fl = [f for f in self.page(p)["fields"] if f.get("kind") in ("text", "inline", "comb", "space") and not f.get("readonly")]
            if fl:
                wr = fl[0]
                break
        if wr is None:
            return {"name": "書き込みの試し", "ok": True, "agent": False, "detail": "はじめの 50 ページに書ける欄が見つからないので、試しませんでした", "items": []}
        with tempfile.TemporaryDirectory(prefix="todoku-audit-") as td:
            try:
                res = api.fill(self.path, {wr["id"]: "テスト"}, str(Path(td) / "t.pdf"), force=True, password=self.password)
            except ValueError as e:
                return {"name": "書き込みの試し", "ok": False, "agent": False, "detail": f"{wr['id']} に書けませんでした：{e}", "items": []}
        return {"name": "書き込みの試し", "ok": bool(res.get("proof_ok") and res.get("read_back_ok")), "agent": False,
                "detail": f"{wr['id']} に試しの値を書き、{res.get('proof')}", "items": res.get("unexpected") or []}
