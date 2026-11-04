"""文書に埋め込んだ表を、頼まれたときに読む：大きいから読み減らす。読めない所は作らない。

窓に入る分（limits.CONTEXT_BUDGET）の埋め込みは、いままでどおり窓に Excel の窓として出す。入らない埋め込みは、窓を開くときには
解析せず、シートの名前と読み方だけを出す（deferred）。中身は read_window に sheet: "E1!シート" と rows を渡すと、その行だけを解析して返す。
list_fields も同じ（api._embed_range_fields）。書くときの欄の名前（E1!シート!セル）は変わらない。

入るかどうかは、ひとり立ちの Excel のブックを開くときと同じ見積もり（lazy/xlsxdoc.XlsxLazy._estimate）で決める。
この決まりは窓を開く道（api.window）の中だけ（window_room）。検査（docx.audit）は、すべての埋め込みを読んで確かめる。
"""

from __future__ import annotations

import contextlib
import contextvars
import hashlib
import re

from .. import limits

_BOOKS_MAX = 4
_ROOM: contextvars.ContextVar[list | None] = contextvars.ContextVar("todoku_embed_room", default=None)
_WHOLE: contextvars.ContextVar[bool] = contextvars.ContextVar("todoku_embed_whole", default=False)


@contextlib.contextmanager
def window_room():
    """この中で作る窓では、窓に入らない埋め込みを後回しにする。"""
    token = _ROOM.set(None if _WHOLE.get() else [limits.CONTEXT_BUDGET])
    try:
        yield
    finally:
        _ROOM.reset(token)


@contextlib.contextmanager
def whole():
    """この中では、すべての埋め込みを読む（時間に縛られない長い仕事の検査）。"""
    token = _WHOLE.set(True)
    try:
        yield
    finally:
        _WHOLE.reset(token)


def spend(chars: int) -> None:
    """窓に chars 字を使った（本文・窓に出した埋め込み）。"""
    room = _ROOM.get()
    if room is not None:
        room[0] -= chars


def book_bytes(doc, e: dict) -> tuple[str, bytes | int] | None:
    """埋め込みのブック：("xlsx", .xlsx のバイト) か ("xls", ブックのストリームのバイト数)。Excel のブックでなければ None。"""
    from ..cfb import CFB, SubStorage

    if getattr(doc, "kind", "") == "doc":
        c = SubStorage(doc.cfb, e["storage"])
    else:
        data = doc.zip.read(e["storage"])
        if data[:4] == b"PK\x03\x04":
            return "xlsx", data
        if data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            return None
        c = CFB(data)
    names = {n.lower(): n for n in c.streams()}
    if "package" in names:
        return "xlsx", c.read(names["package"])
    for n in ("workbook", "book"):
        if n in names:
            return "xls", len(c.read(names[n]))
    return None


_BOOKS: dict = {}  # 同じ埋め込みを読み進めるたびに、シートの索引を作り直さない


def _xlsx_lazy(data: bytes, name: str):
    from ..lazy.xlsxdoc import XlsxLazy

    key = (name, hashlib.sha256(data).digest())
    got = _BOOKS.get(key)
    if got is None:
        while len(_BOOKS) >= _BOOKS_MAX:
            _BOOKS.pop(next(iter(_BOOKS)))
        got = _BOOKS[key] = XlsxLazy(None, name, blob=data)
    return got


def later(doc, e: dict) -> dict | None:
    """窓に入らない埋め込みなら、窓に出す概要。入るなら None（いままでどおり窓を作る）。"""
    room = _ROOM.get()
    if room is None:
        return None
    try:
        got = book_bytes(doc, e)
        if got is None:
            return None
        if got[0] == "xlsx":
            lz = _xlsx_lazy(got[1], e["storage"])
            est = lz.est_chars
        else:
            lz, est = None, got[1]  # BIFF のセルは、窓の字とおよそ同じバイト
    except Exception:  # noqa: BLE001  見積もれない：いままでの道で、開けない理由を伝える
        return None
    if est <= room[0] and room[0] > 0:
        return None
    room[0] -= est
    out = {"n": e["n"], "label": e["label"], "prog": e["prog"], "storage": e["storage"], "deferred": True, "est_chars": est, "fields": [], "writable": None}
    wb = doc.embedded_book(e)
    out["sheets"] = [sh.name for sh in wb.sheets if not sh.state.startswith("non-")] if wb is not None else []
    out["drawn"] = _drawn(wb) if wb is not None else []
    return out


def _drawn(wb) -> list[str]:
    """ブックのグラフ・図の埋め込み（窓の節）：行の範囲では読めないので、後回しにせず、ここで出す。セルの解析はしない。"""
    from .excel import add_to_window, sheet_markdown

    sheets = [{} for _ in wb.sheets]
    add_to_window(wb, {"sheets": sheets})
    return [ln for d in sheets for ln in sheet_markdown(d)]


def markdown(e: dict) -> list[str]:
    """後回しにした埋め込みの、窓の節。"""
    n = e["n"]
    first = (e.get("sheets") or ["シート"])[0]
    return [f"- 本文の〔埋め込み{n}：…〕の所にある表です。大きいので、ここでは中身を読んでいません（窓は見積もりで {e['est_chars']:,} 字）。読めない所はありません："
            f"read_window に sheet: \"E{n}!{first}\" と rows: \"1-500\" を渡すと、その行を読めます（結果の next をたどれば、すべての行を読めます）。"
            f"list_fields も同じ頼み方で、その行の欄を返します。書くときの欄の名前は `E{n}!シート!セル` です。"
            "書けるかどうか（表の絵を描き直せるか）は、書くとき（と audit_document）に確かめます。",
            "- シート：" + ("、".join(e.get("sheets") or []) or "（セルのあるシートはありません）"), ""] \
        + [ln.replace("### ", "#### ", 1) if ln.startswith("### ") else ln for ln in e.get("drawn") or []]


_SHEET = re.compile(r"E(\d+)(?:!(.+))?", re.S)


def split_sheet(sheet) -> tuple[int, str | None] | None:
    """「E1!シート」→ (1, "シート")。埋め込みの指定でなければ None。"""
    m = _SHEET.fullmatch(str(sheet)) if sheet else None
    return (int(m[1]), m[2]) if m else None


def read(doc, n: int, sheet: str | None, rng: dict, max_chars: int):
    """埋め込み n の表の、頼まれた行だけを解析した 1 回分（lazy.base.Page）。"""
    from ..lazy.base import Page, RangeError

    e = next((x for x in getattr(doc, "embeds", []) or [] if x.get("n") == n), None)
    got = book_bytes(doc, e) if e is not None else None
    if got is None:
        raise RangeError(f"埋め込み E{n} の表がありません")
    want = {k: v for k, v in rng.items() if k in ("rows", "cursor")}
    if sheet:
        want["sheet"] = sheet
    if got[0] == "xlsx":
        page = _xlsx_lazy(got[1], e["storage"]).read(want, max_chars)
    else:
        page = _read_book(doc.embedded_book(e), want, max_chars)
    head = f"〔埋め込み{n} の表。書くときの欄の名前は、ここの［シート!C4］や［C4］に「E{n}!」とシートを付けて E{n}!シート!C4〕"
    nxt = None
    if page.next:
        nxt = dict(page.next)
        nxt["sheet"] = f"E{n}!{nxt.get('sheet') or sheet or ''}".rstrip("!")
    return Page(head + "\n" + page.text, f"埋め込み{n}：{page.covered}", nxt, page.remaining, page.body)


def _read_book(wb, rng: dict, max_chars: int):
    """.xls の表：シート 1 枚だけを解析して、行の範囲を読む（BIFF のブックは、もともと大きくなれない）。"""
    from ..lazy.base import Page, RangeError, comma, parse_range
    from ..xlsx import window as xw

    if wb is None:
        raise RangeError("この埋め込みのブックを開けません")
    sheets = [s for s in wb.sheets if not s.state.startswith("non-")]
    name = rng.get("sheet")
    sh = next((s for s in sheets if s.name == name), None) if name else (sheets[0] if sheets else None)
    if sh is None:
        raise RangeError(f"「{name}」というシートはありません（シート：" + "、".join(s.name for s in sheets[:30]) + "）")
    sh.__dict__[xw._MEMO] = {}
    try:
        lines = xw._md_sheet(xw._sheet_doc(wb, sh, xw._Options(wb)))
    finally:
        sh.__dict__.pop(xw._MEMO, None)
    if rng.get("rows") is not None:
        a, b = parse_range(rng["rows"], "行")
        kept = []
        for ln in lines:
            m = re.match(r"- (\d+): ", ln)
            if m is None or (a or 1) <= int(m[1]) <= (b or 10 ** 9):
                kept.append(ln)
        lines = kept
    text = "\n".join(lines)
    at = int(str(rng.get("cursor") or "chars:0").rsplit(":", 1)[-1])
    end = min(at + max_chars, len(text))
    if end < len(text):
        nl = text.rfind("\n", at, end)
        end = nl + 1 if nl > at else end
    piece = text[at:end]
    cov = f"シート「{sh.name}」の {at + 1:,}〜{end:,} 字目（全 {len(text):,} 字）"
    if end >= len(text):
        return Page(piece, cov, None, "", [piece])
    nxt = {"sheet": sh.name, "cursor": f"chars:{end}"}
    if rng.get("rows") is not None:
        nxt["rows"] = rng["rows"]
    return Page(piece, cov, nxt, f"{comma(len(text) - end)} 字", [piece])
