"""書いたことで行が折り返すかの見積もり（目安）：Word・.doc・ODT・RTF・HWPX・PowerPoint の段落と、Excel・.xls・.ods の折り返して表示するセル。

書く前と書いた後の字を、同じ幅（段落なら、ページの幅から余白・字下げを引いた幅か、表のマスの幅からマスの余白を引いた幅。
セルなら、列の幅（結合したセルはその範囲）からセルの余白を引いた幅）と同じ字の大きさで並べ、行の数を数える。書いた後の行が
増えるなら、書いたことで行が折り返す見込み：選んだ字（「○１」）が次の行の字（「種」）と離れて見える、表のマスが高くなる、など。
証明と読み戻しは字が正しいことしか見ないので、見た目のこの変化は、警告として本人に伝える（書き込みは止めない）。

字の幅は textwidth（Unicode の東アジアの幅）：全角は字の大きさ、ほかは半分。書体の字ごとの幅・字詰め・禁則の細かい規則は
見ないので、言い方は「見込み（目安）」で、見積もりのずれを広い側にも狭い側にも見る（wraps）：見積もりの幅で行が増える・書いた後の字が
幅に近く、字を少し広く測るソフトでは増える・書く前の字が幅に近く、字を少し狭く測るソフトでは増える、の 3 つを言い分ける。行の終わりの空白と句読点（、。）は、行からはみ出しても折り返さない（Word のぶら下げ）。
英字の語は、語の途中では切らない（語の初めから次の行へ）。
"""

from __future__ import annotations

from collections.abc import Callable

from .textwidth import cells, east_asian_context

_HANG = set(" 　、。，．,.")  # 行の終わりではみ出しても、次の行へ送らない字（空白・句読点のぶら下げ）
MESSAGE = "書いたことで行が折り返す見込み（目安）"


def _latin(ch: str) -> bool:
    return ch.isascii() and (ch.isalnum() or ch in "'-_@.")


def count_lines(text: str, avail: float, size_at: Callable[[int], float] | float, first_indent: float = 0.0,
                tab: float = 0.0, extra_at: Callable[[int], float] | None = None) -> int:
    """字の並び text を幅 avail（pt）に並べたときの行の数（改行 \\n・\\v は行を分ける）。size_at：字の番号 → 字の大きさ（pt）。
    first_indent：段落の 1 行目の字下げ（pt。ぶら下げは負）。tab：タブの間隔（pt。0 ならタブは字の大きさ 1 つ分）。
    extra_at：字の番号 → その字に足す幅（pt。字を囲む線など。なければ足さない）。"""
    if avail <= 0:
        return 1
    size = (lambda i: float(size_at)) if not callable(size_at) else size_at  # noqa: E731
    wide = east_asian_context(text)
    lines = 1
    x = first_indent
    word_start = None  # いまの行の、英字の語の初め（字の番号, その位置）
    word_w = 0.0
    for i, ch in enumerate(text):
        if ch in "\n\v\r":
            lines += 1
            x, word_start, word_w = 0.0, None, 0.0
            continue
        s = size(i)
        if ch == "\t":
            step = tab if tab > 0 else s
            w = step - (x % step) if step > 0 else s
        else:
            w = 0.5 * s * cells(ch, wide)
        if extra_at is not None:
            w += extra_at(i)
        if _latin(ch):
            if word_start is None:
                word_start, word_w = x, 0.0
            word_w += w
        else:
            word_start, word_w = None, 0.0
        if x + w > avail + 0.01 and x > 0 and ch not in _HANG:
            lines += 1
            if word_start is not None and word_start > 0 and word_w <= avail:
                x = word_w  # 英字の語は、語の初めから次の行へ
                word_start = 0.0
            else:
                x = w
                word_start = 0.0 if word_start is not None else None
            continue
        x += w
    return lines


NARROW = 0.88  # ほかのソフト（LibreOffice など）で幅が狭く測られるときの見積もり（列の幅・字の幅の違いの目安）


def wraps(before: str, after: str, avail: float, size_before, size_after, first_indent: float = 0.0,
          tab: float = 0.0, extra_before=None, extra_after=None) -> tuple[int, int, bool | None] | None:
    """書く前と書いた後の行の数と、見積もりの幅で増えるか（True）・書いた後の字が幅に近く、字の幅を少し広く測るソフトでは
    増えるか（False）・書く前の字が幅に近く、字の幅を少し狭く測るソフトでは増えるか（None）。増えないときは None。
    見積もりのずれは、広い側にも狭い側にもある：書く前の字が、見積もりでは幅を少し出て 2 行でも、元の文書では 1 行に入って
    いることがある（そのとき、書いた後の 2 行は、増えた行）。"""
    for k, sure in ((1.0, True), (NARROW, False), (1 / NARROW, None)):
        nb = count_lines(before, avail * k, size_before, first_indent * k, tab, extra_before)
        na = count_lines(after, avail * k, size_after, first_indent * k, tab, extra_after)
        if na > nb:
            return nb, na, sure
    return None


def written_places(pa: list, pb: list, marks: dict, controls: dict | None = None) -> dict[int, tuple[list[str], int]]:
    """字が変わった段落の番号 → (所の名前, 初めて字が違う所の字の番号)。pa・pb：書いた後・書く前の段落（同じ並び）。
    何を頼まれたかでなく、文書に起きたこと（段落の字が変わった）で決める。名前は、その段落にある書いた所の印（marks：名前 →
    (段落, 初め, 終わり)）。印のない段落（入力欄の中に書いた段落）は、中の字が変わった入力欄の id（controls：id(欄) → 名前）。"""
    index = {id(p): i for i, p in enumerate(pa)}
    named: dict[int, list[str]] = {}
    for name, (q, _a, _b) in marks.items():
        i = index.get(id(q))
        if i is not None:
            named.setdefault(i, []).append(name)
    out: dict[int, tuple[list[str], int]] = {}
    for i, (a, b) in enumerate(zip(pa, pb)):
        ta, tb = a.text, b.text
        if ta == tb or _flow(a) == _flow(b):
            continue
        names = named.get(i) or []
        if not names and controls:
            names = [controls[id(c)] for c, old in zip(_fields(a), _fields(b))
                     if id(c) in controls and _shown(c) != _shown(old)]
        at = next((k for k, (x, y) in enumerate(zip(ta, tb)) if x != y), min(len(ta), len(tb)))
        out[i] = (list(dict.fromkeys(names)), at)
    return out


# 自分の箱の中に値を描く部品（フォームの部品）：値が長くなっても、段落の行は変わらない（箱の中で切れる）
BOXED = frozenset({"hwpx-form", "hwpx-radio", "odf-form", "odf-radio", "odf-label", "odf-other"})


def _flow(p) -> str:
    """段落の行に流れる字（箱の部品の値を除く）。"""
    return "".join(sg.text for sg in p.segs if not sg.hidden
                   and getattr(getattr(sg, "field", None), "source", "") not in BOXED)


def _fields(p) -> list:
    out: list = []
    for sg in p.segs:
        f = getattr(sg, "field", None)
        if f is not None and not any(f is x for x in out):
            out.append(f)
    return out


def _shown(c) -> str:
    return "".join(s.text for s in c.result_segs if not getattr(s, "hidden", False))


def place_name(names: list[str]) -> str:
    return "・".join(names) or "字が変わった段落"


class WrapNote(str):
    """警告の言葉と、増える行の数（added）。ページの増えの事実（pagecount）に渡す。"""

    added: int = 0


def say(place: str, shown: str, nb: int, na: int, basis: str, sure: bool | None = True) -> str:
    """警告の言葉：所・字・行の数・見積もりの元。"""
    s = shown.replace("\n", " ").strip()
    s = s if len(s) <= 40 else s[:40] + "…"
    how = (f"「{s}」が {nb} 行から {na} 行になります" if sure else
           f"「{s}」が幅いっぱいに近く、字の幅を少し広く測るソフト（LibreOffice など）や印刷では {nb} 行から {na} 行になります"
           if sure is False else
           f"「{s}」は、書く前の字が幅いっぱいに近く、字の幅を少し狭く測るソフトや印刷では {nb} 行から {na} 行になります")
    note = WrapNote(f"{place}: {MESSAGE}：{how}（{basis}からの見積もり）。選んだ字や書いた字が次の行の字と離れて見えることがあるので、"
                    "字を短くするか、本人に確かめてください")
    note.added = na - nb
    return note
