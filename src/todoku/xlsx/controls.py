"""Excel のフォームの部品（オプションボタン・ドロップダウン・リストボックス・スピンボタン・スクロールバー）を、入力欄として読み書きする。

規則は Microsoft の仕様書 MS-XLSX 2.6.65 CT_FormControlPr と、VML の x:ClientData のとおり：
- ドロップダウン・リストボックス：選んだ項目の番号（1から。0 は選んでいない）が sel。つながったセル（fmlaLink）にも同じ番号が入る。
  選択肢は fmlaRange のセル（またはファイルの中の itemLst）。複数選択のリストボックスは multiSel で、セルにはつながらない
- スピンボタン・スクロールバー：いまの値が val（min〜max）。つながったセルにも同じ値が入る
- オプションボタン：同じ組のうち選んだもの1つだけが checked。つながったセルには、組の中で何番目か（1から）が入る。
  組は、グループボックスの中のボタン（グループボックスの fmlaGroup があれば、それがつながったセル）と、
  どのグループボックスにも入らないボタン。組の中の順は、ファイルの中の部品の順（作った順）とする。
  仕様書は組の順をはっきり決めていないので、保存された「つながったセルの値」と「選ばれたボタン」が、この順と食い違う組は
  「規則では決められない」として示し、書き込まない
- 古い Excel（.xls）では、組と組の中の順がファイルに書かれている（MS-XLS 2.5.153 FtRboData：組の先頭の印 fFirstBtn と、
  次のボタン idRadNext のつながり）。そのときは、グループボックスの囲みではなく、このつながりを組とする。
  先頭の印がない・輪になっているなど、つながりから組の順が決まらないものは「規則では決められない」とする
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from ..formula.parser import addr
from .workbook import Control, Sheet, Workbook

CHOICE_KINDS = ("Drop", "List")
VALUE_KINDS = ("Spin", "Scroll")


@dataclass
class RadioGroup:
    sheet: str
    radios: list[Control]
    box: Control | None = None
    link: tuple[str, int, int] | None = None
    consistent: bool = True  # 保存された値と、選ばれたボタンが、組の順と合っている
    why: str = ""

    @property
    def options(self) -> list[str]:
        return [r.text.strip() or r.name for r in self.radios]

    @property
    def selected(self) -> int:
        """選ばれたボタンの番号（1から）。選ばれていなければ 0。"""
        for i, r in enumerate(self.radios, 1):
            if r.checked == "Checked":
                return i
        return 0

    @property
    def id(self) -> str:
        if self.link:
            return f"{self.link[0]}!{addr(self.link[1], self.link[2])}"
        return f"{self.sheet}!{self.radios[0].ref_name}"


def _rect(c: Control):
    if c.anchor is None:
        return None
    r1, c1 = c.anchor
    r2, c2 = c.anchor_to or c.anchor
    return (r1, c1, r2, c2)


def _inside(inner, outer) -> bool:
    return outer[0] <= inner[0] and inner[2] <= outer[2] and outer[1] <= inner[1] and inner[3] <= outer[3]


def _area(c: Control) -> int:
    r = _rect(c)
    return (r[2] - r[0] + 1) * (r[3] - r[1] + 1)


def _chains(radios: list[Control]) -> list[tuple[list[Control], str]]:
    """古い Excel（.xls）の組：組の先頭（fFirstBtn）から次のボタン（idRadNext）をたどった順。(ボタン, 決められない理由)。"""
    by_id = {r.shape_id: r for r in radios}
    seen: set[str] = set()
    out: list[tuple[list[Control], str]] = []

    def follow(start: Control, why: str) -> None:
        chain = []
        cur: Control | None = start
        while cur is not None:
            seen.add(cur.shape_id)
            chain.append(cur)
            nxt = by_id.get(cur.next_button or "")
            if nxt is None or nxt is cur or nxt.first_button or nxt is start:
                break  # 次がない（0・自分・ほかの組の先頭・先頭に戻る）
            if nxt.shape_id in seen:
                why = why or "ボタンのつながり（idRadNext）が、ほかの組のボタンにつながっています"
                break
            cur = nxt
        out.append((chain, why))

    for r in radios:
        if r.first_button and r.shape_id not in seen:
            follow(r, "")
    rest = [r for r in radios if r.shape_id not in seen]
    pointed = {r.next_button for r in rest if r.next_button != r.shape_id}  # 自分を指すのは「次がない」
    for r in rest:
        if r.shape_id not in seen and r.shape_id not in pointed:
            follow(r, "組の先頭のボタン（fFirstBtn）の印がありません")
    for r in rest:
        if r.shape_id not in seen:
            follow(r, "ボタンのつながり（idRadNext）が輪になっていて、組の先頭が分かりません")
    return out


def radio_groups(wb: Workbook, sh: Sheet) -> list[RadioGroup]:
    from .workbook import _link_cell

    radios = [c for c in sh.controls if c.kind == "Radio"]
    if not radios:
        return []
    boxes = [c for c in sh.controls if c.kind == "GBox" and _rect(c) is not None]
    groups: dict[str, RadioGroup] = {}
    if any(r.next_button is not None for r in radios):
        # 古い Excel（.xls）：ファイルに書かれた組。グループボックスは見出しにだけ使う（すべてのボタンを囲む、いちばん小さいもの）
        for k, (chain, why) in enumerate(_chains(radios)):
            rects = [_rect(r) for r in chain]
            inside = [b for b in boxes if all(rr is not None and _inside(rr, _rect(b)) for rr in rects)]
            g = RadioGroup(sh.name, chain, min(inside, key=_area) if inside else None)
            if why:
                g.consistent, g.why = False, why
            groups[str(k)] = g
    else:
        for r in radios:
            rr = _rect(r)
            inside = [b for b in boxes if rr is not None and _inside(rr, _rect(b))]
            box = min(inside, key=_area) if inside else None
            key = box.shape_id if box is not None else "(sheet)"
            g = groups.setdefault(key, RadioGroup(sh.name, [], box))
            g.radios.append(r)
    out = []
    for g in groups.values():
        link = None
        if g.box is not None and g.box.fmla_group:
            link = _link_cell(g.box.fmla_group, sh.name)
        if link is None:
            link = next((r.link_cell for r in g.radios if r.link_cell), None)
        if link is not None and wb.has_sheet(link[0]):
            g.link = (wb.sheet(link[0]).name, link[1], link[2])
        links = {r.link for r in g.radios if r.link}
        if not g.consistent:
            pass  # つながり（idRadNext）から組の順が決まらない
        elif len(links) > 1 and not (g.box is not None and g.box.fmla_group):
            g.consistent, g.why = False, f"同じ組のボタンが、違うセル（{'、'.join(sorted(links))}）につながっています"
        elif sum(1 for r in g.radios if r.checked == "Checked") > 1:
            g.consistent, g.why = False, "同じ組の中で、2つ以上のボタンが選ばれています"
        elif g.link is not None:
            cell = wb.sheet(g.link[0]).cells.get(g.link[1:])
            v = cell.value if cell is not None else None
            sel = g.selected
            if not ((v in (None, "", 0) and sel == 0) or (isinstance(v, (int, float)) and not isinstance(v, bool) and int(v) == sel
                                                          and v == int(v))):
                g.consistent, g.why = False, (f"つながったセル {addr(*g.link[1:])} の値（{v!r}）と、選ばれたボタン（{sel} 番目）が、"
                                              "ファイルの中の部品の順では合いません")
        out.append(g)
    _set_places(out)
    return out


def _set_places(groups: list[RadioGroup]) -> None:
    """同じシートで、ボタンの名前が別の組とぶつかる組だけ、ボタンの id に組の位置（先頭のボタンの左上のセル）を付ける。ぶつからない組は変えない。
    ファイルは変えない（Excel の組の動きはグループボックスで決まる。名前は id のためだけに区別する）。"""
    owners: dict[str, set[int]] = {}
    for i, g in enumerate(groups):
        for r in g.radios:
            owners.setdefault(r.name, set()).add(i)
    clashing = {i for who in owners.values() if len(who) > 1 for i in who}
    taken: dict[str, int] = {}
    for i, g in enumerate(groups):
        if i not in clashing:
            continue
        head = g.radios[0]
        place = addr(*head.anchor) if head.anchor else f"s{head.shape_id}"
        taken[place] = taken.get(place, 0) + 1
        if taken[place] > 1:
            place = f"{place}.{head.shape_id}"  # 先頭のボタンが同じセルにある組：図形の番号で分ける
        for r in g.radios:
            r.place = place


def _set_own_places(sh: Sheet) -> None:
    """オプションボタンのほかの、値を書く部品（チェックボックス・ドロップダウン・リストボックス・スピンボタン・スクロールバー）で、
    同じシートのほかの部品と名前がぶつかるものだけ、id に自分の位置（左上のセル）を付ける。ぶつからない部品は変えない。
    ファイルは変えない（Excel は、写して貼った部品に同じ名前を付けることがある：名前は id のためだけに区別する）。
    ActiveX の部品の名前は、Excel がシートの中で 1 つに決めるので、付けない。"""
    count: dict[str, int] = {}
    for c in sh.controls:
        count[c.name] = count.get(c.name, 0) + 1
    todo = [c for c in sh.controls if c.kind in ("CheckBox",) + CHOICE_KINDS + VALUE_KINDS and count[c.name] > 1
            and getattr(c, "ax", None) is None and not c.place]
    if not todo:
        return
    used = {c.ref_name for c in sh.controls}
    for c in todo:
        place = addr(*c.anchor) if c.anchor else f"s{c.shape_id}"
        if f"{c.name}@{place}" in used:
            place = f"{place}.{c.shape_id}"  # 同じ名前の部品が同じセルにある：図形の番号で分ける
        c.place = place
        used.add(c.ref_name)


def ensure_places(wb: Workbook, sh: Sheet) -> None:
    """同じ名前の部品があるシートで、id の位置を入れる（オプションボタンは、組を求めるまで分からない）。"""
    names = [c.name for c in sh.controls if c.kind == "Radio"]
    if len(names) != len(set(names)):
        radio_groups(wb, sh)
    if len(sh.controls) > 1:
        _set_own_places(sh)


def items_of(wb: Workbook, sh: Sheet, ctl: Control, calc=None) -> list[str]:
    """ドロップダウン・リストボックスの選択肢（fmlaRange のセルの見た目の文字、なければ itemLst）。"""
    if ctl.items is not None:
        return list(ctl.items)
    if not ctl.fmla_range:
        return []
    from ..formula.parser import parse
    from ..formula.values import ExcelError, to_text
    from .calc import Calc

    calc = calc or Calc(wb)
    try:
        from ..formula.evaluator import Evaluator

        f = ctl.fmla_range.lstrip("=")
        v = Evaluator(calc, sh.name, 1, 1).ev(parse(f))
        if isinstance(v, ExcelError):
            return []
        g = Evaluator(calc, sh.name, 1, 1).grid(v)
    except Exception:  # noqa: BLE001
        return []
    return [str(x) if isinstance(x, ExcelError) else to_text(x) if x is not None else "" for row in g.data for x in row]


def state(ctl: Control) -> tuple:
    """部品の状態（書き込みで変えてよいもの）。"""
    return (ctl.checked, ctl.sel, ctl.val, tuple(ctl.multi_sel or ()))


# ====================================================================== 書く
@dataclass
class ControlWrite:
    cells: dict[tuple[str, int, int], object] = field(default_factory=dict)  # つながったセルに書く値
    controls: list[tuple[str, Control, dict]] = field(default_factory=list)  # (シート, 部品, 変える状態)


def _index(value, options: list[str], what: str) -> int:
    from ..formula.values import to_text

    if isinstance(value, bool):
        raise ValueError(f"{what}は、選択肢の文字か番号（1から）で指定してください")
    if isinstance(value, (int, float)) and value == int(value):
        i = int(value)
        if 0 <= i <= len(options):
            return i
        raise ValueError(f"{what}の番号 {i} は、選択肢（{len(options)} 個）の範囲外です")
    t = to_text(value).strip()
    if t.isdigit() and t not in options:
        return _index(int(t), options, what)
    hits = [i for i, o in enumerate(options, 1) if o.strip() == t]
    if len(hits) == 1:
        return hits[0]
    if len(hits) > 1:
        return hits[0]  # Excel も、同じ文字の選択肢は先頭のものとして扱う
    shown = "、".join(f"「{o}」" for o in options[:15])
    raise ValueError(f"{what}に「{t}」はありません。選べるのは {shown}")


def plan(wb: Workbook, sheet: str, ctl: Control, value, calc=None) -> ControlWrite:
    """部品 ctl に value を書くときの、つながったセルの値と部品の状態。"""
    out = ControlWrite()
    sh = wb.sheet(sheet)
    if ctl.kind == "Radio":
        g = next(g for g in radio_groups(wb, sh) if ctl in g.radios)
        return plan_group(wb, g, True if value is True else value, chosen=ctl)
    if ctl.kind in CHOICE_KINDS:
        opts = items_of(wb, sh, ctl, calc)
        if ctl.kind == "List" and (ctl.sel_type or "single") != "single":
            vals = value if isinstance(value, (list, tuple)) else [value]
            idx = sorted({_index(v, opts, f"リストボックス「{ctl.name}」") for v in vals} - {0})
            out.controls.append((sheet, ctl, {"multi_sel": idx}))
            return out
        i = _index(value, opts, f"「{ctl.name}」")
        out.controls.append((sheet, ctl, {"sel": i}))
        if ctl.link_cell and wb.has_sheet(ctl.link_cell[0]):
            out.cells[(wb.sheet(ctl.link_cell[0]).name, ctl.link_cell[1], ctl.link_cell[2])] = float(i)
        return out
    if ctl.kind in VALUE_KINDS:
        from ..formula.values import to_number

        try:
            v = to_number(value)
        except Exception:  # noqa: BLE001
            raise ValueError(f"「{ctl.name}」には数を指定してください")
        lo = ctl.vmin if ctl.vmin is not None else 0
        hi = ctl.vmax if ctl.vmax is not None else 100
        if v != int(v) or not (lo <= v <= hi):
            raise ValueError(f"「{ctl.name}」の値は {lo}〜{hi} の整数です（{value!r} は範囲外）")
        out.controls.append((sheet, ctl, {"val": int(v)}))
        if ctl.link_cell and wb.has_sheet(ctl.link_cell[0]):
            out.cells[(wb.sheet(ctl.link_cell[0]).name, ctl.link_cell[1], ctl.link_cell[2])] = float(int(v))
        return out
    raise ValueError(f"「{ctl.name}」（{ctl.kind}）は、値を書く部品ではありません")


def plan_group(wb: Workbook, g: RadioGroup, value, chosen: Control | None = None) -> ControlWrite:
    if not g.consistent:
        raise ValueError(f"オプションボタンの組（{g.id}）は、組の順を規則では決められません：{g.why}")
    if chosen is not None and value is True:
        i = g.radios.index(chosen) + 1
    elif chosen is not None and value is False:
        i = 0 if g.selected == g.radios.index(chosen) + 1 else g.selected
    else:
        i = _index(value, g.options, "オプションボタン")
    out = ControlWrite()
    for k, r in enumerate(g.radios, 1):
        out.controls.append((g.sheet, r, {"checked": "Checked" if k == i else None}))
    if g.link is not None:
        out.cells[g.link] = float(i)
    return out


def controls_linked_to(wb: Workbook) -> dict[tuple[str, int, int], tuple[str, object]]:
    """つながったセル → ("radio", 組) / ("control", (シート, 部品))（チェックボックス以外の、値を書く部品）。"""
    out: dict = {}
    for sh in wb.sheets:
        for g in radio_groups(wb, sh):
            if g.link:
                out[g.link] = ("radio", g)
        for c in sh.controls:
            if c.kind in CHOICE_KINDS + VALUE_KINDS and c.link_cell and wb.has_sheet(c.link_cell[0]):
                out[(wb.sheet(c.link_cell[0]).name, c.link_cell[1], c.link_cell[2])] = ("control", (sh.name, c))
    return out


# ---------------------------------------------------------------------- 部品の XML を書き換える
_ATTR = {"checked": "checked", "sel": "sel", "val": "val", "multi_sel": "multiSel"}
_VML_TAG = {"checked": "Checked", "sel": "Sel", "val": "Val", "multi_sel": "MultiSel"}


def _fmt(key: str, v) -> str | None:
    if v is None:
        return None
    if key == "multi_sel":
        return ",".join(str(x) for x in v) if v else None
    return str(v)


# CT_FormControlPr の属性の順（MS-XLSX 2.6.65 のスキーマの順。Excel もこの順に書く）
_ATTR_ORDER = ["objectType", "checked", "colored", "dropLines", "dropStyle", "dx", "firstButton", "fmlaGroup", "fmlaLink",
               "fmlaRange", "fmlaTxbx", "horiz", "inc", "justLastX", "lockText", "max", "min", "multiSel", "noThreeD",
               "noThreeD2", "page", "sel", "seltype", "selType", "textHAlign", "textVAlign", "val", "widthMin", "editVal",
               "multiLine", "verticalBar", "passwordEdit"]


def set_ctrlprop(xml: bytes, changes: dict) -> bytes:
    """ctrlProp（formControlPr）の状態の属性だけを変える。ある属性は値だけを替え、ない属性はスキーマの順の位置に足す。"""
    m = re.search(rb"<(?:\w+:)?formControlPr\b[^>]*?/?>", xml)
    if not m:
        raise ValueError("コントロールの設定（formControlPr）が見つかりません")
    tag = m[0]
    for key, v in changes.items():
        name = _ATTR[key]
        s = _fmt(key, v)
        pat = re.compile(rb"(\s+)" + name.encode() + rb'="[^"]*"')
        if pat.search(tag):
            tag = pat.sub((rb"\g<1>" + name.encode() + b'="' + s.encode() + b'"') if s is not None else b"", tag, count=1)
            continue
        if s is None:
            continue
        # 足す位置：スキーマの順で、この属性より前に来る属性のうち、いちばん後ろにあるものの後ろ
        rank = _ATTR_ORDER.index(name)
        pos = None
        for am in re.finditer(rb'\s+([\w:]+)="[^"]*"', tag):
            a = am[1].decode()
            if a in _ATTR_ORDER and _ATTR_ORDER.index(a) < rank:
                pos = am.end()
        ins = b" " + name.encode() + b'="' + s.encode() + b'"'
        if pos is None:
            pos = re.search(rb"(\s*/?>)$", tag).start()
        tag = tag[:pos] + ins + tag[pos:]
    return xml[:m.start()] + tag + xml[m.end():]


def set_vml(vml: bytes, shape_id: str, changes: dict) -> bytes:
    """VML の図形の、状態の要素（x:Checked・x:Sel・x:Val・x:MultiSel）だけを変える。"""
    sm = None
    for m in re.finditer(rb"<v:shape\b[^>]*>.*?</v:shape>", vml, re.S):
        head = re.match(rb"<v:shape\b[^>]*>", m[0])[0]
        if re.search(rb'\b(?:id|o:spid)="_x0000_s' + shape_id.encode() + rb'"', head):
            sm = m
            break
    if sm is None:
        raise ValueError(f"VML に図形 {shape_id} が見つかりません")
    block = sm[0]
    for key, v in changes.items():
        tagname = _VML_TAG[key].encode()
        s = _fmt(key, v)
        if key == "checked":
            s = "1" if v == "Checked" else ("2" if v == "Mixed" else None)
        pat = rb"(<x:" + tagname + rb">)[^<]*(</x:" + tagname + rb">)"
        if re.search(pat, block):
            if s is None:
                block = re.sub(rb"[ \t]*<x:" + tagname + rb">[^<]*</x:" + tagname + rb">[ \t]*\r?\n?", b"", block, count=1)
            else:
                block = re.sub(pat, rb"\g<1>" + s.encode() + rb"\g<2>", block, count=1)
        elif s is not None:
            # </x:ClientData> の前に、ほかの要素と同じ改行・字下げで足す
            end = block.find(b"</x:ClientData>")
            if end < 0:
                raise ValueError(f"VML の図形 {shape_id} に状態を書く場所がありません")
            elem = b"<x:" + tagname + b">" + s.encode() + b"</x:" + tagname + b">"
            am = re.search(rb"(\r?\n)[ \t]*$", block[:end])
            prev = re.findall(rb"(\r?\n)([ \t]*)<x:", block[:end])
            if am and prev:
                nl, ind = prev[-1]
                block = block[:am.start()] + nl + ind + elem + block[am.start():]
            else:
                block = block[:end] + elem + block[end:]
    return vml[:sm.start()] + block + vml[sm.end():]
