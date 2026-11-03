"""HWPX の書き込みの証明：変わったのは、書いた欄のある段落（と、書いたフォームの部品の値）だけであることを、書く前と後の
両方を読み直して確かめる。

- パッケージの部品：書き換えた本文の節（Contents/sectionN.xml）のほかは、ヘッダーと圧縮されたバイト列まで同じ。並びも同じ
- 書き換えた節：書いた段落を除いたバイト列が同じ。標準の XML の読み手（expat）で読める
- 書いた段落：字が「元の字の空欄を、書いた値に置き換えたもの」と同じ。足した要素は字（hp:t）・改行（hp:lineBreak）・
  しおり（hp:ctrl の hp:bookmark。名前は「_todoku_」で始まる）・字の並び（hp:run。空の段落に書いたとき）だけ、除いた要素は
  空欄の字（hp:t と、空白の要素）と行の配置の控え（hp:linesegarray・hp:lineseg）だけ。文字の形（charPrIDRef）は元の段落に
  あったものだけ。누름틀の印（fieldBegin）は dirty のほか同じ。フォームの部品は、値の属性（value・selectedValue）と字（hp:text）
  のほか同じ
- 書いていない段落：字が同じ。段落の数が同じ。しおりの名前が重ならない
"""

from __future__ import annotations

import hashlib
import re
from collections import Counter
from dataclasses import dataclass, field

from ..docx.document import iter_paras
from ..docx.proof import _cut
from ..docx.window import Spots, WriteError, resolve_write
from ..odf.textproof import _expected
from ..odf.xml import El, set_attr, well_formed
from ..zipraw import RawZip
from .document import MARK_PREFIX, HwpxDocument
from .package import L

ADD_OK = {"t", "lineBreak", "ctrl", "bookmark", "run", "text"}
DROP_OK = {"t", "linesegarray", "lineseg", "fwSpace", "nbSpace", "lineBreak"}
FORM_TAGS = {"checkBtn", "radioBtn", "comboBox", "listBox", "edit"}


@dataclass
class HwpxProof:
    original_sha256: str
    new_sha256: str
    entries_total: int = 0
    entries_identical: int = 0
    paragraphs_total: int = 0
    paragraphs_changed: int = 0
    controls_changed: int = 0
    unexpected: list[str] = field(default_factory=list)
    lint_errors: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.unexpected and not self.lint_errors

    def headline(self) -> str:
        ctl = f"・フォームの部品 {self.controls_changed} 個" if self.controls_changed else ""
        return (f"パッケージの部品 {self.entries_identical}/{self.entries_total} がバイト一致。段落 {self.paragraphs_total:,} のうち、"
                f"変わったのは書いた欄のある {self.paragraphs_changed} 段落{ctl}だけ"
                + ("" if not self.unexpected else f"（想定外の変更 {len(self.unexpected)} 件）"))


def package_check(pr, original: bytes, new: bytes, allowed: set[str]) -> list[str]:
    za, zb = RawZip(original), RawZip(new)
    if za.names() != zb.names():
        pr.unexpected.append("パッケージの部品の並びが変わっています")
        return []
    changed = []
    for name in za.names():
        pr.entries_total += 1
        ea, eb = za.get(name), zb.get(name)
        if ea.data == eb.data and ea.local == eb.local and ea.descriptor == eb.descriptor:
            pr.entries_identical += 1
            continue
        changed.append(name)
        if name not in allowed:
            pr.unexpected.append(f"書き換えない部品 {name} が変わっています")
    if za.names()[:1] == ["mimetype"] and zb.get("mimetype").method != za.get("mimetype").method:
        pr.lint_errors.append("mimetype の部品が、先頭で圧縮しないで置かれていません")
    return changed


def _tags(el: El) -> Counter:
    return Counter(L(x) or x.tag for x in el.iter() if x is not el)


def _plain(p) -> str:
    return "".join(sg.text for sg in p.segs if not sg.hidden and not (sg.node is not None and sg.node.kind == "ctl"))


def _strip_value(tag: bytes, names) -> bytes:
    for n in names:
        tag = set_attr(tag, n, None)
    return tag


def prove(original: bytes, new: bytes, written: dict) -> HwpxProof:
    pr = HwpxProof(hashlib.sha256(original).hexdigest(), hashlib.sha256(new).hexdigest())
    try:
        a = HwpxDocument(original)
        b = HwpxDocument(new)
    except Exception as e:  # noqa: BLE001
        pr.lint_errors.append(f"書き込み後の文書を読めません: {e}")
        return pr
    changed = package_check(pr, original, new, set(a.parts))
    for part in changed:
        if part in b.parts:
            err = well_formed(b.pkg.xml(part))
            if err:
                pr.lint_errors.append(f"{part}：XML として読めません（{err}）")
    pa = [p for _k, _part, bl in a.all_blocks() for p in iter_paras(bl)]
    pb = [p for _k, _part, bl in b.all_blocks() for p in iter_paras(bl)]
    pr.paragraphs_total = len(pa)
    if len(pa) != len(pb):
        pr.unexpected.append(f"段落の数が変わっています（{len(pa)} → {len(pb)}）")
        return pr
    sp = Spots(a)
    index = {id(p): i for i, p in enumerate(pa)}
    targets: dict[int, list] = {}
    form_paras: set[int] = set()
    used: dict = {}
    for fid, value in written.items():
        if isinstance(value, dict) and "@" in value:
            try:
                it = resolve_write(sp, value["@"], used)
            except WriteError as e:
                pr.unexpected.append(f"書いた所 {fid} の目印が、元の文書で決まりません：{e}")
                continue
            if it["id"] != fid or it["value"] != value["value"]:
                pr.unexpected.append(f"書いた所 {fid} が、目印から決まる所（{it['id']}「{it['value']}」）と違います")
                continue
            value = value["value"]
        else:
            it = sp.items.get(fid)
        if it is None:
            pr.unexpected.append(f"書いた欄 {fid} が、元の文書にありません")
            continue
        c = it.get("control")
        if c is not None and c.source in ("hwpx-form", "hwpx-radio"):
            pr.controls_changed += 1
            for q in (getattr(c, "radio_paras", None) or [c.para]):
                if q is not None and id(q) in index:
                    form_paras.add(index[id(q)])
                    targets.setdefault(index[id(q)], [])
            continue
        p = it.get("para") if c is None else c.para
        if p is None or id(p) not in index:
            pr.unexpected.append(f"書いた欄 {fid} の段落が見つかりません")
            continue
        targets.setdefault(index[id(p)], []).append((fid, it, value))
    for part in changed:
        xa, xb = a.pkg.xml(part), b.pkg.xml(part)
        sa = [(p.el.s, p.el.e) for i, p in enumerate(pa) if p.part == part and i in targets]
        sb = [(p.el.s, p.el.e) for i, p in enumerate(pb) if p.part == part and i in targets]
        if _cut(xa, sa) != _cut(xb, sb):
            pr.unexpected.append(f"{part}：書いた欄のある段落の外が変わっています")
        names = re.findall(rb'<(?:\w+:)?bookmark\b[^>]*?\bname="([^"]*)"', xb)
        if len(set(names)) != len(names):
            pr.lint_errors.append(f"{part}：しおりの名前が重なっています")
    for i, items in targets.items():
        p_old, p_new = pa[i], pb[i]
        pr.paragraphs_changed += 1
        xa, xb = a.pkg.xml(p_old.part), b.pkg.xml(p_new.part)
        if items:
            want = _expected(p_old, items)
            if want is not None and p_new.text != want:
                pr.unexpected.append(f"段落 {i + 1}：書いた後の字が「{p_new.text[:60]}」で、思っていた「{want[:60]}」と違います")
        ta, tb = _tags(p_old.el), _tags(p_new.el)
        added = {k for k, n in (tb - ta).items()}
        dropped = {k for k, n in (ta - tb).items()}
        if not added <= ADD_OK:
            pr.unexpected.append(f"段落 {i + 1}：元にない要素（{', '.join(sorted(added - ADD_OK))}）が足されています")
        if not dropped <= DROP_OK:
            pr.unexpected.append(f"段落 {i + 1}：元の要素（{', '.join(sorted(dropped - DROP_OK))}）が除かれています")
        ba = set(re.findall(rb'bookmark\b[^>]*?\bname="([^"]*)"', xa[p_old.el.s:p_old.el.e]))
        bb = set(re.findall(rb'bookmark\b[^>]*?\bname="([^"]*)"', xb[p_new.el.s:p_new.el.e]))
        if not ba <= bb or any(not n.startswith(MARK_PREFIX.encode()) for n in bb - ba):
            pr.unexpected.append(f"段落 {i + 1}：しおりが、書き込みの印（_todoku_…）のほかに変わっています")
        ca = set(re.findall(rb'charPrIDRef="([^"]*)"', xa[p_old.el.s:p_old.el.e]))
        cb = set(re.findall(rb'charPrIDRef="([^"]*)"', xb[p_new.el.s:p_new.el.e]))
        if not cb <= (ca or {b"0"}):
            pr.unexpected.append(f"段落 {i + 1}：元の段落にない文字の形（{', '.join(x.decode() for x in cb - ca)}）が使われています")
        fa = [_strip_value(xa[x.s:x.se], ["dirty"]) for x in p_old.el.iter() if L(x) == "fieldBegin"]
        fb = [_strip_value(xb[x.s:x.se], ["dirty"]) for x in p_new.el.iter() if L(x) == "fieldBegin"]
        if fa != fb:
            pr.unexpected.append(f"段落 {i + 1}：누름틀・字の欄の印（fieldBegin）が、dirty のほかに変わっています")
        ga = [(_strip_value(xa[x.s:x.se], ["value", "selectedValue"]), L(x)) for x in p_old.el.iter() if L(x) in FORM_TAGS]
        gb = [(_strip_value(xb[x.s:x.se], ["value", "selectedValue"]), L(x)) for x in p_new.el.iter() if L(x) in FORM_TAGS]
        if ga != gb:
            pr.unexpected.append(f"段落 {i + 1}：フォームの部品が、値のほかに変わっています")
        if i in form_paras and not items and _plain_no_forms(p_old) != _plain_no_forms(p_new):
            pr.unexpected.append(f"段落 {i + 1}：フォームの部品のほかの字が変わっています")
    for i, (x, y) in enumerate(zip(pa, pb)):
        if i not in targets and _plain(x) != _plain(y):
            pr.unexpected.append(f"書いていない段落 {i + 1} の字が変わっています（「{x.text[:30]}」→「{y.text[:30]}」）")
    return pr


def _plain_no_forms(p) -> str:
    return "".join(sg.text for sg in p.segs if not sg.hidden and not (sg.node is not None and sg.node.kind == "ctl")
                   and not (sg.field is not None and sg.field.source in ("hwpx-fixed",) and sg.node is not None
                            and L(sg.node.el) in FORM_TAGS))
