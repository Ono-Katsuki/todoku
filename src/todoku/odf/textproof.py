"""OpenDocument の文書の書き込みの証明：変わったのは、書いた欄のある段落（と、書いたフォームの部品の定義）だけであることを、
書く前と後の両方を読み直して確かめる。

- パッケージの部品：書き換えた部品（content.xml・styles.xml）以外は、圧縮されたバイト列まで同じ
- 書き換えた部品：書いた段落（と部品の定義）を除いたバイト列が同じ。標準の XML の読み手（expat）で読める
- 書いた段落：字が「元の字の空欄を、書いた値に置き換えたもの」と同じ。使っている書式（スタイルの名前）は、元の段落にあったものだけ
- 書いていない段落：字が同じ。段落の数が同じ。しおりの名前が重ならない
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field

from ..docx.document import iter_paras
from ..docx.proof import _cut, _expected_text
from ..docx.window import Spots, WriteError, resolve_write
from .package import OdfPackage
from .text import OdtDocument
from .xml import well_formed


@dataclass
class OdfProof:
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
    """部品の並びと、書き換えてよい部品のほかが、圧縮されたバイト列まで同じか。変わった部品の名前を返す。"""
    pa, pb = OdfPackage(original), OdfPackage(new)
    if pa.flat or pb.flat:
        pr.entries_total = pr.entries_identical = 1 if original == new else 0
        pr.entries_total = 1
        return ["(1つの XML)"] if original != new else []
    if pa.names != pb.names:
        pr.unexpected.append("パッケージの部品の並びが変わっています")
        return []
    changed = []
    for name in pa.names:
        pr.entries_total += 1
        ea, eb = pa.zip.get(name), pb.zip.get(name)
        if ea.data == eb.data and ea.read() == eb.read():
            pr.entries_identical += 1
            continue
        changed.append(name)
        if name not in allowed:
            pr.unexpected.append(f"書き換えない部品 {name} が変わっています")
    if pa.zip.names()[:1] == ["mimetype"] and (pb.zip.get("mimetype").method != 0 or pb.zip.names()[0] != "mimetype"):
        pr.lint_errors.append("mimetype の部品が、先頭で圧縮しないで置かれていません（ODF 1.3 第2部 3.3）")
    return changed


def prove(original: bytes, new: bytes, written: dict) -> OdfProof:
    pr = OdfProof(hashlib.sha256(original).hexdigest(), hashlib.sha256(new).hexdigest())
    changed = package_check(pr, original, new, {"content.xml", "styles.xml"})
    try:
        a = OdtDocument(original)
        b = OdtDocument(new)
    except Exception as e:  # noqa: BLE001
        pr.lint_errors.append(f"書き込み後の文書を読めません: {e}")
        return pr
    for part in changed:
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
    form_a: list = []  # 書いたフォームの部品の、書く前の要素
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
        if c is not None and c.source in ("odf-form", "odf-radio"):
            form_a += list(getattr(c, "radios", None) or [c.node])
            pr.controls_changed += 1
            continue
        p = it.get("para") if c is None else c.para
        if p is None or id(p) not in index:
            pr.unexpected.append(f"書いた欄 {fid} の段落が見つかりません")
            continue
        targets.setdefault(index[id(p)], []).append((fid, it, value))
    # 部品ごとに、書いた段落と部品の定義を除いたバイト列を比べる
    added_styles: set[str] = set()
    for part in changed:
        xa, xb = a.pkg.xml(part), b.pkg.xml(part)
        sa = [(p.el.s, p.el.e) for i, p in enumerate(pa) if p.part == part and i in targets and getattr(p, "el", None) is not None]
        sb = [(p.el.s, p.el.e) for i, p in enumerate(pb) if p.part == part and i in targets and getattr(p, "el", None) is not None]
        # 書いた字の書式を打ち消す自動スタイル（textfill.reset_style）：office:automatic-styles の終わりに足しただけで、
        # 打ち消す書式（LOOK_RESET）のほかを持たないこと
        st = _added_styles(a, b, part, pr)
        if st:
            span_b = (st[0][1], st[-1][2])
            sa.append((span_b[0], span_b[0]))
            sb.append(span_b)
            added_styles |= {n for n, _s, _e in st}
        if form_a and part == a.part:
            fa = _form_elements(a)
            fb = _form_elements(b)
            keep = {id(x) for x in form_a}
            idx = [k for k, x in enumerate(fa) if id(x) in keep]
            sa += [(fa[k].s, fa[k].e) for k in idx]
            sb += [(fb[k].s, fb[k].e) for k in idx if k < len(fb)]
        if _cut(xa, sa) != _cut(xb, sb):
            pr.unexpected.append(f"{part}：書いた欄のある段落（と部品）の外が変わっています")
        names = re.findall(rb'bookmark-start\b[^>]*?:name="([^"]*)"', xb)
        if len(set(names)) != len(names):
            pr.lint_errors.append(f"{part}：しおりの名前が重なっています")
    for i, items in targets.items():
        p_old, p_new = pa[i], pb[i]
        pr.paragraphs_changed += 1
        want = _expected(p_old, items)
        if want is not None and p_new.text != want:
            pr.unexpected.append(f"段落 {i + 1}：書いた後の字が「{p_new.text[:60]}」で、思っていた「{want[:60]}」と違います")
        st_a = set(re.findall(rb':style-name="([^"]*)"', a.pkg.xml(p_old.part)[p_old.el.s:p_old.el.e]))
        st_b = set(re.findall(rb':style-name="([^"]*)"', b.pkg.xml(p_new.part)[p_new.el.s:p_new.el.e]))
        st_a |= {n.encode() for n in added_styles}
        if not st_b <= st_a:
            pr.unexpected.append(f"段落 {i + 1}：元の段落にない書式（{', '.join(x.decode('utf-8', 'replace') for x in st_b - st_a)}）が使われています")
    for i, (x, y) in enumerate(zip(pa, pb)):
        if i not in targets and _plain(x) != _plain(y):
            pr.unexpected.append(f"書いていない段落 {i + 1} の字が変わっています（「{x.text[:30]}」→「{y.text[:30]}」）")
    return pr


def _added_styles(a: OdtDocument, b: OdtDocument, part: str, pr: OdfProof) -> list[tuple[str, int, int]]:
    """書いた後の部品の office:automatic-styles に足された style:style（名前・始まり・終わり）。足されたものは、続けて
    終わりに置かれ、名前が todoku_w…、字の書式を打ち消す属性（LOOK_RESET）だけを持つこと。"""
    from .textfill import LOOK_RESET, STYLE_PREFIX
    from .xml import NS, top

    def autos(doc):
        try:
            sec = top(doc.pkg.tree(part)).find("office:automatic-styles")
        except Exception:  # noqa: BLE001
            return None, []
        return sec, ([el for el in sec.els()] if sec is not None else [])

    sa, ea = autos(a)
    sb, eb = autos(b)
    if sb is None or sa is None or len(eb) <= len(ea):
        return []
    added = eb[len(ea):]
    ok = {(qn, v) for resets in LOOK_RESET.values() for qn, v in resets}
    out = []
    for el in added:
        name = el.get("style:name") or ""
        kids = el.els()
        props = kids[0] if len(kids) == 1 and kids[0].tag == "{" + NS["style"] + "}text-properties" else None
        bad = (el.tag != "{" + NS["style"] + "}style" or not name.startswith(STYLE_PREFIX) or el.get("style:family") != "text"
               or props is None or props.els() or len(el.attrs) != 2 or not props.attrs)
        if not bad:
            for k, v in props.attrs.items():
                ns, local = k[1:].split("}", 1)
                pre = next((p for p, u in NS.items() if u == ns), None)
                if (f"{pre}:{local}", v) not in ok:
                    bad = True
        if bad:
            pr.unexpected.append(f"{part}：足したスタイル「{name}」が、書いた字の書式を打ち消すだけのものではありません")
        out.append((name, el.s, el.e))
    if [x.s for x in eb[:len(ea)]] != [x.s for x in ea] or added[-1].e != sb.es:
        pr.unexpected.append(f"{part}：スタイルが、自動スタイルの終わりに足したのでない形で変わっています")
    return out


def _plain(p) -> str:
    """段落の字（フォームの部品の印 ☐☑◉○ と部品の値は除く。部品の状態は、部品の定義のバイト列で比べる）。"""
    return "".join(sg.text for sg in p.segs if not sg.hidden and not (sg.node is not None and sg.node.kind == "ctl"))


def _form_elements(doc: OdtDocument) -> list:
    forms = doc.office_text.find("office:forms")
    if forms is None:
        return []
    return [el for f in forms.iter("form:form") for el in f.els() if el.local not in ("form", "properties", "event-listeners")]


def _expected(p, items) -> str | None:
    """書いた段落の、書いた後の字。入力の欄（text:placeholder など）は、その欄の字を値に置き換えたもの。"""
    plain = [(fid, it, v) for fid, it, v in items if it.get("control") is None]
    ctl = [(fid, it, v) for fid, it, v in items if it.get("control") is not None]
    if not ctl:
        return _expected_text(p, plain)
    text = p.text
    edits = []
    for _fid, it, v in ctl:
        c = it["control"]
        if c.kind == "checkbox":
            return None
        pos = 0
        a = b = None
        for sg in p.segs:
            if sg.hidden:
                continue
            if sg.field is c:
                a = pos if a is None else a
                b = pos + len(sg.text)
            pos += len(sg.text)
        if a is None:
            return None
        edits.append((a, b, str(v)))
    for _fid, it, v in plain:
        if it["kind"] != "inline" or "span" not in it:
            return None
        a, b = it["span"]
        edits.append((a, b, str(v)))
    for a, b, v in sorted(edits, reverse=True):
        text = text[:a] + v + text[b:]
    return text
