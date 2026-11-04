"""PowerPoint（.pptx）：エージェントが「全て正確に読み書きできるか」を、決まった処理で確かめる。

読む: 窓とは別の読み方（Python の xml.etree で部品を読み、段落・表を作らずに a:t の字を集めるだけの読み手）で、
      スライド・ノート・スライドの型（プレースホルダーでない図形）の字を数え、窓に1字残らず出ていることを確かめる。
      絵（p:pic・絵で塗った図形）の数と説明、グラフ・SmartArt の字も、XML から直接拾って照らす。
書く: 窓が示したすべての入力欄に試しの値を一度に書き込み、証明（書いた段落のあるスライドの部品のほかは圧縮したバイトまで一致・
      書いた段落の外は一致）と読み戻し（書いた後のファイルを読み直した段落の字）を確かめる。字を目印にした書き込みも試す。
"""

from __future__ import annotations

import posixpath
import re
import xml.etree.ElementTree as ET
from collections import Counter

from .. import chars as _chars
from ..docx.audit import _sample, macro_check
from ..docx.fill import FillError
from ..xlsx.audit import Audit, Check, agent_places_check
from ..zipraw import RawZip
from .document import PptxDocument, PptxError
from .fill import fill, prove, written_of
from .window import build, markdown_parts

# チェックの箱の字（chars：印のない箱・印のついた箱・印のついた箱かもしれない字。どの形式の窓とも同じ字の集まり）
BOXES = _chars.BOXES | _chars.MAYBE_CHECKED
_A = ("{http://schemas.openxmlformats.org/drawingml/2006/main}", "{http://purl.oclc.org/ooxml/drawingml/main}")
_MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _rels(z: RawZip, part: str) -> list[tuple[str, str, str]]:
    """(Id, 種類の終わり, 先の部品)。"""
    d, base = posixpath.split(part)
    rp = posixpath.join(d, "_rels", base + ".rels")
    if not z.has(rp):
        return []
    out = []
    for r in ET.fromstring(z.read(rp)):
        if r.get("TargetMode") == "External":
            continue
        t = r.get("Target", "")
        name = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(d, t))
        out.append((r.get("Id", ""), r.get("Type", "").rsplit("/", 1)[-1], name))
    return out


def _walk(el, template: bool, out: list[str], pics: list[str]) -> None:
    """a:t の字を集める。代わりの中身（mc:Fallback）は読まない。template：プレースホルダーの図形・絵・表を除く。"""
    for ch in el:
        loc = _local(ch.tag)
        if ch.tag == _MC + "Fallback":
            continue
        if template and loc in ("pic", "graphicFrame"):
            continue
        if loc == "sp":
            ph = any(_local(x.tag) == "ph" for x in ch.iter())
            if template and ph:
                continue
            sppr = next((x for x in ch if _local(x.tag) == "spPr"), None)
            if not template and sppr is not None and any(_local(x.tag) == "blipFill" for x in sppr):
                pics.append(_descr(ch))
        if loc == "pic" and not template:
            pics.append(_descr(ch))
        if loc == "oleObj":
            continue  # 埋め込みの代わりの絵（p:pic）は、埋め込みの印として数える
        if any(ch.tag == a + "t" for a in _A):
            out.append(ch.text or "")
            continue
        _walk(ch, template, out, pics)


def _descr(el) -> str:
    c = next((x for x in el.iter() if _local(x.tag) == "cNvPr"), None)
    return ((c.get("descr") or "") if c is not None else "").strip()


def independent_text(blob: bytes) -> tuple[list[str], list[str], list[str]]:
    """(字の並び, 絵の説明, 読んだ部品)：presentation.xml の p:sldIdLst の順のスライドと、その関係のノート・型。"""
    z = RawZip(blob)
    root_rels = ET.fromstring(z.read("_rels/.rels")) if z.has("_rels/.rels") else []
    main = next((r.get("Target", "").lstrip("/") for r in root_rels if r.get("Type", "").endswith("/officeDocument")), "ppt/presentation.xml")
    pres = ET.fromstring(z.read(main))
    prels = {i: n for i, _t, n in _rels(z, main)}
    texts: list[str] = []
    pics: list[str] = []
    parts: list[str] = []
    templates: list[str] = []
    lst = next((x for x in pres if _local(x.tag) == "sldIdLst"), None)
    for sid in (list(lst) if lst is not None else []):
        part = prels.get(sid.get(_R + "id") or "")
        if not part or not z.has(part):
            continue
        sld = ET.fromstring(z.read(part))
        parts.append(part)
        _walk(_tree(sld), False, texts, pics)
        for _i, kind, name in _rels(z, part):
            if kind == "notesSlide" and z.has(name):
                parts.append(name)
                _walk(_tree(ET.fromstring(z.read(name))), False, texts, pics)
            elif kind == "slideLayout" and z.has(name) and sld.get("showMasterSp") not in ("0", "false"):
                templates.append(name)
                lay = ET.fromstring(z.read(name))
                if lay.get("showMasterSp") not in ("0", "false"):
                    templates += [n for _i2, k2, n in _rels(z, name) if k2 == "slideMaster" and z.has(n)]
    for name in dict.fromkeys(templates):
        parts.append(name)
        _walk(_tree(ET.fromstring(z.read(name))), True, texts, pics)
    return texts, pics, parts


def _tree(root):
    for el in root.iter():
        if _local(el.tag) == "spTree":
            return el
    return ET.Element("none")


def _count(texts) -> Counter:
    c = Counter()
    for t in texts:
        for ch in t:
            if not ch.isspace() and ch not in BOXES:
                c[ch] += 1
    return c


def check_reading(doc: PptxDocument, win: dict) -> list[Check]:
    texts, pics, parts = independent_text(doc.blob)
    raw = _count(texts)
    shown = Counter()
    for st in win["stories"]:
        shown += _count(st["_plain"])
    missing = [f"「{ch}」が {n} 字、窓に出ていません" for ch, n in (raw - shown).items()]
    extra = [f"「{ch}」が {n} 字、元にないのに窓に出ています" for ch, n in (shown - raw).items()]
    out = [Check("読み落とし（スライドの字）", not missing and not extra,
                 f"スライド・ノート・スライドの型の部品 {len(parts)} 個の字 {sum(raw.values()):,} 字を、窓とは別の読み方"
                 "（xml.etree で a:t の字を集めるだけの読み手）で数え、すべて窓に出ていることを確かめました", missing + extra)]
    mine = sorted((p.get("alt") or "") for p in win.get("pictures") or [])
    probs = []
    if len(mine) != len(pics):
        probs.append(f"絵の数が違います（ファイル {len(pics)}・窓 {len(mine)}）")
    elif mine != sorted(pics):
        probs.append("絵の説明が、ファイルの cNvPr の descr と違います")
    out.append(Check("読み落とし（絵の説明）", not probs,
                     f"絵 {len(pics)} 個（p:pic と、絵で塗った図形）の説明（descr）を XML から直接拾い、窓の絵の一覧と同じことを確かめました", probs))
    out += _object_checks(doc, win, parts)
    return out


def _object_checks(doc: PptxDocument, win: dict, parts: list[str]) -> list[Check]:
    """グラフ・SmartArt：スライドの部品の c:chart・dgm:relIds から、関係で部品を直接たどり、字が窓に出ているか。"""
    from ..embeds.audit import _missing, _smartart_check, chart_xml_texts
    from ..embeds.word import item_markdown

    z = doc.zip
    objs = win.get("objects") or []
    charts, dgms = [], []
    for part in parts:
        xml = re.sub(rb"<mc:Fallback\b.*?</mc:Fallback>", b"", z.read(part), flags=re.S)
        rels = {i: n for i, _k, n in _rels(z, part)}
        for m in re.finditer(rb'<(?:c|cx):chart\b[^>]*\br:id="([^"]+)"', xml):
            t = rels.get(m[1].decode())
            if t and z.has(t):
                charts.append(t)
        for m in re.finditer(rb'<(?:\w+:)?relIds\b[^>]*\br:dm="([^"]+)"', xml):
            t = rels.get(m[1].decode())
            if t and z.has(t):
                dgms.append(t)
    out = []
    shown_c = [o for o in objs if o["kind"] == "chart"]
    if charts or shown_c:
        texts = [t for p in charts for t in chart_xml_texts(z.read(p))]
        miss = _missing(texts, "\n".join(x for o in shown_c for x in item_markdown(o)))
        ok = not miss and len(charts) == len(shown_c)
        out.append(Check("読み落とし（グラフ）", ok, f"グラフ {len(charts)} 個の字と値 {len(texts):,} 個を、窓とは別の読み方で拾い、窓のグラフの表 "
                         f"{len(shown_c)} 個に" + ("すべて出ていることを確かめました" if ok else "出ていないものがあります"),
                         ([f"グラフの数が違います（ファイル {len(charts)}・窓 {len(shown_c)}）"] if len(charts) != len(shown_c) else []) + miss[:30]))
    if dgms:
        out.append(_smartart_check(z, sorted(set(dgms)), "\n".join(x for o in objs if o["kind"] == "smartart" for x in item_markdown(o))))
    return out


def check_writing(blob: bytes, filename: str, win: dict) -> list[Check]:
    targets = {f["id"]: _sample(f) for f in win["fields"] if not f.get("readonly")}
    try:
        res = fill(PptxDocument(blob, filename), targets)
    except FillError as e:
        return [Check("全入力欄への書き込み", False, f"{len(targets)} 欄に書き込めませんでした", str(e).splitlines()[:30])]
    pr = prove(blob, res.blob, written_of(res.patches))
    return [Check("全入力欄への書き込み", not pr.unexpected and not pr.lint_errors,
                  f"入力欄 {len(targets):,} か所すべてに試しの値を書き込みました。{pr.headline()}", [*pr.unexpected, *pr.lint_errors]),
            Check("読み戻し", not pr.readback,
                  f"書き込み後の文書を読み直し、段落 {pr.paragraphs_total:,} のすべてで、字が「元の字の、書いた所だけを替えたもの」"
                  "になっていることを確かめました" if not pr.readback else "書き込み後の文書を読み直すと、思っていた字と違う段落があります",
                  pr.readback[:30])]


def anchored_check(blob: bytes, filename: str) -> Check:
    """字を目印にした書き込み：文書の中に1つしかない4字の後ろに試しの字を書き、証明と読み戻し（その段落の字）を確かめる。"""
    from ..docx.document import iter_paras
    from ..docx.window import _in_field

    doc = PptxDocument(blob, filename)
    paras = [q for _k, _p, bl in doc.all_blocks() for q in iter_paras(bl)]
    full = "\n".join(q.text for q in paras)
    tried, last = 0, ""
    for n, q in enumerate(paras, 1):
        t = q.text
        pos = q.positions()
        for i in range(0, len(t) - 3):
            frag = t[i:i + 4]
            if frag != frag.strip(" 　\t") or full.count(frag) != 1 or any(_in_field(q, j) for j in range(i, i + 5)):
                continue
            if any(q.segs[pos[k][0]].kind != "t" for k in range(i, min(i + 5, len(pos)))):
                continue
            tried += 1
            try:
                res = fill(PptxDocument(blob, filename), {"@1": {"old": frag, "new": frag + "（試し）"}})
            except FillError as e:
                last = str(e)
                if tried >= 5:
                    break
                continue
            pr = prove(blob, res.blob, written_of(res.patches))
            return Check("字を目印にした書き込み", pr.ok,
                         f"規則で示していない所への書き込み（段落 {n} の「{frag}」の後ろに「（試し）」）を {res.patches[0].id} として書き、"
                         f"{pr.headline()}。読み戻し：{'その段落の字が思っていたとおり' if not pr.readback else '思っていた字と違う'}",
                         [*pr.unexpected, *pr.lint_errors, *pr.readback])
        if tried >= 5:
            break
    if not tried:
        return Check("字を目印にした書き込み", True, "文書の中に1つしかない字がないので、試していません（同じ字の所は near・occurrence で指す）")
    return Check("字を目印にした書き込み", False, f"字を目印にした書き込みを {tried} か所で試し、書けませんでした", [last])


def run(blob: bytes, filename: str = "") -> Audit:
    au = Audit(filename)
    try:
        doc = PptxDocument(blob, filename)
    except PptxError as e:
        au.checks.append(Check("開く", False, f"PowerPoint の文書として読めません：{e}"))
        return au
    win = build(doc)
    au.checks += check_reading(doc, win)
    if doc.unread:
        au.checks.append(Check("読めない部分", False, "中身を読んでいない部分があります", list(doc.unread)))
    parts = markdown_parts(win)
    biggest = max(len(t) for _, t in parts)
    au.checks.append(Check("窓の大きさ", biggest <= 150_000,
                           f"Markdown {sum(len(t) for _, t in parts):,} 文字" + (f"を {len(parts)} 部分に分け、いちばん大きい部分も {biggest:,} 文字"
                                                                             if len(parts) > 1 else "") + "（AIが1回で読める目安 15万文字以内）"))
    au.checks += check_writing(blob, filename, win)
    ap = agent_places_check(win["fields"])
    if ap is not None:
        au.checks.append(ap)
    au.checks.append(anchored_check(blob, filename))
    au.checks.append(macro_check(doc, win))
    return au
