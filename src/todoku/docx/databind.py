"""コンテンツコントロールのデータの結び付け（w:dataBinding。ECMA-376 第1部 17.5.2.6）。

結び付けたコンテンツコントロールは、Word が開くときに、文書の中の字（w:sdtContent）でなく、ユーザー定義の XML の部品
（customXml/itemN.xml。w:storeItemID が customXml/itemPropsN.xml の ds:itemID と同じもの）の、w:xpath の指す所の値で
表示し直す。だから書くときは、その XML の値も書く（書かないと、Word で開いたとき書いた値が消える）。

- 読む：w:xpath（w:prefixMappings の接頭辞）を、その XML の部品でたどり、指す所の値を読む
- 書く：指す所の字（要素の中身・属性の値）のバイトだけを変える（ほかのバイトはそのまま）
- たどれる XPath は、/接頭辞:名前[番号]/… と、終わりの @属性 だけ（Word の作る形）。ほかの形・たどれない・文字コードが
  UTF-8 でない部品は、推測で書かずに、理由をつけて断る（BindError）
"""

from __future__ import annotations

import datetime as dt
import re
import xml.parsers.expat as expat
from dataclasses import dataclass

from .xmlpos import _tag_end

DS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"


class BindError(ValueError):
    pass


@dataclass
class Binding:
    xpath: str
    store: str  # w:storeItemID（{…}）。ないときは ""
    prefixes: dict[str, str]
    map_as: str = ""  # 日付の欄：w:storeMappedDataAs（text / date / dateTime）


@dataclass
class Target:
    part: str  # customXml/itemN.xml
    kind: str  # element / attr
    s: int  # 値のバイトの範囲（要素の中身・属性の値。空要素 <x/> は要素全体）
    e: int
    value: str
    empty_tag: bytes = b""  # 空要素（<x/>）のとき、その要素の名前（書くときに <x>値</x> にする）
    quote: bytes = b'"'


def binding_of(ctl) -> Binding | None:
    """コンテンツコントロールの w:dataBinding（なければ None）。"""
    if getattr(ctl, "source", "") != "sdt":
        return None
    pr = ctl.node.find("sdtPr")
    db = pr.find("dataBinding") if pr is not None else None
    if db is None:
        return None
    pm = db.get("prefixMappings", "") or ""
    prefixes = {m[1]: m[3] for m in re.finditer(r"xmlns:([\w.-]+)\s*=\s*(['\"])(.*?)\2", pm)}
    map_as = ""
    if pr is not None:
        sm = pr.find("date")
        sm = sm.find("storeMappedDataAs") if sm is not None else None
        map_as = sm.get("val", "") if sm is not None else ""
    return Binding(db.get("xpath", "") or "", db.get("storeItemID", "") or "", prefixes, map_as)


def _norm_id(s: str) -> str:
    return s.strip().strip("{}").lower()


def store_parts(zipf) -> dict[str, str]:
    """ユーザー定義の XML の部品：ds:itemID（小文字、括弧なし） → 部品の名前（customXml/itemN.xml）。
    部品と性質の部品（itemPropsN.xml）の結び付きは、部品の関係（customXml/_rels/itemN.xml.rels）でたどる。"""
    names = set(zipf.names())
    out: dict[str, str] = {}
    for rp in names:
        m = re.fullmatch(r"(.*/)?_rels/([^/]+)\.rels", rp)
        if not m or not (m[1] or "").lower().startswith("customxml/"):
            continue
        item = (m[1] or "") + m[2]
        if item not in names:
            continue
        for t in re.findall(rb'Target="([^"]+)"', zipf.read(rp)):
            target = t.decode("utf-8", "replace").lstrip("/")
            props = target if target.lower().startswith("customxml/") else (m[1] or "") + target
            if props not in names:
                continue
            got = re.search(rb'itemID="([^"]+)"', zipf.read(props))
            if got:
                out[_norm_id(got[1].decode("utf-8", "replace"))] = item
    return out


def property_parts(zipf) -> list[str]:
    """文書のプロパティの部品（docProps/core.xml・docProps/app.xml）：パッケージの関係（_rels/.rels）の種類が
    core-properties・extended-properties のもの（ECMA-376 第2部 11・第1部 15.2.12.3）。Word は、題名・作成者・会社などの
    プロパティを、ユーザー定義の XML の部品と同じに、結び付けの先として見せる（決まった w:storeItemID で指す）。"""
    names = set(zipf.names())
    if "_rels/.rels" not in names:
        return []
    out = []
    for m in re.finditer(rb"<(?:[\w.-]+:)?Relationship\b[^>]*>", zipf.read("_rels/.rels")):
        kind = re.search(rb'\bType="([^"]*)"', m[0])
        target = re.search(rb'\bTarget="([^"]*)"', m[0])
        if not kind or not target or re.search(rb'\bTargetMode="External"', m[0]):
            continue
        if kind[1].rsplit(b"/", 1)[-1] not in (b"core-properties", b"extended-properties"):
            continue
        part = target[1].decode("utf-8", "replace").lstrip("/")
        if part in names and part not in out:
            out.append(part)
    return out


# ------------------------------------------------------------------ XPath
_STEP = re.compile(r"(?:([\w.-]+):)?([\w.-]+)(?:\[(\d+)\])?")


def _steps(xpath: str) -> tuple[list[tuple[str | None, str, int | None]], tuple[str | None, str] | None]:
    x = xpath.strip()
    if not x.startswith("/") or x.startswith("//"):
        raise BindError(f"XPath「{xpath}」は、根からたどる形（/名前[番号]/…）ではありません")
    parts = x[1:].split("/")
    attr = None
    if parts and parts[-1].startswith("@"):
        m = re.fullmatch(r"@(?:([\w.-]+):)?([\w.-]+)", parts[-1])
        if not m:
            raise BindError(f"XPath「{xpath}」の属性の指定を読めません")
        attr = (m[1], m[2])
        parts = parts[:-1]
    steps = []
    for p in parts:
        m = _STEP.fullmatch(p)
        if not m:
            raise BindError(f"XPath「{xpath}」の「{p}」は、確かめられる形（名前[番号]）ではありません")
        steps.append((m[1], m[2], int(m[3]) if m[3] else None))
    if not steps:
        raise BindError(f"XPath「{xpath}」が要素を指していません")
    return steps, attr


class _N:
    __slots__ = ("tag", "attrs", "s", "se", "es", "e", "kids", "text", "raw_attrs")

    def __init__(self, tag, attrs, s):
        self.tag, self.attrs, self.s = tag, attrs, s
        self.se = self.es = self.e = -1
        self.kids: list[_N] = []
        self.text = ""


def _parse(xml: bytes) -> _N:
    p = expat.ParserCreate(namespace_separator="}")
    p.buffer_text = True
    root = _N("#root", {}, 0)
    stack = [root]

    def start(tag, attrs):
        s = p.CurrentByteIndex
        n = _N(tag, attrs, s)
        n.se = _tag_end(xml, s)
        stack[-1].kids.append(n)
        stack.append(n)

    def end(tag):
        n = stack.pop()
        if xml[n.se - 2:n.se] == b"/>":
            n.es = n.e = n.se
        else:
            n.es = p.CurrentByteIndex
            n.e = _tag_end(xml, n.es)

    def chars(data):
        stack[-1].text += data

    p.StartElementHandler = start
    p.EndElementHandler = end
    p.CharacterDataHandler = chars
    p.Parse(xml, True)
    return root


def _is_utf8(xml: bytes) -> bool:
    if xml[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return False
    m = re.match(rb"(?:\xef\xbb\xbf)?\s*<\?xml[^>]*?encoding\s*=\s*['\"]([^'\"]+)['\"]", xml)
    return m is None or m[1].decode("ascii", "replace").lower().replace("_", "-") in ("utf-8", "utf8")


def resolve(xml: bytes, b: Binding, part: str) -> Target:
    """w:xpath の指す所（値とバイトの範囲）。たどれなければ BindError。"""
    if not _is_utf8(xml):
        raise BindError(f"ユーザー定義の XML の部品 {part} の文字コードが UTF-8 でないので、値をバイトのまま書き換えられません")
    try:
        root = _parse(xml)
    except expat.ExpatError as e:
        raise BindError(f"ユーザー定義の XML の部品 {part} を XML として読めません（{e}）") from None
    steps, attr = _steps(b.xpath)
    cur = root
    for pre, local, idx in steps:
        if pre is not None and pre not in b.prefixes:
            raise BindError(f"XPath「{b.xpath}」の接頭辞「{pre}」が、w:prefixMappings にありません")
        want = (b.prefixes[pre] + "}" + local) if pre is not None else local
        hits = [k for k in cur.kids if k.tag == want]
        if idx is None:
            if len(hits) != 1:
                raise BindError(f"XPath「{b.xpath}」の「{local}」に当たる要素が {len(hits)} 個あり、どれか決まりません")
            cur = hits[0]
        else:
            if not (1 <= idx <= len(hits)):
                raise BindError(f"XPath「{b.xpath}」の指す要素が、{part} にありません")
            cur = hits[idx - 1]
    if attr is not None:
        pre, local = attr
        if pre is not None and pre not in b.prefixes:
            raise BindError(f"XPath「{b.xpath}」の接頭辞「{pre}」が、w:prefixMappings にありません")
        key = (b.prefixes[pre] + "}" + local) if pre is not None else local
        if key not in cur.attrs:
            raise BindError(f"XPath「{b.xpath}」の指す属性が、{part} にありません")
        tag = xml[cur.s:cur.se]
        if pre is None:
            qn = rb""
        else:
            names = sorted({k for k, v in _decls(xml).items() if v == b.prefixes[pre]})
            if not names:
                raise BindError(f"XPath「{b.xpath}」の属性の接頭辞を、{part} の中で見つけられません")
            qn = rb"(?:" + rb"|".join(re.escape(n.encode()) for n in names) + rb"):"
        m = list(re.finditer(rb"(?<![\w.:-])" + qn + re.escape(local.encode()) + rb"\s*=\s*(['\"])(.*?)\1", tag, re.S))
        if len(m) != 1:
            raise BindError(f"XPath「{b.xpath}」の指す属性を、{part} の中で1つに決められません")
        s = cur.s + m[0].start(2)
        return Target(part, "attr", s, s + len(m[0][2]), cur.attrs[key], quote=m[0][1])
    if cur.kids:
        raise BindError(f"XPath「{b.xpath}」の指す要素に子の要素があり、値の字だけを書き換えられません")
    inner = xml[cur.se:cur.es]
    if b"<!--" in inner or b"<?" in inner:
        raise BindError(f"XPath「{b.xpath}」の指す要素の中に、注釈・処理命令があり、値の字だけを書き換えられません")
    if cur.se == cur.e:
        name = re.match(rb"<([^\s/>]+)", xml[cur.s:cur.se])[1]
        return Target(part, "element", cur.s, cur.e, "", empty_tag=name)
    return Target(part, "element", cur.se, cur.es, cur.text)


def _decls(tag: bytes) -> dict[str, str]:
    return {m[1].decode(): m[3].decode("utf-8", "replace")
            for m in re.finditer(rb"xmlns:([\w.-]+)\s*=\s*(['\"])(.*?)\2", tag)}


def locate(doc, ctl) -> Target | None:
    """結び付いた値の所（結び付けがなければ None）。たどれなければ BindError。"""
    b = binding_of(ctl)
    if b is None:
        return None
    stores = store_parts(doc.zip)
    if b.store:
        part = stores.get(_norm_id(b.store))
        if part is None:
            # ユーザー定義の XML の部品にない番号：文書のプロパティの部品を指すことがある。番号の表は持たず、XPath（要素の名前と
            # 名前空間）でたどれるプロパティの部品が 1 つだけのときに、その部品と決める
            found = []
            for prop in property_parts(doc.zip):
                try:
                    found.append(resolve(doc.zip.read(prop), b, prop))
                except BindError:
                    continue
            if len(found) != 1:
                raise BindError(f"結び付けた XML の部品（w:storeItemID {b.store}）が、文書の中にありません")
            return found[0]
        return resolve(doc.zip.read(part), b, part)
    # w:storeItemID のない結び付け：XPath のたどれる部品が1つだけのときに決める
    found = []
    for part in sorted(set(stores.values())):
        try:
            found.append(resolve(doc.zip.read(part), b, part))
        except BindError:
            continue
    if len(found) != 1:
        raise BindError(f"結び付けた XML の部品（w:storeItemID なし）が、{len(found)} 個の部品に当たり、決まりません")
    return found[0]


def edit(xml: bytes, t: Target, value: str) -> bytes:
    """値の字だけを変えた部品のバイト列。"""
    v = value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\r", "&#13;")
    if t.kind == "attr":
        v = v.replace(t.quote.decode(), "&quot;" if t.quote == b'"' else "&apos;").replace("\n", "&#10;").replace("\t", "&#9;")
        return xml[:t.s] + v.encode("utf-8") + xml[t.e:]
    if t.empty_tag:
        if not value:
            return xml
        tag = xml[t.s:t.e]
        return xml[:t.s] + tag[:-2].rstrip() + b">" + v.encode("utf-8") + b"</" + t.empty_tag + b">" + xml[t.e:]
    return xml[:t.s] + v.encode("utf-8") + xml[t.e:]


# ------------------------------------------------------------------ 値の形（欄の種類ごと）
def _listitems(ctl) -> list[tuple[str, str]]:
    pr = ctl.node.find("sdtPr")
    out = []
    for k in pr.kids if pr is not None else []:
        if k.local in ("dropDownList", "comboBox"):
            for li in k.findall("listItem"):
                d, v = li.get("displayText"), li.get("value")
                out.append(((d if d is not None else v) or "", (v if v is not None else d) or ""))
    return out


def stored_value(ctl, shown, b: Binding, current: str, display: str = "") -> str:
    """欄に書く値 shown（表示する字。チェックの欄は bool、日付は datetime.date）を、結び付けた XML に書く値にする。
    display：日付の欄の、書式を通した表示の字（w:storeMappedDataAs が text のとき、その字を書く）。"""
    if ctl.kind == "checkbox":
        if current.strip() in ("0", "1"):
            return "1" if shown else "0"
        return "true" if shown else "false"
    if ctl.kind in ("dropdown", "combo"):
        for d, v in _listitems(ctl):
            if d == shown:
                return v
        if ctl.kind == "combo":
            return str(shown)
        raise BindError(f"「{shown}」は選択肢にありません")
    if ctl.kind == "date":
        d: dt.date = shown
        how = b.map_as or _date_form(current)
        if how == "dateTime":
            return d.strftime("%Y-%m-%dT00:00:00") + ("Z" if current.strip().endswith("Z") else "")
        if how == "date":
            return d.strftime("%Y-%m-%d")
        if how == "text":
            return display
        raise BindError("結び付けた日付の欄の値の形（w:storeMappedDataAs）が決まりません")
    if ctl.kind == "text":
        return str(shown)
    raise BindError(f"この種類のコンテンツコントロール（{ctl.kind}）の結び付けた値の形を、確かめていません")


def _date_form(current: str) -> str:
    c = current.strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}T[\d:.]+Z?", c):
        return "dateTime"
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", c):
        return "date"
    return ""


def shown_value(ctl, stored: str, b: Binding):
    """結び付けた XML の値 → Word が欄に表示する値（読む窓の値）。読めない形なら BindError。"""
    if ctl.kind == "checkbox":
        t = stored.strip().lower()
        if t in ("true", "1"):
            return True
        if t in ("false", "0", ""):
            return False
        raise BindError(f"結び付けたチェックの欄の値「{stored}」を、はい／いいえとして読めません")
    if ctl.kind in ("dropdown", "combo"):
        for d, v in _listitems(ctl):
            if v == stored:
                return d
        if ctl.kind == "combo" or stored == "":
            return stored
        raise BindError(f"結び付けた値「{stored}」が、選択肢の値にありません")
    if ctl.kind == "date":
        if stored.strip() == "":
            return ""
        how = b.map_as or _date_form(stored)
        if how in ("date", "dateTime"):
            m = re.match(r"(\d{4})-(\d{2})-(\d{2})", stored.strip())
            if not m:
                raise BindError(f"結び付けた日付の値「{stored}」を読めません")
            from .fill import format_word_date

            return format_word_date(dt.date(int(m[1]), int(m[2]), int(m[3])), ctl.date_format or "yyyy/MM/dd",
                                    ctl.date_lid, ctl.date_calendar)
        return stored
    if ctl.kind == "text":
        return stored
    raise BindError(f"この種類のコンテンツコントロール（{ctl.kind}）の結び付けた値を、確かめていません")


def bound_value(doc, ctl):
    """(Word が開いたとき欄に表示する値, 結び付けの問題)。結び付けがなければ (None, None) と Target None。"""
    try:
        t = locate(doc, ctl)
    except BindError as e:
        return None, str(e), None
    if t is None:
        return None, None, None
    try:
        return shown_value(ctl, t.value, binding_of(ctl)), None, t
    except Exception as e:  # noqa: BLE001
        return None, str(e), t
