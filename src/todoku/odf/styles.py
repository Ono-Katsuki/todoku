"""OpenDocument のスタイル（ODF 1.3 第3部 16 章）：自動スタイル（部品ごと）・共通スタイル・既定のスタイル・数の書式・リストのスタイル。

性質は、スタイルの親（style:parent-style-name）をたどり、最後に既定のスタイル（style:default-style）を見る。
content.xml の要素は content.xml の自動スタイルと共通スタイルを、styles.xml の要素（ヘッダーなど）は styles.xml の
自動スタイルと共通スタイルを使う（同じ名前の自動スタイルが、部品ごとに別にありうる）。
"""

from __future__ import annotations

from .. import symfont
from .package import FLAT, OdfPackage
from .xml import NS, El, q, top


def _family(v: str) -> str:
    """書体の名前（svg:font-family・fo:font-family）の 1 つ目：引用符を外す。"""
    return (v or "").split(",")[0].strip().strip("'\"").strip()


class Styles:
    def __init__(self, pkg: OdfPackage):
        self.pkg = pkg
        self.auto: dict[str, dict[tuple[str, str], El]] = {}  # 部品 → (family, 名前) → style:style
        self.common: dict[tuple[str, str], El] = {}
        self.defaults: dict[str, El] = {}
        self.numbers: dict[str, dict[str, El]] = {}  # 部品 → 名前 → number:*-style
        self.lists: dict[str, dict[str, El]] = {}  # 部品 → 名前 → text:list-style
        self.data: dict[str, bytes] = {}
        self.outline: El | None = None
        # 書体の宣言（style:font-face）の名前 → (書体の名前（svg:font-family の 1 つ目）, 「記号の書体」と宣言したか（style:font-charset="x-symbol"）)
        # 同じ書体の名前に、記号の書体の宣言と、そうでない宣言の 2 つがあることがある：宣言の名前ごとに持つ
        self.faces: dict[str, tuple[str, bool]] = {}
        parts = [FLAT] if pkg.flat else [p for p in ("styles.xml", "content.xml") if pkg.has(p)]
        for part in parts:
            root = pkg.tree(part)
            if root is None:
                continue
            self.data[part] = pkg.xml(part)
            doc = top(root)
            self.auto[part] = {}
            self.numbers.setdefault(part, {})
            self.lists.setdefault(part, {})
            decls = doc.find("office:font-face-decls")
            for el in (decls.els() if decls is not None else []):
                if el.tag == q("style:font-face") and el.get("style:name"):
                    fam = _family(el.get("svg:font-family") or el.get("style:name"))
                    self.faces.setdefault(el.get("style:name"), (fam, el.get("style:font-charset") == "x-symbol"))
            for sec_name in ("styles", "automatic-styles"):
                sec = doc.find("office:" + sec_name)
                if sec is None:
                    continue
                for el in sec.els():
                    loc = el.local
                    if el.tag == q("style:style"):
                        key = (el.get("style:family", ""), el.get("style:name", ""))
                        if sec_name == "styles":
                            self.common[key] = el
                        else:
                            self.auto[part][key] = el
                    elif el.tag == q("style:default-style"):
                        self.defaults[el.get("style:family", "")] = el
                    elif el.tag.startswith("{" + NS["number"] + "}") and loc.endswith("-style"):
                        self.numbers[part][el.get("style:name", "")] = el
                    elif el.tag == q("text:list-style"):
                        self.lists[part][el.get("style:name", "")] = el
                    elif el.tag == q("text:outline-style"):
                        self.outline = el

        self._sym_memo: dict = {}

    def _part(self, part: str) -> str:
        return FLAT if self.pkg.flat else part

    def face(self, tp: El) -> str | None:
        """字の書式（style:text-properties）の決める書体が、記号の書体（対応表のある書体か、文書が記号の書体と宣言した書体）なら、
        その名前。記号の書体でなければ ""。書体を決めていなければ None。
        書体は style:font-name（書体の宣言の名前）か fo:font-family。記号の書体の宣言は、書体の宣言か、字の書式そのものの
        style:font-charset="x-symbol"。"""
        declared = tp.get("style:font-charset") == "x-symbol"
        name = tp.get("style:font-name")
        if name is not None:
            fam, sym = self.faces.get(name, (name, False))
            declared = declared or sym
        else:
            fam = tp.get("fo:font-family")
            if fam is None:
                return None
            fam = _family(fam)
        return fam if declared or symfont.known(fam) else ""

    def symbol_font(self, part: str, family: str, name: str | None) -> str | None:
        """スタイル（親・既定のスタイルもたどる）の決める書体が、記号の書体（対応表のある書体か、文書が記号の書体と宣言した書体）なら、
        その名前。記号の書体でなければ ""。どこも書体を決めていなければ None。"""
        key = (self._part(part), family, name)
        if key in self._sym_memo:
            return self._sym_memo[key]
        got = None
        els = list(self.chain(part, family, name)) if name else []
        if self.defaults.get(family) is not None:
            els.append(self.defaults[family])
        tpq = q("style:text-properties")
        for el in els:
            for pe in el.kids:
                if isinstance(pe, El) and pe.tag == tpq:
                    got = self.face(pe)
                    if got is not None:
                        break
            if got is not None:
                break
        self._sym_memo[key] = got
        return got

    def get(self, part: str, family: str, name: str) -> El | None:
        part = self._part(part)
        el = self.auto.get(part, {}).get((family, name))
        if el is None:
            el = self.common.get((family, name))
        return el

    def chain(self, part: str, family: str, name: str | None):
        """スタイルと、その親をたどった並び。"""
        seen = set()
        while name and name not in seen:
            seen.add(name)
            el = self.get(part, family, name)
            if el is None:
                break
            yield el
            name = el.get("style:parent-style-name")

    def prop(self, part: str, family: str, name: str | None, kind: str, attr: str) -> str | None:
        """性質（kind は「style:text-properties」など、attr は「fo:border」など）。親・既定のスタイルの順に見る。"""
        k, a = q(kind), q(attr)
        for el in self.chain(part, family, name):
            for pe in el.kids:
                if isinstance(pe, El) and pe.tag == k and a in pe.attrs:
                    return pe.attrs[a]
        d = self.defaults.get(family)
        if d is not None:
            for pe in d.kids:
                if isinstance(pe, El) and pe.tag == k and a in pe.attrs:
                    return pe.attrs[a]
        return None

    def attr(self, part: str, family: str, name: str | None, attr: str) -> str | None:
        """スタイルの要素そのものの属性（style:data-style-name・style:list-style-name など）。親もたどる。"""
        a = q(attr)
        for el in self.chain(part, family, name):
            if a in el.attrs:
                return el.attrs[a]
        return None

    def number_style(self, part: str, name: str | None) -> El | None:
        if not name:
            return None
        part = self._part(part)
        return self.numbers.get(part, {}).get(name) or self.numbers.get(self._part("styles.xml"), {}).get(name)

    def list_style(self, part: str, name: str | None) -> El | None:
        if not name:
            return None
        part = self._part(part)
        return self.lists.get(part, {}).get(name) or self.lists.get(self._part("styles.xml"), {}).get(name)

    # ------------------------------------------------------------------ よく使う性質
    def underline(self, part: str, family: str, name: str | None) -> bool:
        u = self.prop(part, family, name, "style:text-properties", "style:text-underline-style")
        return u not in (None, "none")

    def hidden(self, part: str, family: str, name: str | None) -> bool:
        return self.prop(part, family, name, "style:text-properties", "text:display") == "none"
