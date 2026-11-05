"""タグ付きの PDF（文書の構造の木がある）で、ページの描画に足した字を、構造の木にも足す。

構造の木を読む読み手（Mac のプレビューの読み上げなど）は、木にない字を読まない（実物で測った）。そこで：
- 足した字を、番号（/MCID）つきの印の範囲にする（pagedraw・circle）
- /Span の要素を足し、見出しの字を持つ要素の中の、その字のすぐ後ろに入れる（読む順が見出しの隣になる）
- ページの番号 → 要素の一覧（ParentTree）に足す。元の要素・一覧の項目は変えない。足すだけ

どの要素の隣に入れるかは、文書の事実だけで決める：
  1. 書いた所の見出しとして窓が読んだ字を持つ範囲のうち、いちばん近いもの（how：見出し。○は囲んだ字）
  2. 同じ行の左にあるいちばん近い字（how：左）
  3. 同じ列の上にあるいちばん近い字（how：上）
  4. ページのいちばん近い字（how：近く）
  5. ページに番号のついた字がない・木のその所がつじつまの合わないとき：文書の要素の終わり（how：文書の終わり。結果で知らせる）
位置だけで決めるので、見出しから離れた要素の隣に入ることがある（結果の structure と warnings で伝える）。1 つの範囲に
見出しと空きが並ぶ行（「令和　年　月　日」）では、値はその範囲の後ろにまとめて入る。範囲の途中には入れられない。
Alt・ActualText を持つ要素（と、その下）や図の中には入れない（読み手が、その説明で中身を置き換えるので、値が読まれなくなる）。
その親の段落のすぐ後ろに入れる。その分、段落の途中の空きの値も、段落の後ろで読まれる。

足せないとき（木の入口が間接参照でない・ParentTree が読めない形・ページの印が閉じていない）は、番号をつけずに描き、理由を返す
（黙って推測で直さない）。
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from .objects import Name, Ref

MAX_DEPTH = 64
OPAQUE = {"Figure", "Formula", "Form", "Artifact"}  # 中の字を読まず、それ自身として読まれる要素


@dataclass
class Placed:
    """足した要素の 1 つ：読み手にどこで読まれるかの事実。"""
    id: str
    page: int
    how: str  # 見出し・左・上・近く・文書の終わり
    near: str  # 隣に入れた要素が持つ字（短く）
    gap_pt: float | None


def _clean(text: str) -> str:
    return re.sub(r"[\s　]", "", text or "")


def rank(r, s) -> tuple[int, float]:
    """箱 s から箱 r への近さ（circle._rank と同じ考え方。同じ行・同じ列・そのほか）。"""
    if min(r[3], s[3]) - max(r[1], s[1]) > 0:
        return 0, max(0.0, max(r[0], s[0]) - min(r[2], s[2]))
    if min(r[2], s[2]) - max(r[0], s[0]) > 0:
        return 1, max(0.0, max(r[1], s[1]) - min(r[3], s[3]))
    dx = max(0.0, max(r[0], s[0]) - min(r[2], s[2]))
    dy = max(0.0, max(r[1], s[1]) - min(r[3], s[3]))
    return 2, (dx * dx + dy * dy) ** 0.5


class Tree:
    """元の文書の構造の木（読むだけ）。足す手順は Adder。"""

    def __init__(self, doc):
        self.doc = doc
        cat = doc.resolve(doc.trailer.get("Root")) or {}
        self.ref = cat.get("StructTreeRoot") if isinstance(cat, dict) else None
        self.root = doc.resolve(self.ref) if self.ref is not None else None
        self.present = isinstance(self.root, dict)
        self.why: str | None = None
        self.entries: dict[int, object] = {}  # ParentTree の鍵 → 値（要素か一覧。解決前のまま）
        self.leaves: list[list] = []  # 葉の、根からの道：[(節の参照か None, 節)...]
        self.rolemap = {}
        if not self.present:
            return
        rm = doc.resolve(self.root.get("RoleMap"))
        self.rolemap = rm if isinstance(rm, dict) else {}
        if not isinstance(self.ref, Ref):
            self.why = "文書の構造の木の入口が、別のオブジェクトになっていません"
            return
        pt = self.root.get("ParentTree")
        if pt is None:
            return  # ParentTree がない：作って足せる（元の項目はないので、失うものがない）
        try:
            self._walk(pt, [], set())
        except ValueError as e:
            self.why = f"文書の構造の木の一覧（ParentTree）が読めません（{e}）"

    def _walk(self, node, path: list, seen: set) -> None:
        d = self.doc.resolve(node)
        if not isinstance(d, dict):
            raise ValueError("節が辞書ではありません")
        if isinstance(node, Ref):
            if node.num in seen or len(seen) > 100000:
                raise ValueError("節が自分を指しています")
            seen.add(node.num)
        elif path:
            raise ValueError("子の節が別のオブジェクトになっていません")
        path = path + [(node if isinstance(node, Ref) else None, d)]
        kids = self.doc.resolve(d.get("Kids"))
        nums = self.doc.resolve(d.get("Nums"))
        if isinstance(kids, list) and kids:
            for k in kids:
                self._walk(k, path, seen)
            return
        if nums is None:
            nums = []
        if not isinstance(nums, list) or len(nums) % 2:
            raise ValueError("Nums が鍵と値の組の一覧ではありません")
        if not isinstance(d.get("Nums"), list) and d.get("Nums") is not None:
            raise ValueError("Nums が別のオブジェクトになっています")
        for i in range(0, len(nums), 2):
            k = self.doc.resolve(nums[i])
            if not isinstance(k, int) or isinstance(k, bool):
                raise ValueError("鍵が整数ではありません")
            self.entries[k] = nums[i + 1]
        self.leaves.append(path)

    def std_type(self, elem: dict) -> str:
        t = self.doc.resolve(elem.get("S"))
        name = t.name if isinstance(t, Name) else ""
        seen = set()
        while name in self.rolemap and name not in seen:
            seen.add(name)
            m = self.doc.resolve(self.rolemap[name])
            if not isinstance(m, Name):
                break
            name = m.name
        return name

    def chain(self, ref) -> list | None:
        """要素から根までの要素（根は含まない）：[(参照, 辞書)...]。根に着かなければ None（つじつまが合わない）。"""
        out = []
        cur = ref
        seen = set()
        while len(out) < MAX_DEPTH:
            if not isinstance(cur, Ref) or cur.num in seen:
                return None
            seen.add(cur.num)
            if cur.num == self.ref.num:
                return out
            d = self.doc.resolve(cur)
            if not isinstance(d, dict):
                return None
            out.append((cur, d))
            cur = d.get("P")
        return None


def tree_of(doc) -> Tree:
    t = doc.__dict__.get("_structadd_tree")
    if t is None:
        t = doc.__dict__["_structadd_tree"] = Tree(doc)
    return t


def open_marks(data: bytes) -> int:
    """描画の終わりで閉じていない印の範囲（BDC・BMC）の数。"""
    from .content import tokens

    d = 0
    for kind, val in tokens(data):
        if kind == "o":
            if val in ("BDC", "BMC"):
                d += 1
            elif val == "EMC" and d > 0:
                d -= 1
    return d


def first_mcid(doc, page: dict) -> tuple[int | None, str | None]:
    """このページで足す字に最初につける番号。返す：(番号 | None, 足せない理由 | None)。タグのない文書は (None, None)。
    書き込みと証明が同じ決め方で使う（元の文書で決まる）。"""
    tree = tree_of(doc)
    if not tree.present:
        return None, None
    if tree.why:
        return None, tree.why
    data = doc.page_content(page)
    if open_marks(data):
        return None, "ページの描画の印の範囲が閉じていません"
    used = [int(m[1]) for m in re.finditer(rb"/MCID\s+(\d+)", data)]
    res = doc.resolve(page.get("Resources"))
    props = doc.resolve(res.get("Properties")) if isinstance(res, dict) else None
    for v in (props.values() if isinstance(props, dict) else ()):
        v = doc.resolve(v)
        n = doc.resolve(v.get("MCID")) if isinstance(v, dict) else None
        if isinstance(n, int):
            used.append(n)
    key = doc.resolve(page.get("StructParents"))
    n = max(used, default=-1) + 1
    if isinstance(key, int) and key in tree.entries:
        arr = doc.resolve(tree.entries[key])
        if not isinstance(arr, list):
            return None, "ページの要素の一覧（ParentTree の値）が一覧ではありません"
        n = max(n, len(arr))
    elif page.get("StructParents") is not None and not isinstance(key, int):
        return None, "ページの StructParents が整数ではありません"
    return n, None


def marked_chars(doc, page: dict) -> list[tuple]:
    """番号のついた印の中の、ページの字：(持ち主の鍵, MCID, 字, 見た目の箱)。持ち主の鍵は ParentTree の鍵（ページか、フォームの部品）。"""
    from .content import Interpreter
    from .layout import View

    view = View(doc, page)
    pc = Interpreter(doc).run_page(page)
    pk = doc.resolve(page.get("StructParents"))
    pk = pk if isinstance(pk, int) else None
    out = []
    for g in pc.glyphs:
        if g.mc is None or not g.text.strip():
            continue
        owner, mcid = g.mc
        key = pk if owner == "page" else owner
        if not isinstance(key, int):
            continue
        out.append((key, mcid, g.text, view.rect(g.x0, g.y0, g.x1, g.y1), owner == "page"))
    return out


def candidates(chars: list[tuple], rect, label: str) -> list[tuple]:
    """隣に入れる範囲の候補（よい順）：(鍵, MCID, how, 字, 箱との距離)。"""
    runs: dict[tuple, dict] = {}
    for key, mcid, text, box, _pg in chars:
        r = runs.setdefault((key, mcid), {"text": [], "boxes": []})
        r["text"].append(text)
        r["boxes"].append(box)
    lab = _clean(label)
    scored = []
    for k, r in runs.items():
        t = _clean("".join(r["text"]))
        best = min(rank(b, rect) for b in r["boxes"])
        left = [rect[0] - b[2] for b in r["boxes"] if min(b[3], rect[3]) - max(b[1], rect[1]) > 0 and b[2] <= rect[0] + 0.5]
        up = [rect[1] - b[3] for b in r["boxes"] if min(b[2], rect[2]) - max(b[0], rect[0]) > 0 and b[3] <= rect[1] + 0.5]
        if lab and t and (lab in t or (len(t) >= 2 and t in lab)):
            scored.append(((0, best), k, "見出し", t, best[1]))
        if left:
            scored.append(((1, (0, min(left))), k, "左", t, max(0.0, min(left))))
        if up:
            scored.append(((2, (0, min(up))), k, "上", t, max(0.0, min(up))))
        scored.append(((3, best), k, "近く", t, best[1]))
    scored.sort(key=lambda s: (s[0], s[1]))
    seen = set()
    out = []
    for _s, k, how, t, gap in scored:
        if k not in seen:
            seen.add(k)
            out.append((k[0], k[1], how, t, gap))
    return out


class Adder:
    """元の文書に、足す。ページごとに add_page を呼び、最後に finish。"""

    def __init__(self, up, doc):
        self.up = up
        self.doc = doc
        self.tree = Tree(doc)  # 足すと変わるので、読むだけの控え（tree_of）とは別に
        self.placed: list[Placed] = []
        self.new_elems: set[int] = set()
        self.doc_elem: Ref | None = None
        nk = doc.resolve(self.tree.root.get("ParentTreeNextKey"))
        self.next_key = max([nk if isinstance(nk, int) else 0] + [k + 1 for k in self.tree.entries])
        self.key_changed = False

    def _root(self) -> dict:
        return self.up.get(self.tree.ref)

    def _pt_node(self, entry) -> dict:
        """ParentTree の節（書き換える形）。entry：(節の参照か None, 節)。"""
        ref, _d = entry
        if ref is not None:
            return self.up.get(ref)
        return self._root()["ParentTree"]

    def _alloc_key(self) -> int:
        k = self.next_key
        self.next_key += 1
        self.key_changed = True
        return k

    def _new_entry(self, key: int, value) -> None:
        """鍵を ParentTree に足す（小さい順に。範囲（Limits）も広げる）。"""
        tree = self.tree
        if self.tree.root.get("ParentTree") is None and not tree.leaves:
            node = {"Nums": []}
            self._root()["ParentTree"] = self.up.new(node)
            tree.root = self._root()
            ref = self._root()["ParentTree"]
            tree.leaves.append([(ref, node)])
        if not tree.leaves:
            raise ValueError("no leaf")
        def low(path):
            nums = self.doc.resolve(path[-1][1].get("Nums")) or []
            return self.doc.resolve(nums[0]) if nums else None
        # 鍵より小さい鍵を持つ、いちばん後ろの葉（なければ最初の葉）
        path = tree.leaves[0]
        for p in tree.leaves:
            lo = low(p)
            if lo is not None and lo <= key:
                path = p
        node = self._pt_node(path[-1])
        nums = list(self.doc.resolve(node.get("Nums")) or [])
        i = 0
        while i < len(nums) and self.doc.resolve(nums[i]) < key:
            i += 2
        nums[i:i] = [key, value]
        node["Nums"] = nums
        for e in path:
            lim = self.doc.resolve(e[1].get("Limits"))
            if isinstance(lim, list) and len(lim) == 2 and not self.doc.resolve(lim[0]) <= key <= self.doc.resolve(lim[1]):
                self._pt_node(e)["Limits"] = [min(self.doc.resolve(lim[0]), key), max(self.doc.resolve(lim[1]), key)]

    def _array_for(self, page_no: int, page: dict) -> tuple[list, int]:
        """ページの要素の一覧（書き換える形）。"""
        tree = self.tree
        key = self.doc.resolve(page.get("StructParents"))
        pd = self.up.get(page["_ref"])
        if not isinstance(key, int):
            key = self._alloc_key()
            pd["StructParents"] = key
        elif key >= self.next_key:
            self.next_key = key + 1
            self.key_changed = True
        cur = tree.entries.get(key)
        if isinstance(cur, Ref):
            return self.up.get(cur), key
        if isinstance(cur, list):
            # 一覧が葉の中に直接あるとき：葉の中の一覧を書き換える
            for path in tree.leaves:
                nums = self.doc.resolve(path[-1][1].get("Nums")) or []
                for i in range(0, len(nums), 2):
                    if self.doc.resolve(nums[i]) == key:
                        node = self._pt_node(path[-1])
                        arr = list(nums[i + 1])
                        new_nums = list(nums)
                        new_nums[i + 1] = arr
                        node["Nums"] = new_nums
                        return arr, key
        arr: list = []
        ref = self.up.new(arr)
        self._new_entry(key, ref)
        tree.entries[key] = ref
        return arr, key

    def _kids(self, ref: Ref) -> list:
        """要素の /K（書き換える一覧）。"""
        e = self.up.get(ref)
        k = e.get("K")
        if isinstance(k, Ref):
            kv = self.doc.resolve(k) if k.num not in self.new_elems else None
            if isinstance(kv, list):
                return self.up.get(k)
        if isinstance(k, list):
            return k
        k = [] if k is None else [k]
        e["K"] = k
        return k

    def _element_of(self, key: int, mcid: int, page_ref: Ref) -> tuple | None:
        """番号の範囲を持つ要素の、足す所：(親の参照, 親の /K で、その後ろに入れる項目の見分け)。つじつまが合わなければ None。"""
        doc = self.doc
        arr = doc.resolve(self.tree.entries.get(key))
        if not isinstance(arr, list) or not 0 <= mcid < len(arr) or not isinstance(arr[mcid], Ref):
            return None
        ref = arr[mcid]
        chain = self.tree.chain(ref)
        if not chain:
            return None
        if not self._holds_mcid(chain[0][1], mcid, chain):
            return None
        # 中身を説明で置き換える要素の下には入れない。そういう要素の親（段落など）にも入れない：親の要素の字が、ページの描画の順の
        # 範囲として読まれ、後ろに足した字までのページの字を丸ごと読む読み手がある（Mac のプレビューで測った）。親の、すぐ後ろに入れる
        at = None
        for i, (_r, d) in enumerate(chain):
            if d.get("Alt") is not None or d.get("ActualText") is not None or self.tree.std_type(d) in OPAQUE:
                at = i
        if at is None:
            return ref, ("mcid", mcid)
        if at + 2 >= len(chain):
            return None  # 親の親が根：文書の要素の終わりへ
        return chain[at + 2][0], ("elem", chain[at + 1][0].num)

    def _holds_mcid(self, elem: dict, mcid: int, chain) -> bool:
        doc = self.doc
        k = doc.resolve(elem.get("K"))
        for x in (k if isinstance(k, list) else [k]):
            x = doc.resolve(x) if isinstance(x, Ref) else x
            if x == mcid and not isinstance(x, bool):
                return True
            if isinstance(x, dict) and doc.resolve(x.get("MCID")) == mcid:
                return True
        return False

    def _insert(self, parent: Ref, anchor: tuple | None, new: Ref) -> None:
        kids = self._kids(parent)
        pos = len(kids)
        if anchor is not None:
            for i, x in enumerate(kids):
                xv = self.doc.resolve(x) if isinstance(x, Ref) and x.num not in self.new_elems else x
                hit = (anchor[0] == "mcid" and (xv == anchor[1] or (isinstance(xv, dict) and self.doc.resolve(xv.get("MCID")) == anchor[1]))) \
                    or (anchor[0] == "elem" and isinstance(x, Ref) and x.num == anchor[1])
                if hit:
                    pos = i + 1
                    break
            # 同じ所に足した字は、読む順（足す順）に並べる
            while pos < len(kids) and isinstance(kids[pos], Ref) and kids[pos].num in self.new_elems:
                pos += 1
        kids.insert(pos, new)

    def _document_element(self) -> Ref:
        """文書の要素（木のいちばん上の要素）。なければ作る。"""
        if self.doc_elem is not None:
            return self.doc_elem
        doc = self.doc
        k = doc.resolve(self.tree.root.get("K"))
        items = k if isinstance(k, list) else ([] if k is None else [self.tree.root.get("K")])
        refs = [x for x in items if isinstance(x, Ref) and isinstance(doc.resolve(x), dict)]
        for x in refs:
            if self.tree.std_type(doc.resolve(x)) == "Document":
                self.doc_elem = x
                return x
        if refs and len(refs) == 1 and self.tree.std_type(doc.resolve(refs[0])) not in OPAQUE:
            self.doc_elem = refs[0]
            return refs[0]
        new = self.up.new({"Type": Name("StructElem"), "S": Name("Document"), "P": self.tree.ref, "K": []})
        self.new_elems.add(new.num)
        root = self._root()
        cur = root.get("K")
        if isinstance(cur, Ref) and isinstance(doc.resolve(cur), list):
            self.up.get(cur).append(new)
        elif isinstance(cur, list):
            root["K"] = list(cur) + [new]
        elif cur is None:
            root["K"] = new
        else:
            root["K"] = [cur, new]
        self.doc_elem = new
        return new

    def add_page(self, page_no: int, page: dict, seq: list[dict]) -> None:
        """seq：足した順の {id, mcid, rect（見た目）, label, alt}。"""
        if not seq:
            return
        arr, _key = self._array_for(page_no, page)
        pref = page["_ref"]
        chars = marked_chars(self.doc, page)
        for it in seq:
            parent, anchor = None, None
            how, near, gap = "文書の終わり", "", None
            for key, mcid, h, t, g in candidates(chars, it["rect"], it.get("label") or ""):
                got = self._element_of(key, mcid, pref)
                if got is not None:
                    parent, anchor = got
                    how, near, gap = h, t, g
                    break
            if parent is None:
                parent = self._document_element()
            d = {"Type": Name("StructElem"), "S": Name("Span"), "P": parent, "Pg": pref, "K": it["mcid"]}
            new = self.up.new(d)
            self.new_elems.add(new.num)
            self._insert(parent, anchor, new)
            arr.extend([None] * (it["mcid"] + 1 - len(arr)))
            arr[it["mcid"]] = new
            self.placed.append(Placed(it["id"], page_no, how, near[:16], round(gap, 1) if gap is not None else None))

    def finish(self) -> None:
        if self.key_changed:
            self._root()["ParentTreeNextKey"] = self.next_key


def note(placed: list[Placed]) -> str | None:
    """本人とエージェントに伝える、足した所の限り（事実）。"""
    if not placed:
        return None
    count: dict[str, int] = {}
    for p in placed:
        count[p.how] = count.get(p.how, 0) + 1
    parts = "・".join(f"{h} {n}" for h, n in count.items())
    weak = [p for p in placed if p.how in ("近く", "文書の終わり")]
    msg = (f"この PDF はタグ（文書の構造の木）があるので、書いた所 {len(placed)} か所を木にも足しました（読み上げで読まれるように）。"
           f"読む順の所は字の位置だけで決めたので（{parts}）、見出しから離れた所で読まれることがあります（structure に、どの字の隣かがあります）")
    if weak:
        msg += "。見出しの字も同じ行・列の字も見つからなかった所：" + "、".join(
            f"{p.id}（{p.how}{'：' + p.near if p.near else ''}）" for p in weak[:12]) + (" ほか" if len(weak) > 12 else "")
    return msg


def read_back(doc, page_no: int, page: dict, spans: list[dict]) -> list[str]:
    """書いた後の文書で、足した範囲が要素につながっているか：番号 → ParentTree の一覧 → 要素（/K が番号・/P から根まで）。"""
    tree = tree_of(doc)
    out = []
    key = doc.resolve(page.get("StructParents"))
    arr = doc.resolve(tree.entries.get(key)) if isinstance(key, int) else None
    for s in spans:
        m = s["mcid"]
        e = doc.resolve(arr[m]) if isinstance(arr, list) and 0 <= m < len(arr) else None
        if not isinstance(e, dict):
            out.append(f"{s['id']}：番号 {m} の要素が、ページの一覧（ParentTree）にありません")
            continue
        if doc.resolve(e.get("K")) != m:
            out.append(f"{s['id']}：番号 {m} の要素の /K が番号と違います")
            continue
        if e.get("ActualText") is not None:
            out.append(f"{s['id']}：要素に ActualText があります")
            continue
        ref = arr[m]
        chain = tree.chain(ref)
        if not chain:
            out.append(f"{s['id']}：要素から木の根へたどれません")
            continue
        par = doc.resolve(e.get("P"))
        k = doc.resolve(par.get("K")) if isinstance(par, dict) else None
        if not any(isinstance(x, Ref) and x.num == ref.num for x in (k if isinstance(k, list) else [par.get("K") if isinstance(par, dict) else None])):
            out.append(f"{s['id']}：親の要素の /K に、足した要素がありません")
    return out
