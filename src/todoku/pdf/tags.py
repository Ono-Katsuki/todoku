"""PDF のタグ（文書の構造の木）から、画像や図の説明（代わりの字）を探す。

画像の説明は、段階（エージェントが読めるか）とは別の、アクセシビリティの注意として示す（WCAG 1.1.1・PDF/UA）。
説明があるとみなすのは、ファイルの構造に書かれているときだけ：
- 画像を包む印（BDC）の Alt・ActualText
- 画像の構造の要素（か、その親）の Alt・ActualText（MCID は ParentTree で、画像そのものは StructParent で引く）
- 飾り（Artifact）の印：説明は要らない
"""

from __future__ import annotations

from .objects import Name, PdfString, decode_text


class StructTree:
    def __init__(self, doc):
        self.doc = doc
        cat = doc.resolve(doc.trailer.get("Root")) or {}
        root = doc.resolve(cat.get("StructTreeRoot")) if isinstance(cat, dict) else None
        self.tagged = isinstance(root, dict)
        self.root = root if self.tagged else {}
        rm = doc.resolve(self.root.get("RoleMap"))
        self.rolemap = rm if isinstance(rm, dict) else {}
        self.parents: dict[int, object] = {}
        if self.tagged:
            self._numtree(self.root.get("ParentTree"), 0)

    def _numtree(self, node, depth: int) -> None:
        node = self.doc.resolve(node)
        if not isinstance(node, dict) or depth > 32:
            return
        nums = self.doc.resolve(node.get("Nums"))
        if isinstance(nums, list):
            for i in range(0, len(nums) - 1, 2):
                k = self.doc.resolve(nums[i])
                if isinstance(k, int):
                    self.parents[k] = nums[i + 1]
        for kid in self.doc.resolve(node.get("Kids")) or []:
            self._numtree(kid, depth + 1)

    def elem_for(self, mc: tuple | None, objparent: int | None, page_key: int | None):
        """画像の構造の要素（辞書）。見つからなければ None。"""
        if objparent is not None:
            e = self.doc.resolve(self.parents.get(objparent))
            if isinstance(e, dict):
                return e
        if mc is not None:
            owner, mcid = mc
            key = page_key if owner == "page" else owner
            arr = self.doc.resolve(self.parents.get(key)) if isinstance(key, int) else None
            if isinstance(arr, list) and 0 <= mcid < len(arr):
                e = self.doc.resolve(arr[mcid])
                if isinstance(e, dict):
                    return e
        return None

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

    def description(self, elem: dict) -> str | None:
        """要素か、その親に書かれた説明（Alt・ActualText）。空の説明は、説明がないものとして扱う。"""
        e, guard = elem, 0
        while isinstance(e, dict) and guard < 64:
            for k in ("Alt", "ActualText"):
                v = self.doc.resolve(e.get(k))
                if isinstance(v, (bytes, PdfString)):
                    t = decode_text(v).strip()
                    if t:
                        return t
            if self.std_type(e) == "Artifact":
                return None
            e = self.doc.resolve(e.get("P"))
            guard += 1
        return None


def describe_pictures(doc, page: dict, pictures, tree: StructTree | None = None) -> list[dict]:
    """ページの画像ごとの説明：{"alt": 説明 | None, "decorative": 飾りか, "tagged": タグの要素があるか, "role": 要素の種類}。"""
    tree = tree or StructTree(doc)
    pk = doc.resolve(page.get("StructParents"))
    pk = pk if isinstance(pk, int) else None
    out = []
    for p in pictures:
        d = {"alt": p.alt, "decorative": bool(p.artifact), "tagged": False, "role": ""}
        if not d["alt"] and not d["decorative"] and tree.tagged:
            e = tree.elem_for(p.mc, p.objparent, pk)
            if e is not None:
                d["tagged"] = True
                d["role"] = tree.std_type(e)
                d["alt"] = tree.description(e)
        out.append(d)
    return out


