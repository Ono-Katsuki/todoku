"""JSON（RFC 8259）と JSON Lines（1行に1つの JSON）を読み書きする。値の場所は JSON Pointer（RFC 6901）で指す。

読む：値ごとに「元の字の範囲」・種類・値を持つ（数は元の字のまま。1.0 と 1 を区別する）。
書く：書いた値の字だけを置き換える。ほかの値・空白・改行・キーの順はそのまま。文字列は、ファイルの書き方（非 ASCII の字を
      \\u でエスケープするか）に合わせる。ない名前（キー）に書くと、そのオブジェクトの終わりに、ほかの名前と同じ字下げで足す。
      配列の終わりに足すときは /配列/- と書く。
証明：書く前と後を読み直し、書いた値だけが変わり、ほかのすべての字（空白も）が同じことを確かめる。
"""

from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass, field

from .textenc import Decoded, EncodingError, decode, kept_encoding


class JsonError(ValueError):
    pass


@dataclass
class JNode:
    kind: str  # object / array / string / number / true / false / null
    a: int
    b: int
    path: str
    value: object = None  # 字・数・真偽・None（オブジェクトと配列は None）
    raw: str = ""  # 数の元の字
    members: list = field(default_factory=list)  # オブジェクト：[(名前, 名前の始まり, 名前の終わり, JNode)]
    items: list = field(default_factory=list)  # 配列：[JNode]

    @property
    def is_leaf(self) -> bool:
        return self.kind not in ("object", "array")


def esc_token(t: str) -> str:
    """JSON Pointer の1つの名前のエスケープ（RFC 6901：~ → ~0、/ → ~1）。"""
    return t.replace("~", "~0").replace("/", "~1")


def unesc_token(t: str) -> str:
    return t.replace("~1", "/").replace("~0", "~")


_NUM = re.compile(r"-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?")
_ESC = {'"': '"', "\\": "\\", "/": "/", "b": "\b", "f": "\f", "n": "\n", "r": "\r", "t": "\t"}


class _Parser:
    def __init__(self, text: str, base: str = ""):
        self.t = text
        self.i = 0
        self.base = base
        self.dups: list[str] = []
        self.escaped_nonascii = 0  # 非 ASCII の字を \u で書いた数
        self.raw_nonascii = 0  # 非 ASCII の字をそのまま書いた数

    def err(self, msg: str):
        line = self.t.count("\n", 0, self.i) + 1
        raise JsonError(f"{line} 行目（{self.i + 1} 字目）：{msg}")

    def ws(self) -> None:
        t = self.t
        n = len(t)
        while self.i < n and t[self.i] in " \t\n\r":
            self.i += 1

    def value(self, path: str) -> JNode:
        self.ws()
        t = self.t
        if self.i >= len(t):
            self.err("値がありません")
        ch = t[self.i]
        a = self.i
        if ch == "{":
            node = JNode("object", a, a, path)
            self.i += 1
            self.ws()
            seen = set()
            if t.startswith("}", self.i):
                self.i += 1
                node.b = self.i
                return node
            while True:
                self.ws()
                if not t.startswith('"', self.i):
                    self.err("名前（文字列）がありません")
                ka = self.i
                key = self.string()
                kb = self.i
                self.ws()
                if not t.startswith(":", self.i):
                    self.err("「:」がありません")
                self.i += 1
                if key in seen:
                    self.dups.append(path + "/" + esc_token(key))
                seen.add(key)
                v = self.value(path + "/" + esc_token(key))
                node.members.append((key, ka, kb, v))
                self.ws()
                if t.startswith(",", self.i):
                    self.i += 1
                    continue
                if t.startswith("}", self.i):
                    self.i += 1
                    break
                self.err("「,」か「}」がありません")
            node.b = self.i
            return node
        if ch == "[":
            node = JNode("array", a, a, path)
            self.i += 1
            self.ws()
            if t.startswith("]", self.i):
                self.i += 1
                node.b = self.i
                return node
            while True:
                v = self.value(f"{path}/{len(node.items)}")
                node.items.append(v)
                self.ws()
                if t.startswith(",", self.i):
                    self.i += 1
                    continue
                if t.startswith("]", self.i):
                    self.i += 1
                    break
                self.err("「,」か「]」がありません")
            node.b = self.i
            return node
        if ch == '"':
            s = self.string()
            return JNode("string", a, self.i, path, s)
        for word, kind, val in (("true", "true", True), ("false", "false", False), ("null", "null", None)):
            if t.startswith(word, self.i):
                self.i += len(word)
                return JNode(kind, a, self.i, path, val)
        m = _NUM.match(t, self.i)
        if m and m.end() > self.i:
            self.i = m.end()
            raw = m[0]
            num = int(raw) if re.fullmatch(r"-?\d+", raw) else float(raw)
            return JNode("number", a, self.i, path, num, raw)
        self.err(f"値として読めない字「{ch}」があります")

    def string(self) -> str:
        t = self.t
        self.i += 1
        out = []
        n = len(t)
        while True:
            j = self.i
            while j < n and t[j] not in '"\\' and t[j] >= " ":
                j += 1
            chunk = t[self.i:j]
            if chunk:
                out.append(chunk)
                if not chunk.isascii():
                    self.raw_nonascii += 1
            self.i = j
            if j >= n:
                self.err("文字列が閉じていません")
            ch = t[j]
            if ch == '"':
                self.i += 1
                return "".join(out)
            if ch < " ":
                self.err("文字列の中に、エスケープしていない制御文字があります")
            # エスケープ
            if j + 1 >= n:
                self.err("文字列が閉じていません")
            e = t[j + 1]
            if e in _ESC:
                out.append(_ESC[e])
                self.i = j + 2
                continue
            if e == "u":
                h = t[j + 2:j + 6]
                if not re.fullmatch(r"[0-9a-fA-F]{4}", h):
                    self.err("\\u のあとに16進の4桁がありません")
                code = int(h, 16)
                self.i = j + 6
                if 0xD800 <= code < 0xDC00 and t.startswith("\\u", self.i) and re.fullmatch(r"[0-9a-fA-F]{4}", t[self.i + 2:self.i + 6]):
                    low = int(t[self.i + 2:self.i + 6], 16)
                    if 0xDC00 <= low < 0xE000:
                        code = 0x10000 + ((code - 0xD800) << 10) + (low - 0xDC00)
                        self.i += 6
                if code >= 0x80:
                    self.escaped_nonascii += 1
                out.append(chr(code))
                continue
            self.err(f"知らないエスケープ「\\{e}」があります")


@dataclass
class Doc:
    """1つの JSON の値（JSON Lines では1行）。"""

    root: JNode
    line: int = 0  # JSON Lines の行の番号（1から）。JSON は 0
    offset: int = 0  # ファイルの文字列の中の、この値の字の始まり（JSON Lines の行の始まり）


class JsonFile:
    def __init__(self, blob: bytes, filename: str = "", kind: str = "json", encoding: str | None = None):
        self.blob = blob
        self.filename = filename
        self.kind = kind
        try:
            self.dec: Decoded = decode(blob, encoding)
        except UnicodeDecodeError as e:
            raise JsonError(f"文字コードを読めません（{e.reason}）") from None
        self.text = self.dec.text
        self.docs: list[Doc] = []
        self.duplicates: list[str] = []
        esc = raw = 0
        if kind == "jsonl":
            pos = 0
            for k, line in enumerate(re.split(r"(?<=\n)", self.text), 1):
                body = line.rstrip("\r\n")
                if body.strip():
                    p = _Parser(body)
                    try:
                        root = p.value("")
                        p.ws()
                        if p.i != len(body):
                            p.err("値のあとに余分な字があります")
                    except JsonError as e:
                        raise JsonError(f"{k} 行目：{e}") from None
                    _shift(root, pos)
                    self.docs.append(Doc(root, k, pos))
                    self.duplicates += [f"L{k}{d}" for d in p.dups]
                    esc += p.escaped_nonascii
                    raw += p.raw_nonascii
                pos += len(line)
            self.lines = self.text.count("\n") + (0 if self.text.endswith("\n") or not self.text else 1)
        else:
            p = _Parser(self.text)
            root = p.value("")
            p.ws()
            if p.i != len(self.text):
                p.err("値のあとに余分な字があります")
            self.docs.append(Doc(root))
            self.duplicates = p.dups
            esc, raw = p.escaped_nonascii, p.raw_nonascii
        # 非 ASCII の字の書き方：エスケープしている（Python の json.dumps の既定など）か、そのままか
        self.escape_nonascii = esc > 0 and raw == 0
        nl = self.text.count("\r\n")
        self.newline = "\r\n" if nl and nl >= self.text.count("\n") - nl else "\n"

    def sha256(self) -> str:
        return hashlib.sha256(self.blob).hexdigest()

    # ------------------------------------------------------------------ 値をたどる
    def leaves(self):
        """(id, JNode, Doc)：すべての値（オブジェクト・配列は除く）を、ファイルの順に。"""
        for d in self.docs:
            pre = f"L{d.line}" if self.kind == "jsonl" else ""
            stack = [d.root]
            while stack:
                n = stack.pop()
                if n.is_leaf:
                    yield (pre + n.path) if (pre + n.path) else "/", n, d
                elif n.kind == "object":
                    stack.extend(reversed([m[3] for m in n.members]))
                else:
                    stack.extend(reversed(n.items))

    def nodes(self):
        for d in self.docs:
            pre = f"L{d.line}" if self.kind == "jsonl" else ""
            stack = [d.root]
            while stack:
                n = stack.pop()
                yield pre + n.path, n, d
                if n.kind == "object":
                    stack.extend(reversed([m[3] for m in n.members]))
                elif n.kind == "array":
                    stack.extend(reversed(n.items))

    def split_id(self, key: str) -> tuple[Doc, str]:
        """id → (どの値か, JSON Pointer)。JSON は「/a/b」（全体は「」か「/」）、JSON Lines は「L3/a/b」（行の全体は「L3」）。"""
        k = str(key).strip()
        if self.kind == "jsonl":
            m = re.fullmatch(r"[Ll](\d+)(/.*)?", k)
            if not m:
                raise JsonError(f"「{key}」は値の id ではありません（L行/名前 の形。例 L3/name）")
            line = int(m[1])
            d = next((x for x in self.docs if x.line == line), None)
            if d is None:
                raise JsonError(f"{line} 行目に JSON の値がありません")
            return d, m[2] or ""
        if k == "/":
            return self.docs[0], ""
        if k and not k.startswith("/"):
            raise JsonError(f"「{key}」は JSON Pointer ではありません（/ で始めます。例 /applicant/name）")
        return self.docs[0], k

    def find(self, key: str) -> tuple[JNode | None, JNode | None, str, Doc]:
        """(値, 親, 最後の名前, Doc)。値がなく、親がオブジェクト（または配列の「-」）なら、値は None。"""
        d, ptr = self.split_id(key)
        if ptr == "":
            return d.root, None, "", d
        toks = [unesc_token(t) for t in ptr.split("/")[1:]]
        cur = d.root
        parent = None
        for i, tok in enumerate(toks):
            parent = cur
            last = i == len(toks) - 1
            if cur.kind == "object":
                hits = [m for m in cur.members if m[0] == tok]
                if len(hits) > 1:
                    raise JsonError(f"「{key}」：同じ名前「{tok}」が {len(hits)} つあり、どれか決まりません")
                if not hits:
                    if last:
                        return None, cur, tok, d
                    raise JsonError(f"「{key}」：名前「{tok}」がありません")
                cur = hits[0][3]
            elif cur.kind == "array":
                if tok == "-" and last:
                    return None, cur, tok, d
                if not re.fullmatch(r"0|[1-9][0-9]*", tok) or int(tok) >= len(cur.items):
                    raise JsonError(f"「{key}」：配列に {tok} 番目の値がありません（0〜{len(cur.items) - 1}。終わりに足すときは「-」）")
                cur = cur.items[int(tok)]
            else:
                raise JsonError(f"「{key}」：{cur.path or '/'} は値（{cur.kind}）で、その中はありません")
        return cur, parent, toks[-1], d


def _shift(node: JNode, by: int) -> None:
    stack = [node]
    while stack:
        n = stack.pop()
        n.a += by
        n.b += by
        if n.kind == "object":
            n.members = [(k, ka + by, kb + by, v) for k, ka, kb, v in n.members]
            stack.extend(m[3] for m in n.members)
        elif n.kind == "array":
            stack.extend(n.items)


# ====================================================================== 書く
def literal(value, f: JsonFile, old: JNode | None, notes: list[str], key: str) -> str:
    """書く値の JSON の字。いまの値の種類に合わせる（数の所に数の字が来たら数に、真偽の所に true/false が来たら真偽に）。"""
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=f.escape_nonascii, separators=(", ", ": "))
    kind = old.kind if old is not None else None
    if isinstance(value, str):
        t = value.strip()
        if kind == "number" and _NUM.fullmatch(t):
            return t
        if kind in ("true", "false") and t.lower() in ("true", "false"):
            return t.lower()
        if kind == "number":
            notes.append(f"{key}：数の値の所に、文字列「{value}」を書きました")
        elif kind in ("true", "false"):
            notes.append(f"{key}：真偽（true/false）の値の所に、文字列「{value}」を書きました")
        return json.dumps(value, ensure_ascii=f.escape_nonascii)
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value.is_integer() and kind == "number" and old is not None and "." not in old.raw \
                and "e" not in old.raw.lower():
            return str(int(value))
        return json.dumps(value)
    return json.dumps(str(value), ensure_ascii=f.escape_nonascii)


@dataclass
class JsonPatch:
    id: str
    value: object  # 書いた値（JSON として読み直した値）
    text: str  # 書いた JSON の字
    added: bool = False  # 名前・配列の値を足した
    key: str = ""


@dataclass
class JsonFillResult:
    blob: bytes
    patches: list[JsonPatch] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def _indent_before(text: str, pos: int) -> str | None:
    """pos の字の前の、行の頭からの空白（その前に改行がなければ None：1行に書いたもの）。"""
    j = pos
    while j > 0 and text[j - 1] in " \t":
        j -= 1
    if j > 0 and text[j - 1] in "\r\n":
        return text[j:pos]
    return None


def fill(f: JsonFile, values: dict) -> JsonFillResult:
    edits: list[tuple[int, int, str]] = []  # 文字列の中の位置で
    patches: list[JsonPatch] = []
    notes: list[str] = []
    problems: list[str] = []
    added_at: dict[int, list[str]] = {}
    appended: dict[int, int] = {}  # 配列（id(JNode)）→ 終わりに足した数
    for key, v in values.items():
        try:
            node, parent, last, d = f.find(key)
        except JsonError as e:
            problems.append(str(e))
            continue
        pre = f"L{d.line}" if f.kind == "jsonl" else ""
        if node is not None:
            if not node.is_leaf and not isinstance(v, (dict, list)):
                notes.append(f"{key}：{'オブジェクト' if node.kind == 'object' else '配列'}を、値1つに置き換えました")
            lit = literal(v, f, node, notes, key)
            edits.append((node.a, node.b, lit))
            patches.append(JsonPatch((pre + node.path) or "/", json.loads(lit), lit, False, str(key)))
            continue
        # ない名前・配列の終わり：親の終わりに、ほかの値と同じ字下げで足す
        lit = literal(v, f, None, notes, key)
        t = f.text
        kids = [m[3] for m in parent.members] if parent.kind == "object" else parent.items
        if parent.kind == "object":
            if any(p_.added and p_.id == pre + parent.path + "/" + esc_token(last) for p_ in patches):
                problems.append(f"{key}：同じ名前に2回書こうとしています")
                continue
            canon = pre + parent.path + "/" + esc_token(last)
            member = json.dumps(last, ensure_ascii=f.escape_nonascii) + _colon(f, parent) + lit
        else:
            k_ = appended.get(id(parent), 0)
            appended[id(parent)] = k_ + 1
            canon = pre + parent.path + "/" + str(len(kids) + k_)
            member = lit
        if kids:
            first_start = parent.members[0][1] if parent.kind == "object" else kids[0].a
            ind = _indent_before(t, first_start)
            sep = ("," + f.newline + ind) if ind is not None else ", "
            at = kids[-1].b
        else:
            sep = ", " if added_at.get(parent.a + 1) else ""
            at = parent.a + 1
        added_at.setdefault(at, []).append(sep + member)
        patches.append(JsonPatch(canon, json.loads(lit), lit, True, str(key)))
    if problems:
        raise JsonError("\n".join(problems))
    for at, parts in added_at.items():
        edits.append((at, at, "".join(parts)))
    edits.sort(key=lambda e: (e[0], e[1]))
    for (a1, b1, _x), (a2, _b2, _y) in zip(edits, edits[1:]):
        if a2 < b1:
            raise JsonError("同じ所に2つの書き込みが重なりました（オブジェクトと、その中の値に同時に書いています）")
    try:
        out = f.dec.rewrite(f.blob, edits)
    except EncodingError as e:
        raise JsonError(str(e)) from None
    return JsonFillResult(out, patches, notes)


def _colon(f: JsonFile, obj: JNode) -> str:
    """名前と値の間の書き方（「: 」か「:」）：同じオブジェクトの、ほかの名前に合わせる。"""
    if obj.members:
        _k, _ka, kb, v = obj.members[-1]
        return f.text[kb:v.a]
    return ": "


# ====================================================================== 証明
@dataclass
class JsonProof:
    original_sha256: str
    new_sha256: str
    values_total: int = 0
    values_changed: int = 0
    unexpected: list[str] = field(default_factory=list)
    lint_errors: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.unexpected and not self.lint_errors

    def headline(self) -> str:
        return (f"値 {self.values_total:,} のうち、変わったのは書いた {self.values_changed} 個だけ。ほかの値・空白・改行・名前の順はそのまま"
                + ("" if not self.unexpected else f"（想定外の変更 {len(self.unexpected)} 件）"))


def prove(original: bytes, new: bytes, written: dict[str, object], filename: str = "", kind: str = "json",
          encoding: str | None = None) -> JsonProof:
    """written：id（JSON Pointer。JSON Lines は L行/…）→ 書いた値。書く前と後を読み直して比べる。"""
    pr = JsonProof(hashlib.sha256(original).hexdigest(), hashlib.sha256(new).hexdigest())
    a = JsonFile(original, filename, kind, encoding)
    try:
        b = JsonFile(new, filename, kind, encoding)
    except JsonError as e:
        pr.lint_errors.append(f"書いた後のファイルを JSON として読めません：{e}")
        return pr
    if not kept_encoding(a.dec, b.dec, new) and not a.dec.ascii_only:
        pr.unexpected.append(f"文字コードが変わっています（{a.dec.label} → {b.dec.label}）")
    na = {i: n for i, n, _d in a.nodes()}
    nb = {i: n for i, n, _d in b.nodes()}
    w = {(k if k != "/" else ""): v for k, v in written.items()}

    def inside(i: str) -> bool:
        return any(i == k or i.startswith(k + "/") for k in w)

    for i, n in na.items():
        if inside(i):
            continue
        if not n.is_leaf:
            if i not in nb or nb[i].kind != n.kind:
                pr.unexpected.append(f"{i or '/'}：書いていない所の形が変わっています")
            continue
        pr.values_total += 1
        m = nb.get(i)
        if m is None or m.kind != n.kind or a.text[n.a:n.b] != b.text[m.a:m.b]:
            pr.unexpected.append(f"{i or '/'}：書いていない値が変わっています")
    for k, v in w.items():
        pr.values_total += 1
        pr.values_changed += 1
        m = nb.get(k)
        if m is None:
            pr.unexpected.append(f"{k or '/'}：書いた値が、書いた後のファイルにありません")
            continue
        got = _plain(m)
        if not _equal(got, v):
            pr.unexpected.append(f"{k or '/'}：書いた後の値が {json.dumps(got, ensure_ascii=False)[:60]} で、"
                                 f"書いた {json.dumps(v, ensure_ascii=False)[:60]} と違います")
    extra = [i for i in nb if i not in na and not inside(i)]
    if extra:
        pr.unexpected.append("書いていない名前・値が増えています：" + "、".join(extra[:10]))
    # 字の比べ：書いた値（足した名前・値は、前の値の終わりからその値の終わりまで）を除くと、書く前と後の字がすべて同じ
    added = {k for k in w if k not in na}
    if _cut(a, {k for k in w if k in na}, na, added, before=True) != _cut(b, set(w) - added, nb, added, before=False):
        pr.unexpected.append("書いた値の外の字（空白・改行を含む）が変わっています")
    return pr


def _plain(n: JNode):
    if n.is_leaf:
        return n.value
    if n.kind == "object":
        return {k: _plain(v) for k, _ka, _kb, v in n.members}
    return [_plain(x) for x in n.items]


def _equal(x, y) -> bool:
    if isinstance(x, dict) and isinstance(y, dict):
        return x.keys() == y.keys() and all(_equal(x[k], y[k]) for k in x)
    if isinstance(x, list) and isinstance(y, list):
        return len(x) == len(y) and all(_equal(p, q) for p, q in zip(x, y))
    if isinstance(x, bool) or isinstance(y, bool):
        return type(x) is type(y) and x == y
    if isinstance(x, (int, float)) and isinstance(y, (int, float)):
        return x == y
    return type(x) is type(y) and x == y


def _parent_of(i: str) -> str:
    return i.rsplit("/", 1)[0] if "/" in i else i


def _cut(f: JsonFile, replaced: set, nodes: dict, added: set, before: bool) -> str:
    """書いた値の字を「\\0」に置き換えた文字列。足した名前・値は、書く前は親の終わり（最後の値の後ろ）の位置に「\\0」、
    書いた後は前の値の終わりからその値の終わりまでを「\\0」にする。"""
    spans = [(nodes[k].a, nodes[k].b) for k in replaced if k in nodes]
    for k in added:
        if before:
            p = nodes.get(_parent_of(k))
            if p is None:
                continue
            kids = [m[3] for m in p.members] if p.kind == "object" else p.items
            at = kids[-1].b if kids else p.a + 1
            spans.append((at, at))
        else:
            n = nodes.get(k)
            p = nodes.get(_parent_of(k))
            if n is None or p is None:
                continue
            kids = [m[3] for m in p.members] if p.kind == "object" else p.items
            idx = next((j for j, x in enumerate(kids) if x is n), 0)
            spans.append((kids[idx - 1].b if idx > 0 else p.a + 1, n.b))
    out, at = [], 0
    for a, b in sorted(spans):
        if a < at:
            continue
        out.append(f.text[at:a])
        out.append("\0")
        at = b
    out.append(f.text[at:])
    return "".join(out)
