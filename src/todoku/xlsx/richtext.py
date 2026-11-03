"""字ごとに書式の違う字のセル（リッチテキスト：<r> の並び）と、ふりがな（<rPh>・<phoneticPr>）を、字の一部を書き換えても残す。

字の中の空欄に書く・箱の字に印をつける・字を目印に書き換える（old/new）のは、セルの字の一部を書き換える書き込み。セルを
ただの字（<is><t>…</t></is>）として書き直すと、取り消し線・赤い字などの字ごとの書式が消え、「旧住所」を消した線が消える。

- 元の字と書き換えた後の字を並べて（difflib）、変えなかった字は元の <r> の書式（<rPr> のバイト）のまま、書き換えた字・書き足した
  字は、置き換えた元の字の <r> に入れる。書き足した所が、書式の違う2つの <r> のちょうど境目なら、どちらの書式か決められないので
  書かない（ValueError）
- ふりがな <rPh sb eb>：書き換えなかった字のふりがなは、字の位置をずらして残す。書き換えた字に掛かるふりがなは、読みが合わなく
  なるので外す（<phoneticPr> は元のまま）
- 証明（kept_problems）：書いた後のセルを読み直し、書き換えなかった字の書式が元と同じかを、字ごとに確かめる
"""

from __future__ import annotations

import difflib
import re
from dataclasses import dataclass, field

_ENT = re.compile(r"&(amp|lt|gt|quot|apos|#\d+|#x[0-9A-Fa-f]+);")
_XESC = re.compile(r"_x([0-9A-Fa-f]{4})_")


def _unent(s: str) -> str:
    def one(m):
        e = m[1]
        if e[0] == "#":
            return chr(int(e[2:], 16) if e[1] in "xX" else int(e[1:]))
        return {"amp": "&", "lt": "<", "gt": ">", "quot": '"', "apos": "'"}[e]

    return _XESC.sub(lambda m: chr(int(m[1], 16)), _ENT.sub(one, s))


def escape_text(s: str) -> bytes:
    """<t> の中身：XML の字の置き換えと、「_xHHHH_」の形の字をそのまま残す書き方（_x005F_xHHHH_。ECMA-376 22.9.2.19）。"""
    from xml.sax.saxutils import escape

    s = _XESC.sub(lambda m: "_x005F" + m[0], s)
    return escape(s).encode("utf-8")


@dataclass
class Rst:
    """<si>・<is> の中身。runs：(字, <rPr> のバイト（ない <r> は b""、<r> でない <t> は None）)。"""

    prefix: bytes
    runs: list[tuple[str, bytes | None]] = field(default_factory=list)
    rph: list[tuple[int, int, bytes]] = field(default_factory=list)  # (sb, eb, <rPh> のバイト)
    ppr: bytes = b""  # <phoneticPr …/>

    @property
    def text(self) -> str:
        return "".join(t for t, _p in self.runs)

    @property
    def rich(self) -> bool:
        """字ごとの書式があるか（<r> が2つ以上か、書式のある <r>）か、ふりがながあるか。"""
        rs = [p for _t, p in self.runs if p is not None]
        return len(self.runs) > 1 or any(rs) or bool(self.rph)

    def char_props(self) -> list[bytes | None]:
        out: list[bytes | None] = []
        for t, p in self.runs:
            out += [p] * len(t)
        return out


def parse_rst(inner: bytes, prefix: bytes = b"") -> Rst:
    """<si> か <is> の中身（子の要素のバイト）を読む。"""
    p = re.escape(prefix)
    out = Rst(prefix)
    pos = 0
    pat = re.compile(rb"<" + p + rb"(t|r|rPh|phoneticPr)\b[^>]*?(?:/>|>(.*?)</" + p + rb"\1>)", re.S)
    for m in pat.finditer(inner):
        if inner[pos:m.start()].strip():
            raise ValueError("字の要素のほかの中身があります")
        pos = m.end()
        tag, body = m[1], m[2] or b""
        if tag == b"t":
            out.runs.append((_unent(body.decode("utf-8")), None))
        elif tag == b"r":
            rpr = re.search(rb"<" + p + rb"rPr\b[^>]*?(?:/>|>.*?</" + p + rb"rPr>)", body, re.S)
            t = re.search(rb"<" + p + rb"t\b[^>]*?(?:/>|>(.*?)</" + p + rb"t>)", body, re.S)
            out.runs.append((_unent((t[1] or b"").decode("utf-8")) if t else "", rpr[0] if rpr else b""))
        elif tag == b"rPh":
            head = m[0][:m[0].index(b">")]
            sb = re.search(rb'\bsb="(\d+)"', head)
            eb = re.search(rb'\beb="(\d+)"', head)
            out.rph.append((int(sb[1]) if sb else 0, int(eb[1]) if eb else 0, m[0]))
        else:
            out.ppr = m[0]
    if inner[pos:].strip():
        raise ValueError("字の要素のほかの中身があります")
    return out


def _cell_span(xml: bytes, prefix: bytes, a: str):
    p = re.escape(prefix)
    m = re.search(rb"<" + p + rb"c\b[^>]*?\br=\"" + a.encode() + rb"\"[^>]*?(?:/>|>(.*?)</" + p + rb"c>)", xml, re.S)
    return m


def cell_rst(wb, sheet: str, r: int, c: int) -> Rst | None:
    """.xlsx のセルの字（共有文字列か、セルの中の字）の <si>・<is> を読む。字のセルでなければ None。"""
    from ..formula.parser import addr

    sh = wb.sheet(sheet)
    xml = wb.zip.read(sh.path)
    pm = re.search(rb"<(\w+:)?sheetData\b", xml)
    prefix = (pm[1] if pm else None) or b""
    m = _cell_span(xml, prefix, addr(r, c))
    if m is None or m[1] is None:
        return None
    head = m[0][:m[0].index(b">")]
    t = re.search(rb'\bt="(\w+)"', head)
    t = t[1] if t else b""
    p = re.escape(prefix)
    if t == b"inlineStr":
        im = re.search(rb"<" + p + rb"is\b[^>]*?(?:/>|>(.*?)</" + p + rb"is>)", m[1], re.S)
        return parse_rst(im[1] or b"", prefix) if im else None
    if t == b"s":
        vm = re.search(rb"<" + p + rb"v>\s*(\d+)\s*</" + p + rb"v>", m[1])
        path = getattr(wb, "sst_path", None)
        if vm is None or not path or not wb.zip.has(path):
            return None
        sst = wb.zip.read(path)
        sm = re.search(rb"<(\w+:)?sst\b", sst)
        sp = (sm[1] if sm else None) or b""
        q = re.escape(sp)
        want = int(vm[1])
        for i, si in enumerate(re.finditer(rb"<" + q + rb"si\b[^>]*?(?:/>|>(.*?)</" + q + rb"si>)", sst, re.S)):
            if i == want:
                runs, rph, ppr = _reprefix(parse_rst(si[1] or b"", sp), sp, prefix)
                return Rst(prefix, runs, rph, ppr)
        return None
    return None


def _reprefix(rst: Rst, old: bytes, new: bytes):
    """共有文字列の部品の名前空間の接頭辞（ふつうはなし）を、シートの接頭辞にそろえる。"""
    if old == new:
        return rst.runs, rst.rph, rst.ppr

    def fix(b: bytes | None):
        if not b:
            return b
        return re.sub(rb"(</?)" + re.escape(old), rb"\1" + new, b) if old else re.sub(rb"(</?)(?=\w)", rb"\1" + new, b)

    return [(t, fix(p)) for t, p in rst.runs], [(a, b, fix(x)) for a, b, x in rst.rph], fix(rst.ppr) or b""


def _alignment(old: str, new: str):
    return difflib.SequenceMatcher(None, old, new, autojunk=False).get_opcodes()


def splice(rst: Rst, new_text: str) -> tuple[bytes, list[bytes | None], list[str], list[bytes | None]]:
    """元の <si>/<is> の書式を残して、字を new_text にした <is> の中身。
    (中身, 書いた後の字ごとの書式, 外したふりがなの字, 書き足した字の書式（<rPr>）の並び)。"""
    old = rst.text
    props = rst.char_props()
    run_of = []
    for n, (t, _p) in enumerate(rst.runs):
        run_of += [n] * len(t)
    new_run: list[int] = []
    old_to_new: dict[int, int] = {}
    touched: list[tuple[int, int]] = []  # 書き換えた元の字の範囲（書き足しは (i, i)）
    written_runs: set[int] = set()
    for op, i1, i2, j1, j2 in _alignment(old, new_text):
        if op == "equal":
            for k in range(i2 - i1):
                old_to_new[i1 + k] = len(new_run)
                new_run.append(run_of[i1 + k])
            continue
        touched.append((i1, i2))
        if op == "delete":
            continue
        if i2 > i1:
            src = run_of[i1]
        else:
            before = run_of[i1 - 1] if i1 > 0 else None
            after = run_of[i1] if i1 < len(old) else None
            if before is not None and after is not None and before != after and props[i1 - 1] != props[i1]:
                raise ValueError(f"書き足した字「{new_text[j1:j2]}」が、書式の違う字の境目（「{old[max(0, i1 - 6):i1]}｜"
                                 f"{old[i1:i1 + 6]}」）に入るので、どちらの書式（取り消し線・色など）にするか決められません")
            src = before if before is not None else (after if after is not None else 0)
        new_run += [src] * (j2 - j1)
        if j2 > j1:
            written_runs.add(src)
    # 同じ <r> に入る字をまとめる
    pre = rst.prefix
    out = bytearray()
    segs: list[tuple[int, str]] = []
    for ch, n in zip(new_text, new_run):
        if segs and segs[-1][0] == n:
            segs[-1] = (n, segs[-1][1] + ch)
        else:
            segs.append((n, ch))
    plain = all(p is None for _t, p in rst.runs)
    if plain:
        out += _t(pre, new_text)
    else:
        for n, s in segs:
            rpr = rst.runs[n][1] or b""
            out += b"<" + pre + b"r>" + rpr + _t(pre, s) + b"</" + pre + b"r>"
    dropped = []
    for sb, eb, raw in rst.rph:
        hit = any((i1 < eb and i2 > sb) if i2 > i1 else (sb < i1 < eb) for i1, i2 in touched)
        if hit or sb not in old_to_new or (eb > sb and eb - 1 not in old_to_new):
            dropped.append(old[sb:eb])
            continue
        nsb = old_to_new[sb]
        neb = old_to_new[eb - 1] + 1 if eb > sb else nsb
        raw = re.sub(rb'\bsb="\d+"', b'sb="' + str(nsb).encode() + b'"', raw, count=1)
        raw = re.sub(rb'\beb="\d+"', b'eb="' + str(neb).encode() + b'"', raw, count=1)
        out += raw
    out += rst.ppr
    after_props = [rst.runs[n][1] for n in new_run] if rst.runs else [None] * len(new_text)
    return bytes(out), after_props, dropped, [rst.runs[n][1] for n in sorted(written_runs) if n < len(rst.runs)]


def _t(pre: bytes, s: str) -> bytes:
    body = escape_text(s)
    sp = b' xml:space="preserve"' if s != s.strip() or "\n" in s or not s else b""
    return b"<" + pre + b"t" + sp + b">" + body + b"</" + pre + b"t>"


def kept_problems(before: Rst, after: Rst | None) -> list[str]:
    """書き換えなかった字の書式（<rPr>）が、書いた後も同じか。違う所の説明の並び。"""
    if after is None:
        return ["書いた後のセルの字ごとの書式を読み直せません"]
    a, b = before.text, after.text
    pa, pb = before.char_props(), after.char_props()
    out = []
    for op, i1, i2, j1, _j2 in _alignment(a, b):
        if op != "equal":
            continue
        for k in range(i2 - i1):
            if (pa[i1 + k] or b"") != (pb[j1 + k] or b""):
                out.append(f"書き換えなかった字「{a[i1 + k]}」（{i1 + k + 1} 文字目）の書式が変わりました")
                break
    if before.ppr != after.ppr:
        out.append("ふりがなの設定（phoneticPr）が変わりました")
    return out


def describe_props(rpr: bytes | None) -> list[str]:
    """<rPr> のうち、書いた字が消したように・見えなく印刷される書式。"""
    if not rpr:
        return []
    out = []
    m = re.search(rb"<(?:\w+:)?strike\b([^>]*)/?>", rpr)
    if m and not re.search(rb'val="(0|false)"', m[1]):
        out.append("取り消し線")
    c = re.search(rb'<(?:\w+:)?color\b[^>]*\brgb="([0-9A-Fa-f]{6,8})"', rpr)
    if c:
        rgb = c[1].decode()[-6:].upper()
        if rgb == "FFFFFF":
            out.append("白い字")
        elif rgb != "000000":
            out.append(f"字の色（#{rgb}）")
    return out
