"""XPS（.xps・.oxps、ECMA-388）：ページの字（Glyphs の UnicodeString）を読む（読むだけ）。

- 入れ物は OPC の ZIP。パッケージの関係 fixedrepresentation → FixedDocumentSequence → FixedDocument（PageContent）→ FixedPage。
  部品が小さな部品（/[0].piece …）に分かれていれば、つないで読む（ECMA-388 第 9 節）
- 字：Glyphs の UnicodeString（先頭の「{}」は字の中の「{」を逃がす印、ECMA-388 12.1.4）。位置は OriginX・OriginY と
  RenderTransform（Canvas の入れ子の行列）から決め、上から下・左から右の順に、同じ高さのものを 1 行にする（配置からの推定）。
- UnicodeString がなく Indices（字形の番号）だけの Glyphs は、字に直せないので読まず、窓に数を書く（推測しない）
- 絵（ImageBrush の ImageSource）の中身は読まない。view_page は PDF のページだけ（XPS は窓で読む）
"""

from __future__ import annotations

import html
import re
import xml.etree.ElementTree as ET
from posixpath import dirname, join, normpath

from .. import readonlydoc as rd
from ..zipraw import RawZip

KIND_JA = "XPS（.xps・.oxps。読むだけ）"
WRITE = "XPS は印刷の形で、字を入れる欄がありません。とどくは書きません。書く必要があれば、Word・Excel・PDF の様式を発行元に頼んでください"
_REP = re.compile(r"/fixedrepresentation$", re.I)


class XpsError(ValueError):
    pass


def is_xps(z: RawZip, names: set) -> bool:
    """パッケージの関係（_rels/.rels）に fixedrepresentation があるか（XPS 1.0 と OpenXPS の両方の名前）。"""
    if "_rels/.rels" not in names:
        return False
    try:
        rel = z.read("_rels/.rels").decode("utf-8", "replace")
    except Exception:  # noqa: BLE001
        return False
    return "fixedrepresentation" in rel


def _read_part(z: RawZip, names: set, part: str) -> bytes:
    part = part.lstrip("/")
    if part in names:
        return z.read(part)
    low = {n.lower(): n for n in names}
    if part.lower() in low:
        return z.read(low[part.lower()])
    pieces = [n for n in names if n.lower().startswith(part.lower() + "/[") and n.lower().endswith(".piece")]
    if pieces:
        key = lambda n: int(re.search(r"\[(\d+)\]", n.rsplit("/", 1)[-1])[1])  # noqa: E731
        return b"".join(z.read(n) for n in sorted(pieces, key=key))
    raise XpsError(f"部品 {part} がありません")


def _resolve(base: str, src: str) -> str:
    return src.lstrip("/") if src.startswith("/") else normpath(join(dirname(base), src))


def _ln(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _xml(data: bytes, part: str) -> ET.Element:
    try:
        return ET.fromstring(data)
    except ET.ParseError as e:
        raise XpsError(f"部品 {part} を XML として読めません（{e}）") from None


def pages(blob: bytes) -> tuple[RawZip, set, list[str]]:
    z = RawZip(blob)
    names = set(z.names())
    root = _xml(z.read("_rels/.rels"), "_rels/.rels")
    seq = next((r.get("Target") for r in root if _REP.search(r.get("Type") or "")), None)
    if not seq:
        raise XpsError("FixedDocumentSequence への関係がありません")
    seq = seq.lstrip("/")
    out = []
    for dr in _xml(_read_part(z, names, seq), seq):
        if _ln(dr.tag) != "DocumentReference":
            continue
        fd = _resolve(seq, dr.get("Source") or "")
        for pc in _xml(_read_part(z, names, fd), fd):
            if _ln(pc.tag) == "PageContent":
                out.append(_resolve(fd, pc.get("Source") or ""))
    return z, names, out


Matrix = tuple[float, float, float, float, float, float]
IDENT: Matrix = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


def _mul(m: Matrix, n: Matrix) -> Matrix:
    """n を先に当て、次に m を当てる行列（XAML の入れ子：親が m）。"""
    a, b, c, d, e, f = n
    A, B, C, D, E, F = m
    return (a * A + b * C, a * B + b * D, c * A + d * C, c * B + d * D, e * A + f * C + E, e * B + f * D + F)


def _matrix(el: ET.Element) -> Matrix:
    text = el.get("RenderTransform")
    if not text:
        for k in el:
            if _ln(k.tag).endswith(".RenderTransform"):
                mt = next((x for x in k if _ln(x.tag) == "MatrixTransform"), None)
                if mt is not None:
                    text = mt.get("Matrix")
    if not text or text.startswith("{"):  # リソースの参照：ページの辞書までは読まない（位置が少しずれるだけで、字は落ちない）
        return IDENT
    try:
        v = [float(x) for x in re.split(r"[,\s]+", text.strip())]
    except ValueError:
        return IDENT
    return tuple(v) if len(v) == 6 else IDENT  # type: ignore[return-value]


def _unicode(s: str) -> str:
    return s[2:] if s.startswith("{}") else s


def read_page(data: bytes, part: str) -> tuple[list[str], int, int]:
    """(行, UnicodeString のない Glyphs の数, 絵の数)。"""
    root = _xml(data, part)
    runs: list[tuple[float, float, float, str]] = []
    no_text = [0]
    images = [0]

    def walk(el: ET.Element, m: Matrix, depth: int):
        if depth > 200:
            raise XpsError("Canvas の入れ子が深すぎます")
        for k in el:
            t = _ln(k.tag)
            if t in ("Canvas", "Glyphs", "Path"):
                mm = _mul(m, _matrix(k))
                if t == "Glyphs":
                    s = k.get("UnicodeString")
                    if s:
                        x, y = float(k.get("OriginX") or 0), float(k.get("OriginY") or 0)
                        em = abs(float(k.get("FontRenderingEmSize") or 12))
                        runs.append((x * mm[0] + y * mm[2] + mm[4], x * mm[1] + y * mm[3] + mm[5], em, _unicode(s)))
                    elif k.get("Indices"):
                        no_text[0] += 1
                if t == "Path" or t == "Canvas" or t == "Glyphs":
                    for b in k.iter():
                        if _ln(b.tag) == "ImageBrush" and t != "Canvas":
                            images[0] += 1
                if t == "Canvas":
                    walk(k, mm, depth + 1)

    walk(root, IDENT, 0)
    runs.sort(key=lambda r: (r[1], r[0]))
    lines: list[list[tuple[float, float, float, str]]] = []
    for r in runs:
        if lines and abs(lines[-1][-1][1] - r[1]) <= 0.25 * max(r[2], lines[-1][-1][2]):
            lines[-1].append(r)
        else:
            lines.append([r])
    out = []
    for ln in lines:
        ln.sort(key=lambda r: r[0])
        text = ""
        prev_end = None
        for x, _y, em, s in ln:
            if prev_end is not None and x - prev_end > em:
                text += " ｜ "
            text += s
            prev_end = x + len(s) * em * 0.5  # 字の幅は書体の中：半角の幅で少なめに見積もる（大きな空きだけを区切る）
        out.append(text)
    return out, no_text[0], images[0]


def build(blob: bytes, name: str) -> dict:
    try:
        z, names, plist = pages(blob)
    except KeyError as e:
        raise XpsError(f"部品がありません（{e}）") from None
    sections = []
    n_no, n_img = 0, 0
    for i, part in enumerate(plist, 1):
        lines, no, img = read_page(_read_part(z, names, part), part)
        n_no += no
        n_img += img
        sections.append((f"{i}ページ", lines))
    unread = []
    if n_no:
        unread.append(f"字の番号（Indices）だけで字のない Glyphs が {n_no} か所あります。字に直せないので読んでいません（推測しません）")
    if n_img:
        unread.append(f"絵 {n_img} 個の中身は字として読めません")
    return rd.make("xps", KIND_JA, name, blob, sections,
                   "元のファイルが正本。ページの字を、位置から上から下・左から右に並べました（読み順は配置からの推定）。" + WRITE, unread,
                   images=n_img, glyphs_without_text=n_no)


def grade(blob: bytes, name: str):
    from ..grade import Reason, _locked

    try:
        win = build(blob, name)
    except (XpsError, ValueError) as e:
        return _locked("読めない", str(e), "読めないので、書けません")
    if not rd.window_text(win).replace("ページ", "").strip("0123456789") and (win["images"] or win["glyphs_without_text"]):
        return _locked("字のない XPS", "どのページにも字の情報がありません（絵か、字に直せない字形だけです）。発行元に、字の入った PDF か Word を頼んでください",
                       "読めないので、書けません")
    more = [Reason("B", "読み順は配置からの推定", "字は UnicodeString のまま読みますが、行の組み立ては位置から決めたので、表の並びはエージェントが読んで判断します")]
    return rd.grade(win, "A", "字はファイルにそのまま", "ページの字（Glyphs の UnicodeString）を、そのまま読みます", "XPS には書かない", WRITE, more)


def audit(blob: bytes, name: str):
    from ..xlsx.audit import Audit, Check

    try:
        win = build(blob, name)
        z, names, plist = pages(blob)
    except (XpsError, ValueError) as e:
        au = Audit(name)
        au.checks.append(Check("読む", False, str(e)))
        return au
    raw = []
    for part in plist:
        rawb = _read_part(z, names, part)
        data = rawb.decode("utf-16" if rawb[:2] in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig", "replace")
        for m in re.finditer(r'UnicodeString\s*=\s*"([^"]*)"', data):
            raw.append(_unicode(html.unescape(m[1])))
    return rd.audit(win, name, raw, f"{len(plist)} ページの UnicodeString", WRITE)
