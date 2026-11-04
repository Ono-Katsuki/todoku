"""文字だけのファイル（.txt）に書き込む：書いた所の字のバイトだけを、そのファイルの文字コードで作ったバイト列に置き換える。

- 行の中の空欄（「令和　年」の空白など）：空欄の字を値に置き換える（Word・RTF と同じ）
- □：■（☐ は ☑）にする
- 字を目印にした書き込み（{"@1": {"old": …, "new": …}}）：old と new の違う所だけを置き換える
ほかのバイト（ほかの字・改行・BOM）はそのまま。ファイル全体を文字列にして書き直すことはしない（Shift_JIS には、同じ字に
2つのバイト列があるものがあるので）。改行は書かない（行の数と、行の id が変わるので）。
文字だけのファイルには見えない印（しおり）を置けないので、書いた所は、証明の中で書いた後のファイルを読み直して確かめる。
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass, field

from ..datafile.textenc import EncodingError
from ..docx.fill import LINE_BREAKS, FillError, _as_bool, _is_text_value, written_of  # 証明に渡す、書いた所 → 値（Word と同じ）
from ..docx.proof import _expected_text
from ..docx.window import Spots, WriteError, box_mark, resolve_write
from .document import NEWLINE_JA, TxtDocument, TxtError


@dataclass
class Patch:
    id: str
    kind: str
    value: object
    anchor: dict | None = None
    key: str = ""
    label: str = ""


@dataclass
class FillResult:
    blob: bytes
    patches: list[Patch] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def _edit_of(it: dict, value) -> tuple[int, int, str] | None:
    """欄と書く値 → 行の中の (始め, 終わり, 置き換える字)。チェックの字が変わらないとき（□ のチェックを外すなど）は None。"""
    p = it["para"]
    a, b = it["span"]
    if it["kind"] == "checkbox":
        new = box_mark(p.text[a:b], _as_bool(value))
        return None if new == p.text[a:b] else (a, b, new)
    if it["kind"] != "inline":
        raise FillError(f"{it['id']}: この種類の欄（{it['kind']}）には書けません")
    return a, b, str(value)


def plan(doc: TxtDocument, values: dict) -> tuple[list[tuple[int, int, str, int]], list[Patch]]:
    """書く所：(ファイル全体の文字列の中の始め, 終わり, 置き換える字, 行番号) と、書いた所の一覧。"""
    sp = Spots(doc)
    used: dict = {}
    edits: list[tuple[int, int, str, int]] = []
    patches: list[Patch] = []
    problems: list[str] = []
    for key, value in values.items():
        try:
            if key.startswith("@"):
                if not isinstance(value, dict):
                    raise FillError(f'{key}: 字を目印にした書き込みは {{"old": …, "new": …}} の形で指定してください')
                try:
                    it = resolve_write(sp, value, used)
                except WriteError as e:
                    raise FillError(f"{key}: {e}") from None
                text = it["value"]
                patch = Patch(it["id"], "input", text, anchor=dict(value), key=key, label=it.get("label", ""))
            else:
                it = sp.items.get(key)
                if it is None:
                    raise FillError(f"{key}: そのような欄はありません{sp.items.hint(key)}")
                if it["kind"] == "checkbox":
                    text = _as_bool(value)
                    patch = Patch(key, "control", text, key=key, label=it.get("label", ""))
                else:
                    text = _is_text_value(value)
                    patch = Patch(key, "input", text, key=key, label=it.get("label", ""))
            if isinstance(text, str) and any(ch in text for ch in LINE_BREAKS):  # str.splitlines が行を分ける字すべて
                raise FillError(f"{key}: 改行・行を分ける字（CR・LF・U+2028・U+0085・改ページの字など）は書けません（文字だけのファイルでは、"
                                "読み手によって行が増えて、ほかの行の id が変わります）。"
                                "1行ずつ、それぞれの行の欄に書いてください")
            e = _edit_of(it, text)
            if e is not None:
                a, b, v = e
                ln = it["para"]
                for x, y, _v, _n in edits:
                    if ln.a + a < y and x < ln.a + b or (a == b and x <= ln.a + a <= y) or (x == y and ln.a + a <= x <= ln.a + b):
                        raise FillError(f"{key}: 同じ所に2つの書き込みが重なりました")
                edits.append((ln.a + a, ln.a + b, v, ln.number))
            patches.append(patch)
        except FillError as e:
            problems.append(str(e))
    if problems:
        raise FillError("\n".join(problems))
    return edits, patches


def apply(doc: TxtDocument, edits: list[tuple[int, int, str, int]]) -> bytes:
    """書く所の字のバイトだけを、ファイルの文字コードで作ったバイト列に置き換える。"""
    out = bytearray()
    at = 0
    for a, b, v, _n in sorted(edits, key=lambda e: (e[0], e[1])):
        x, y = doc.dec.span(doc.blob, a, b)
        try:
            enc = doc.dec.encode(v)
        except EncodingError as e:
            raise FillError(str(e)) from None
        out += doc.blob[at:x]
        out += enc
        at = y
    out += doc.blob[at:]
    return bytes(out)


def fill(doc: TxtDocument, values: dict, *, force: bool = False) -> FillResult:
    edits, patches = plan(doc, values)
    warnings = []
    if doc.dec.ascii_only and any(not v.isascii() for _a, _b, v, _n in edits):
        warnings.append("もとのファイルは ASCII の字だけで文字コードが決まらないので、UTF-8（BOM なし）で書きました")
    return FillResult(apply(doc, edits), patches, warnings)


# ====================================================================== 証明
@dataclass
class TxtProof:
    original_sha256: str
    new_sha256: str
    lines_total: int = 0
    lines_changed: int = 0
    bytes_replaced: int = 0
    unexpected: list[str] = field(default_factory=list)
    lint_errors: list[str] = field(default_factory=list)
    readback: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.unexpected and not self.lint_errors and not self.readback

    def headline(self) -> str:
        return (f"行 {self.lines_total:,} のうち、変わったのは書いた所のある {self.lines_changed} 行だけ。"
                f"書いた所の字（元の {self.bytes_replaced} バイト）のほかのバイト（ほかの字・改行・BOM）は、書く前と1バイトも違いません"
                + ("" if self.ok else f"（想定外 {len(self.unexpected) + len(self.lint_errors) + len(self.readback)} 件）"))


def prove(original: bytes, new: bytes, written: dict, encoding: str | None = None) -> TxtProof:
    """書き込み処理の結果を使わずに、元のファイルで書く所を決め直し、前後の 2 つのファイルを読み直して比べる。
    - バイト：書いた後のファイルは、元のファイルの「書く所の字のバイト」だけを、値をファイルの文字コードにしたバイト列に
      置き換えたものと、1バイトも違わない
    - 読み直し：文字コード・BOM・行の数・行ごとの改行が同じで、書いていない行の字は同じ。書いた行の字は
      「元の字の、書いた所だけを値に替えたもの」"""
    pr = TxtProof(hashlib.sha256(original).hexdigest(), hashlib.sha256(new).hexdigest())
    a = TxtDocument(original, encoding=encoding)
    sp = Spots(a)
    used: dict = {}
    by_line: dict[int, list] = {}
    edits: list[tuple[int, int, str]] = []
    for fid, value in written.items():
        if isinstance(value, dict) and "@" in value:
            try:
                it = resolve_write(sp, value["@"], used)
            except WriteError as e:
                pr.unexpected.append(f"書いた所 {fid} の目印が、元のファイルで決まりません：{e}")
                continue
            if it["id"] != fid or it["value"] != value["value"]:
                pr.unexpected.append(f"書いた所 {fid} が、目印から決まる所（{it['id']}「{it['value']}」）と違います")
                continue
            value = value["value"]
        else:
            it = sp.items.get(fid)
        if it is None or "span" not in it:
            pr.unexpected.append(f"書いた欄 {fid} が、元のファイルにありません")
            continue
        ln = it["para"]
        by_line.setdefault(ln.number, []).append((fid, it, value))
        try:
            e = _edit_of(it, value)
        except FillError as ex:
            pr.unexpected.append(str(ex))
            continue
        if e is not None:
            edits.append((ln.a + e[0], ln.a + e[1], e[2]))
    # バイト：元のファイルに、書く所の置き換えだけをしたもの
    want = bytearray()
    at = 0
    for x0, y0, v in sorted(edits):
        x, y = a.dec.span(original, x0, y0)
        if x < at:
            pr.unexpected.append("書く所が重なっています")
            break
        try:
            enc = a.dec.encode(v)
        except EncodingError as ex:
            pr.unexpected.append(str(ex))
            continue
        want += original[at:x] + enc
        pr.bytes_replaced += y - x
        at = y
    want += original[at:]
    if bytes(want) != new:
        k = next((i for i in range(min(len(want), len(new))) if want[i] != new[i]), min(len(want), len(new)))
        pr.unexpected.append(f"書いた後のファイルが、元のファイルの書く所だけを置き換えたものと違います（{k} バイト目から）")
    # 読み直し
    try:
        b = TxtDocument(new, encoding=encoding)
    except TxtError as e:
        pr.lint_errors.append(f"書き込み後のファイルを読めません：{e}")
        return pr
    if (a.dec.encoding, a.dec.bom) != (b.dec.encoding, b.dec.bom):
        pr.unexpected.append(f"文字コードが変わっています（{a.dec.label} → {b.dec.label}）")
    pr.lines_total = len(a.lines)
    if len(a.lines) != len(b.lines):
        pr.readback.append(f"行の数が変わりました（{len(a.lines)} → {len(b.lines)}）")
        return pr
    for x, y in zip(a.lines, b.lines):
        if x.terminator != y.terminator:
            pr.unexpected.append(f"{x.number} 行目の改行が変わっています（{NEWLINE_JA[x.terminator]} → {NEWLINE_JA[y.terminator]}）")
        items = by_line.get(x.number)
        if items is None:
            if x.text != y.text:
                pr.unexpected.append(f"書いていない {x.number} 行目の字が変わっています（「{x.text[:30]}」→「{y.text[:30]}」）")
            continue
        pr.lines_changed += 1
        expect = _expected_text(x, items)
        if expect is None or y.text != expect:
            pr.readback.append(f"{x.number} 行目：読み直すと「{y.text[:40]}」で、思っていた「{(expect or '')[:40]}」と違います")
    return pr


# ====================================================================== api から
def api_fill(blob: bytes, name: str, values: dict, force: bool = False, encoding: str | None = None) -> dict:
    from ..openkinds import OpenFormatError

    try:
        doc = TxtDocument(blob, name, encoding)
        res = fill(doc, values, force=force)
    except (FillError, TxtError) as e:
        raise OpenFormatError(str(e)) from None
    pr = prove(blob, res.blob, written_of(res.patches), encoding)
    result = {"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings),
              "unexpected": (pr.unexpected + pr.readback)[:20], "structure_errors": pr.lint_errors[:20],
              "text_read_back": pr.readback[:30],
              "ids_after_writing": "文字だけのファイルには、書いた所の印を置けません。書いた後のファイルでは、書いた字も印刷された字に"
                                   "数えるので、同じ行の後ろの空欄の id の字数が増えます（P3.c2 に「8」と書くと、P3.c3 だった空欄が P3.c4 になる）。"
                                   "続けて書くときは、書いた後のファイルを"
                                   "開き直して id を確かめるか、元のファイルに、まとめて書いてください"}
    anchored = {p.key: p.id for p in res.patches if p.anchor}
    if anchored:
        result["anchored"] = anchored
    told = {p.id: p.value for p in res.patches}
    # まだ空いている欄：書いた後のファイルでは id の番号が詰まるので、元のファイルの欄から、書いた所を除いたもの
    from .window import build

    win_fields = build(doc)["fields"]
    rest = [{"id": f["id"], "label": f.get("label", ""), "kind": f.get("kind", ""), "value": f.get("value"),
             **({"agent": True} if f.get("agent") else {})} for f in win_fields if f["id"] not in told]
    return {"blob": res.blob, "result": result, "readback": {}, "labels": {p.id: p.label for p in res.patches},
            "told": told, "rest": rest}
