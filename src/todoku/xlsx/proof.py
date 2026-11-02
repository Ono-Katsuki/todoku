"""書き込み後のファイルが、元から「書き込んだ箇所だけ」変わっていることを独立に確かめる。

書き込み処理の内部情報は使わず、2つのファイルを読み直して比べる。
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field

from ..formula.parser import addr
from .calc import values_equal
from .workbook import Workbook


@dataclass
class Proof:
    original_sha256: str
    new_sha256: str
    entries_total: int
    entries_identical: int  # 圧縮データまでバイト一致したエントリ
    entries_changed: list[str]
    cell_changes: list[dict] = field(default_factory=list)  # 値か数式が変わったセル
    unexpected: list[str] = field(default_factory=list)  # 許可していない変更
    xml_bytes_changed: int = 0  # 変わったXMLのうち、差し替わったバイト数
    xml_bytes_total: int = 0
    lint_errors: list[str] = field(default_factory=list)  # Excelが開くための構造の条件に反するもの
    lint_parts: int = 0
    control_changes: list[dict] = field(default_factory=list)  # チェックボックスの状態が変わったもの
    container: str = "zip"  # zip（xlsx）/ cfb（xls）
    records_total: int = 0  # xls：ブックのストリームのレコード数
    records_identical: int = 0

    @property
    def ok(self) -> bool:
        return not self.unexpected and not self.lint_errors

    def headline(self) -> str:
        n_in = sum(1 for c in self.cell_changes if c["kind"] == "input")
        n_f = sum(1 for c in self.cell_changes if c["kind"] == "formula")
        n_c = len(self.control_changes)
        parts = f"入力 {n_in}・計算結果 {n_f}" + (f"・チェックボックス {n_c}" if n_c else "")
        if self.container == "cfb":
            return (f"変更 {len(self.cell_changes) + n_c} か所（{parts}）。"
                    f"それ以外は元とバイト一致（複合ファイルのストリーム {self.entries_identical}/{self.entries_total} が完全一致、"
                    f"ブックのレコードも {self.records_identical:,}/{self.records_total:,} がバイト一致）")
        return (f"変更 {len(self.cell_changes) + n_c} か所（{parts}）。"
                f"それ以外は元とバイト一致（ZIPのエントリ {self.entries_identical}/{self.entries_total} が完全一致、"
                f"変わったXMLも {self.xml_bytes_total - self.xml_bytes_changed:,}/{self.xml_bytes_total:,} バイトが一致）")


def _common_affix(a: bytes, b: bytes) -> tuple[int, int]:
    n = min(len(a), len(b))
    i = 0
    while i < n and a[i] == b[i]:
        i += 1
    j = 0
    while j < n - i and a[len(a) - 1 - j] == b[len(b) - 1 - j]:
        j += 1
    return i, j


_CELL_SEG = re.compile(rb"<(?:\w+:)?c\b[^>]*?\br=\"([A-Z]+\d+)\"[^>]*?(?:/>|>.*?</(?:\w+:)?c>)", re.S)


def _segments(xml: bytes) -> list[tuple[bytes, bytes]]:
    """XMLを「セル要素」と「その間」に分ける。どちらも直前のセル番地を鍵にするので、鍵は重ならない。"""
    out = []
    last = 0
    prev = b"^"
    for m in _CELL_SEG.finditer(xml):
        out.append((b"gap:" + prev, xml[last:m.start()]))
        out.append((b"c:" + m[1], m[0]))
        prev = m[1]
        last = m.end()
    out.append((b"gap:" + prev, xml[last:]))
    return out


_VML_SHAPE = re.compile(rb"<v:shape\b[^>]*?\bid=\"([^\"]+)\"[^>]*>.*?</v:shape>", re.S)


def _vml_segments(xml: bytes) -> list[tuple[bytes, bytes]]:
    out = []
    last = 0
    prev = b"^"
    for m in _VML_SHAPE.finditer(xml):
        out.append((b"gap:" + prev, xml[last:m.start()]))
        out.append((b"shape:" + m[1], m[0]))
        prev = m[1]
        last = m.end()
    out.append((b"gap:" + prev, xml[last:]))
    return out


def _changed_bytes(a: bytes, b: bytes) -> int:
    """元のバイトのうち、書き換え後に残っていないバイト数。"""
    import difflib

    if b"<v:shape" in a[:200_000] and b"urn:schemas-microsoft-com:vml" in a[:2000]:
        # VML：図形ごとに比べる（図形の中だけを、前後の一致を除いて数える）
        mb = dict(_vml_segments(b))
        changed = 0
        for key, seg in _vml_segments(a):
            other = mb.get(key)
            if other == seg:
                continue
            if other is None:
                changed += len(seg)
                continue
            i, j = _common_affix(seg, other)
            changed += len(seg) - i - j
        return changed
    sa, sb = _segments(a), _segments(b)
    if len(sa) > 2:
        # セル単位で比べる（鍵が一意なので、同じ鍵のバイト列が同じかを見るだけでよい）
        mb = dict(sb)
        kept = sum(len(seg) for key, seg in sa if mb.get(key) == seg)
        return len(a) - kept
    i, j = _common_affix(a, b)
    ma, mb = a[i:len(a) - j], b[i:len(b) - j]
    # タグの区切り（>）ごとに分けて比べる（1バイトずつ比べると、大きな部品で時間がかかりすぎる）
    ta = re.split(rb"(?<=>)", ma)
    tb = re.split(rb"(?<=>)", mb)
    sm = difflib.SequenceMatcher(None, ta, tb, autojunk=False)
    kept = sum(sum(len(x) for x in ta[blk.a:blk.a + blk.size]) for blk in sm.get_matching_blocks())
    return len(ma) - kept


_CHECKED_MARKUP = [
    re.compile(rb'\s+checked="[^"]*"'),  # ctrlProp
    re.compile(rb"[ \t]*<x:Checked>[^<]*</x:Checked>[ \t]*\r?\n?"),  # VML
]


def _without_checked(data: bytes) -> bytes:
    for pat in _CHECKED_MARKUP:
        data = pat.sub(b"", data)
    return data


def prove(original: bytes, new: bytes, allowed_inputs: set[tuple[str, str]] | None = None,
          allowed_controls: set[tuple[str, str]] | None = None) -> Proof:
    if original[:4] == b"\xd0\xcf\x11\xe0":
        from ..xls.proof import prove as prove_xls

        return prove_xls(original, new, allowed_inputs, allowed_controls)
    from ..zipraw import RawZip

    za, zb = RawZip(original), RawZip(new)
    diff = {n for n in za.names() if n in set(zb.names()) and
            (za.get(n).local, za.get(n).data, za.get(n).descriptor) != (zb.get(n).local, zb.get(n).data, zb.get(n).descriptor)}
    # 変わったのがシートのXMLだけなら、そのシートだけを読み直して比べる（ほかはバイト一致なので中身も同じ）
    only = diff if diff and all(n.startswith("xl/worksheets/sheet") and n.endswith(".xml") for n in diff) else None
    a = Workbook(original, only_parts=only)
    b = Workbook(new, only_parts=only)
    names_a = a.zip.names()
    names_b = b.zip.names()
    unexpected = []
    if names_a != names_b:
        unexpected.append(f"ZIPのエントリ構成が変わりました: {sorted(set(names_a) ^ set(names_b))}")
    identical = 0
    changed = []
    xml_total = 0
    xml_changed = 0
    for n in names_a:
        if n not in names_b:
            continue
        ea, eb = a.zip.get(n), b.zip.get(n)
        if ea.local == eb.local and ea.data == eb.data and ea.descriptor == eb.descriptor:
            identical += 1
            continue
        changed.append(n)
        ra, rb = ea.read(), eb.read()
        xml_total += len(ra)
        xml_changed += _changed_bytes(ra, rb)
        if n.endswith(".vml") or "/ctrlProps/" in n:
            # コントロールの部品は、チェックの状態のほかは1バイトも変わっていてはいけない
            if _without_checked(ra) != _without_checked(rb):
                unexpected.append(f"{n}: チェックの状態のほかにも変更があります")

    changes = []
    for sa in a.sheets:
        try:
            sb = b.sheet(sa.name)
        except KeyError:
            unexpected.append(f"シート {sa.name} がなくなりました")
            continue
        keys = set(sa.cells) | set(sb.cells)
        for k in sorted(keys):
            ca, cb = sa.cells.get(k), sb.cells.get(k)
            va = ca.value if ca else None
            vb = cb.value if cb else None
            fa = ca.formula if ca else None
            fb = cb.formula if cb else None
            sty_a = ca.style if ca else sa.style_of(*k)
            sty_b = cb.style if cb else sb.style_of(*k)
            if fa != fb:
                unexpected.append(f"{sa.name}!{addr(*k)}: 数式が変わりました")
            if sty_a != sty_b:
                unexpected.append(f"{sa.name}!{addr(*k)}: 書式が変わりました")
            same = values_equal(va, vb) and type(va) is type(vb)
            if not same:
                kind = "formula" if fa is not None or (ca and ca.array_master) else "input"
                changes.append({"sheet": sa.name, "cell": addr(*k), "kind": kind, "before": va, "after": vb})
                if kind == "input" and allowed_inputs is not None and (sa.name, addr(*k)) not in allowed_inputs:
                    unexpected.append(f"{sa.name}!{addr(*k)}: 書き込みを指示していないセルが変わりました")
        if sa.merges != sb.merges:
            unexpected.append(f"{sa.name}: セル結合が変わりました")
        if len(sa.validations) != len(sb.validations):
            unexpected.append(f"{sa.name}: 入力規則が変わりました")
    control_changes = []
    for sa in a.sheets:
        if not b.has_sheet(sa.name):
            continue
        sb = b.sheet(sa.name)
        after = {c.shape_id: c for c in sb.controls}
        if len(sa.controls) != len(sb.controls):
            unexpected.append(f"{sa.name}: フォームコントロールの数が変わりました")
        for ca in sa.controls:
            cb = after.get(ca.shape_id)
            if cb is None:
                continue
            if (ca.kind, ca.link, ca.name) != (cb.kind, cb.link, cb.name):
                unexpected.append(f"{sa.name}: コントロール「{ca.name}」の設定が変わりました")
            if ca.checked == cb.checked:
                continue
            control_changes.append({"sheet": sa.name, "control": ca.name, "kind": "control", "link": ca.link,
                                    "before": ca.checked or "Unchecked", "after": cb.checked or "Unchecked"})
            if allowed_controls is not None and (sa.name, ca.name) not in allowed_controls:
                unexpected.append(f"{sa.name}: 指示していないチェックボックス「{ca.name}」が変わりました")
            if cb.link_cell and b.has_sheet(cb.link_cell[0]):
                lc = b.sheet(cb.link_cell[0]).cells.get(cb.link_cell[1:])
                lv = lc.value if lc else None
                if bool(lv is True) != cb.is_checked:
                    unexpected.append(f"{sa.name}: チェックボックス「{ca.name}」の表示と、つながったセル {cb.link} の値（{lv!r}）が食い違っています")
    if [d.text for d in a.names] != [d.text for d in b.names]:
        unexpected.append("名前付き範囲が変わりました")
    from .lint import lint

    lr = lint(new, parts=changed)
    before = set(lint(original, parts=changed).errors) if lr.errors else set()
    return Proof(
        lint_errors=[e for e in lr.errors if e not in before],  # 書き込みで新しく生じたものだけ
        lint_parts=lr.checked_parts,
        original_sha256=hashlib.sha256(original).hexdigest(),
        new_sha256=hashlib.sha256(new).hexdigest(),
        entries_total=len(names_a),
        entries_identical=identical,
        entries_changed=changed,
        cell_changes=changes,
        control_changes=control_changes,
        unexpected=unexpected,
        xml_bytes_changed=xml_changed,
        xml_bytes_total=xml_total,
    )
