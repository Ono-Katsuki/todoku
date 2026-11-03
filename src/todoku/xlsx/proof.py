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
    control_changes: list[dict] = field(default_factory=list)  # フォームの部品（チェックボックスなど）の状態が変わったもの
    sheets_outside_identical: int = 0  # 書き換えたシートの XML のうち、書いたセル・計算し直したセルのほかがバイト一致したもの
    sheets_changed: int = 0
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
        parts = f"入力 {n_in}・計算結果 {n_f}" + (f"・フォームの部品 {n_c}" if n_c else "")
        if self.container == "cfb":
            return (f"変更 {len(self.cell_changes) + n_c} か所（{parts}）。"
                    f"それ以外は元とバイト一致（複合ファイルのストリーム {self.entries_identical}/{self.entries_total} が完全一致、"
                    f"ブックのレコードも {self.records_identical:,}/{self.records_total:,} がバイト一致）")
        return (f"変更 {len(self.cell_changes) + n_c} か所（{parts}）。"
                f"それ以外は元とバイト一致（ZIPのエントリ {self.entries_identical}/{self.entries_total} が完全一致、"
                f"変わったXMLも {self.xml_bytes_total - self.xml_bytes_changed:,}/{self.xml_bytes_total:,} バイトが一致"
                + (f"、書き換えたシート {self.sheets_outside_identical}/{self.sheets_changed} は、書いたセルと計算し直したセルを"
                   "取り除くとバイト一致" if self.sheets_changed else "") + "）")


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
    # 部品の状態（チェック・選んだ番号・値・複数選択）。ctrlProp の属性と、VML の要素
    re.compile(rb'\s+(?:checked|sel|val|multiSel)="[^"]*"'),
    re.compile(rb"[ \t]*<x:(Checked|Sel|Val|MultiSel)>[^<]*</x:\1>[ \t]*\r?\n?"),
    # ActiveX の部品の控えの絵の参照（共有していた絵を分けたとき。どの絵を指すかは ActiveX の証明で確かめる）
    re.compile(rb'(?<=<v:imagedata )o:relid="[^"]*"'),
]


_ROW_EMPTY = re.compile(rb"<((?:\w+:)?row\b[^>]*)></(?:\w+:)?row>")
_ROW_BARE = re.compile(rb'<(?:\w+:)?row r="(\d+)"/>')
_ROW_NUM = re.compile(rb'<(?:\w+:)?row\b[^>]*?\br="(\d+)"')
_CONTROL_PR = re.compile(rb"<(?:\w+:)?controlPr\b[^>]*>")
_RID = re.compile(rb'\br:id="[^"]*"')
_AX_REL = re.compile(rb'<Relationship Id="rIdAx\d+" Type="[^"]*/image" Target="[^"]*"/>')
_DATA_EMPTY = re.compile(rb"<((?:\w+:)?sheetData)></(?:\w+:)?sheetData>")


def _outside_cells(a: bytes, b: bytes, keys: set[str]) -> str | None:
    """シートの XML から、書いてよいセル（書いたセル・計算し直したセル）の要素を両方で取り除き、残りがバイト一致かを見る。
    セルの要素がなかった所に書いたときにできる行の入れ物（新しい行・中身のなかった行）だけは、同じに扱う。
    一致すれば None、違えば最初に違う所を返す。"""
    def strip(xml: bytes) -> bytes:
        out, last = bytearray(), 0
        for m in _CELL_SEG.finditer(xml):
            if m[1].decode() in keys:
                out += xml[last:m.start()]
                last = m.end()
        out += xml[last:]
        x = _ROW_EMPTY.sub(rb"<\1/>", bytes(out))
        return _DATA_EMPTY.sub(rb"<\1/>", x)

    # ActiveX の部品の控えの絵を分けたとき：controlPr の r:id が、足した関係（rIdAx…）に変わる（どの絵かは ActiveX の証明で確かめる）
    ta, tb = _CONTROL_PR.findall(a), _CONTROL_PR.findall(b)
    if len(ta) == len(tb) and ta != tb:
        fix = {}
        for x, y in zip(ta, tb):
            if x != y and re.sub(rb'(\br:id=")rIdAx\d+"', lambda m: _RID.search(x)[0] if _RID.search(x) else m[0], y) == x:
                fix[y] = x
        if fix:
            b = _CONTROL_PR.sub(lambda m: fix.get(m[0], m[0]), b)
    sa, sb = strip(a), strip(b)
    rows_a = set(_ROW_NUM.findall(sa))
    sb = _ROW_BARE.sub(lambda m: b"" if m[1] not in rows_a else m[0], sb)
    sb = _DATA_EMPTY.sub(rb"<\1/>", sb)
    if sa == sb:
        return None
    i, _j = _common_affix(sa, sb)
    return sa[max(0, i - 60):i + 60].decode("utf-8", "replace")


def _without_full_calc(xml: bytes) -> bytes:
    """workbook.xml から、開いたときに計算し直す印（calcPr の fullCalcOnLoad、とその印だけの calcPr）を除く。"""
    xml = re.sub(rb'<(?:\w+:)?calcPr fullCalcOnLoad="1"/>', b"", xml)
    return re.sub(rb'\s+fullCalcOnLoad="[^"]*"', b"", xml)


def _without_checked(data: bytes) -> bytes:
    for pat in _CHECKED_MARKUP:
        data = pat.sub(b"", data)
    return data


def _rc_(a: str) -> tuple[int, int]:
    from ..formula.parser import parse_cell

    return parse_cell(a)


def _state_text(c) -> str:
    if c.kind == "Radio":
        return "選んだ" if c.checked == "Checked" else "選んでいない"
    if c.kind in ("Drop", "List"):
        return f"項目 {','.join(map(str, c.multi_sel))}" if c.multi_sel else f"項目 {c.sel or 0}"
    return f"値 {c.val}"


def _linked_controls_agree(b) -> list[str]:
    """書き込み後のファイルで、部品の表示と、つながったセルの値が合っているか（オプションボタン・ドロップダウン・スピンなど）。"""
    from .controls import CHOICE_KINDS, VALUE_KINDS, radio_groups

    out = []
    for sh in b.sheets:
        for g in radio_groups(b, sh):
            if g.link and g.consistent is False and "合いません" in g.why:
                out.append(f"{sh.name}: オプションボタンの表示と、つながったセルの値が食い違っています（{g.why}）")
        for c in sh.controls:
            if c.kind not in CHOICE_KINDS + VALUE_KINDS or not c.link_cell or not b.has_sheet(c.link_cell[0]):
                continue
            if c.kind == "List" and (c.sel_type or "single") != "single":
                continue
            lc = b.sheet(c.link_cell[0]).cells.get(c.link_cell[1:])
            lv = lc.value if lc else None
            want = (c.sel or 0) if c.kind in CHOICE_KINDS else c.val
            if want is None:
                continue
            ok = (lv in (None, "") and want == 0) or (isinstance(lv, (int, float)) and not isinstance(lv, bool) and lv == want)
            if not ok:
                out.append(f"{sh.name}: 部品「{c.name}」の表示（{want}）と、つながったセル {c.link} の値（{lv!r}）が食い違っています")
    return out


def _xls_rich_problems(original: bytes, new: bytes, rich) -> list[str]:
    """.xls の、字ごとの書式を残して書いたセル：書き換えなかった字の字体（FormatRun）が、書いた後も同じか。"""
    from ..xls.book import XlsWorkbook
    from .richtext import props_kept, xls_runs

    a, b = XlsWorkbook(original), XlsWorkbook(new)
    out = []
    for sn, cell in sorted(rich):
        k = (sn, *_rc_(cell))
        ia, ib = a.rich_cells.get(k), b.rich_cells.get(k)
        if ia is None:
            out.append(f"{sn}!{cell}: 元のセルの字ごとの書式を読み直せません")
            continue
        ta = a.shared_strings[ia]
        pa = xls_runs(ta, a.sst_runs[ia])
        if ib is None:
            cb = b.sheet(sn).cells.get(k[1:])
            tb = cb.value if cb is not None and isinstance(cb.value, str) else ""
            pb = [None] * len(tb)
        else:
            tb = b.shared_strings[ib]
            pb = xls_runs(tb, b.sst_runs[ib])
        out += [f"{sn}!{cell}: {x}" for x in props_kept(ta, pa, tb, pb)]
    return out


def _table_changes(a: Workbook, b: Workbook, unexpected: list) -> tuple[dict, set, set]:
    """テーブルの部品の変わり方を確かめる。返すのは (列の名前の書き換え {テーブル（小文字）: {元（小文字）: 新}},
    確かめたテーブルの部品, 構造化参照を書き換えたセル（あとで足す）)。

    - 列の名前・集計の行の見出し（totalsRowLabel）のほかの、テーブルの部品の変更は認めない
    - テーブルの部品の XML は、元の XML の列の属性と数式の構造化参照を、その書き換えで直したものとバイト一致
    - 書いた後のブックで、見出しのセルの字と列の名前、集計の行のセルの字と集計の行の見出しが合っている
      （元のブックで合っていた所）。数式の構造化参照は、元のブックで指せていた列を、書いた後も指せている"""
    from ..formula.values import to_text
    from . import tablerename as tr

    renames: dict = {}
    parts: set = set()
    if not getattr(a, "tables", None):
        return renames, parts, set()
    col_changes: dict = {}
    for key, ta in a.tables.items():
        tb = b.tables.get(key)
        if tb is None or not ta.part:
            if tb is None:
                unexpected.append(f"テーブル「{ta.name}」がなくなりました")
            continue
        xa = a.zip.read(ta.part) if a.zip.has(ta.part) else b""
        xb = b.zip.read(tb.part) if b.zip.has(tb.part) else b""
        if xa == xb:
            continue
        aa, ab = tr.column_attrs(xa), tr.column_attrs(xb)
        if len(aa) != len(ab) or (ta.r1, ta.c1, ta.r2, ta.c2) != (tb.r1, tb.c1, tb.r2, tb.c2):
            unexpected.append(f"テーブル「{ta.name}」の範囲か列の数が変わりました")
            continue
        cols: dict = {}
        for i, (x, y) in enumerate(zip(aa, ab)):
            for k in set(x) | set(y):
                if x.get(k) == y.get(k):
                    continue
                if k not in ("name", "totalsRowLabel") or y.get(k) is None:
                    unexpected.append(f"テーブル「{ta.name}」の列 {i + 1} の定義（{k}）が変わりました")
                    continue
                cols.setdefault(i, {})[k] = y[k]
                if k == "name":
                    renames.setdefault(key, {})[x["name"].lower()] = y["name"]
        col_changes[ta.part] = (ta, cols)
    for part, (ta, cols) in col_changes.items():
        if tr.table_xml(a.zip.read(part), ta.name, cols, renames) != b.zip.read(part):
            unexpected.append(f"{part}: テーブルの列の名前・集計の行の見出しと、構造化参照の書き換えのほかにも変更があります")
        parts.add(part)
    # ほかのテーブルの部品も、構造化参照の書き換え（calculatedColumnFormula など）で変わることがある
    for key, ta in a.tables.items():
        if ta.part in parts or not ta.part or not a.zip.has(ta.part) or not b.zip.has(ta.part):
            continue
        xa, xb = a.zip.read(ta.part), b.zip.read(ta.part)
        if xa != xb and renames and tr.table_xml(xa, ta.name, {}, renames) == xb:
            parts.add(ta.part)
    # 書いた後のブックで、見出し・集計の行のセルの字が、テーブルの定義と合っているか（元で合っていた所）
    for key, ta in a.tables.items():
        tb = b.tables.get(key)
        if tb is None or not b.has_sheet(tb.sheet) or not tb.part or not b.zip.has(tb.part):
            continue
        attrs_a = tr.column_attrs(a.zip.read(ta.part)) if a.zip.has(ta.part) else []
        attrs_b = tr.column_attrs(b.zip.read(tb.part))
        for i in range(min(len(attrs_a), len(attrs_b))):
            c = tb.c1 + i

            def text(wb, sheet, r):
                cell = wb.sheet(sheet).cells.get((r, c))
                return to_text(cell.value) if cell is not None and cell.value is not None else ""

            if tb.header_rows:
                if text(a, ta.sheet, ta.r1) == attrs_a[i].get("name") and text(b, tb.sheet, tb.r1) != attrs_b[i].get("name"):
                    unexpected.append(f"{tb.sheet}!{addr(tb.r1, c)}: 表の見出しの字「{text(b, tb.sheet, tb.r1)}」が、テーブル「{tb.name}」の"
                                      f"列の名前「{attrs_b[i].get('name')}」と違います（Excel は開くときに修復を求めます）")
            if tb.totals_rows and attrs_b[i].get("totalsRowLabel") is not None:
                if text(a, ta.sheet, ta.r2) == attrs_a[i].get("totalsRowLabel") and \
                        text(b, tb.sheet, tb.r2) != attrs_b[i].get("totalsRowLabel"):
                    unexpected.append(f"{tb.sheet}!{addr(tb.r2, c)}: 集計の行の字「{text(b, tb.sheet, tb.r2)}」が、テーブル「{tb.name}」の"
                                      f"集計の行の見出し「{attrs_b[i].get('totalsRowLabel')}」と違います")
    if renames:
        unexpected += _dangling_struct_refs(a, b)
    return renames, parts, set()


def _renamed_formula(a: Workbook, sheet: str, k: tuple, fa: str, fb: str, renames: dict) -> bool:
    """書いた後の数式 fb が、元の数式 fa の構造化参照の列の名前を書き換えただけのものか（構文木で比べる）。"""
    from ..formula.parser import parse
    from . import tablerename as tr

    here = tr.table_at(a, sheet, *k)
    try:
        return tr.rename_ast(parse(fa), renames, here.name if here else None) == parse(fb)
    except Exception:  # noqa: BLE001  読めない数式：書き換えだけかを確かめられない
        return False


def _dangling_struct_refs(a: Workbook, b: Workbook) -> list[str]:
    """元のブックで、ある列を指していた構造化参照が、書いた後のブックでは、ない列を指している（列の名前を書き換えたのに、
    参照を書き換えていない）。"""
    from ..formula.parser import parse
    from . import tablerename as tr

    def bad(wb) -> set:
        out = set()
        cols = {k: {c.lower() for c in t.columns} for k, t in wb.tables.items()}
        for sh in wb.sheets:
            for (r, c), cell in sh.cells.items():
                f = cell.formula
                if not f or "[" not in f:
                    continue
                try:
                    ast = parse(f)
                except Exception:  # noqa: BLE001
                    continue
                here = tr.table_at(wb, sh.name, r, c)
                for ref in tr.struct_refs(ast):
                    tn = (ref.table or (here.name if here else "")).lower()
                    for col in (ref.col1, ref.col2):
                        if col is not None and tn in cols and col.lower() not in cols[tn]:
                            out.add((sh.name, addr(r, c), ref.table or tn, col))
        return out

    before = bad(a)
    return [f"{s}!{c}: 数式の構造化参照 {t}[{col}] が、テーブルにない列を指しています（列の名前と参照がずれています）"
            for s, c, t, col in sorted(bad(b) - before)]


def prove(original: bytes, new: bytes, allowed_inputs: set[tuple[str, str]] | None = None,
          allowed_controls: set[tuple[str, str]] | None = None, overwrite: set[tuple[str, str]] | None = None,
          rich: set[tuple[str, str]] | None = None) -> Proof:
    """overwrite：数式が値に置き換わってよいセル（セルの値を変えたときに動くマクロが、数式のセルに値を書いたもの）。
    rich：字ごとの書式（取り消し線・色）を残して字の一部を書き換えたセル {(シート, 番地)}。書き換えなかった字の書式が、
    書いた後も同じかを字ごとに確かめる（xlsx/richtext.kept_problems）。
    allowed_controls：書き込みを指示した部品 {(シート, 部品の名前)}。{(シート, 部品の名前): 書いた状態} でもよく、
    古いExcel（.xls）では、部品がその状態になっているかも確かめる。"""
    if original[:4] == b"\xd0\xcf\x11\xe0":
        from ..xls.proof import prove as prove_xls

        pr = prove_xls(original, new, allowed_inputs, allowed_controls)
        if rich:
            pr.unexpected += _xls_rich_problems(original, new, rich)
        return pr
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
    added = [n for n in names_b if n not in set(names_a)]  # ActiveX の部品の、共有していた控えの絵から分けた絵（下で確かめる）
    if [n for n in names_a if n not in set(names_b)] or names_b[:len(names_a)] != names_a or \
            any("/media/" not in n for n in added):
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
        if "/activeX/" in n or "/media/" in n:
            continue  # ActiveX の部品の中身と控えの絵（バイナリ）：下の ActiveX の証明で確かめる
        xml_total += len(ra)
        xml_changed += _changed_bytes(ra, rb)
        if n.endswith(".vml") or "/ctrlProps/" in n:
            # コントロールの部品は、チェックの状態のほかは1バイトも変わっていてはいけない
            if _without_checked(ra) != _without_checked(rb):
                unexpected.append(f"{n}: 部品の状態（チェック・選んだ項目・値）のほかにも変更があります")

    # テーブルの見出し・集計の行の見出しを書き換えたとき（force）：列の名前・集計の行の見出しと、その列を指す構造化参照だけが、
    # 見出しのセルの字に合わせて変わったこと（tablerename）
    renames, table_parts, renamed_cells = _table_changes(a, b, unexpected)
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
            if fa != fb and not (fb is None and overwrite and (sa.name, addr(*k)) in overwrite):
                if renames and fa is not None and fb is not None and _renamed_formula(a, sa.name, k, fa, fb, renames):
                    renamed_cells.add((sa.name, addr(*k)))
                else:
                    unexpected.append(f"{sa.name}!{addr(*k)}: 数式が変わりました")
            if sty_a != sty_b:
                unexpected.append(f"{sa.name}!{addr(*k)}: 書式が変わりました")
            if cb is not None and cb.schema_error and not (ca is not None and ca.schema_error == cb.schema_error):
                # 値が一つに決まらないセル（<v> が二つなど）：Excel と LibreOffice で見える値が違うので、書き込みとして通さない
                unexpected.append(f"{sa.name}!{addr(*k)}: セルの形が規格（CT_Cell）に合いません（{cb.schema_error}）")
            same = values_equal(va, vb) and type(va) is type(vb)
            if not same or (fa != fb and fb is None):
                kind = "formula" if (fa is not None and fb is not None) or (ca and ca.array_master) else "input"
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
            if getattr(ca, "ax", None) is not None:
                continue  # ActiveX の部品は、下の ActiveX の証明で確かめる
            from .controls import state

            if state(ca) == state(cb):
                continue
            if ca.kind != "CheckBox":
                control_changes.append({"sheet": sa.name, "control": ca.name, "kind": "control", "link": ca.link,
                                        "before": _state_text(ca), "after": _state_text(cb)})
                if allowed_controls is not None and (sa.name, ca.name) not in allowed_controls:
                    unexpected.append(f"{sa.name}: 指示していない部品「{ca.name}」が変わりました")
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
    # 書き換えた部品のバイトの比べ：シートの XML は、書いてよいセルの要素のほかはバイト一致。workbook.xml は開いたときに計算し直す印
    # だけ。ほかに変わってよい部品は、calcChain・フォームの部品（ctrlProps・VML・ActiveX・控えの絵）だけ（それぞれ下と上で確かめる）
    by_path = {sa.path: sa.name for sa in a.sheets}
    wb_path = getattr(a, "wb_path", None)
    n_sheets = n_same = 0
    for n in changed:
        if n in by_path:
            keys = {c["cell"] for c in changes if c["sheet"] == by_path[n]}
            keys |= {cell for sh, cell in (allowed_inputs or ()) if sh == by_path[n]}
            keys |= {cell for sh, cell in (overwrite or ()) if sh == by_path[n]}
            keys |= {cell for sh, cell in renamed_cells if sh == by_path[n]}
            n_sheets += 1
            where = _outside_cells(a.zip.read(n), b.zip.read(n), keys)
            if where is None:
                n_same += 1
            else:
                unexpected.append(f"{n}: 書いたセル・計算し直したセルのほかのバイトが変わっています（{where!r} のあたり）")
        elif n == wb_path:
            if _without_full_calc(a.zip.read(n)) != _without_full_calc(b.zip.read(n)):
                unexpected.append(f"{n}: 開いたときに計算し直す印のほかにも変更があります")
        elif n.endswith(".rels"):
            # ActiveX の部品の控えの絵を分けたときに、分けた絵への関係を足すことだけ（どの絵かは ActiveX の証明で確かめる）
            if _AX_REL.sub(b"", b.zip.read(n)) != a.zip.read(n):
                unexpected.append(f"{n}: 分けた控えの絵への関係を足すほかにも変更があります")
        elif n in table_parts:
            pass  # テーブルの部品：_table_changes で確かめた
        elif not (n == getattr(a, "calc_chain", None) or n.endswith(".vml") or "/ctrlProps/" in n or "/activeX/" in n
                  or "/media/" in n):
            unexpected.append(f"{n}: 書き込みで変わってよい部品ではありません")
    # calcChain.xml：変わってよいのは、数式を値に置き換えたセルの記録を取り除くことだけ
    chain = getattr(a, "calc_chain", None)
    if chain and a.zip.has(chain) and b.zip.has(chain) and a.zip.read(chain) != b.zip.read(chain):
        from .edit import calc_chain_entries

        ea, eb = set(calc_chain_entries(a.zip.read(chain))), set(calc_chain_entries(b.zip.read(chain)))
        ids = {sh.sheet_id: sh.name for sh in a.sheets}
        removed = ea - eb
        if eb - ea:
            unexpected.append("calcChain.xml に、記録が足されました")
        for sid, cell in sorted(removed):
            sn = ids.get(sid)
            ca_ = a.sheet(sn).cells.get(_rc_(cell)) if sn else None
            cb_ = b.sheet(sn).cells.get(_rc_(cell)) if sn and b.has_sheet(sn) else None
            if not (ca_ is not None and ca_.formula is not None and (cb_ is None or cb_.formula is None)):
                unexpected.append(f"calcChain.xml から、数式の残っているセル {sn}!{cell} の記録が取り除かれました")
    # ActiveX の部品（MS-OFORMS）：指示した部品の、指示した性質のバイトと大きさの記録だけが変わり、控えの絵は書いた状態の絵
    if any(getattr(c, "ax", None) is not None for s in a.sheets for c in s.controls):
        from ..axfile import prove_excel

        ax_changes, ax_bad, _ax_parts = prove_excel(a, b, changed, allowed_controls, added)
        control_changes += ax_changes
        unexpected += ax_bad
    if rich:
        # 字ごとの書式を残したセル：元の書式（<rPr>）が、書き換えなかった字に残っているか
        from .richtext import cell_rst, kept_problems

        fa, fb = Workbook(original), Workbook(new)
        for sn, cell in sorted(rich):
            r_, c_ = _rc_(cell)
            try:
                before_r, after_r = cell_rst(fa, sn, r_, c_), cell_rst(fb, sn, r_, c_)
            except (ValueError, KeyError) as e:
                unexpected.append(f"{sn}!{cell}: 字ごとの書式を読み直せません（{e}）")
                continue
            if before_r is None:
                unexpected.append(f"{sn}!{cell}: 元のセルの字ごとの書式を読み直せません")
                continue
            unexpected += [f"{sn}!{cell}: {x}" for x in kept_problems(before_r, after_r)]
    before_bad = set(_linked_controls_agree(a))
    unexpected += [x for x in _linked_controls_agree(b) if x not in before_bad]  # 書き込みで新しく生じた食い違いだけ
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
        sheets_outside_identical=n_same,
        sheets_changed=n_sheets,
    )
