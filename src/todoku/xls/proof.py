"""古いExcel（.xls）への書き込みの証明と構造の検査、窓とは別の読み方による文字の取り出し。

証明は、書き込み処理の内部情報を使わず、前後の2つのファイルを読み直して比べる。
- 複合ファイルの中のほかのストリームは、バイト一致
- ブックのストリームのレコードは、書き込んだセル・計算結果・フォームの部品の状態・位置を持つレコード
  （BOUNDSHEET・INDEX・DBCELL・EXTSST・SST の追加・ROW の列の範囲）のほかは、バイト一致
- 部品の OBJ（と続く CONTINUE）は、状態のバイト（チェック・選んだ項目・値・複数選択の印。どこにあるかは
  MS-XLS 2.4.181 Obj のとおりに読み直して決める）のほかはバイト一致。状態が変わった部品は、書き込みを指示した
  部品であり、指示した状態になっていて、つながったセルの値と合っていること
- 計算結果が変わった数式は、元のブックに書いた値を入れて計算し直した結果と同じであること。計算し直すと変わるはずの数式が
  変わっていなければ、それも知らせる（書いたセルに依存する数式の結果だけが、正しい値に変わったことを確かめる）
- Excel 5.0/95（BIFF5）のブックも同じように確かめる。Excel 97 と 95 の両方の形（Workbook と Book）を持つブックは、
  両方のストリームをそれぞれ確かめ、書いたセルと変わった数式の値が、二つの形で同じであることも確かめる
"""

from __future__ import annotations

import hashlib
import struct

from ..cfb import CFB
from ..formula.parser import addr
from ..xlsx.calc import values_equal
from ..xlsx.controls import state
from ..xlsx.proof import Proof, _linked_controls_agree, _state_text
from .book import ERRORS, Seg, XlsWorkbook, records, xlstring
from .obj import without_state

CELL_T = {0x0203, 0x027E, 0x00FD, 0x0204, 0x00D6, 0x0201, 0x0205, 0x0006, 0x00BD, 0x00BE}
POINTER_T = {0x0085, 0x020B, 0x00D7, 0x00FF, 0x01C1}  # 位置を持つレコード（と、再計算の印）


def _is_globals(bof: bytes) -> bool:
    return len(bof) >= 4 and struct.unpack_from("<H", bof, 2)[0] == 0x0005


def split(data: bytes, starts: list[int] | None = None):
    """ブックのストリームを (ブック全体のレコード, [シートのレコード]) に分ける。starts を渡すと、シートの部分の BOF の位置を足す。"""
    out_glob: list = []
    out_sheets: list[list] = []
    depth = 0
    target = None
    for _pos, t, b in records(data):
        if t == 0x0809 and depth == 0:
            target = out_glob if _is_globals(b) else []
            if target is not out_glob:
                out_sheets.append(target)
                if starts is not None:
                    starts.append(_pos)
        if t == 0x0809:
            depth += 1
        if target is not None:
            target.append((t, b))
        if t == 0x000A:
            depth -= 1
            if depth == 0:
                target = None
    return out_glob, out_sheets


def worksheet_parts(wb) -> tuple[list, list[tuple[str, list]], list[list]]:
    """(ブック全体のレコード, [(ワークシートの名前, その部分のレコード)], [ほかの部分のレコード])。シートと部分は、
    BOUNDSHEET の位置（lbPlyPos）で結ぶ（グラフのシートなど、セルのないシートの部分が間にあっても、ずれない）。"""
    starts: list[int] = []
    glob, parts = split(wb.data, starts)
    by_pos = dict(zip(starts, range(len(parts))))
    out, used = [], set()
    for sh in wb.sheets:
        if sh.state.startswith("non-") or not str(sh.path).startswith("#"):
            continue
        k = by_pos.get(int(str(sh.path)[1:]))
        if k is not None:
            out.append((sh.name, parts[k]))
            used.add(k)
    return glob, out, [p for k, p in enumerate(parts) if k not in used]


def _cells(recs) -> dict[tuple[int, int], tuple]:
    """シートのセル → (種類, 値に関わる中身)。書式番号の違いも比べる。"""
    out = {}
    depth = 0
    last = None
    for t, b in recs:
        if t == 0x0809:
            depth += 1
            continue
        if t == 0x000A:
            depth -= 1
            continue
        if depth != 1:
            continue
        if t in (0x00BD, 0x00BE):
            r, cf = struct.unpack_from("<HH", b, 0)
            w = 6 if t == 0x00BD else 2
            n = (len(b) - 6) // w
            for k in range(n):
                item = b[4 + w * k:4 + w * (k + 1)]
                out[(r, cf + k)] = ("rk" if t == 0x00BD else "blank", item)
        elif t in CELL_T and len(b) >= 6:
            r, c = struct.unpack_from("<HH", b, 0)
            if t == 0x027E:
                out[(r, c)] = ("rk", b[4:])
            elif t == 0x0201:
                out[(r, c)] = ("blank", b[4:6])
            elif t == 0x0006:
                cce = struct.unpack_from("<H", b, 20)[0]
                out[(r, c)] = ("formula", b[4:6], b[6:14], b[14:16], b[22:22 + cce], b[22 + cce:])
                last = (r, c)
            else:
                out[(r, c)] = (hex(t), b[4:])
        elif t == 0x0207 and last is not None:
            out[last] = out[last] + (("string", b),)
    return out


def prove(original: bytes, new: bytes, allowed_inputs: set[tuple[str, str]] | None = None,
          allowed_controls: set[tuple[str, str]] | None = None) -> Proof:
    pr = _prove_one(original, new, allowed_inputs, allowed_controls, None)
    a = XlsWorkbook(original)
    if not a.dual_format:
        return pr
    # Excel 97 と 95 の両方の形：Book（BIFF5）も同じように確かめ、二つの形の値が合うことを確かめる
    book = next(n for n in CFB(original).streams() if n.lower() == "book")
    if book in pr.entries_changed:
        pr.unexpected = [u for u in pr.unexpected if not u.startswith(f"ストリーム {book} が変わりました")]
        p5 = _prove_one(original, new, allowed_inputs, allowed_controls, "Book")
        pr.unexpected += [f"Book（Excel 95 の形）: {u}" for u in p5.unexpected]
        pr.lint_errors += [f"Book（Excel 95 の形）: {u}" for u in p5.lint_errors]
        pr.records_total += p5.records_total
        pr.records_identical += p5.records_identical
    b8, b5 = XlsWorkbook(new), XlsWorkbook(new, stream="Book")
    keys = {(c["sheet"], c["cell"]) for c in pr.cell_changes}
    for sheet, a1 in sorted(keys):
        from ..formula.parser import parse_cell

        r, c = parse_cell(a1)
        x = b8.sheet(sheet).cells.get((r, c)) if b8.has_sheet(sheet) else None
        y = b5.sheet(sheet).cells.get((r, c)) if b5.has_sheet(sheet) else None
        vx, vy = (x.value if x else None), (y.value if y else None)
        if not (values_equal(vx, vy) or (vx in (None, "") and vy in (None, ""))):
            pr.unexpected.append(f"{sheet}!{a1}: 書いたあとの値が Workbook（{vx!r}）と Book（{vy!r}）で食い違います")
    return pr


def _recalc_check(a: XlsWorkbook, cell_changes: list[dict]) -> list[str]:
    """変わった数式の結果が、元のブックに書いた値を入れて計算し直した結果と同じか（書き込み処理とは別に計算する）。"""
    from ..formula.parser import parse_cell
    from ..xlsx.calc import recalc

    inputs = {}
    for ch in cell_changes:
        if ch["kind"] == "input":
            r, c = parse_cell(ch["cell"])
            inputs[(ch["sheet"], r, c)] = ch["after"]
    if not inputs:
        return [f"{ch['sheet']}!{ch['cell']}: 書いたセルがないのに、数式の計算結果が変わりました"
                for ch in cell_changes if ch["kind"] == "formula"]
    want, problems = recalc(a, inputs)
    out = []
    got = {}
    for ch in cell_changes:
        if ch["kind"] != "formula":
            continue
        r, c = parse_cell(ch["cell"])
        got[(ch["sheet"], r, c)] = ch["after"]
    for k, v in got.items():
        if k in problems:
            continue
        if k not in want:
            out.append(f"{k[0]}!{addr(k[1], k[2])}: 書いた値に依存しない数式の計算結果が変わりました")
        elif not values_equal(want[k], v):
            out.append(f"{k[0]}!{addr(k[1], k[2])}: 数式の計算結果が {v!r} ですが、計算し直すと {want[k]!r} です")
    for k, v in want.items():
        if k not in got and k not in problems:
            out.append(f"{k[0]}!{addr(k[1], k[2])}: 書いた値で計算結果が {v!r} に変わるはずの数式が、変わっていません")
    return out


def _prove_one(original: bytes, new: bytes, allowed_inputs, allowed_controls, stream: str | None) -> Proof:
    a_cfb, b_cfb = CFB(original), CFB(new)
    unexpected: list[str] = []
    sa, sb = a_cfb.streams(), b_cfb.streams()
    if set(sa) != set(sb):
        unexpected.append(f"複合ファイルのストリームの構成が変わりました: {sorted(set(sa) ^ set(sb))}")
    a = XlsWorkbook(original, stream=stream)
    b = XlsWorkbook(new, stream=stream)
    identical = sum(1 for n in sa if n in sb and a_cfb.read(n) == b_cfb.read(n))
    changed = [n for n in sa if n in sb and a_cfb.read(n) != b_cfb.read(n)]
    ax_seen: set = set()
    if any(getattr(c, "ax", None) is not None for s_ in a.sheets for c in s_.controls):
        # ActiveX の部品（Ctls・MBD の入れ物）：書いた部品の、書いた性質だけが変わっていること
        from ..axbin import prove_xls

        ax_changes, ax_bad, ax_seen = prove_xls(a, b, allowed_controls)
        unexpected += ax_bad
    for n in changed:
        if n != a.stream_name and n not in ax_seen and stream is None:
            unexpected.append(f"ストリーム {n} が変わりました（ブック以外は変えないはずです）")
    ga, sha, other_a = worksheet_parts(a)
    gb, shb, other_b = worksheet_parts(b)
    # ブック全体の部分：位置のレコードと SST（文字列の追加）のほかは同じであること
    def norm_glob(g):
        out = []
        for t, body in g:
            if t in POINTER_T:
                out.append((t, body[4:] if t == 0x0085 else b""))  # BOUNDSHEET は位置以外を比べる
            elif t in (0x00FC, 0x003C):
                continue
            else:
                out.append((t, body))
        return out

    if norm_glob(ga) != norm_glob(gb):
        unexpected.append("ブック全体の設定のレコードが変わりました")
    # 共有文字列：元の文字列は同じ番号のまま残っていること（足すのは後ろだけ）
    if b.shared_strings[:len(a.shared_strings)] != a.shared_strings:
        unexpected.append("共有文字列の並びが変わりました")
    cell_changes, control_changes = [], []
    if [n for n, _r in sha] != [n for n, _r in shb]:
        unexpected.append("ワークシートの並びが変わりました")
    if other_a != other_b:
        unexpected.append("ワークシートでない部分（グラフのシートなど）が変わりました")
    for (name, ra), (_nb, rb) in zip(sha, shb):
        ca, cb = _cells(ra), _cells(rb)
        sh_a, sh_b = a.sheet(name), b.sheet(name)
        for k in sorted(set(ca) | set(cb)):
            x, y = ca.get(k), cb.get(k)
            if x == y:
                continue
            r, c = k[0] + 1, k[1] + 1
            va = sh_a.cells.get((r, c))
            vb = sh_b.cells.get((r, c))
            if x is not None and x[0] == "formula":
                if y is None or y[0] != "formula" or y[4] != x[4] or y[5] != x[5] or y[1] != x[1]:
                    unexpected.append(f"{name}!{addr(r, c)}: 数式か書式が変わりました")
                    continue
                cell_changes.append({"sheet": name, "cell": addr(r, c), "kind": "formula",
                                     "before": va.value if va else None, "after": vb.value if vb else None})
                continue
            xa = x[1][:2] if x and x[0] not in ("rk", "blank") else (x[1][:2] if x else None)
            ya = y[1][:2] if y and y[0] not in ("rk", "blank") else (y[1][:2] if y else None)
            if xa is not None and ya is not None and xa != ya:
                unexpected.append(f"{name}!{addr(r, c)}: 書式が変わりました")
            same = va is not None and vb is not None and values_equal(va.value, vb.value) and type(va.value) is type(vb.value)
            if same:
                continue  # 同じ値の書き方が変わっただけ（MULBLANK の分け方など）
            cell_changes.append({"sheet": name, "cell": addr(r, c), "kind": "input",
                                 "before": va.value if va else None, "after": vb.value if vb else None})
            if allowed_inputs is not None and (name, addr(r, c)) not in allowed_inputs:
                unexpected.append(f"{name}!{addr(r, c)}: 書き込みを指示していないセルが変わりました")
        # セル以外のレコード（部品の状態・位置のレコード・ROW の列の範囲のほか）は同じであること
        def others(rs):
            out = []
            k = 0
            while k < len(rs):
                t, body = rs[k]
                k += 1
                if t in CELL_T or t in (0x0207, 0x00D7, 0x020B):
                    continue
                if t == 0x0208:
                    out.append((t, body[:2] + body[6:]))  # 列の範囲（colMic・colMac）は変わってよい
                elif t == 0x005D:
                    # 部品：続く CONTINUE もつなぎ、状態（チェック・選んだ項目・値・複数選択）のバイトのほかを比べる
                    parts = [body]
                    while k < len(rs) and rs[k][0] == 0x003C:
                        parts.append(rs[k][1])
                        k += 1
                    joined = without_state(b"".join(parts))
                    from ..axbin import xls_pict

                    pict = xls_pict(joined)
                    if pict is not None and "flags_at" in pict:
                        # 書いた部品だけに立てる、絵を作り直させる印（fPrintCalc・fAutoLoad）：ActiveX の証明で確かめる
                        from ..axbin import FLAGS_REFRESH

                        j = bytearray(joined)
                        struct.pack_into("<H", j, pict["flags_at"], pict["flags"] & ~FLAGS_REFRESH)
                        joined = bytes(j)
                    if pict is not None and "pos_at" in pict:
                        # ActiveX の部品の位置と大きさ（Ctls の中）：上の ActiveX の証明で、指す中身を読み直して確かめる
                        j = bytearray(joined)
                        j[pict["pos_at"]:pict["pos_at"] + 4] = bytes(4)
                        if "size_at" in pict:
                            j[pict["size_at"]:pict["size_at"] + 4] = bytes(4)
                        joined = bytes(j)
                    out.append((t, tuple(len(p) for p in parts), joined))
                else:
                    out.append((t, body))
            return out

        oa, ob = others(ra), others(rb)
        if [x for x in oa if x[0] != 0x0208] != [x for x in ob if x[0] != 0x0208]:
            unexpected.append(f"{name}: セル以外のレコードが変わりました")
        from ..xlsx.controls import ensure_places

        ensure_places(a, sh_a)  # 同じ名前の組：指示した組と別の組を、id で見分ける
        ctl_a = {c.shape_id: c for c in sh_a.controls}
        for cb_ in sh_b.controls:
            ca_ = ctl_a.get(cb_.shape_id)
            if getattr(cb_, "ax", None) is not None:
                continue  # ActiveX の部品は、上の ActiveX の証明で確かめた
            if ca_ is not None and state(ca_) != state(cb_):
                control_changes.append(_control_change(name, ca_, cb_))
                unexpected += _control_problems(b, name, ca_, cb_, allowed_controls)
    # 部品の表示と、つながったセルの値が合っていること（書き込みで新しく生じた食い違いだけ）
    before_bad = set(_linked_controls_agree(a))
    unexpected += [x for x in _linked_controls_agree(b) if x not in before_bad]
    unexpected += _recalc_check(a, cell_changes)
    lint_errors = lint(new, stream).errors
    before = set(lint(original, stream).errors) if lint_errors else set()
    ra_all = [(t, body) for _p, t, body in records(a.data)]
    rb_all = [(t, body) for _p, t, body in records(b.data)]
    from collections import Counter

    common = Counter(ra_all) & Counter(rb_all)
    if ax_seen:
        control_changes += ax_changes
    return Proof(
        container="cfb", records_total=len(ra_all), records_identical=sum(common.values()),
        original_sha256=hashlib.sha256(original).hexdigest(), new_sha256=hashlib.sha256(new).hexdigest(),
        entries_total=len(sa), entries_identical=identical, entries_changed=changed,
        cell_changes=cell_changes, control_changes=control_changes, unexpected=unexpected,
        xml_bytes_changed=0, xml_bytes_total=0,
        lint_errors=[e for e in lint_errors if e not in before], lint_parts=1)


def _control_change(sheet: str, ca, cb) -> dict:
    if ca.kind == "CheckBox":
        before, after = ca.checked or "Unchecked", cb.checked or "Unchecked"
    else:
        before, after = _state_text(ca), _state_text(cb)
    return {"sheet": sheet, "control": ca.ref_name, "kind": "control", "link": ca.link, "before": before, "after": after}


def _control_problems(b: XlsWorkbook, sheet: str, ca, cb, allowed_controls) -> list[str]:
    """状態が変わった部品の確かめ：書き込みを指示した部品か。指示した状態（allowed_controls が
    {(シート, 部品): 変える状態} のとき）になっているか。チェックボックスなら、つながったセルの値と合っているか。"""
    out = []
    what = "チェックボックス" if ca.kind == "CheckBox" else "部品"
    if allowed_controls is not None and (sheet, ca.ref_name) not in allowed_controls:
        out.append(f"{sheet}: 指示していない{what}「{ca.ref_name}」が変わりました")
        return out
    want = allowed_controls.get((sheet, ca.ref_name)) if isinstance(allowed_controls, dict) else None
    if want is not None and not isinstance(want, dict):
        want = {"checked": want}  # チェックボックス：書いた状態（"Checked" / None）
    for key, v in (want or {}).items():
        got = getattr(cb, key, None)
        same = tuple(got or ()) == tuple(v or ()) if key == "multi_sel" else got == v
        if not same:
            out.append(f"{sheet}: {what}「{ca.name}」の状態（{key}）が、書き込みを指示した {v!r} ではなく {got!r} になっています")
    if ca.kind == "CheckBox" and cb.link_cell and b.has_sheet(cb.link_cell[0]):
        lc = b.sheet(cb.link_cell[0]).cells.get(cb.link_cell[1:])
        lv = lc.value if lc else None
        if bool(lv is True) != cb.is_checked:
            out.append(f"{sheet}: チェックボックス「{ca.name}」の表示と、つながったセル {cb.link} の値（{lv!r}）が食い違っています")
    return out


# ====================================================================== 構造の検査
class _Lint:
    def __init__(self):
        self.errors: list[str] = []
        self.checked_parts = 1

    @property
    def ok(self) -> bool:
        return not self.errors


def lint(blob: bytes, stream: str | None = None) -> _Lint:
    """Excelが修復なしに開くための条件：レコードの長さ・BOF と EOF の対応・シートの位置・索引・共有文字列の数。"""
    rep = _Lint()
    try:
        wb = XlsWorkbook(blob, stream=stream)
    except Exception as e:  # noqa: BLE001
        rep.errors.append(f"ブックとして読めません（{e}）")
        return rep
    data = wb.data
    offs = {}
    depth = 0
    for pos, t, b in records(data):
        offs[pos] = t
        if len(b) > (2080 if wb.biff == 5 else 8224):
            rep.errors.append(f"レコード 0x{t:04X} が長すぎます（{len(b)} バイト）")
        if t == 0x0809:
            depth += 1
        elif t == 0x000A:
            depth -= 1
            if depth < 0:
                rep.errors.append("EOF が BOF より多くあります")
    if depth != 0:
        rep.errors.append("BOF と EOF の数が合いません")
    for sh in wb.sheets:
        if sh.path.startswith("#") and offs.get(int(sh.path[1:])) != 0x0809:
            rep.errors.append(f"シート「{sh.name}」の位置（BOUNDSHEET）が、シートの始まりを指していません")
    for pos, t, b in records(data):
        if t == 0x020B:
            head = 12 if wb.biff == 5 else 16  # INDEX の、DBCELL の位置の並びの前の大きさ（BIFF5 は行が 2 バイトずつ）
            n = (len(b) - head) // 4
            for k in range(n):
                p = struct.unpack_from("<I", b, head + 4 * k)[0]
                if offs.get(p) != 0x00D7:
                    rep.errors.append("索引（INDEX）が DBCELL を指していません")
                    break
        if t == 0x00FF:
            ents = (len(b) - 2) // 8
            for k in range(ents):
                ib = struct.unpack_from("<I", b, 2 + 8 * k)[0]
                within = struct.unpack_from("<H", b, 6 + 8 * k)[0]
                if offs.get(ib - within) not in (0x00FC, 0x003C):
                    rep.errors.append("共有文字列の索引（EXTSST）がずれています")
                    break
    rep.errors += _dbcell_errors(data)
    # 共有文字列の数
    for pos, t, b in records(data):
        if t == 0x00FC:
            unique = struct.unpack_from("<I", b, 4)[0]
            if unique != len(wb.shared_strings):
                rep.errors.append("共有文字列の数が合いません")
            break
    return rep


def _dbcell_errors(data: bytes) -> list[str]:
    """DBCELL（行のまとまりの索引）の位置をたどって確かめる。最初の値は、DBCELL から最初の ROW までの距離。
    続く値は、最初のものは 2 番目の ROW の頭（最初の ROW の終わり）から、ほかは一つ前の値の位置からの距離で、行の最初のセルに着く。
    値の並びは、OpenOffice.org の資料（4.7.4）では ROW ごと（セルのない行は次の行のセルを指す）、MS-XLS（2.4.78）と Excel が
    保存したファイルでは、セルのある行ごと。どちらかの読み方でちょうど合えばよい（位置の控えなので、どちらでも Excel は読める）。"""
    recs = {pos: (t, b) for pos, t, b in records(data)}
    out = []
    for pos, (t, b) in recs.items():
        if t != 0x00D7 or len(b) < 4:
            continue
        n = (len(b) - 4) // 2
        if n == 0:
            continue  # 空のまとまり
        first = pos - struct.unpack_from("<I", b, 0)[0]
        if recs.get(first, (None,))[0] != 0x0208:
            out.append("行の索引（DBCELL）が、行のまとまりの最初の ROW を指していません")
            continue
        rows, at = [], first
        while recs.get(at, (None,))[0] == 0x0208:
            rows.append(struct.unpack_from("<H", recs[at][1], 0)[0])
            at += 4 + len(recs[at][1])
        firsts: dict[int, int] = {}
        p = at
        while p in recs and p != pos:
            tt, bb = recs[p]
            if tt in CELL_T and len(bb) >= 2:
                firsts.setdefault(struct.unpack_from("<H", bb, 0)[0], p)
            p += 4 + len(bb)
        offs = [struct.unpack_from("<H", b, 4 + 2 * k)[0] for k in range(n)]
        start = first + 4 + len(recs[first][1])

        def walk(want: list[int | None]) -> bool:
            q = start
            for k, off in enumerate(offs):
                q += off
                if k < len(want) and want[k] is not None and q != want[k]:
                    return False
            return True

        cell_rows = [r for r in rows if r in firsts]
        by_cells = walk([firsts[r] for r in cell_rows]) and n >= len(cell_rows)
        by_rows = n == len(rows) and walk([firsts.get(r) for r in rows])
        if not (by_cells or by_rows):
            out.append("行の索引（DBCELL）の、行の最初のセルの位置がずれています")
    return out


# ====================================================================== 窓とは別の読み方
def raw_texts(blob: bytes) -> dict[str, list[tuple[str, str]]]:
    """レコードを直接読み、すべての文字と数値を拾う（窓を作る処理とは別の読み方）。"""
    wb = XlsWorkbook(blob)
    if wb.biff == 5:
        from .biff5 import raw_texts5

        return raw_texts5(wb)
    out: dict[str, list[tuple[str, str]]] = {"セルの文字": [], "セルの数値": [], "図形の文字": [], "メモ": [],
                                             "チェックボックスなどの文字": [], "入力規則の説明": [], "ヘッダー・フッター": [],
                                             "ActiveX の部品の文字": []}
    from ..axfile import texts as ax_texts

    for s_ in wb.sheets:
        for c in s_.controls:
            if getattr(c, "ax", None) is not None:
                out["ActiveX の部品の文字"] += [(f"{s_.name}（ActiveX）", t) for t in ax_texts(c.ax)]
    sst = wb.shared_strings
    _glob, parts, _others = worksheet_parts(wb)
    for name, recs in parts:
        depth = 0
        last = None
        objs = {}
        cur_ot = None
        notes = []
        for i, (t, b) in enumerate(recs):
            if t == 0x0809:
                depth += 1
                continue
            if t == 0x000A:
                depth -= 1
                continue
            if depth != 1:
                continue
            if t == 0x00FD:
                r, c, _x, k = struct.unpack_from("<HHHI", b, 0)
                if k < len(sst) and sst[k]:
                    out["セルの文字"].append((f"{name}!{addr(r + 1, c + 1)}", sst[k]))
            elif t in (0x0204, 0x00D6):
                r, c = struct.unpack_from("<HH", b, 0)
                s = xlstring(b, 6)[0]
                if s:
                    out["セルの文字"].append((f"{name}!{addr(r + 1, c + 1)}", s))
            elif t == 0x0203:
                r, c = struct.unpack_from("<HH", b, 0)
                out["セルの数値"].append((f"{name}!{addr(r + 1, c + 1)}", repr(struct.unpack_from("<d", b, 6)[0])))
            elif t == 0x027E:
                from .book import _rk

                r, c, _x, v = struct.unpack_from("<HHHI", b, 0)
                out["セルの数値"].append((f"{name}!{addr(r + 1, c + 1)}", repr(_rk(v))))
            elif t == 0x00BD:
                from .book import _rk

                r, c0 = struct.unpack_from("<HH", b, 0)
                for k in range((len(b) - 6) // 6):
                    v = struct.unpack_from("<I", b, 6 + 6 * k)[0]
                    out["セルの数値"].append((f"{name}!{addr(r + 1, c0 + k + 1)}", repr(_rk(v))))
            elif t == 0x0205:
                r, c, _x, v, e = struct.unpack_from("<HHHBB", b, 0)
                out["セルの文字"].append((f"{name}!{addr(r + 1, c + 1)}", ERRORS.get(v, "#N/A") if e else ("TRUE" if v else "FALSE")))
            elif t == 0x0006:
                r, c = struct.unpack_from("<HH", b, 0)
                res = b[6:14]
                where = f"{name}!{addr(r + 1, c + 1)}"
                last = where
                if res[6:8] == b"\xff\xff":
                    if res[0] == 1:
                        out["セルの文字"].append((where, "TRUE" if res[2] else "FALSE"))
                    elif res[0] == 2:
                        out["セルの文字"].append((where, ERRORS.get(res[2], "#N/A")))
                else:
                    out["セルの数値"].append((where, repr(struct.unpack_from("<d", res, 0)[0])))
            elif t == 0x0207 and last is not None:
                s = xlstring(b, 0)[0]
                if s:
                    out["セルの文字"].append((last, s))
                last = None
            elif t == 0x005D:
                pos = 0
                while pos + 4 <= len(b):
                    ft, cb = struct.unpack_from("<HH", b, pos)
                    if ft == 0x15:
                        cur_ot, oid = struct.unpack_from("<HH", b, pos + 4)
                        objs[oid] = cur_ot
                    pos += 4 + cb
                    if ft in (0x00, 0x13):
                        break
            elif t == 0x01B6:
                cch = struct.unpack_from("<H", b, 10)[0]
                if cch and i + 1 < len(recs) and recs[i + 1][0] == 0x003C:
                    parts = []
                    k = i + 1
                    while k < len(recs) and recs[k][0] == 0x003C:
                        parts.append(recs[k][1])
                        k += 1
                    seg = Seg(parts)
                    high = bool(seg.u8() & 1)
                    try:
                        text = seg.chars(cch, high)
                    except Exception:  # noqa: BLE001
                        text = ""
                    kind = {0x19: "メモ", 0x0B: "チェックボックスなどの文字", 0x0C: "チェックボックスなどの文字",
                            0x07: "チェックボックスなどの文字", 0x0E: "チェックボックスなどの文字",
                            0x13: "チェックボックスなどの文字"}.get(cur_ot, "図形の文字")  # 0x13 はグループボックスの見出し
                    if text.strip():
                        notes.append((kind, text))
            elif t == 0x01BE:
                pos = 4
                for key in ("promptTitle", "errorTitle", "prompt", "error"):
                    s, pos = xlstring(b, pos)
                    if s not in ("", "\x00"):
                        out["入力規則の説明"].append((f"{name}（入力規則）", s))
            elif t in (0x0014, 0x0015) and len(b) >= 3:
                s = xlstring(b, 0)[0]
                if s.strip():
                    out["ヘッダー・フッター"].append((f"{name}（ヘッダー）", s))
        for kind, text in notes:
            out[kind].append((f"{name}（{kind}）", text))
    return out
