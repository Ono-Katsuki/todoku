""".xlsb への書き込み（読み戻しと証明つき）。

- 値のあるセル：レコードの大きさを変えずに上書きする。字数を変える・共有文字列のセルを変える書き込みはしない
  （.xlsx で保存し直せば、ふつうに書ける）
- 空のセル：シートの部品にレコードを足し、位置の索引（binaryIndex）を合わせて直す（grow.py）。Excel で開く確認はできないので、warnings で伝える

書くもの：数（BrtCellReal・BrtCellRk。Rk は、その形でちょうど表せる数だけ）・真偽（BrtCellBool）・同じ字数の字（BrtCellSt）。
書いたセルに連なる数式の保存値（BrtFmlaNum の 8 バイト）も、同じ大きさのまま計算し直した値にする。字の結果の数式が変わるときは書かない。
"""

from __future__ import annotations

import struct

from ..formula.values import ExcelError
from ..xls.book import _rk
from ..xlsx.edit import resolve_target
from ..zipraw import RawZip
from .book import C_BLANK, C_BOOL, C_ISST, C_REAL, C_RK, C_RSTR, C_ST, F_NUM, XlsbWorkbook

CANT_GROW = ("空のセルに値を入れると、レコードの大きさが変わり、シートの部品の中の位置を指す索引（binaryIndex）がずれます。"
             "とどくは大きさの変わる書き込みをしません。.xlsx で保存し直してもらえば書けます")


class XlsbWriteError(ValueError):
    pass


def _number(v) -> float | None:
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        return None


def _encode(rid: int, old_len: int, v, where: str) -> bytes:
    if rid == C_BLANK:
        raise XlsbWriteError(f"{where}：{CANT_GROW}")
    if rid in (C_ISST, C_RSTR):
        raise XlsbWriteError(f"{where}：共有の字のセルは、同じ字を使うほかのセルや字の並びを変えるので、上書きできません。"
                             ".xlsx で保存し直してもらえば書けます")
    if rid == C_REAL or rid == C_RK:
        x = _number(v)
        if x is None:
            raise XlsbWriteError(f"{where}：数のセルです。数でない値（{v!r}）を書くとレコードの形が変わるので書けません")
        if rid == C_REAL:
            return struct.pack("<d", x)
        if x != int(x) or not -(1 << 29) <= x < (1 << 29):
            raise XlsbWriteError(f"{where}：このセルの数の形（Rk）では {v!r} をちょうど表せません（大きさが変わるので書けません）")
        raw = (int(x) << 2 | 2) & 0xFFFFFFFF
        if _rk(raw) != x:
            raise XlsbWriteError(f"{where}：このセルの数の形（Rk）では {v!r} をちょうど表せません")
        return struct.pack("<I", raw)
    if rid == C_BOOL:
        if isinstance(v, bool):
            return bytes([1 if v else 0])
        raise XlsbWriteError(f"{where}：真偽のセルです。true か false を書いてください")
    if rid == C_ST:
        s = str(v)
        enc = s.encode("utf-16-le")
        if 4 + len(enc) != old_len:
            raise XlsbWriteError(f"{where}：字のセルは、同じ字数（UTF-16 で {(old_len - 4) // 2} 字）でないと書けません（大きさが変わります）。"
                                 ".xlsx で保存し直してもらえば書けます")
        return struct.pack("<I", len(s)) + enc
    raise XlsbWriteError(f"{where}：このセルの形（レコード {rid}）には書けません")


def _check_rules(wb, sh, r: int, c: int, v, where: str, force: bool) -> None:
    if sh.protected and wb.style(sh.style_of(r, c)).locked and not force:
        raise XlsbWriteError(f"{where}：保護されたシートのロックされたセルです（Excel でも書けません）")
    for dv in sh.validations:
        if dv.covers(r, c) and dv.type == "list" and (dv.formula1 or "").startswith('"') and not force:
            opts = dv.formula1.strip('"').split(",")
            if str(v) not in opts:
                raise XlsbWriteError(f"{where}：選択肢（{'／'.join(opts)}）にない値です")


def fill(blob: bytes, name: str, values: dict, force: bool = False, notes: dict | None = None):
    """(書いたブック, 書いたセル {(シート, 行, 列): 値}, 計算し直した数式 {…}, 書き換えた範囲 {部品: [(始め, 終わり)]})。
    notes を渡すと、伝えること（replaced_text・written_formatting・warnings）を入れる。"""
    from ..formula.parser import addr
    from ..xlsx.cellcheck import written_cells_format
    from ..xlsx.edit import EditError, _both_boxes_marked, check_box_in_text, explicit_overwrite, is_printed_text, replaced_printed

    wb = XlsbWorkbook(blob, name)
    parts: dict[str, bytearray] = {}
    spans: dict[str, list[tuple[int, int]]] = {}
    written = {}
    grows: dict[str, list[tuple[int, int, int, object]]] = {}  # シートの部品 → (行, 列, 書式, 値)。0 始まり
    claimed: dict[tuple[str, int, int], str] = {}
    keys: dict[tuple[str, int, int], str] = {}
    notes = notes if notes is not None else {}
    for key, v in values.items():
        try:
            sheet, r, c = resolve_target(wb, key)
        except Exception as e:  # noqa: BLE001
            raise XlsbWriteError(f"{key}：セルが分かりません（{e}）") from None
        if (sheet, r, c) in claimed:
            # 同じセルを2つの指示で書かない（B9 と b9）：後の指示が前の指示を黙って消す
            raise XlsbWriteError(f"「{claimed[(sheet, r, c)]}」と「{key}」は、同じセル {sheet}!{addr(r, c)} を指しています。"
                                 "後の指示が前の指示を消してしまうので、書いていません。1つのセルには1つの指示で書いてください")
        claimed[(sheet, r, c)] = str(key)
        keys[(sheet, r, c)] = str(key)
        sh = wb.sheet(sheet)
        where = f"{sheet}!{key.rsplit('!', 1)[-1]}"
        cell = sh.cells.get((r, c))
        if cell is not None and (cell.has_formula or cell.array_master):
            raise XlsbWriteError(f"{where}：数式のセルには書きません")
        # 印刷された字のあるセル：Excel の書き手と同じ決まり（字を消す書き込みは、明示がなければ証明で示す）
        try:
            v, overwrite_text = explicit_overwrite(key, v)
            if is_printed_text(cell) and not overwrite_text:
                if isinstance(v, bool):
                    v, note = check_box_in_text(cell.value, v, f"{sheet}!{addr(r, c)}", anchors=False)
                    notes.setdefault("warnings", []).extend([note] + _both_boxes_marked(cell.value, v, f"{sheet}!{addr(r, c)}"))
                else:
                    gone = replaced_printed(wb, sh, sheet, r, c, cell, v, anchors=False)
                    if gone:
                        notes.setdefault("replaced_text", []).append(gone)
        except EditError as e:
            raise XlsbWriteError(str(e)) from None
        rec = wb.value_records.get((sheet, r, c))
        _check_rules(wb, sh, r, c, v, where, force)
        if rec is None or rec[1] == C_BLANK:
            if sh.state.startswith("non-worksheet"):
                raise XlsbWriteError(f"{where}：シートでない部品には書きません")
            shown, _notes = _interpret(wb, sheet, r, c, v, where)
            grows.setdefault(sh.path, []).append((r - 1, c - 1, sh.style_of(r, c), shown))
            written[(sheet, r, c)] = shown
            continue
        part, rid, pos, ln = rec
        new = _encode(rid, ln, v, where)
        buf = parts.setdefault(part, bytearray(wb.zip.read(part)))
        buf[pos:pos + ln] = new
        spans.setdefault(part, []).append((pos, pos + ln))
        written[(sheet, r, c)] = v
    # 書いたセルの書式：見えない字（字と塗りが同じ色など）は書かない。取り消し線などは伝える（Excel の書き手と同じ）
    try:
        notes["written_formatting"] = written_cells_format(wb, written, keys, force)
    except ValueError as e:
        raise XlsbWriteError(str(e)) from None
    recalculated = _recalc(wb, blob, name, parts, spans, grows)
    return _assemble(blob, wb, parts, grows), written, recalculated, spans


def _interpret(wb, sheet: str, r: int, c: int, v, where: str):
    """Excel でそのセルに入力したときと同じ規則で、数・日付・真偽・字にする（xlsx と同じ）。"""
    from ..xlsx.edit import EditError, interpret_input

    try:
        out, notes = interpret_input(wb, sheet, r, c, v)
    except EditError as e:
        raise XlsbWriteError(str(e)) from None
    if out is None or out == "":
        raise XlsbWriteError(f"{where}：空の値は書けません")
    if isinstance(out, float) and (out != out or out in (float("inf"), float("-inf"))):
        raise XlsbWriteError(f"{where}：値 {v!r} は数ではありません")
    return out, notes


def index_part(wb, sheet_part: str) -> str | None:
    """シートの関係にある、位置の索引の部品の名前。"""
    for typ, target in wb._rels(sheet_part).values():
        if typ.lower().endswith("binaryindex") and wb.zip.has(target):
            return target
    return None


def _assemble(blob: bytes, wb, parts: dict, grows: dict) -> bytes:
    from .grow import GrowError, plan, rewrite_index, splice

    z = RawZip(blob)
    done = set()
    try:
        for part, items in grows.items():
            raw = bytes(parts[part]) if part in parts else wb.zip.read(part)
            edits = plan(raw, items)
            new = splice(raw, edits)
            z.replace(part, new)
            done.add(part)
            ip = index_part(wb, part)
            if ip:
                z.replace(ip, rewrite_index(wb.zip.read(ip), edits, new, [(r, c) for r, c, _x, _v in items]))
    except (GrowError, struct.error, KeyError) as e:
        raise XlsbWriteError(f"シートのレコードを足せません（{e}）") from None
    for part, buf in parts.items():
        if part not in done:
            z.replace(part, bytes(buf))
    return z.build()


def _recalc(wb, blob: bytes, name: str, parts: dict, spans: dict, grows: dict | None = None) -> dict:
    """書いたあとのブックを計算し直し、書く前は保存値と合っていた数式で合わなくなったものの、保存値を同じ大きさで直す。"""
    from ..xlsx.calc import verify

    def rebuilt() -> XlsbWorkbook:
        return XlsbWorkbook(_assemble(blob, wb, parts, grows or {}), name)

    before = {(r.sheet, r.address) for r in verify(wb).results if r.status == "mismatch"}
    out = {}
    for _round in range(8):  # 数式が読む数式：直した値で、もう一度確かめる
        after = rebuilt()
        pending = [r for r in verify(after).results if r.status == "mismatch" and (r.sheet, r.address) not in before]
        if not pending:
            return out
        for r in pending:
            from ..formula.parser import parse_cell

            row, col = parse_cell(r.address)
            rec = wb.value_records.get((r.sheet, row, col))
            if rec is None or rec[1] != F_NUM or isinstance(r.actual, (str, bool, ExcelError)) or r.actual is None:
                raise XlsbWriteError(f"{r.sheet}!{r.address}：書いた値で、この数式の結果が数でなくなります（字の結果は大きさが変わるので、"
                                     "保存値を直せません）。.xlsx で保存し直してもらえば書けます")
            part, _rid, pos, ln = rec
            buf = parts.setdefault(part, bytearray(wb.zip.read(part)))
            buf[pos:pos + ln] = struct.pack("<d", float(r.actual))
            spans.setdefault(part, []).append((pos, pos + ln))
            out[f"{r.sheet}!{r.address}"] = float(r.actual)
    raise XlsbWriteError("数式を計算し直しても、保存値が定まりません")


def prove(old: bytes, new: bytes, spans: dict, grown: dict | None = None, indexes: dict | None = None) -> list[str]:
    """ZIP の部品のうち、書いたシートのほかがバイト一致か（圧縮したバイトまで）、書いたシートで変わったのが値のバイトだけか。問題の一覧を返す。
    grown：レコードを足した部品 → {(行, 列)}（0 始まり）。indexes：その部品の索引の部品の名前。読み直して、順と索引を確かめる（grow.py）。"""
    from .grow import GrowError, prove_index, prove_sheet

    a, b = RawZip(old), RawZip(new)
    problems = []
    grown = grown or {}
    indexes = indexes or {}
    index_parts = set(indexes.values())
    if a.names() != b.names():
        problems.append("部品の一覧が変わっています")
    by = {e.name: e for e in b.entries}
    for ea in a.entries:
        eb = by.get(ea.name)
        if eb is None:
            problems.append(f"{ea.name} がなくなっています")
        elif ea.name in grown:
            try:
                problems += [f"{ea.name}：{x}" for x in prove_sheet(a.read(ea.name), b.read(ea.name), grown[ea.name], spans.get(ea.name, []))]
                if ea.name in indexes:
                    ip = indexes[ea.name]
                    problems += [f"{ip}：{x}" for x in prove_index(a.read(ip), b.read(ip), b.read(ea.name), a.read(ea.name), grown[ea.name])]
            except (GrowError, struct.error, ValueError, KeyError) as e:
                problems.append(f"{ea.name}：読み直せません（{e}）")
        elif ea.name in index_parts:
            pass  # 上で確かめる
        elif ea.name in spans:
            x, y = a.read(ea.name), b.read(ea.name)
            if len(x) != len(y):
                problems.append(f"{ea.name} の大きさが変わっています")
                continue
            allowed = sorted(spans[ea.name])
            i = 0
            for k in (k for k in range(len(x)) if x[k] != y[k]):
                while i < len(allowed) and allowed[i][1] <= k:
                    i += 1
                if i >= len(allowed) or k < allowed[i][0]:
                    problems.append(f"{ea.name}：値のほかのバイト（位置 {k}）が変わっています")
                    break
        elif ea.data != eb.data or ea.local != eb.local:
            problems.append(f"{ea.name} のバイトが変わっています")
    return problems


def api_fill(blob: bytes, name: str, values: dict, force: bool = False) -> dict:
    from ..formula.parser import addr

    notes: dict = {}
    new, written, recalculated, spans = fill(blob, name, values, force, notes)
    old_wb = XlsbWorkbook(blob, name)
    grown: dict[str, set] = {}
    for (s, r, c) in written:
        rec = old_wb.value_records.get((s, r, c))
        if rec is None or rec[1] == C_BLANK:
            grown.setdefault(old_wb.sheet(s).path, set()).add((r - 1, c - 1))
    indexes = {p: index_part(old_wb, p) for p in grown if index_part(old_wb, p)}
    problems = prove(blob, new, spans, grown, indexes)
    problems += _same_cells(old_wb, XlsbWorkbook(new, name), written, recalculated)
    readback = {f"{s}!{addr(r, c)}": v for (s, r, c), v in written.items()}
    if problems:
        head = "証明で、想定外の変更が見つかりました"
    elif grown:
        n = sum(len(v) for v in grown.values())
        head = (f"{len(written)} セルを書きました（うち空のセル {n} 個はレコードを足し、位置の索引も直しました）。ほかの部品は圧縮したバイトまで同じで、"
                "元のレコードは順のまま残り、索引の位置はすべてセルの始まりを指しています")
    else:
        head = "%d セルをレコードの大きさを変えずに上書きしました。ほかの部品は圧縮したバイトまで同じで、シートの部品で変わったのは値のバイトだけです" % len(written)
    from ..xlsx.cellcheck import formatting_notes
    from .grow import UNVERIFIED

    result = {"proof_ok": not problems, "proof": head, "warnings": ([UNVERIFIED] if grown else []) + list(notes.get("warnings") or []),
              "unexpected": problems[:20], "structure_errors": [], "recalculated": recalculated, "written_as": {k: str(v) for k, v in readback.items()}}
    if notes.get("replaced_text"):
        result["replaced_text"] = notes["replaced_text"]  # 書き込みの終わり（api）で、証明の想定外の変更に入れる
    if notes.get("written_formatting"):
        result["written_formatting"] = notes["written_formatting"]
        result["warnings"] += formatting_notes(notes["written_formatting"])
    return {"blob": new, "readback": readback, "labels": {}, "result": result}


def _same_cells(old, new, written: dict, recalculated: dict) -> list[str]:
    """読み直したブックで、書いたセルと計算し直した数式のほかのセルが、値も書式も元のままか。書いたセルが書いた値か。"""
    problems = []
    for sh in old.sheets:
        if sh.state.startswith("non-worksheet"):
            continue
        ns = new.sheet(sh.name)
        for (r, c), cell in sh.cells.items():
            key = (sh.name, r, c)
            n = ns.cells.get((r, c))
            if key in written:
                continue
            if n is None or n.style != cell.style:
                problems.append(f"{sh.name} の {r} 行 {c} 列のセルが変わっています")
            elif n.value != cell.value and n.value == n.value and f"{sh.name}!{_addr(r, c)}" not in recalculated:
                problems.append(f"{sh.name} の {r} 行 {c} 列の値が変わっています")
        extra = set(ns.cells) - set(sh.cells) - {(r, c) for s, r, c in written if s == sh.name}
        if extra:
            problems.append(f"{sh.name} に、書いていないセルが増えています")
    for (s, r, c), v in written.items():
        got = new.sheet(s).cells.get((r, c))
        if got is None or got.value != v:
            problems.append(f"{s} の {r} 行 {c} 列を読み直すと、書いた値と違います")
    return problems


def _addr(r: int, c: int) -> str:
    from ..formula.parser import addr

    return addr(r, c)


def read_cells(blob: bytes, name: str, keys) -> dict:
    wb = XlsbWorkbook(blob, name)
    out = {}
    for key in keys:
        try:
            s, r, c = resolve_target(wb, key)
        except Exception:  # noqa: BLE001
            continue
        cell = wb.sheet(s).cells.get((r, c))
        out[key] = cell.value if cell is not None else None
    return out
