""".xlsb への書き込み：すでに値のあるセルを、レコードの大きさを変えずに上書きする（読み戻しと証明つき）。

レコードの大きさを変える書き込み（空のセルに値を入れる・字数を変える・共有文字列のセルを変える）はしない：
Excel が作るブックには、シートの部品の中の位置を指す索引（xl/worksheets/binaryIndexN.bin の BrtIndexRowBlock など）があり、
大きさを変えるとずれる。索引を作り直したファイルが Excel で開けることを確かめる手段がないので、確かめられない書き込みをしない。
.xlsx で保存し直せば、ふつうに書ける。

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


def fill(blob: bytes, name: str, values: dict, force: bool = False):
    """(書いたブック, 書いたセル {(シート, 行, 列): 値}, 計算し直した数式 {…}, 書き換えた範囲 {部品: [(始め, 終わり)]})。"""
    wb = XlsbWorkbook(blob, name)
    parts: dict[str, bytearray] = {}
    spans: dict[str, list[tuple[int, int]]] = {}
    written = {}
    for key, v in values.items():
        try:
            sheet, r, c = resolve_target(wb, key)
        except Exception as e:  # noqa: BLE001
            raise XlsbWriteError(f"{key}：セルが分かりません（{e}）") from None
        sh = wb.sheet(sheet)
        where = f"{sheet}!{key.rsplit('!', 1)[-1]}"
        cell = sh.cells.get((r, c))
        if cell is not None and (cell.has_formula or cell.array_master):
            raise XlsbWriteError(f"{where}：数式のセルには書きません")
        rec = wb.value_records.get((sheet, r, c))
        if rec is None:
            raise XlsbWriteError(f"{where}：{CANT_GROW}")
        _check_rules(wb, sh, r, c, v, where, force)
        part, rid, pos, ln = rec
        new = _encode(rid, ln, v, where)
        buf = parts.setdefault(part, bytearray(wb.zip.read(part)))
        buf[pos:pos + ln] = new
        spans.setdefault(part, []).append((pos, pos + ln))
        written[(sheet, r, c)] = v
    recalculated = _recalc(wb, blob, name, parts, spans)
    z = RawZip(blob)
    for part, buf in parts.items():
        z.replace(part, bytes(buf))
    return z.build(), written, recalculated, spans


def _recalc(wb, blob: bytes, name: str, parts: dict, spans: dict) -> dict:
    """書いたあとのブックを計算し直し、書く前は保存値と合っていた数式で合わなくなったものの、保存値を同じ大きさで直す。"""
    from ..xlsx.calc import verify

    def rebuilt() -> XlsbWorkbook:
        z = RawZip(blob)
        for part, buf in parts.items():
            z.replace(part, bytes(buf))
        return XlsbWorkbook(z.build(), name)

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


def prove(old: bytes, new: bytes, spans: dict) -> list[str]:
    """ZIP の部品のうち、書いたシートのほかがバイト一致か（圧縮したバイトまで）、書いたシートで変わったのが値のバイトだけか。問題の一覧を返す。"""
    a, b = RawZip(old), RawZip(new)
    problems = []
    if a.names() != b.names():
        problems.append("部品の一覧が変わっています")
    by = {e.name: e for e in b.entries}
    for ea in a.entries:
        eb = by.get(ea.name)
        if eb is None:
            problems.append(f"{ea.name} がなくなっています")
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

    new, written, recalculated, spans = fill(blob, name, values, force)
    problems = prove(blob, new, spans)
    readback = {f"{s}!{addr(r, c)}": v for (s, r, c), v in written.items()}
    head = (f"{len(written)} セルをレコードの大きさを変えずに上書きしました。ほかの部品は圧縮したバイトまで同じで、シートの部品で変わったのは値のバイトだけです"
            if not problems else "証明で、想定外の変更が見つかりました")
    return {"blob": new, "readback": readback, "labels": {},
            "result": {"proof_ok": not problems, "proof": head, "warnings": [], "unexpected": problems[:20], "structure_errors": [],
                       "recalculated": recalculated, "written_as": {k: str(v) for k, v in readback.items()}}}


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
