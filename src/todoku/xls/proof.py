"""古いExcel（.xls）への書き込みの証明と構造の検査、窓とは別の読み方による文字の取り出し。

証明は、書き込み処理の内部情報を使わず、前後の2つのファイルを読み直して比べる。
- 複合ファイルの中のほかのストリームは、バイト一致
- ブックのストリームのレコードは、書き込んだセル・計算結果・チェックボックスの状態・位置を持つレコード
  （BOUNDSHEET・INDEX・DBCELL・EXTSST・SST の追加・ROW の列の範囲）のほかは、バイト一致
"""

from __future__ import annotations

import hashlib
import struct

from ..cfb import CFB
from ..formula.parser import addr
from ..xlsx.calc import values_equal
from ..xlsx.proof import Proof
from .book import ERRORS, Seg, XlsWorkbook, records, xlstring

CELL_T = {0x0203, 0x027E, 0x00FD, 0x0204, 0x00D6, 0x0201, 0x0205, 0x0006, 0x00BD, 0x00BE}
POINTER_T = {0x0085, 0x020B, 0x00D7, 0x00FF, 0x01C1}  # 位置を持つレコード（と、再計算の印）


def _is_globals(bof: bytes) -> bool:
    return len(bof) >= 4 and struct.unpack_from("<H", bof, 2)[0] == 0x0005


def split(data: bytes):
    """ブックのストリームを (ブック全体のレコード, [シートのレコード]) に分ける。"""
    out_glob: list = []
    out_sheets: list[list] = []
    depth = 0
    target = None
    for _pos, t, b in records(data):
        if t == 0x0809 and depth == 0:
            target = out_glob if _is_globals(b) else []
            if target is not out_glob:
                out_sheets.append(target)
        if t == 0x0809:
            depth += 1
        if target is not None:
            target.append((t, b))
        if t == 0x000A:
            depth -= 1
            if depth == 0:
                target = None
    return out_glob, out_sheets


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
    a_cfb, b_cfb = CFB(original), CFB(new)
    unexpected: list[str] = []
    sa, sb = a_cfb.streams(), b_cfb.streams()
    if set(sa) != set(sb):
        unexpected.append(f"複合ファイルのストリームの構成が変わりました: {sorted(set(sa) ^ set(sb))}")
    a = XlsWorkbook(original)
    b = XlsWorkbook(new)
    identical = sum(1 for n in sa if n in sb and a_cfb.read(n) == b_cfb.read(n))
    changed = [n for n in sa if n in sb and a_cfb.read(n) != b_cfb.read(n)]
    for n in changed:
        if n != a.stream_name:
            unexpected.append(f"ストリーム {n} が変わりました（ブック以外は変えないはずです）")
    ga, sha = split(a.data)
    gb, shb = split(b.data)
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
    names = [s.name for s in a.sheets if not s.state.startswith("non-")]
    for name, ra, rb in zip(names, sha, shb):
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
        # セル以外のレコード（チェックボックスの状態・位置のレコード・ROW の列の範囲のほか）は同じであること
        def others(rs):
            out = []
            for t, body in rs:
                if t in CELL_T or t in (0x0207, 0x00D7, 0x020B):
                    continue
                if t == 0x0208:
                    out.append((t, body[:2] + body[6:]))  # 列の範囲（colMic・colMac）は変わってよい
                elif t == 0x005D:
                    out.append((t, _obj_without_checked(body)))
                else:
                    out.append((t, body))
            return out

        oa, ob = others(ra), others(rb)
        if [x for x in oa if x[0] != 0x0208] != [x for x in ob if x[0] != 0x0208]:
            unexpected.append(f"{name}: セル以外のレコードが変わりました")
        ctl_a = {c.shape_id: c for c in sh_a.controls}
        for cb_ in sh_b.controls:
            ca_ = ctl_a.get(cb_.shape_id)
            if ca_ is not None and ca_.checked != cb_.checked:
                control_changes.append({"sheet": name, "control": ca_.name, "kind": "control", "link": ca_.link,
                                        "before": ca_.checked or "Unchecked", "after": cb_.checked or "Unchecked"})
                if allowed_controls is not None and (name, ca_.name) not in allowed_controls:
                    unexpected.append(f"{name}: 指示していないチェックボックス「{ca_.name}」が変わりました")
    lint_errors = lint(new).errors
    before = set(lint(original).errors) if lint_errors else set()
    ra_all = [(t, body) for _p, t, body in records(a.data)]
    rb_all = [(t, body) for _p, t, body in records(b.data)]
    from collections import Counter

    common = Counter(ra_all) & Counter(rb_all)
    return Proof(
        container="cfb", records_total=len(ra_all), records_identical=sum(common.values()),
        original_sha256=hashlib.sha256(original).hexdigest(), new_sha256=hashlib.sha256(new).hexdigest(),
        entries_total=len(sa), entries_identical=identical, entries_changed=changed,
        cell_changes=cell_changes, control_changes=control_changes, unexpected=unexpected,
        xml_bytes_changed=0, xml_bytes_total=0,
        lint_errors=[e for e in lint_errors if e not in before], lint_parts=1)


def _obj_without_checked(b: bytes) -> bytes:
    out = bytearray(b)
    pos = 0
    while pos + 4 <= len(out):
        ft, cb = struct.unpack_from("<HH", out, pos)
        if ft == 0x12 and cb >= 2:
            struct.pack_into("<H", out, pos + 4, 0)
        if ft in (0x00, 0x13):
            break
        pos += 4 + cb
    return bytes(out)


# ====================================================================== 構造の検査
class _Lint:
    def __init__(self):
        self.errors: list[str] = []
        self.checked_parts = 1

    @property
    def ok(self) -> bool:
        return not self.errors


def lint(blob: bytes) -> _Lint:
    """Excelが修復なしに開くための条件：レコードの長さ・BOF と EOF の対応・シートの位置・索引・共有文字列の数。"""
    rep = _Lint()
    try:
        wb = XlsWorkbook(blob)
    except Exception as e:  # noqa: BLE001
        rep.errors.append(f"ブックとして読めません（{e}）")
        return rep
    data = wb.data
    offs = {}
    depth = 0
    for pos, t, b in records(data):
        offs[pos] = t
        if len(b) > 8224:
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
            n = (len(b) - 16) // 4
            for k in range(n):
                p = struct.unpack_from("<I", b, 16 + 4 * k)[0]
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
    # 共有文字列の数
    for pos, t, b in records(data):
        if t == 0x00FC:
            unique = struct.unpack_from("<I", b, 4)[0]
            if unique != len(wb.shared_strings):
                rep.errors.append("共有文字列の数が合いません")
            break
    return rep


# ====================================================================== 窓とは別の読み方
def raw_texts(blob: bytes) -> dict[str, list[tuple[str, str]]]:
    """レコードを直接読み、すべての文字と数値を拾う（窓を作る処理とは別の読み方）。"""
    wb = XlsWorkbook(blob)
    out: dict[str, list[tuple[str, str]]] = {"セルの文字": [], "セルの数値": [], "図形の文字": [], "メモ": [],
                                             "チェックボックスなどの文字": [], "入力規則の説明": [], "ヘッダー・フッター": []}
    sst = wb.shared_strings
    _glob, sheets = split(wb.data)
    names = [s.name for s in wb.sheets if not s.state.startswith("non-")]
    for name, recs in zip(names, sheets):
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
                            0x07: "チェックボックスなどの文字", 0x0E: "チェックボックスなどの文字"}.get(cur_ot, "図形の文字")
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
