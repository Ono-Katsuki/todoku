"""古い形式（.xls・.doc）の ActiveX の部品：置き場所を読み、値を書き、証明する。

- .xls（MS-XLS 2.4.181 Obj・2.5.150 FtPictFmla・2.5.151 FtPioGrbit・2.5.202 PictFmlaEmbedInfo・2.5.203 PictFmlaKey）：
  絵の OBJ（ot=8）で pictFlags.fCtl=1。strClass に ProgID（Forms.TextBox.1 など）。fPrstm=1 なら "Ctls" ストリームの
  lPosInCtlStm から cbBufInCtlStm バイト、0 なら "MBD" と8桁の16進の名前の入れ物（contents。フレームは f と o）。
  部品の大きさが変わるときは、Ctls の後ろの部品の lPosInCtlStm と、その部品の cbBufInCtlStm を OBJ の中で直す
  （4バイトずつの、長さの変わらない書き換え）。つながったセル・選択肢の範囲は fmlaLinkedCell・fmlaListFillRange
- .doc（MS-DOC 2.1.4 ObjectPool・2.9.165 ODT・2.6.1 sprmCPicLocation）：CONTROL のフィールドの区切りの字の
  sprmCPicLocation の値の前に「_」を付けた、ObjectPool の中の入れ物。ODT の fOCX が立ち、fStream なら
  "\\003OCXDATA"、そうでなければ "contents" が部品の保存の形。部品の名前は "\\003OCXNAME"（UTF-16）
"""

from __future__ import annotations

import struct

from . import activex as ax
from .axfile import (CAPTION_KINDS, CHECK_KINDS, NUM_KINDS, AxRef, _cell_value_for, _read_control, _state_text,
                     _struct_kind, to_str)


# ====================================================================== .xls
def xls_pict(raw: bytes) -> dict | None:
    """絵の OBJ の FtPioGrbit と FtPictFmla を読む。ActiveX の部品でなければ None。
    lPosInCtlStm・cbBufInCtlStm の、OBJ の中身（続く CONTINUE もつないだもの）の中の位置も返す。"""
    pos = 0
    flags = None
    out: dict = {}
    while pos + 4 <= len(raw):
        ft, cb = struct.unpack_from("<HH", raw, pos)
        body = pos + 4
        if ft == 0x00:
            break
        if ft == 0x08:  # FtPioGrbit
            flags = struct.unpack_from("<H", raw, body)[0]
            out["flags_at"] = body
        elif ft == 0x09:  # FtPictFmla
            fl = flags or 0
            p = body
            cb_fmla = struct.unpack_from("<H", raw, p)[0]
            p += 2
            end_fmla = p + cb_fmla
            tbl = False
            if cb_fmla:
                cce = struct.unpack_from("<H", raw, p)[0] & 0x7FFF
                rgce = raw[p + 6:p + 6 + cce]
                q = p + 6 + cce
                tbl = rgce[:1] == b"\x02"  # PtgTbl
                if tbl and q + 3 <= end_fmla and raw[q] == 0x03:  # PictFmlaEmbedInfo（ttb = 3）
                    n = raw[q + 1]
                    if n:
                        high = raw[q + 3] & 1
                        s = raw[q + 4:q + 4 + n * (2 if high else 1)]
                        out["progid"] = s.decode("utf-16-le" if high else "latin-1", errors="replace")
            p = end_fmla
            if tbl:
                out["pos_at"] = p
                out["pos"] = struct.unpack_from("<I", raw, p)[0]
                p += 4
                if fl & 0x20:  # fPrstm：Ctls の中の大きさ
                    out["size_at"] = p
                    out["size"] = struct.unpack_from("<I", raw, p)[0]
                    p += 4
            if fl & 0x10 and p + 4 <= body + cb:  # fCtl：PictFmlaKey（cbKey・keyBuf・fmlaLinkedCell・fmlaListFillRange）
                cb_key = struct.unpack_from("<I", raw, p)[0]
                p += 4 + cb_key
                for key in ("link", "fill"):
                    if p + 2 > body + cb:
                        break
                    n = struct.unpack_from("<H", raw, p)[0]
                    if n:
                        cce = struct.unpack_from("<H", raw, p + 2)[0] & 0x7FFF
                        out[key] = bytes(raw[p + 8:p + 8 + cce])
                    p += 2 + n
        pos = body + cb
    if flags is None or not flags & 0x10:
        return None
    out["flags"] = flags
    out["prstm"] = bool(flags & 0x20)
    return out


def xls_control(wb, sh, ob: dict, anchor):
    """.xls の ActiveX の部品（xlsx.workbook.Control）。ActiveX でなければ None。"""
    from .xlsx.workbook import Control, _link_cell

    info = xls_pict(ob.get("raw") or b"")
    if info is None:
        return None
    name = (anchor or {}).get("name") or f"Control {ob['id']}"
    ref = AxRef("xls", name, progid=info.get("progid", ""))
    ref.kind = ax.kind_of(progid=ref.progid)
    ref.xls = info
    cf = wb.cfb
    try:
        if info["prstm"]:
            ref.persistence = "ctls"
            ref.part = ref.data_part = "Ctls"
            data = cf.read("Ctls") if cf.has("Ctls") else b""
            seg = data[info["pos"]:info["pos"] + info.get("size", 0)]
            clsid = next((g for g, k in ax.CLSIDS.items() if k == ref.kind), None)
            prefix = seg[:16] if clsid and seg[:16] == ax.guid_bytes(clsid) else b""
            ref.start = info["pos"] + len(prefix)
            ref.length = info.get("size", 0) - len(prefix)
            ref.prefix = len(prefix)
            _read_control(ref, data)
        else:
            from .cfb import SubStorage
            from .cfbwrite import storage_clsid

            ref.persistence = "storage"
            ref.part = f"MBD{info.get('pos', 0):08X}"
            sub = SubStorage(cf, ref.part)
            names = {n.lower(): n for n in sub.streams()}
            cls = storage_clsid(cf, ref.part)
            if cls:
                ref.clsid = cls
                ref.kind = ax.kind_of(cls) or ref.kind
            if "f" in names and "contents" not in names:
                ref.kind = "Frame"
                ref.frame = ax.parse_frame(sub.read(names["f"]), sub.read(names["o"]) if "o" in names else b"")
            elif "contents" in names:
                ref.stream = names["contents"]
                ref.data_part = f"{ref.part}/{ref.stream}"
                _read_control(ref, sub.read(ref.stream))
            else:
                ref.error = "部品の中身（contents）がありません"
    except (ax.AxError, KeyError, struct.error) as e:
        ref.error = str(e) or type(e).__name__
    pos = (anchor["r1"], anchor["c1"]) if anchor else None
    ctl = Control(kind="ActiveX " + (ref.kind or "Control"), name=name, shape_id=str(ob["id"]), anchor=pos,
                  anchor_to=(anchor["r2"], anchor["c2"]) if anchor else None,
                  from_off=(anchor or {}).get("off1", 0.0), to_off=(anchor or {}).get("off2", 0.0))
    for key, attr in (("link", "link"), ("fill", "fmla_range")):
        if info.get(key):
            try:
                setattr(ctl, attr, wb.decoder.decode(info[key]).replace("$", ""))
            except Exception:  # noqa: BLE001
                setattr(ctl, attr, "#REF!")
    if ctl.link:
        ctl.link_cell = _link_cell(ctl.link, sh.name)
    ctl.ax = ref
    fill_ctl(ctl, ref)
    return ctl


def fill_ctl(ctl, ref: AxRef) -> None:
    """読んだ部品の見出し・状態を、Excel の部品（Control）の欄に入れる。"""
    c = ref.control
    if c is not None:
        if ref.kind in CAPTION_KINDS:
            ctl.text = c.caption
        if ref.kind in CHECK_KINDS + ("OptionButton",):
            st = c.checked
            ctl.checked = "Checked" if st else ("Mixed" if st is None else None)
        if ref.kind in NUM_KINDS:
            lo, hi = ax.value_range(c)
            ctl.val, ctl.vmin, ctl.vmax = c.get("Position"), lo, hi
    elif ref.frame is not None:
        ctl.text = ref.frame.caption


def apply_xls(wb, ed, writes: list) -> tuple[dict[str, bytes], list[dict]]:
    """.xls の部品の性質を書く。(ストリームの名前 → 中身, 書いたことの記録)。OBJ の位置の直しは ed（XlsEditor）に入れる。"""
    from .cfb import SubStorage

    streams: dict[str, bytes] = {}
    done = []
    edits: list[tuple[int, int, bytes]] = []  # Ctls の (元の始まり, 元の終わり, 新しいバイト列)
    from .axfile import merge_writes

    for sheet, ctl, changes, _items in merge_writes(writes):
        ref = ctl.ax
        if ref.control is None:
            raise ax.AxError(f"「{ctl.name}」の中身を読めないので書けません：{ref.error}")
        if all(ref.control.get(k) == v or (v in (None, "") and ref.control.get(k) == ax.default_of(ref.control.struct, k))
               for k, v in changes.items()):
            continue
        kind = _struct_kind(ref)
        if ref.persistence == "ctls":
            data = wb.cfb.read("Ctls")
            edits.append((ref.start, ref.start + ref.length, ax.set_props(data, kind, changes, ref.start, ref.start + ref.length)))
        else:
            old = streams.get(ref.data_part) or SubStorage(wb.cfb, ref.part).read(ref.stream)
            streams[ref.data_part] = ax.set_props(old, kind, changes)
        done.append({"sheet": sheet, "control": ctl.name, "kind": ref.kind, "changes": dict(changes),
                     "before": {k: ref.control.get(k) for k in changes}, "parts": [ref.data_part]})
    # 書いた部品の FtPioGrbit に、fPrintCalc（印刷のときに絵を直す）と fAutoLoad（開いたときに部品を読み込む）を立てる。
    # .xls の控えの絵（描画のまとまりの BLIP）は描き直さないので、Excel が部品から絵を作り直すようにする（MS-XLS 2.5.151）
    written = {(d["sheet"], d["control"]) for d in done}
    for sh in wb.sheets:
        for c in sh.controls:
            r = getattr(c, "ax", None)
            if r is None or (sh.name, c.name) not in written or "flags_at" not in r.xls:
                continue
            _patch_obj(ed, sh.name, c.shape_id, r.xls["flags_at"], struct.pack("<H", r.xls["flags"] | FLAGS_REFRESH))
    if not edits:
        return streams, done
    data = wb.cfb.read("Ctls")
    edits.sort()
    out, at = bytearray(), 0
    for a, z, new in edits:
        out += data[at:a] + new
        at = z
    out += data[at:]
    streams["Ctls"] = bytes(out)

    def moved(p: int) -> int:
        return p + sum(len(new) - (z - a) for a, z, new in edits if z <= p)

    # 部品の位置と大きさ（FtPictFmla の lPosInCtlStm・cbBufInCtlStm）を直す：OBJ の中の同じ長さの書き換え
    for sh in wb.sheets:
        for c in sh.controls:
            r = getattr(c, "ax", None)
            if r is None or r.persistence != "ctls" or "size_at" not in r.xls:
                continue
            start, size = r.xls["pos"], r.xls.get("size", 0)
            grow = sum(len(new) - (z - a) for a, z, new in edits if start <= a and z <= start + size)
            new_pos, new_size = moved(start), size + grow
            if (new_pos, new_size) == (start, size):
                continue
            _patch_obj(ed, sh.name, c.shape_id, r.xls["pos_at"], struct.pack("<II", new_pos, new_size))
    return streams, done


FLAGS_REFRESH = 0x0004 | 0x0200  # FtPioGrbit の fPrintCalc（ビット2）と fAutoLoad（ビット9）


def _patch_obj(ed, sheet: str, shape_id: str, at: int, raw: bytes) -> None:
    """OBJ（と続く CONTINUE）の中身の at から、raw を同じ長さで書き換える。"""
    from .xls.write import Rec

    idx = ed._obj_at(sheet, shape_id)
    parts = [ed.recs[k].b for k in idx]
    blob = bytearray(b"".join(parts))
    blob[at:at + len(raw)] = raw
    at2 = 0
    for k, part in zip(idx, parts):
        nb = bytes(blob[at2:at2 + len(part)])
        at2 += len(part)
        if nb != part:
            ed.recs[k] = Rec(ed.recs[k].t, nb, False)
            ed.changed = True


def write_streams(blob: bytes, streams: dict[str, bytes]) -> bytes:
    """複合ファイルのストリームを差し替える（大きさが変わってもよい）。"""
    from .cfbwrite import replace_stream

    for name, data in streams.items():
        blob = replace_stream(blob, name, data)
    return blob


def _body(wb, r: AxRef) -> tuple[bytes, bytes]:
    """部品の (頭の CLSID, 保存の形)。"""
    from .cfb import SubStorage

    if r.persistence == "ctls":
        data = wb.cfb.read("Ctls")
        return data[r.start - getattr(r, "prefix", 0):r.start], data[r.start:r.start + r.length]
    return b"", SubStorage(wb.cfb, r.part).read(r.stream)


def prove_xls(a, b, allowed_controls) -> tuple[list[dict], list[str], set[str]]:
    """.xls の ActiveX の部品の証明：(部品の変化, 想定外, 確かめたストリーム)。部品の位置は、読み直した位置で比べる。"""
    from .xlsx.calc import values_equal

    changes_out, bad, seen = [], [], set()
    for sa in a.sheets:
        if not b.has_sheet(sa.name):
            continue
        after = {c.shape_id: c for c in b.sheet(sa.name).controls}
        for ca in sa.controls:
            ra = getattr(ca, "ax", None)
            if ra is None:
                continue
            cb = after.get(ca.shape_id)
            rb = getattr(cb, "ax", None) if cb is not None else None
            seen.add(ra.data_part or ra.part)
            what = f"{sa.name}: {ra.label}「{ca.name}」"
            if rb is None or ra.control is None or rb.control is None:
                if rb is None or (ra.error or "") != (rb.error or ""):
                    bad.append(f"{what}：読めなくなりました")
                continue
            want = None
            if isinstance(allowed_controls, dict):
                w = allowed_controls.get((sa.name, ca.name))
                want = w.get("ax") if isinstance(w, dict) else None
            elif allowed_controls is not None and (sa.name, ca.name) in allowed_controls:
                want = {k: rb.control.get(k) for k in ("Value", "Position") if ra.control.get(k) != rb.control.get(k)}
            (pa, body_a), (pb, body_b) = _body(a, ra), _body(b, rb)
            if pa != pb:
                bad.append(f"{what}：部品の頭の CLSID が変わりました")
            fa, fb = ra.xls.get("flags", 0), rb.xls.get("flags", 0)
            if fa != fb and not (want and fb == fa | FLAGS_REFRESH):
                bad.append(f"{what}：FtPioGrbit の印が、書いた部品の絵を作り直させる印（fPrintCalc・fAutoLoad）のほかにも変わりました")
            if want:
                bad += [f"{what}：{e}" for e in ax.verify(body_a, body_b, _struct_kind(ra), want)]
            elif body_a != body_b:
                bad.append(f"{what}：書き込みを指示していないのに、部品の中身が変わりました")
            if ax.state(ra.control) != ax.state(rb.control):
                changes_out.append({"sheet": sa.name, "control": ca.name, "kind": "control", "link": ca.link,
                                    "before": _state_text(ax.state(ra.control)), "after": _state_text(ax.state(rb.control))})
                if want is None:
                    bad.append(f"{what}：指示していないのに変わりました")
                elif cb.link_cell and b.has_sheet(cb.link_cell[0]):
                    lc = b.sheet(cb.link_cell[0]).cells.get(cb.link_cell[1:])
                    lv = lc.value if lc is not None else None
                    expect = _cell_value_for(rb, {k: rb.control.get(k) for k in want}, None, None)
                    if not values_equal(lv, expect) and not (isinstance(expect, str) and lv is not None and to_str(lv) == expect):
                        bad.append(f"{what}：つながったセルの値（{lv!r}）が、部品の値 {expect!r} と合いません")

    def gaps(wb) -> list[bytes]:
        """Ctls の中で、どの部品も指していないバイト。"""
        rs = sorted((r.start - getattr(r, "prefix", 0), r.start + r.length) for s in wb.sheets for c in s.controls
                    if (r := getattr(c, "ax", None)) is not None and r.persistence == "ctls")
        data, out, at = wb.cfb.read("Ctls"), [], 0
        for x, y in rs:
            out.append(data[at:x])
            at = max(at, y)
        out.append(data[at:])
        return out

    if a.cfb.has("Ctls") and b.cfb.has("Ctls"):
        if gaps(a) != gaps(b):
            bad.append("Ctls の、部品の外のバイトが変わりました")
        seen.add("Ctls")
    return changes_out, bad, seen


# ====================================================================== .doc
def doc_ref(doc, storage: str, progid: str) -> AxRef:
    """CONTROL のフィールドが指す ObjectPool の入れ物の部品を読む（ODT の fOCX・fStream、OCXNAME、contents か OCXDATA）。"""
    from .cfb import SubStorage
    from .cfbwrite import storage_clsid

    ref = AxRef("doc", "", progid=progid, part=storage, persistence="storage")
    ref.kind = ax.kind_of(progid=progid)
    try:
        sub = SubStorage(doc.cfb, storage)
        names = {n.lower(): n for n in sub.streams()}
        if not names:
            ref.error = f"入れ物 {storage} がありません"
            return ref
        if "\x03ocxname" in names:
            ref.name = sub.read(names["\x03ocxname"]).decode("utf-16-le", errors="replace").split("\x00")[0]
        cls = storage_clsid(doc.cfb, storage)
        if cls:
            ref.clsid = cls
            ref.kind = ax.kind_of(cls) or ref.kind
        odt = sub.read(names["\x03objinfo"]) if "\x03objinfo" in names else b""
        bits = struct.unpack_from("<H", odt, 0)[0] if len(odt) >= 2 else 0
        if odt and not bits >> 12 & 1:
            ref.error = "ODT の fOCX が立っていない入れ物です"
        stream_mode = bool(bits >> 13 & 1)  # fStream：部品は \003OCXDATA の1つのストリーム
        key = "\x03ocxdata" if stream_mode and "\x03ocxdata" in names else "contents"
        if "f" in names and key not in names:
            ref.kind = "Frame"
            ref.frame = ax.parse_frame(sub.read(names["f"]), sub.read(names["o"]) if "o" in names else b"")
            return ref
        if key not in names:
            ref.error = "部品の中身（contents）がありません"
            return ref
        ref.stream = names[key]
        ref.data_part = f"{storage}/{ref.stream}"
        data = sub.read(ref.stream)
        if key == "\x03ocxdata":
            prefix, _b = ax.split_stream_init(data, next((g for g, k in ax.CLSIDS.items() if k == ref.kind), None))
            ref.start = len(prefix)
        _read_control(ref, data)
    except (ax.AxError, KeyError, struct.error) as e:
        ref.error = str(e) or type(e).__name__
    return ref


def _doc_seg(blob: bytes, r: AxRef) -> tuple[bytes, bytes]:
    from .cfb import CFB

    data = CFB(blob).read(r.data_part)
    return data[:r.start], data[r.start:]


def apply_doc(doc, blob: bytes, writes: list) -> bytes:
    """.doc の部品の性質を書く：ObjectPool の入れ物の contents（か OCXDATA）を差し替える（大きさが変わってもよい）。"""
    from .cfb import CFB
    from .cfbwrite import replace_stream

    for ctl, changes in writes:
        ref = ctl.ax
        data = CFB(blob).read(ref.data_part)
        new = data[:ref.start] + ax.set_props(data, _struct_kind(ref), changes, ref.start)
        blob = replace_stream(blob, ref.data_part, new)
    return blob


def prove_doc(a, b, written: dict, ids: dict) -> tuple[list[str], set[str]]:
    """.doc の部品の証明：書き込み前後の部品を順に比べる。(想定外, 確かめたストリーム)。"""
    from .cfb import CFB

    bad, seen = [], set()
    ra_all, rb_all = getattr(a, "ax_refs", []), getattr(b, "ax_refs", [])
    if len(ra_all) != len(rb_all):
        return [f"ActiveX の部品の数が変わりました（{len(ra_all)} → {len(rb_all)}）"], seen
    ca, cb = CFB(a.blob), CFB(b.blob)
    for ra, rb in zip(ra_all, rb_all):
        what = f"{ra.label}「{ra.name}」"
        pre = ra.part + "/"
        for n in ca.streams():
            if n.startswith(pre):
                seen.add(n)
                if n != ra.data_part and (n not in cb.streams() or ca.read(n) != cb.read(n)):
                    bad.append(f"{what}：入れ物の中の {n} が変わりました")
        if ra.control is None or rb.control is None:
            if ra.data_part and ca.read(ra.data_part) != cb.read(rb.data_part):
                bad.append(f"{what}：読めない部品の中身が変わりました")
            continue
        fid = ids.get(id(ra))
        (pa, body_a), (pb, body_b) = _doc_seg(a.blob, ra), _doc_seg(b.blob, rb)
        if pa != pb:
            bad.append(f"{what}：部品の頭の CLSID が変わりました")
        if fid is not None and fid in written:
            try:
                changes = ax.plan(ra.control, written[fid], None, force=True)[0]
            except ax.AxError as e:
                bad.append(f"{what}：書いた値 {written[fid]!r} を、部品の決まりで読めません（{e}）")
                continue
            bad += [f"{what}：{e}" for e in ax.verify(body_a, body_b, _struct_kind(ra), changes)]
        elif body_a != body_b:
            bad.append(f"{what}：書き込みを指示していないのに、部品の中身が変わりました")
    return bad, seen
