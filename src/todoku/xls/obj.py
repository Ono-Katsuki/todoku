"""古いExcel（.xls、BIFF8）の OBJ レコード（図形・フォームの部品）を、仕様書 MS-XLS 2.4.181 Obj のとおりに読む。

OBJ は、部品ごとの小さな構造（FtCmo・FtCbls・FtRbo・FtSbs・ObjLinkFmla・FtCblsData・FtRboData・FtLbsData など）を
決まった順に並べたもの。それぞれ ft（種類）と cb（長さ）で始まる。読んだ値のほかに、書き込みで変えてよい
「部品の状態」（チェック・選んだ項目・値）のバイトの位置も返す。書き込み（write.py）と証明（proof.py）は、
この位置のバイトだけを変えてよいものとして扱う。

- FtCmo（2.5.143）：ot（部品の種類）・id（番号）・フラグ（fUIObj：Excel が自動で置いた補助の部品）
- FtCblsData（2.5.141）：fChecked（チェックボックスとオプションボタンの状態。0 なし・1 あり・2 どちらでもない）
- FtRboData（2.5.153）：idRadNext（同じ組の次のオプションボタンの id。0 か自分の id なら次はない）・fFirstBtn（組の先頭）
- FtSbs（2.5.154）：iVal（いまの値）・iMin・iMax・dInc・dPage・fHoriz。スピンボタン・スクロールバーのほか、
  リストボックス・ドロップダウンにもある（そのときの iVal は、上に隠れている行の数で、選んだ項目ではない）
- ObjLinkFmla（2.5.189）：ft（0x14 か 0x0E）のあとに ObjFmla（2.5.187）。ObjFmla は cbFmla（この後ろの長さ）・
  ObjectParsedFormula（2.5.198.22：cce・4バイトの未使用・rgce）・偶数にするための詰め物。ft と cbFmla が、
  ほかの部品の ft と cb にあたる
- FtLbsData（2.5.147）：リストボックス・ドロップダウン。cbFContinued・fmla（選択肢の範囲の ObjFmla）・cLines（項目の数）・
  iSel（選んだ最初の項目。1から、0 は選んでいない）・フラグ（fUseCB・fValidPlex・wListSelType・lct）・idEdit・
  dropData（ドロップダウンだけ。LbsDropData 2.5.171：長さの変わる文字 str と、その長さが奇数のときの1バイト）・
  rgLines（fValidPlex のとき、項目の文字が cLines 個）・bsels（wListSelType が 0 でないとき、項目ごとに選んだか 1バイトずつ）。
  長いときは続く CONTINUE レコードにまたがる（rgLines の文字は1つのレコードの中に収まる）。cbFContinued は長さの
  目安でしかない（Excel は 0x1FEE と書く）ので使わず、頭から順に読む。OBJ はこの部品で終わる（後ろに ftEnd はない）

つないだ中身（OBJ の中身と、続く CONTINUE の中身を順につないだもの）を読み、位置もその先頭からのバイト数で返す。
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field

FT_END, FT_CBLS, FT_RBO, FT_SBS, FT_SBSFMLA, FT_RBODATA, FT_CBLSDATA, FT_LBSDATA, FT_CBLSFMLA, FT_CMO = (
    0x00, 0x0A, 0x0B, 0x0C, 0x0E, 0x11, 0x12, 0x13, 0x14, 0x15)
OT_CHECKBOX, OT_RADIO, OT_SPIN, OT_SCROLL, OT_LIST, OT_GBOX, OT_DROP = 0x0B, 0x0C, 0x10, 0x11, 0x12, 0x13, 0x14

# 部品の種類ごとの、書き込みで変えてよい状態
STATE_FIELDS: dict[int, tuple[str, ...]] = {
    OT_CHECKBOX: ("checked",), OT_RADIO: ("checked",), OT_SPIN: ("val",), OT_SCROLL: ("val",),
    OT_LIST: ("sel", "bsels"), OT_DROP: ("sel",),
}
SEL_TYPES = {0: "single", 1: "multi", 2: "extend"}


class ObjError(ValueError):
    pass


@dataclass
class ObjInfo:
    ot: int | None = None
    id: int | None = None
    flags: int = 0  # FtCmo のフラグ
    checked: int | None = None  # FtCblsData.fChecked
    link: bytes | None = None  # つながったセル（ObjLinkFmla の rgce）
    rad_next: int | None = None  # FtRboData.idRadNext
    first_btn: bool = False  # FtRboData.fFirstBtn
    sbs: tuple[int, int, int, int, int, int] | None = None  # FtSbs の (iVal, iMin, iMax, dInc, dPage, fHoriz)
    fmla: bytes | None = None  # 選択肢の範囲（FtLbsData.fmla の rgce）
    lines: int | None = None  # cLines
    sel: int | None = None  # iSel
    sel_type: int = 0  # wListSelType
    lct: int | None = None  # 何のためのリストか（fUseCB のときだけ）。0 はふつうのドロップダウン、3 はオートフィルターなど
    items: list[str] | None = None  # rgLines
    bsels: bytes | None = None
    drop_text: str | None = None  # LbsDropData.str
    at: dict[str, tuple[int, int]] = field(default_factory=dict)  # 状態 → (位置, バイト数)
    error: str | None = None  # 途中で読めなくなったときの理由

    @property
    def ui_object(self) -> bool:
        """Excel が自動で置いた補助の部品（オートフィルター・入力規則・ピボットテーブルのボタンなど）。入力欄ではない。"""
        return bool(self.flags & 0x0100) or (self.lct is not None and self.lct != 0)


def _fmla(data: bytes, start: int, cb_fmla: int) -> bytes | None:
    """ObjFmla の cbFmla の後ろ（start）から cb_fmla バイトの、ObjectParsedFormula の rgce。"""
    if cb_fmla == 0:
        return None
    cce = struct.unpack_from("<H", data, start)[0] & 0x7FFF
    if 6 + cce > cb_fmla or start + 6 + cce > len(data):
        raise ObjError("部品の数式の長さが合いません")
    return bytes(data[start + 6:start + 6 + cce])


def _xlstr(data: bytes, pos: int) -> tuple[str, int]:
    """XLUnicodeString（2.5.294：文字数 u16・fHighByte・文字）。(文字, 次の位置)。"""
    cch = struct.unpack_from("<H", data, pos)[0]
    high = data[pos + 2] & 0x01
    pos += 3
    n = cch * (2 if high else 1)
    if pos + n > len(data):
        raise ObjError("部品の文字が途中で終わっています")
    raw = bytes(data[pos:pos + n])
    return (raw.decode("utf-16-le") if high else raw.decode("latin-1")), pos + n


def _lbs(data: bytes, p: int, cb_continued: int, info: ObjInfo) -> None:
    if cb_continued == 0:
        return  # ft と cbFContinued のほかの欄はない
    cbf = struct.unpack_from("<H", data, p)[0]
    info.fmla = _fmla(data, p + 2, cbf)
    q = p + 2 + cbf
    info.lines, info.sel, flags, _id_edit = struct.unpack_from("<HHHH", data, q)
    info.at["sel"] = (q + 2, 2)
    q += 8
    info.sel_type = (flags >> 4) & 0x03
    if flags & 0x01:  # fUseCB
        info.lct = flags >> 8
    if info.ot == OT_DROP:  # LbsDropData：wStyle 等・cLine・dxMin・str と、str が奇数バイトのときの1バイト
        q += 6
        start = q
        info.drop_text, q = _xlstr(data, q)
        q += (q - start) % 2
    if flags & 0x02:  # fValidPlex：項目の文字がこの中にある
        items = []
        for _ in range(info.lines):
            s, q = _xlstr(data, q)
            items.append(s)
        info.items = items
    if info.sel_type != 0:
        if q + info.lines > len(data):
            raise ObjError("複数選択の印（bsels）が途中で終わっています")
        info.bsels = bytes(data[q:q + info.lines])
        info.at["bsels"] = (q, info.lines)


def parse(data: bytes) -> ObjInfo:
    """OBJ の中身（続く CONTINUE の中身もつないだもの）を読む。壊れていても、読めたところまでを返す。"""
    info = ObjInfo()
    pos = 0
    try:
        while pos + 4 <= len(data):
            ft, cb = struct.unpack_from("<HH", data, pos)
            body = pos + 4
            if ft == FT_END:
                break
            if ft == FT_CMO:
                info.ot, info.id, info.flags = struct.unpack_from("<HHH", data, body)
            elif ft == FT_CBLSDATA:
                info.checked = struct.unpack_from("<H", data, body)[0]
                info.at["checked"] = (body, 2)
            elif ft == FT_RBODATA:
                info.rad_next, first = struct.unpack_from("<HH", data, body)
                info.first_btn = bool(first)
            elif ft == FT_SBS:
                info.sbs = struct.unpack_from("<hhhhhH", data, body + 4)  # unused1（4バイト）の後ろ
                info.at["val"] = (body + 4, 2)
            elif ft in (FT_CBLSFMLA, FT_SBSFMLA):
                info.link = _fmla(data, body, cb)  # cb は ObjFmla の cbFmla
            elif ft == FT_LBSDATA:
                _lbs(data, body, cb, info)
                break  # FtLbsData は OBJ の最後の部品
            pos = body + cb
    except (struct.error, IndexError, UnicodeDecodeError, ObjError) as e:
        info.error = str(e) or type(e).__name__
    return info


def without_state(data: bytes) -> bytes:
    """部品の状態（書き込みで変えてよいバイト）を 0 にしたもの。証明で、状態のほかが1バイトも変わっていないかを比べる。"""
    info = parse(data)
    out = bytearray(data)
    for key in STATE_FIELDS.get(info.ot, ()):
        if key in info.at:
            off, n = info.at[key]
            out[off:off + n] = bytes(n)
    return bytes(out)


def with_state(data: bytes, change: dict) -> bytes:
    """部品の状態だけを変えた中身。change は {"checked": "Checked"/"Mixed"/None, "sel": 番号, "val": 値,
    "multi_sel": [番号, ...]}（xlsx の部品の書き込みと同じ形）。長さは変わらない。"""
    info = parse(data)
    if info.error:
        raise ObjError(f"部品（OBJ）を読めません：{info.error}")
    allowed = STATE_FIELDS.get(info.ot, ())
    out = bytearray(data)
    for key, v in change.items():
        if key == "checked" and "checked" in allowed and "checked" in info.at:
            struct.pack_into("<H", out, info.at["checked"][0], {"Checked": 1, "Mixed": 2}.get(v, 0))
        elif key == "sel" and "sel" in allowed and "sel" in info.at:
            if not (isinstance(v, int) and 0 <= v <= (info.lines or 0)):
                raise ObjError(f"選んだ項目の番号 {v!r} は、項目の数（{info.lines}）の範囲外です")
            struct.pack_into("<H", out, info.at["sel"][0], v)
        elif key == "val" and "val" in allowed and "val" in info.at:
            lo, hi = info.sbs[1], info.sbs[2]
            if not (isinstance(v, int) and lo <= v <= hi):
                raise ObjError(f"値 {v!r} は、{lo}〜{hi} の範囲外です")
            struct.pack_into("<h", out, info.at["val"][0], v)
        elif key == "multi_sel" and "bsels" in allowed and "bsels" in info.at:
            off, n = info.at["bsels"]
            chosen = set(v or ())
            if any(not (isinstance(i, int) and 1 <= i <= n) for i in chosen):
                raise ObjError(f"選んだ項目の番号 {sorted(chosen)} に、項目の数（{n}）の範囲外のものがあります")
            out[off:off + n] = bytes(1 if k + 1 in chosen else 0 for k in range(n))
        else:
            raise ObjError(f"この部品（種類 0x{info.ot or 0:02X}）には、状態「{key}」を書く場所がありません")
    return bytes(out)
