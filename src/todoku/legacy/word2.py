"""Word for Windows 1.x・2.0（複合ファイルでない .doc）の字を読む（読むだけ）。

FIB の形は Word 6 の資料の先頭と共通：[0:2] 印（0xA59B が Word 1.x、0xA5DB が Word 2.0）、[2:4] nFib、[6:8] lid、[10:12] フラグ（0x0004 高速保存・
0x0100 暗号）、[0x18:0x1C] fcMin、[0x1C:0x20] fcMac、[0x34:] ccpText・ccpFtn・ccpHdd。本文は fcMin から ccpText バイト、続いて脚注・ヘッダー。

字は 1 バイトで、日本語の Word（lid 0x0411）は Shift_JIS、それ以外は Windows-1252。段落の区切りは 0x0D、表のマスの区切りは 0x07、フィールドの印は 0x13〜0x15。
高速保存（断片に分かれた文書）と暗号は読まず、F と次の一手を伝える（並べ方を推測しない）。
"""

from __future__ import annotations

import re
import struct

from .. import readonlydoc as rd

KIND_JA = "古い Word（Word 1.x・2.0 の .doc。読むだけ）"
WRITE = "Word 1.x・2.0 の形には書きません。Word で .docx に保存し直してもらえば書けます"
MAGIC = {0xA59B: "Word 1.x", 0xA5DB: "Word 2.0"}


class Word2Error(ValueError):
    pass


def is_word2(b: bytes) -> bool:
    if len(b) < 0x40:
        return False
    magic = struct.unpack_from("<H", b, 0)[0]
    if magic not in MAGIC:
        return False
    fc_min, fc_mac = struct.unpack_from("<II", b, 0x18)
    return 0 < fc_min <= fc_mac <= len(b)


def _stories(b: bytes) -> tuple[list[tuple[str, bytes]], str]:
    if not is_word2(b):
        raise Word2Error("Word 1.x・2.0 の文書として読めません")
    flags = struct.unpack_from("<H", b, 0x0A)[0]
    if flags & 0x0100:
        raise Word2Error("パスワードで保護された Word 1.x・2.0 の文書です。解いて .docx に保存し直したものか、PDF を頼んでください")
    if flags & 0x0004:
        raise Word2Error("高速保存された Word 1.x・2.0 の文書です。字が断片に分かれていて、並べ方を確かめられないので読みません。Word で開いて、全部を保存し直したもの（.docx か PDF）を頼んでください")
    lid = struct.unpack_from("<H", b, 6)[0]
    fc_min, fc_mac = struct.unpack_from("<II", b, 0x18)
    ccp_text, ccp_ftn, ccp_hdd = struct.unpack_from("<III", b, 0x34)
    if fc_min + ccp_text + ccp_ftn + ccp_hdd > fc_mac:
        raise Word2Error("本文の長さがファイルと合いません（壊れています）")
    enc = "cp932" if lid == 0x0411 else "cp1252"
    p = fc_min
    out = []
    for name, n in (("本文", ccp_text), ("脚注", ccp_ftn), ("ヘッダー・フッター", ccp_hdd)):
        if n:
            out.append((name, b[p:p + n]))
        p += n
    return out, enc


def _clean(s: str) -> str:
    # フィールド：\x13 命令 \x14 結果 \x15 → 結果だけ（結果がなければ消す）
    s = re.sub(r"\x13[^\x13\x14\x15]*\x14", "", s)
    s = re.sub(r"\x13[^\x13\x14\x15]*\x15", "", s)
    s = s.replace("\x15", "").replace("\x01", "")
    return s


def build(blob: bytes, name: str) -> dict:
    stories, enc = _stories(blob)
    secs = []
    for nm, raw in stories:
        text = _clean(raw.decode(enc, "replace"))
        lines = []
        for para in re.split(r"[\r\x0c]", text):
            cells = [c.strip() for c in para.split("\x07")]
            lines.append(" ｜ ".join(c for c in cells if c) if len(cells) > 1 else para.replace("\x0b", " ").strip())
        secs.append((nm, [x for x in lines if x]))
    return rd.make("word2", KIND_JA, name, blob, secs,
                   "元のファイルが正本。字のバイトを文書の順に読みました（字の飾り・表の形・絵は読みません）。" + WRITE, encoding=enc)


def grade(blob: bytes, name: str):
    from ..grade import _locked

    try:
        win = build(blob, name)
    except Word2Error as e:
        return _locked("読めない", str(e), "読めないので、書けません")
    return rd.grade(win, "B", "古い Word は字を読む", "字のバイトを文書の順に読みます。表の形・字の飾り・絵は読まないので、並びの意味はエージェントが判断します",
                    "Word 1.x・2.0 には書かない", WRITE)


def audit(blob: bytes, name: str):
    from ..xlsx.audit import Audit, Check

    try:
        win = build(blob, name)
        stories, enc = _stories(blob)
    except Word2Error as e:
        au = Audit(name)
        au.checks.append(Check("読む", False, str(e)))
        return au
    raw = []
    for _n, x in stories:
        t = _clean(x.decode(enc, "replace"))
        raw.append(re.sub(r"[\x00-\x1f]", "", t))
    return rd.audit(win, name, raw, "本文・脚注・ヘッダーのバイト", WRITE)
