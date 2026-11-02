"""マクロ（VBA）のソースを、vbaProject.bin（Excel・Word のマクロの入れ物）から取り出す。

Microsoft の仕様書 MS-OVBA の手順どおり：
- 複合ファイルの中の VBA/dir ストリーム（圧縮）を解き、モジュールの記録（名前・ストリーム名・ソースの位置・種類）を読む
- 各モジュールのストリームの、ソースの位置から後ろを解き、プロジェクトのコードページで文字にする
圧縮（2.4.1）：先頭が 0x01、4096 バイトごとのかたまり。かたまりの頭の2バイトの下12ビットが大きさ−3、最上位ビットが圧縮の有無。
"""

from __future__ import annotations

import math
import re
import struct
from dataclasses import dataclass, field

from .cfb import CFB


class VbaError(ValueError):
    pass


def decompress(data: bytes) -> bytes:
    """MS-OVBA 2.4.1.3.1 Decompression Algorithm。"""
    if not data or data[0] != 0x01:
        raise VbaError("圧縮されたマクロのデータではありません（先頭が 0x01 でない）")
    out = bytearray()
    cur = 1
    end = len(data)
    while cur < end:
        if cur + 2 > end:
            break
        header = struct.unpack_from("<H", data, cur)[0]
        size = (header & 0x0FFF) + 3
        flag = (header & 0x8000) >> 15
        chunk_start_out = len(out)
        chunk_end = min(end, cur + size)
        cur += 2
        if flag == 0:
            out += data[cur:cur + 4096]  # 圧縮していないかたまり
            cur += 4096
            continue
        while cur < chunk_end:
            flags = data[cur]
            cur += 1
            for i in range(8):
                if cur >= chunk_end:
                    break
                if not (flags >> i) & 1:
                    out.append(data[cur])
                    cur += 1
                    continue
                token = struct.unpack_from("<H", data, cur)[0]
                diff = len(out) - chunk_start_out
                bit_count = max(math.ceil(math.log2(diff)) if diff > 0 else 0, 4)
                length_mask = 0xFFFF >> bit_count
                length = (token & length_mask) + 3
                offset = ((token & (~length_mask & 0xFFFF)) >> (16 - bit_count)) + 1
                src = len(out) - offset
                if src < 0:
                    raise VbaError("圧縮されたマクロのデータが壊れています（コピーの元がない）")
                for k in range(length):  # 重なるコピーもあるので1バイトずつ
                    out.append(out[src + k])
                cur += 2
    return bytes(out)


def decrypt(enc: bytes) -> bytes:
    """MS-OVBA 2.4.3.3 Decryption（PROJECT ストリームの CMG・DPB・GC の値の難読化を解く）。"""
    if len(enc) < 3:
        raise VbaError("短すぎます")
    seed, version_enc, proj_key_enc = enc[0], enc[1], enc[2]
    if seed ^ version_enc != 2:
        raise VbaError("難読化の版が 2 ではありません")
    proj_key = seed ^ proj_key_enc
    u1, e1, e2 = proj_key, proj_key_enc, version_enc
    i = 3
    ignored = (seed & 6) // 2

    def step(b: int) -> int:
        nonlocal u1, e1, e2
        x = b ^ ((e2 + u1) & 0xFF)
        e2, e1, u1 = e1, b, x
        return x

    for _ in range(ignored):
        step(enc[i])
        i += 1
    length = 0
    for k in range(4):
        length += step(enc[i]) << (8 * k)
        i += 1
    data = bytes(step(b) for b in enc[i:i + length])
    if len(data) != length:
        raise VbaError("長さが合いません")
    return data


KINDS = {"procedural": "標準モジュール", "document": "シート・ブックのモジュール", "class": "クラスモジュール",
         "form": "ユーザーフォーム"}


@dataclass
class Module:
    name: str
    stream: str
    offset: int
    kind: str  # procedural / document / class / form
    source: str = ""
    procedures: list[str] = field(default_factory=list)

    @property
    def has_code(self) -> bool:
        """宣言（Attribute・Option）だけでなく、手続きがあるか。"""
        return bool(self.procedures)


@dataclass
class Project:
    code_page: int
    name: str
    modules: list[Module]
    locked: bool = False  # 見るのにパスワードが要る（ソースそのものは読める）

    @property
    def procedures(self) -> list[str]:
        return [f"{m.name}.{p}" for m in self.modules for p in m.procedures]


_PROC = re.compile(r"^\s*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(?:Get|Let|Set))\s+([^\s(]+)",
                   re.I | re.M)


def _records(d: bytes):
    """dir ストリームの記録：(Id, 中身)。PROJECTVERSION（0x0009）だけは、大きさの欄が 4 でも中身は 6 バイト。"""
    i = 0
    while i + 6 <= len(d):
        rid, size = struct.unpack_from("<HI", d, i)
        i += 6
        if rid == 0x0009:
            size = 6
        yield rid, d[i:i + size]
        i += size


def read_project(blob: bytes) -> Project:
    """vbaProject.bin（複合ファイル）からモジュールとソースを取り出す。"""
    c = CFB(blob)
    names = {p.lower(): p for p in c.streams()}
    dir_path = next((p for k, p in names.items() if k.endswith("vba/dir") or k == "dir"), None)
    if dir_path is None:
        raise VbaError("マクロの目録（VBA/dir）がありません")
    base = dir_path[: -len("dir")]
    d = decompress(c.read(dir_path))
    code_page = 1252
    proj_name = ""
    modules: list[Module] = []
    cur: dict | None = None
    for rid, data in _records(d):
        if rid == 0x0003 and len(data) >= 2:
            code_page = struct.unpack_from("<H", data)[0]
        elif rid == 0x0004:
            proj_name = data.decode(_codec(code_page), errors="replace")
        elif rid == 0x0019:  # MODULENAME
            cur = {"name": data.decode(_codec(code_page), errors="replace"), "kind": "procedural"}
        elif rid == 0x0047 and cur is not None:  # MODULENAMEUNICODE
            cur["name"] = data.decode("utf-16-le", errors="replace")
        elif rid == 0x001A and cur is not None:  # MODULESTREAMNAME
            cur["stream"] = data.decode(_codec(code_page), errors="replace")
        elif rid == 0x0032 and cur is not None:  # MODULESTREAMNAME の Unicode
            cur["stream"] = data.decode("utf-16-le", errors="replace")
        elif rid == 0x0031 and cur is not None:  # MODULEOFFSET
            cur["offset"] = struct.unpack_from("<I", data)[0]
        elif rid == 0x0022 and cur is not None:  # MODULETYPE：ドキュメント・クラス・フォーム
            cur["kind"] = "class"
        elif rid == 0x0021 and cur is not None:  # MODULETYPE：標準モジュール
            cur["kind"] = "procedural"
        elif rid == 0x002B and cur is not None:  # モジュールの記録の終わり
            modules.append(Module(cur["name"], cur.get("stream", cur["name"]), cur.get("offset", 0), cur["kind"]))
            cur = None
    project_text = ""
    # PROJECT ストリームは VBA の入れ物の隣（.xlsm の vbaProject.bin なら根、.doc なら Macros/ の中）
    root = dir_path[: -len("VBA/dir")] if dir_path.lower().endswith("vba/dir") else ""
    proj_path = names.get((root + "PROJECT").lower())
    if proj_path is not None:
        project_text = c.read(proj_path).decode(_codec(code_page), errors="replace")
    docs = {m.group(1).strip() for m in re.finditer(r"^Document=([^/\r\n]+)", project_text, re.M)}
    forms = {m.group(1).strip() for m in re.finditer(r"^BaseClass=([^\r\n]+)", project_text, re.M)}
    # 見るためのパスワード：PROJECT ストリームの CMG（保護の状態）を解いて、利用者が保護したか（fUserProtected）を見る
    cmg = re.search(r'^CMG="([0-9A-Fa-f]+)"', project_text, re.M)
    locked = False
    if cmg:
        try:
            state = decrypt(bytes.fromhex(cmg.group(1)))
            locked = bool(state and state[0] & 0x01)
        except (VbaError, ValueError):
            locked = False
    for m in modules:
        if m.name in docs:
            m.kind = "document"
        elif m.name in forms:
            m.kind = "form"
        path = names.get((base + m.stream).lower())
        if path is None:
            continue
        raw = c.read(path)
        try:
            src = decompress(raw[m.offset:])
        except VbaError:
            continue  # ソースのない（p-code だけの）モジュール
        m.source = src.decode(_codec(code_page), errors="replace").replace("\r\n", "\n")
        m.procedures = [f"{x.group(2)}" for x in _PROC.finditer(m.source)
                        if not x.group(0).lstrip().lower().startswith(("declare", "end"))]
    return Project(code_page, proj_name, modules, locked)


def _codec(cp: int) -> str:
    return {932: "cp932", 936: "gbk", 949: "cp949", 950: "cp950", 65001: "utf-8", 1200: "utf-16-le",
            10001: "shift_jis"}.get(cp, f"cp{cp}")


def visible_source(m: Module) -> str:
    """窓に出すソース：Attribute 行（VBE が自動で付け、編集の画面には出さない設定）を除く。"""
    lines = [ln for ln in m.source.split("\n") if not ln.startswith("Attribute ")]
    return "\n".join(lines).strip("\n")
