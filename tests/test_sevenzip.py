"""7z：入れ物を形式の説明（7zFormat.txt）のとおりに手で作って解く。7zz があれば、同じ書庫を 7zz でも解いて照らし、7zz の作った書庫も解く。"""

from __future__ import annotations

import bz2
import hashlib
import lzma
import random
import shutil
import struct
import subprocess
import zlib

import pytest

from pathlib import Path

from todoku import api, archives
from todoku.aes import Aes
from todoku.book import kind_of
from todoku.legacy import branch, ppmd7, sevenzip

from coding_helpers import LzmaRangeEncoder

R = random.Random(9)
WORDS = [bytes(R.randrange(97, 123) for _ in range(R.randrange(2, 10))) for _ in range(300)]
TEXT = b" ".join(R.choice(WORDS) for _ in range(3000))
CSV = "名前,年\n山田,3\n".encode()
FIXTURES = Path(__file__).parent / "fixtures" / "ppmd"
HAS7 = shutil.which("7zz") is not None


# ---------------------------------------------------------------- 書庫を作る
def num(v: int) -> bytes:
    if v < 0x80:
        return bytes((v,))
    if v < 0x4000:
        return bytes((0x80 | (v >> 8), v & 255))
    if v < 0x200000:
        return bytes((0xC0 | (v >> 16),)) + (v & 0xFFFF).to_bytes(2, "little")
    if v < 0x10000000:
        return bytes((0xE0 | (v >> 24),)) + (v & 0xFFFFFF).to_bytes(3, "little")
    return b"\xFF" + v.to_bytes(8, "little")


def coder(cid: bytes, props: bytes = b"", ni: int = 1, no: int = 1) -> bytes:
    flag = len(cid) | (0x10 if ni != 1 or no != 1 else 0) | (0x20 if props else 0)
    o = bytes((flag,)) + cid
    if ni != 1 or no != 1:
        o += num(ni) + num(no)
    if props:
        o += num(len(props)) + props
    return o


def folder(coders: list[bytes], binds: list[tuple[int, int]] = ()) -> bytes:
    return num(len(coders)) + b"".join(coders) + b"".join(num(a) + num(b) for a, b in binds)


def streams(packs: list[list[int]], folders: list[tuple[bytes, list[int], int | None]], subs: list[tuple[list[int], list[int | None]]] | None) -> bytes:
    """folders：(folder のバイト, 出力の大きさ, CRC)。subs：フォルダごとの (大きさ, CRC)。"""
    allp = [x for p in packs for x in p]
    o = b"\x06" + num(0) + num(len(allp)) + b"\x09" + b"".join(num(x) for x in allp) + b"\x00"
    o += b"\x07\x0B" + num(len(folders)) + b"\x00" + b"".join(f[0] for f in folders)
    o += b"\x0C" + b"".join(num(x) for f in folders for x in f[1])
    crcs = [f[2] for f in folders]
    if any(c is not None for c in crcs):
        o += b"\x0A\x01" + b"".join(struct.pack("<I", c or 0) for c in crcs)
    o += b"\x00"
    if subs is not None:
        o += b"\x08\x0D" + b"".join(num(len(s[0])) for s in subs)
        sizes = [x for s in subs for x in s[0][:-1]]
        if sizes:
            o += b"\x09" + b"".join(num(x) for x in sizes)
        crc = [c for s in subs for c in s[1]]
        o += b"\x0A\x01" + b"".join(struct.pack("<I", c) for c in crc)
        o += b"\x00"
    return o + b"\x00"


def names_prop(names: list[str]) -> bytes:
    body = b"\x00" + b"".join(n.encode("utf-16-le") + b"\0\0" for n in names)
    return b"\x11" + num(len(body)) + body


def archive(body: bytes, head: bytes) -> bytes:
    start = b"".join((struct.pack("<QQI", len(body), len(head), zlib.crc32(head) & 0xFFFFFFFF),))
    return b"7z\xbc\xaf\x27\x1c\x00\x04" + struct.pack("<I", zlib.crc32(start) & 0xFFFFFFFF) + start + body + head


def lzma2_props(dic: int = 1 << 20) -> bytes:
    bits = 0
    while (2 | (bits & 1)) << (bits // 2 + 11) < dic:
        bits += 1
    return bytes((bits,))


def lz2(data: bytes, pre: list | None = None) -> bytes:
    return lzma.compress(data, format=lzma.FORMAT_RAW, filters=(pre or []) + [{"id": lzma.FILTER_LZMA2, "dict_size": 1 << 20}])


def single(data: bytes, packed: bytes, coders: list[bytes], sizes: list[int], binds=(), name: str = "a.bin") -> bytes:
    head = b"\x01\x04" + streams([[len(packed)]], [(folder(coders, binds), sizes, zlib.crc32(data) & 0xFFFFFFFF)], None)
    head += b"\x05" + num(1) + names_prop([name]) + b"\x00\x00"
    return archive(packed, head)


def get(blob: bytes, name: str = "a.bin", pw: str | None = None) -> bytes:
    archives.use_password(pw)
    try:
        return archives.extract(blob, name, "7z")
    finally:
        archives.use_password(None)


def seven(tmp_path, blob: bytes, name: str = "a.bin", pw: str | None = None) -> bytes | None:
    """7zz でも解く（なければ None）。"""
    if not HAS7:
        return None
    p = tmp_path / "chk.7z"
    p.write_bytes(blob)
    cmd = ["7zz", "e", "-so", "-bd"] + ([f"-p{pw}"] if pw else []) + [str(p), name]
    r = subprocess.run(cmd, capture_output=True, timeout=60, stdin=subprocess.DEVNULL)
    assert r.returncode == 0, r.stderr.decode(errors="replace")
    return r.stdout


# ---------------------------------------------------------------- 圧縮と変換
def test_lzma2_and_container(tmp_path):
    blob = single(TEXT, lz2(TEXT), [coder(b"\x21", lzma2_props())], [len(TEXT)])
    assert kind_of(blob) == "7z"
    assert get(blob) == TEXT
    assert seven(tmp_path, blob) in (None, TEXT)


def test_lzma1(tmp_path):
    packed = lzma.compress(TEXT, format=lzma.FORMAT_RAW, filters=[{"id": lzma.FILTER_LZMA1, "dict_size": 1 << 20, "lc": 3, "lp": 0, "pb": 2}])
    props = bytes(((2 * 5 + 0) * 9 + 3,)) + (1 << 20).to_bytes(4, "little")
    blob = single(TEXT, packed, [coder(b"\x03\x01\x01", props)], [len(TEXT)])
    assert get(blob) == TEXT
    assert seven(tmp_path, blob) in (None, TEXT)


@pytest.mark.parametrize("name,cid,packed", [
    ("deflate", b"\x04\x01\x08", lambda d: (lambda c: c.compress(d) + c.flush())(zlib.compressobj(9, zlib.DEFLATED, -15))),
    ("bzip2", b"\x04\x02\x02", bz2.compress),
    ("copy", b"\x00", lambda d: d),
])
def test_stdlib_coders(tmp_path, name, cid, packed):
    blob = single(TEXT, packed(TEXT), [coder(cid)], [len(TEXT)])
    assert get(blob) == TEXT
    assert seven(tmp_path, blob) in (None, TEXT)


def test_deflate64_fixed_and_stored():
    # Deflate と同じ形の範囲（固定の符号・無圧縮のブロック）は、zlib の出力を Deflate64 としても解ける
    c = zlib.compressobj(9, zlib.DEFLATED, -15)
    d = c.compress(TEXT) + c.flush()
    assert sevenzip.deflate64(d, 1 << 30) == TEXT
    c = zlib.compressobj(0, zlib.DEFLATED, -15)
    d = c.compress(CSV) + c.flush()
    assert sevenzip.deflate64(d, 1 << 30) == CSV
    # 長さの符号 285（16 ビットの追加）：固定の符号で作る
    w = _Lsb()
    w.bits(1, 1)
    w.bits(1, 2)
    w.fixed(ord("x"))
    w.fixed(285)
    w.bits(1000 - 3, 16)
    w.bits(0, 5)  # 距離の符号 0：固定の符号は上位のビットから 5 ビット
    w.fixed(256)
    assert sevenzip.deflate64(w.done(), 1 << 30) == b"x" * 1001


class _Lsb:
    def __init__(self):
        self.a = 0
        self.n = 0

    def bits(self, v, k):
        self.a |= v << self.n
        self.n += k

    def rev(self, v, k):
        for i in range(k - 1, -1, -1):
            self.bits((v >> i) & 1, 1)

    def fixed(self, s):
        if s < 144:
            self.rev(0x30 + s, 8)
        elif s < 256:
            self.rev(0x190 + s - 144, 9)
        elif s < 280:
            self.rev(s - 256, 7)
        else:
            self.rev(0xC0 + s - 280, 8)

    def done(self):
        return self.a.to_bytes((self.n + 7) // 8, "little")


def _x86_like() -> bytes:
    r = random.Random(1)
    o = bytearray()
    while len(o) < 20000:
        o += bytes((0xE8,)) + r.randrange(0, 60000).to_bytes(4, "little") + bytes(r.randrange(256) for _ in range(r.randrange(0, 20)))
    return bytes(o)


def _words4(pat) -> bytes:
    r = random.Random(2)
    o = bytearray()
    for _ in range(4000):
        o += pat(r) if r.random() < .5 else r.randrange(1 << 32).to_bytes(4, "little")
    return bytes(o)


FILTERS = [
    ("bcj", b"\x03\x03\x01\x03", lzma.FILTER_X86, _x86_like()),
    ("arm", b"\x03\x03\x05\x01", lzma.FILTER_ARM, _words4(lambda r: r.randrange(1 << 24).to_bytes(3, "little") + b"\xEB")),
    ("armt", b"\x03\x03\x07\x01", lzma.FILTER_ARMTHUMB, _words4(lambda r: bytes((r.randrange(256), 0xF0 | r.randrange(8), r.randrange(256), 0xF8 | r.randrange(8))))),
    ("ppc", b"\x03\x03\x02\x05", lzma.FILTER_POWERPC, _words4(lambda r: bytes((0x48 | r.randrange(4), r.randrange(256), r.randrange(256), (r.randrange(64) << 2) | 1)))),
    ("sparc", b"\x03\x03\x08\x05", lzma.FILTER_SPARC, _words4(lambda r: bytes((0x40, r.randrange(64), r.randrange(256), r.randrange(256))))),
    ("ia64", b"\x03\x03\x04\x01", lzma.FILTER_IA64, bytes(random.Random(3).randrange(256) for _ in range(16000))),
]


@pytest.mark.parametrize("name,cid,fid,data", FILTERS, ids=[f[0] for f in FILTERS])
def test_branch_filters(tmp_path, name, cid, fid, data):
    """順方向は標準の lzma（liblzma）の変換：Python で書いた逆と、別の実装で照らす。"""
    packed = lz2(data, [{"id": fid}])
    # フォルダは「LZMA2 → 変換」の 2 つのコーダー。1 番の出力が主
    coders = [coder(cid), coder(b"\x21", lzma2_props())]
    blob = single(data, packed, coders, [len(data), len(data)], [(0, 1)])
    assert get(blob) == data
    assert seven(tmp_path, blob) in (None, data)


def test_delta_and_swap(tmp_path):
    packed = lz2(TEXT, [{"id": lzma.FILTER_DELTA, "dist": 4}])
    blob = single(TEXT, packed, [coder(b"\x03", b"\x03"), coder(b"\x21", lzma2_props())], [len(TEXT), len(TEXT)], [(0, 1)])
    assert get(blob) == TEXT
    assert seven(tmp_path, blob) in (None, TEXT)
    assert branch.swap(b"abcd", 2) == b"badc" and branch.swap(b"abcdef", 4) == b"dcbaef"


def _make_bcj2_streams(data: bytes) -> tuple[bytes, bytes, bytes, bytes]:
    """BCJ2 の順方向（すべての E8・E9・条件の飛びを番地にする）。範囲符号は LZMA の 1 ビットの符号。"""
    p = [1024] * 258
    rc = LzmaRangeEncoder()
    main = bytearray()
    call = bytearray()
    jump = bytearray()
    prev = 0
    i = 0
    n = len(data)

    while i < n:
        b = data[i]
        main.append(b)
        i += 1
        is_j = b & 0xFE == 0xE8 or (prev == 0x0F and b & 0xF0 == 0x80)
        if not is_j:
            prev = b
            continue
        k = prev if b == 0xE8 else 256 if b == 0xE9 else 257
        if i + 4 <= n and data[i + 3] == 0:
            rc.encode_bit(p, k, 1)
            rel = int.from_bytes(data[i:i + 4], "little")
            abs_ = (rel + i + 4) & 0xFFFFFFFF
            (call if b == 0xE8 else jump).extend(abs_.to_bytes(4, "big"))
            prev = data[i + 3]
            i += 4
        else:
            rc.encode_bit(p, k, 0)
            prev = b
    return bytes(main), bytes(call), bytes(jump), rc.finish()


def test_bcj2_hand_built(tmp_path):
    data = _x86_like() + b"\x0F\x84\x10\x00\x00\x00" * 40
    main, call, jump, rc = _make_bcj2_streams(data)
    assert branch.bcj2(main, call, jump, rc, len(data)) == data
    # 入れ物：BCJ2（入力 4）に、4 つのパックのストリームを直接つなぐ（Copy を通さない）
    c = coder(b"\x03\x03\x01\x1b", b"", 4, 1)
    fold = num(1) + c + num(0) + num(1) + num(2) + num(3)  # 束ねなし・パックの入力の番号
    body = main + call + jump + rc
    head = b"\x01\x04\x06" + num(0) + num(4) + b"\x09" + b"".join(num(len(x)) for x in (main, call, jump, rc)) + b"\x00"
    head += b"\x07\x0B" + num(1) + b"\x00" + fold + b"\x0C" + num(len(data)) + b"\x0A\x01" + struct.pack("<I", zlib.crc32(data) & 0xFFFFFFFF) + b"\x00\x00"
    head += b"\x05" + num(1) + names_prop(["a.bin"]) + b"\x00\x00"
    blob = archive(body, head)
    assert get(blob) == data


def _ppmd_restart_data() -> bytes:
    r = random.Random(4)
    return b" ".join(r.choice(WORDS) for _ in range(1500)) + bytes(r.randrange(256) for _ in range(8000))


def test_ppmd_made_by_7zz_with_restart(tmp_path, monkeypatch):
    """見本は 7zz で作った（-m0=ppmd:mem=64k:o=6）。小さなモデルに多くの種類の字：モデルのやり直しと、空きのつなぎ直しを通る。"""
    blob = (FIXTURES / "ppmd_restart.7z").read_bytes()
    restarts = []
    plain = ppmd7.Ppmd7.restart
    monkeypatch.setattr(ppmd7.Ppmd7, "restart", lambda self: (restarts.append(1), plain(self))[1])
    assert get(blob) == _ppmd_restart_data()
    assert len(restarts) > 1  # 初めの 1 回のほかに、モデルを使い切ってやり直した
    assert seven(tmp_path, blob) in (None, _ppmd_restart_data())


def test_ppmd_rejects_huge_model():
    with pytest.raises(ppmd7.PpmdError):
        ppmd7.Ppmd7(1 << 40, 6)


def _aes_props(salt: bytes, iv: bytes, cycles: int) -> bytes:
    b0 = cycles | (0x80 if salt else 0) | (0x40 if iv else 0)
    b1 = ((max(len(salt) - 1, 0)) << 4) | max(len(iv) - 1, 0)
    return bytes((b0, b1)) + salt + iv


def _aes_encrypt(data: bytes, pw: str, salt: bytes, iv: bytes, cycles: int) -> bytes:
    h = hashlib.sha256()
    pb = pw.encode("utf-16-le")
    for i in range(1 << cycles):
        h.update(salt + pb + i.to_bytes(8, "little"))
    return Aes(h.digest()).encrypt_cbc(data + bytes((16 - len(data) % 16) % 16), (iv + bytes(16))[:16])


def test_aes_files(tmp_path):
    salt, iv = b"salt5678", bytes(range(16))
    packed = _aes_encrypt(lz2(TEXT), "ひみつ", salt, iv, 10)
    aes = coder(b"\x06\xf1\x07\x01", _aes_props(salt, iv, 10))
    lz = coder(b"\x21", lzma2_props())
    n = len(lz2(TEXT))
    blob = single(TEXT, packed, [lz, aes], [len(TEXT), n], [(0, 1)])
    assert get(blob, pw="ひみつ") == TEXT
    assert seven(tmp_path, blob, pw="ひみつ") in (None, TEXT)
    with pytest.raises(archives.NeedsPassword):
        get(blob)
    with pytest.raises(archives.WrongPassword):
        get(blob, pw="ちがう")
    p = tmp_path / "e.7z"
    p.write_bytes(blob)
    assert api.summary(str(p))["members"][0]["name"] == "a.bin"
    with pytest.raises(api.PasswordNeeded):
        api.fields(f"{p}!/a.bin")
    with pytest.raises(api.PasswordWrong):
        api.summary(str(p), password="ちがう")
    assert api.summary(str(p), password="ひみつ")["members"][0]["name"] == "a.bin"


def test_encoded_and_encrypted_header(tmp_path):
    inner = b"\x01\x04" + streams([[len(CSV)]], [(folder([coder(b"\x00")]), [len(CSV)], zlib.crc32(CSV) & 0xFFFFFFFF)], None)
    inner += b"\x05" + num(1) + names_prop(["様式/申請.csv"]) + b"\x00\x00"
    # EncodedHeader（縮めただけ）
    packed = lz2(inner)
    enc = b"\x17" + _enc_streams(len(CSV), [[len(packed)]], [(folder([coder(b"\x21", lzma2_props())]), [len(inner)], zlib.crc32(inner) & 0xFFFFFFFF)])
    blob = archive(CSV + packed, enc)
    assert get(blob, "様式/申請.csv") == CSV
    assert seven(tmp_path, blob, "様式/申請.csv") in (None, CSV)
    # 暗号にした見出し
    salt, iv = b"s" * 8, bytes(16)
    p2 = _aes_encrypt(inner, "pw", salt, iv, 8)
    enc = b"\x17" + _enc_streams(len(CSV), [[len(p2)]], [(folder([coder(b"\x06\xf1\x07\x01", _aes_props(salt, iv, 8))]), [len(inner)], zlib.crc32(inner) & 0xFFFFFFFF)])
    blob = archive(CSV + p2, enc)
    assert get(blob, "様式/申請.csv", pw="pw") == CSV
    assert seven(tmp_path, blob, "様式/申請.csv", pw="pw") in (None, CSV)
    p = tmp_path / "h.7z"
    p.write_bytes(blob)
    with pytest.raises(api.PasswordNeeded):
        api.summary(str(p))
    with pytest.raises(api.PasswordWrong):
        api.summary(str(p), password="x")
    assert api.summary(str(p), password="pw")["members"][0]["kind"] == "csv"


def _enc_streams(pack_pos: int, packs, folders) -> bytes:
    allp = [x for p in packs for x in p]
    o = b"\x06" + num(pack_pos) + num(len(allp)) + b"\x09" + b"".join(num(x) for x in allp) + b"\x00"
    o += b"\x07\x0B" + num(len(folders)) + b"\x00" + b"".join(f[0] for f in folders)
    o += b"\x0C" + b"".join(num(x) for f in folders for x in f[1])
    o += b"\x0A\x01" + b"".join(struct.pack("<I", f[2]) for f in folders) + b"\x00\x00"
    return o


def _solid(files: dict[str, bytes], empties: list[str] = (), dirs: list[str] = ()) -> bytes:
    data = b"".join(files.values())
    packed = lz2(data)
    sizes = [len(b) for b in files.values()]
    crcs = [zlib.crc32(b) & 0xFFFFFFFF for b in files.values()]
    head = b"\x01\x04" + streams([[len(packed)]], [(folder([coder(b"\x21", lzma2_props())]), [len(data)], None)], [(sizes, crcs)])
    names = list(files) + list(empties) + list(dirs)
    head += b"\x05" + num(len(names)) + names_prop(names)
    n = len(names)
    ne = len(empties) + len(dirs)
    if ne:
        bits = [False] * len(files) + [True] * ne
        head += b"\x0E" + _bv_prop(bits)
        head += b"\x0F" + _bv_prop([True] * len(empties) + [False] * len(dirs))
    head += b"\x00\x00"
    _ = n
    return archive(packed, head)


def _bv_prop(bits: list[bool]) -> bytes:
    o = bytearray((len(bits) + 7) // 8)
    for i, b in enumerate(bits):
        if b:
            o[i // 8] |= 0x80 >> (i % 8)
    return num(len(o)) + bytes(o)


def test_solid_empty_dirs_through_api(tmp_path):
    blob = _solid({"一/data.csv": CSV, "b.txt": TEXT[:500]}, ["空.txt"], ["フォルダ"])
    p = tmp_path / "a.7z"
    p.write_bytes(blob)
    s = api.summary(str(p))
    assert s["kind"] == "7z" and [m["name"] for m in s["members"]] == ["一/data.csv", "b.txt", "空.txt"]
    assert s["members"][0]["kind"] == "csv"
    assert get(blob, "b.txt") == TEXT[:500] and get(blob, "空.txt") == b""
    assert seven(tmp_path, blob, "一/data.csv") in (None, CSV)
    assert {f["id"] for f in api.fields(f"{p}!/一/data.csv")} >= {"R2.C2"}
    assert api.grade(str(p))["write"]["grade"] == "D"
    with pytest.raises(api.ApiError, match="7z"):
        api.fill(f"{p}!/一/data.csv", {"R2.C2": "4"}, str(tmp_path / "o.7z"))
    k, _d, parts = api.window(str(p))
    assert k == "7z" and "!/" in parts[0][1]


def test_unknown_coder_is_named():
    blob = single(TEXT, TEXT, [coder(b"\x0b")], [len(TEXT)])
    with pytest.raises(archives.ArchiveError, match="RISC-V"):
        get(blob)


def test_crc_mismatch(tmp_path):
    data = bytearray(single(TEXT, TEXT, [coder(b"\x00")], [len(TEXT)]))
    data[40] ^= 1
    with pytest.raises(archives.ArchiveError, match="CRC"):
        get(bytes(data))


def test_huge_declared_dict_and_size_do_not_allocate():
    # 宣言の大きな辞書：解くのに使う辞書は、ファイルの大きさまでにする
    props = bytes((40,))
    blob = single(TEXT, lz2(TEXT), [coder(b"\x21", props)], [len(TEXT)])
    assert get(blob) == TEXT
    # 宣言の大きな展開の大きさ：解く前に止める
    blob = single(TEXT, TEXT, [coder(b"\x21", lzma2_props())], [1 << 62])
    with pytest.raises(ValueError):
        get(blob)


@pytest.mark.parametrize("seed", range(3))
def test_corruption_never_raw(tmp_path, seed):
    r = random.Random(seed)
    base = [_solid({"a.csv": CSV, "b.txt": TEXT[:800]}),
            (FIXTURES / "ppmd_small.7z").read_bytes()]
    p = tmp_path / "c.7z"
    for blob in base:
        for t in range(50):
            b = bytearray(blob)
            if t % 3 == 0:
                b = b[:r.randrange(8, len(b))]
            else:
                for _ in range(r.randrange(1, 6)):
                    b[r.randrange(6, len(b))] = r.randrange(256)
            p.write_bytes(bytes(b))
            for path in (str(p), f"{p}!/a.csv", f"{p}!/a.bin"):
                try:
                    api.summary(path)
                except api.ApiError:
                    pass
            archives.use_password(None)


@pytest.mark.skipif(not HAS7, reason="7zz がない")
@pytest.mark.parametrize("args", [["-m0=lzma2", "-ms=on"], ["-m0=ppmd:mem=1m:o=8"], ["-mf=BCJ2"], ["-m0=arm64", "-m1=lzma2"], ["-m0=deflate64"], ["-pひみつ", "-mhe=on"]])
def test_made_by_7zz(tmp_path, args):
    d = tmp_path / "in"
    (d / "中").mkdir(parents=True)
    r = random.Random(6)
    want = {"中/x.txt": TEXT[:4000], "code.bin": _x86_like()[:8000],
            "a64.bin": b"".join((0x94000000 | r.randrange(1 << 26)).to_bytes(4, "little") if r.random() < .5 else (0x90000000 | r.randrange(1 << 24)).to_bytes(4, "little") for _ in range(1000))}
    for n, b in want.items():
        (d / n).write_bytes(b)
    out = tmp_path / "o.7z"
    subprocess.run(["7zz", "a", "-bd", "-y", str(out)] + args + ["."], cwd=d, check=True, capture_output=True, timeout=120)
    blob = out.read_bytes()
    pw = "ひみつ" if any(a.startswith("-p") for a in args) else None
    for n, b in want.items():
        assert get(blob, n, pw=pw) == b

