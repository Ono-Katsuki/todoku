"""tar.bz2・tar.xz・.bz2・.xz と、ZIP の Deflate64・BZIP2・LZMA。読むだけ。"""

from __future__ import annotations

import bz2
import io
import lzma
import struct
import tarfile
import zipfile
import zlib

import pytest

from todoku import api, archives

BODY = "申請書の本文です。\n".encode("utf-8") * 50


def _tar(mode: str) -> bytes:
    buf = io.BytesIO()
    with tarfile.open(fileobj=buf, mode=mode) as t:
        ti = tarfile.TarInfo("中/申請.txt")
        ti.size = len(BODY)
        t.addfile(ti, io.BytesIO(BODY))
    return buf.getvalue()


@pytest.mark.parametrize("mode,ext", [("w:bz2", "tar.bz2"), ("w:xz", "tar.xz")])
def test_tar_bz2_xz_read(tmp_path, mode, ext):
    p = tmp_path / f"a.{ext}"
    p.write_bytes(_tar(mode))
    win = "\n".join(t for _n, t in api.window(str(p))[2])
    assert "申請.txt" in win
    win = "\n".join(t for _n, t in api.window(f"{p}!/中/申請.txt")[2])
    assert "申請書の本文です" in win


@pytest.mark.parametrize("mode,ext", [("w:bz2", "tar.bz2"), ("w:xz", "tar.xz")])
def test_tar_bz2_xz_write_is_refused_with_the_way_forward(tmp_path, mode, ext):
    p = tmp_path / f"a.{ext}"
    p.write_bytes(_tar(mode))
    with pytest.raises(api.ApiError, match="書きません"):
        api.fill(f"{p}!/中/申請.txt", {"@1": {"old": "本文", "new": "本", "occurrence": 1}}, str(tmp_path / f"o.{ext}"))


@pytest.mark.parametrize("ext,pack", [("bz2", bz2.compress), ("xz", lzma.compress)])
def test_single_file_bz2_xz(tmp_path, ext, pack):
    p = tmp_path / f"申請.txt.{ext}"
    p.write_bytes(pack(BODY))
    assert "申請.txt" in "\n".join(t for _n, t in api.window(str(p))[2])
    assert "申請書の本文です" in "\n".join(t for _n, t in api.window(f"{p}!/申請.txt")[2])


@pytest.mark.parametrize("pack,k", [(bz2.compress, "bz2"), (lzma.compress, "xz")])
def test_concatenated_and_truncated_streams(pack, k):
    assert archives.unpack(pack(b"a") + pack(b"b"), k) == b"ab"
    whole = pack(BODY)
    with pytest.raises(archives.ArchiveError):
        archives.unpack(whole[:len(whole) // 2], k)
    with pytest.raises(archives.ArchiveError, match="読めないバイト"):
        archives.unpack(whole + b"junk", k)


def test_bomb_stops_at_the_budget(monkeypatch):
    from todoku.limits import Budget, LimitExceeded, call_budget

    with call_budget(Budget(part=1 << 20, total=1 << 20)), pytest.raises(LimitExceeded):
        archives.unpack(bz2.compress(bytes(8 << 20)), "bz2")


# ---------------------------------------------------------------- ZIP の別の圧縮
def _zip_with(method: int, data: bytes, raw: bytes, name: str = "a.txt") -> bytes:
    nm = name.encode()
    crc = zlib.crc32(raw) & 0xFFFFFFFF
    local = struct.pack("<IHHHHHIIIHH", 0x04034B50, 20, 0x800, method, 0, 0, crc, len(data), len(raw), len(nm), 0) + nm + data
    cen = struct.pack("<IHHHHHHIIIHHHHHII", 0x02014B50, 20, 20, 0x800, method, 0, 0, crc, len(data), len(raw), len(nm), 0, 0, 0, 0, 0, 0) + nm
    end = struct.pack("<IHHHHIIH", 0x06054B50, 0, 0, 1, 1, len(cen), len(local), 0)
    return local + cen + end


def _deflate_raw(raw: bytes) -> bytes:
    c = zlib.compressobj(9, zlib.DEFLATED, -15)
    return c.compress(raw) + c.flush()


def test_zip_deflate64_reads_a_stream_without_its_extended_codes(tmp_path):
    # 長さの符号 285 と距離の符号 30・31 を使わない Deflate の流れは、Deflate64 としても同じ字になる
    raw = "".join(f"行{i}\n" for i in range(50)).encode("utf-8")
    p = tmp_path / "a.zip"
    p.write_bytes(_zip_with(9, _deflate_raw(raw), raw))
    assert archives.extract(p.read_bytes(), "a.txt", "zip") == raw


def test_zip_deflate64_length_code_285_means_16_extra_bits():
    from todoku.legacy.sevenzip import deflate64

    # 固定の符号のブロック（RFC 1951）で、字「A」と、長さの符号 285・追加 16 ビット 0x0000（長さ 3）・距離の符号 0（距離 1）
    bits: list[int] = [1, 1, 0]  # 終わりのブロック・固定の符号（01 を下位から）

    def code(v: int, n: int) -> None:
        bits.extend((v >> (n - 1 - i)) & 1 for i in range(n))  # 符号は上位から

    def extra(v: int, n: int) -> None:
        bits.extend((v >> i) & 1 for i in range(n))  # 追加のビットは下位から

    code(0x30 + 65, 8)  # 字 65
    code(0xC0 + (285 - 280), 8)  # 285
    extra(0, 16)
    code(0, 5)  # 距離の符号 0
    code(0, 7)  # 終わり（256）
    data = bytearray((len(bits) + 7) // 8)
    for i, b in enumerate(bits):
        data[i // 8] |= b << (i % 8)
    assert deflate64(bytes(data), 100) == b"AAAA"


def test_zip_bzip2_and_lzma(tmp_path):
    raw = BODY
    assert archives.extract(_zip_with(12, bz2.compress(raw), raw), "a.txt", "zip") == raw
    f = {"id": lzma.FILTER_LZMA1, "dict_size": 1 << 16, "lc": 3, "lp": 0, "pb": 2}
    props = lzma._encode_filter_properties(f)
    z = struct.pack("<HH", 0x0914, len(props)) + props + lzma.compress(raw, format=lzma.FORMAT_RAW, filters=[f])
    assert archives.extract(_zip_with(14, z, raw), "a.txt", "zip") == raw


def test_zip_other_method_checks_the_crc():
    raw = b"abc"
    blob = bytearray(_zip_with(12, bz2.compress(raw), raw))
    blob[14] ^= 0xFF  # 中央の一覧にも CRC があるので、両方を変える
    i = bytes(blob).index(b"PK\x01\x02")
    blob[i + 16] ^= 0xFF
    with pytest.raises(archives.ArchiveError, match="CRC"):
        archives.extract(bytes(blob), "a.txt", "zip")
