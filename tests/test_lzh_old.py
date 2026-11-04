"""LZH の見出しのレベル 3 と、とどくが解かない方式（-lh1-〜-lh3-）：見分けて、別の道具の手順を返す。"""

from __future__ import annotations

import random
import shutil
import struct
import subprocess

import pytest

from todoku import api, archives
from todoku.book import kind_of
from todoku.jtd.lha import crc16, lh5
from todoku.legacy import lzh

R = random.Random(21)
WORDS = [bytes(R.randrange(97, 123) for _ in range(R.randrange(2, 9))) for _ in range(200)]
TEXT = b" ".join(R.choice(WORDS) for _ in range(2500))
CSV = "名前,年\n山田,3\n".encode()


def level1(name: bytes, data: bytes, method: bytes, body: bytes) -> bytes:
    h = method + struct.pack("<IIIBB", len(body), len(data), 0, 0x20, 1) + bytes((len(name),)) + name + struct.pack("<H", crc16(data)) + b"M" + struct.pack("<H", 0)
    return bytes((len(h), sum(h) & 0xFF)) + h + body


def level3(name: bytes, data: bytes, method: bytes, body: bytes, folder: bytes = b"") -> bytes:
    chain = [b"\x00\x00\x00", bytes((1,)) + name]  # 先頭は見出しの CRC-16（その欄を 0 にして全体で計算する）
    if folder:
        chain.append(bytes((2,)) + folder + b"\xff")
    raw = b""
    sizes = [len(c) + 4 for c in chain]
    for i, c in enumerate(chain):
        nxt = sizes[i + 1] if i + 1 < len(chain) else 0
        raw += c + struct.pack("<I", nxt)
    total = 32 + len(raw)
    h = b"\x04\x00" + method + struct.pack("<IIIBB", len(body), len(data), 0, 0x20, 3) + struct.pack("<HB", crc16(data), 0x4D)
    h += struct.pack("<II", total, sizes[0]) + raw
    crc = crc16(h)
    h = h[:33] + struct.pack("<H", crc) + h[35:]
    return h + body


@pytest.mark.parametrize("method", [b"-lh1-", b"-lh2-", b"-lh3-"])
def test_old_methods_guide_to_another_tool(tmp_path, method):
    """解かないので、中身は何でもよい（方式の印で決める）。"""
    blob = level1(b"a.csv", CSV, method, b"\x55" * 20) + level1(b"b.csv", CSV, b"-lh0-", CSV) + b"\0"
    assert kind_of(blob) == "lzh"
    p = tmp_path / "古い.lzh"
    p.write_bytes(blob)
    s = api.summary(str(p))
    step = s["needs_tool"]["next_step"]
    text = " ".join(c["command"] for c in step["commands"])
    assert all(t in text for t in ("lha x", "unar", "7z x", "bsdtar"))
    assert step["after"].endswith("もう一度 open_document で開いてください") and "元の形式には戻りません" in step["writes"]
    with pytest.raises(api.ApiError) as e:
        api.window(f"{p}!/a.csv")
    assert method.decode() in str(e.value) and "lha x" in str(e.value)
    g = api.grade(str(p))
    assert g["grade"] == "E" and g["write"]["grade"] == "D"
    assert api.audit(str(p))["needs_agent"]
    assert api.fields(f"{p}!/b.csv")  # 解ける方式のファイルは、同じ書庫でも読める
    with pytest.raises(archives.ArchiveError):
        archives.extract(blob, "a.csv", "lzh", "a.lzh")


def test_readable_methods_have_no_guide(tmp_path):
    p = tmp_path / "a.lzh"
    p.write_bytes(level1(b"b.csv", CSV, b"-lh0-", CSV) + b"\0")
    assert "needs_tool" not in api.summary(str(p))


def test_no_adaptive_decoder_remains():
    for name in ("_Dyn", "undyn", "ready_made", "_Static3"):
        assert not hasattr(lzh, name)


def test_level3_header(tmp_path):
    data = TEXT[:3000]
    blob = level3("名前.txt".encode("cp932"), data, b"-lh5-", lh5(data), "下".encode("cp932")) + level3(b"b.txt", b"plain", b"-lh0-", b"plain") + b"\0"
    assert kind_of(blob, "a.lzh") == "lzh"
    es = lzh.entries(blob)
    assert [e.name for e in es] == ["下/名前.txt", "b.txt"]
    assert lzh.extract(blob, es[0]) == data and lzh.extract(blob, es[1]) == b"plain"
    if shutil.which("bsdtar"):  # libarchive はレベル 3 を読む（7zz は読まない）。名前は ASCII で照らす
        ascii_blob = level3(b"n.txt", data, b"-lh5-", lh5(data), b"sub") + level3(b"b.txt", b"plain", b"-lh0-", b"plain") + b"\0"
        assert [e.name for e in lzh.entries(ascii_blob)] == ["sub/n.txt", "b.txt"]
        p = tmp_path / "t.lzh"
        p.write_bytes(ascii_blob)
        r = subprocess.run(["bsdtar", "-xOf", str(p), "sub/n.txt"], capture_output=True, timeout=30, stdin=subprocess.DEVNULL)
        assert r.returncode == 0 and r.stdout == data, r.stderr
