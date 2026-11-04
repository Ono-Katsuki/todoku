"""壊れたファイル：どの形式でも、わけを字で伝える例外（ValueError の仲間）か段階で返す。IndexError などの生の例外を出さない。"""

from __future__ import annotations

import gzip
import io
import random
import tarfile

import jet_maker
import pytest
import test_archives as ta
import test_more_formats as tm
import xlsb_maker

from todoku import api


def _tar() -> bytes:
    return ta._tar({"a.csv": ta.CSV})


SAMPLES = {
    "a.xlsb": xlsb_maker.sample,
    "a.ppt": tm._ppt,
    "a.xps": lambda: tm._xps(tm.PAGE),
    "a.mdb": jet_maker.sample,
    "a.dbf": tm._dbf,
    "b2.xls": tm._biff2,
    "b4.xls": lambda: tm._biff34(4),
    "a.doc": lambda: tm._word2(b"abc\r"),
    "a.tar": _tar,
    "a.tgz": lambda: gzip.compress(_tar(), mtime=0),
    "a.lzh": lambda: ta._lzh_level0(b"a.csv", ta.CSV) + b"\x00",
    "a.msg": ta._msg,
    "a.eml": ta._eml,
    "a.sxw": lambda: tm._zip({"mimetype": "application/vnd.sun.xml.writer", "content.xml": tm._SXW}),
    "a.pages": lambda: tm._zip({"Index/Document.iwa": tm._iwa([(3, 2001, tm._f(1, 0) + tm._f(3, "本文".encode()))])}),
}


@pytest.mark.parametrize("name", sorted(SAMPLES))
def test_broken_files_never_raise_raw_errors(tmp_path, name):
    blob = SAMPLES[name]()
    rnd = random.Random(name)
    cases = [blob[:len(blob) // 2], blob[:len(blob) * 9 // 10], blob[:40]]
    for _ in range(12):
        b = bytearray(blob)
        for _k in range(8):
            b[rnd.randrange(len(b))] = rnd.randrange(256)
        cases.append(bytes(b))
    p = tmp_path / name
    for c in cases:
        p.write_bytes(c)
        try:
            api.window(str(p))
        except ValueError:  # ApiError と、壊れた ZIP の ZipFormatError：わけを字で伝える例外
            pass
        try:
            assert api.grade(str(p))["grade"] in "SABCDEF"
        except ValueError:
            pass


def test_tar_member_paths_are_only_names(tmp_path):
    """中の名前の「..」を、ファイルの置き場所として使わない（書き込みの一時ファイルも名前だけ）。"""
    buf = io.BytesIO()
    with tarfile.open(fileobj=buf, mode="w") as t:
        ti = tarfile.TarInfo("../../evil.csv")
        ti.size = len(ta.CSV)
        t.addfile(ti, io.BytesIO(ta.CSV))
    p = tmp_path / "a.tar"
    p.write_bytes(buf.getvalue())
    out = tmp_path / "out.tar"
    r = api.fill(f"{p}!/../../evil.csv", {"R2.C2": "9"}, str(out))
    assert r["proof_ok"] and not (tmp_path.parent.parent / "evil.csv").exists()
