"""tar・gzip・LZH・7z/RAR の見分け・メール：中のファイルを「a!/名前」で読む・書く（書いた所のほかはバイト一致）。"""

from __future__ import annotations

import gzip
import io
import struct
import tarfile
from email.message import EmailMessage

import pytest

from todoku import api, archives
from todoku.book import kind_of
from todoku.jtd.lha import crc16, lh5
from todoku.legacy import lzh

CSV = "名前,年\n山田,3\n".encode()


def _tar(files: dict, fmt=tarfile.USTAR_FORMAT) -> bytes:
    buf = io.BytesIO()
    with tarfile.open(fileobj=buf, mode="w", format=fmt, encoding="utf-8") as t:
        for n, b in files.items():
            ti = tarfile.TarInfo(n)
            ti.size = len(b)
            ti.mtime = 1000000000
            ti.uname = "ono"
            t.addfile(ti, io.BytesIO(b))
    return buf.getvalue()


def _p(tmp_path, name, blob):
    p = tmp_path / name
    p.write_bytes(blob)
    return str(p)


def test_tar_read_write_keeps_other_members(tmp_path):
    blob = _tar({"一/data.csv": CSV, "b.txt": "別のファイル\n".encode(), "c.csv": b"x\n1\n"})
    assert kind_of(blob) == "tar"
    p = _p(tmp_path, "a.tar", blob)
    s = api.summary(p)
    assert s["kind"] == "tar" and [m["name"] for m in s["members"]] == ["一/data.csv", "b.txt", "c.csv"]
    assert api.fields(p) == []
    assert {f["id"] for f in api.fields(f"{p}!/一/data.csv")} >= {"R2.C2"}
    out = tmp_path / "b.tar"
    r = api.fill(f"{p}!/一/data.csv", {"R2.C2": "4000 長い値にして大きさが変わる" * 40}, str(out))
    assert r["proof_ok"] and r["read_back_ok"]
    new = out.read_bytes()
    with tarfile.open(fileobj=io.BytesIO(new)) as t:
        assert t.getnames() == ["一/data.csv", "b.txt", "c.csv"]
        assert t.extractfile("b.txt").read() == "別のファイル\n".encode() and t.extractfile("c.csv").read() == b"x\n1\n"
        assert b"4000" in t.extractfile("一/data.csv").read()
        assert t.getmember("一/data.csv").uname == "ono" and t.getmember("一/data.csv").mtime == 1000000000
    assert blob[-1024:] == new[-1024:]
    assert archives.unchanged_except(blob, new, "一/data.csv", "tar") == []
    assert api.audit(p)["ok"] and api.grade(p)["write"]["grade"] != "D"


def test_tar_cp932_names_and_gnu_long_names(tmp_path):
    longname = "d/" + "x" * 150 + ".csv"
    p = _p(tmp_path, "a.tar", _tar({longname: CSV}, tarfile.GNU_FORMAT))
    assert api.summary(p)["members"][0]["name"] == longname
    assert len(api.fields(f"{p}!/{longname}")) >= 1


def test_tgz_is_a_tar_and_write_keeps_gz_header(tmp_path):
    inner = _tar({"data.csv": CSV, "keep.csv": b"a\n"})
    buf = io.BytesIO()
    with gzip.GzipFile(filename="orig.tar", mode="wb", fileobj=buf, mtime=12345) as g:
        g.write(inner)
    blob = buf.getvalue()
    assert kind_of(blob) == "tar"
    p = _p(tmp_path, "a.tgz", blob)
    assert [m["name"] for m in api.summary(p)["members"]] == ["data.csv", "keep.csv"]
    out = tmp_path / "b.tgz"
    r = api.fill(f"{p}!/data.csv", {"R2.C2": "4"}, str(out))
    assert r["proof_ok"] and r["read_back_ok"]
    new = out.read_bytes()
    assert new[:20] == blob[:20]  # 見出し（時刻・名前）はそのまま
    with tarfile.open(fileobj=io.BytesIO(gzip.decompress(new))) as t:
        assert t.extractfile("data.csv").read() == "名前,年\n山田,4\n".encode()
        assert t.extractfile("keep.csv").read() == b"a\n"


def test_gz_single_member(tmp_path):
    buf = io.BytesIO()
    with gzip.GzipFile(filename="統計.csv".encode("cp932").decode("latin-1"), mode="wb", fileobj=buf, mtime=7) as g:
        g.write(CSV)
    blob = buf.getvalue()
    assert kind_of(blob) == "gz"
    p = _p(tmp_path, "x.gz", blob)
    assert api.summary(p)["members"][0]["name"] == "統計.csv"
    assert api.fields(f"{p}!/統計.csv")
    out = tmp_path / "y.gz"
    r = api.fill(f"{p}!/統計.csv", {"R2.C2": "9"}, str(out))
    assert r["proof_ok"] and r["read_back_ok"]
    assert gzip.decompress(out.read_bytes()) == "名前,年\n山田,9\n".encode()
    # FNAME がなければ、ファイルの名前から
    p2 = _p(tmp_path, "y.csv.gz", gzip.compress(CSV, mtime=0))
    assert api.summary(p2)["members"][0]["name"] == "y.csv"
    # 2 つの部分をつないだ gz：読むが、書かない
    p3 = _p(tmp_path, "z.csv.gz", gzip.compress(CSV, mtime=0) + gzip.compress(b"x\n", mtime=0))
    assert len(api.fields(f"{p3}!/z.csv")) >= 1
    with pytest.raises(api.ApiError, match="つないだ"):
        api.fill(f"{p3}!/z.csv", {"R2.C2": "9"}, str(tmp_path / "z2.gz"))
    # 壊れた gz は読まない
    bad = bytearray(blob)
    bad[-5] ^= 0xFF
    with pytest.raises(api.ApiError):
        api.fields(_p(tmp_path, "bad.gz", bytes(bad)) + "!/統計.csv")


# ---------------------------------------------------------------------- LZH
def _lzh_level0(name: bytes, data: bytes, method=b"-lh5-", body=None) -> bytes:
    body = lh5(data) if body is None else body
    h = method + struct.pack("<IIIBB", len(body), len(data), 0, 0x20, 0) + bytes([len(name)]) + name + struct.pack("<H", crc16(data))
    return bytes([len(h), sum(h) & 0xFF]) + h + body


def _lzh_level1(name: bytes, data: bytes, folder: bytes = b"", method=b"-lh0-") -> bytes:
    body = data
    ext = bytes([0x02]) + folder + struct.pack("<H", 0) if folder else b""
    extsize = struct.pack("<H", len(ext)) if folder else struct.pack("<H", 0)
    h = (method + struct.pack("<IIIBB", len(ext) + len(body), len(data), 0, 0x20, 1) + bytes([len(name)]) + name
         + struct.pack("<H", crc16(data)) + b"M" + extsize)
    return bytes([len(h), sum(h) & 0xFF]) + h + ext + body


def _lzh_level2(name: bytes, data: bytes, method=b"-lh5-") -> bytes:
    body = lh5(data) if method == b"-lh5-" else data
    ext_name = bytes([0x01]) + name + struct.pack("<H", 0)
    first = struct.pack("<H", len(ext_name))
    total = 26 + len(ext_name)
    h = struct.pack("<H", total) + method + struct.pack("<IIIBB", len(body), len(data), 0, 0x20, 2) + struct.pack("<HB", crc16(data), 0x4D) + first + ext_name
    return h + body


def test_lzh_levels_and_shift_jis_names(tmp_path):
    long = ("役所の統計 " * 400).encode("cp932")
    blob = (_lzh_level0("申請書.csv".encode("cp932"), CSV) + _lzh_level1(b"b.txt", long, folder=b"sub\xff")
            + _lzh_level2("統計.txt".encode("cp932"), long) + b"\x00")
    assert kind_of(blob) == "lzh"
    p = _p(tmp_path, "a.lzh", blob)
    s = api.summary(p)
    assert [m["name"] for m in s["members"]] == ["申請書.csv", "sub/b.txt", "統計.txt"]
    assert "Shift_JIS（推定）" in "\n".join(t for _n, t in api.window(p)[2])
    assert {f["id"] for f in api.fields(f"{p}!/申請書.csv")} >= {"R2.C2"}
    assert "役所の統計" in "\n".join(t for _n, t in api.window(f"{p}!/統計.txt")[2])
    g = api.grade(p)
    assert g["write"]["grade"] == "D" and any("書かない" in r["title"] for r in g["write"]["reasons"])
    with pytest.raises(api.ApiError, match="LZH"):
        api.fill(f"{p}!/申請書.csv", {"R2.C2": "1"}, str(tmp_path / "b.lzh"))


def test_lzh_rejects_bad_crc_and_unknown_method(tmp_path):
    good = _lzh_level0(b"a.csv", CSV, method=b"-lh0-", body=CSV)
    bad = bytearray(good)
    bad[-1] ^= 1
    with pytest.raises(api.ApiError, match="CRC"):
        api.fields(_p(tmp_path, "b.lzh", bytes(bad)) + "!/a.csv")
    p = _p(tmp_path, "c.lzh", _lzh_level0(b"a.csv", CSV, method=b"-lzs-", body=CSV))
    with pytest.raises(api.ApiError, match="-lzs-"):
        api.fields(p + "!/a.csv")
    assert api.grade(p)["grade"] == "F"


def _bits(*fields) -> bytes:
    s = "".join(format(v, f"0{n}b") for v, n in fields)
    s += "0" * (-len(s) % 8)
    return bytes(int(s[i:i + 8], 2) for i in range(0, len(s), 8))


@pytest.mark.parametrize("method,pbit", [("-lh4-", 4), ("-lh5-", 4), ("-lh6-", 5), ("-lh7-", 5)])
def test_lh_methods_differ_only_in_position_code_width(method, pbit):
    """1 記号だけの表のブロック（n=0）を手で作る：字 2 つのあと、長さ 8・距離 2 の参照。"""
    def block(count, sym, p_only):
        return [(count, 16), (0, 5), (0, 5), (0, 9), (sym, 9), (0, pbit), (p_only, pbit)]
    stream = _bits(*block(1, 65, 0), *block(1, 66, 0), *block(1, 256 + 8 - 3, 1))
    assert lzh.unlh(stream, 10, method) == b"ABABABABAB"


def test_self_extracting_exe_is_read_without_running(tmp_path):
    blob = b"MZ" + bytes(5000) + _lzh_level0(b"a.csv", CSV) + b"\x00"
    assert kind_of(blob, "setup.exe") == "lzh"
    p = _p(tmp_path, "setup.exe", blob)
    assert api.summary(p)["members"][0]["name"] == "a.csv"
    assert kind_of(blob, "other.bin") != "lzh"


def test_lzh_audit_goes_through_members(tmp_path):
    p = _p(tmp_path, "a.lzh", _lzh_level0(b"a.csv", CSV) + b"\x00")
    au = api.audit(p)
    assert au["checks"] and all(c["name"].startswith("a.csv：") for c in au["checks"])


@pytest.mark.parametrize("blob,kind", [(b"7z\xbc\xaf\x27\x1c" + bytes(30), "7z"), (b"Rar!\x1a\x07\x00" + bytes(30), "rar")])
def test_broken_7z_and_rar_say_so(tmp_path, blob, kind):
    assert kind_of(blob) == kind
    p = _p(tmp_path, f"a.{kind}", blob)
    with pytest.raises(api.ApiError, match="壊れて"):
        api.window(p)
    g = api.grade(p)
    assert g["grade"] == "F" and "壊れて" in str(g["reasons"])


# ---------------------------------------------------------------------- メール
def _eml() -> bytes:
    m = EmailMessage()
    m["From"] = "山田 <yamada@example.jp>"
    m["To"] = "b@example.jp"
    m["Subject"] = "申請について"
    m.set_content("本文の一行目\n二行目", charset="iso-2022-jp")
    m.add_attachment(CSV, maintype="text", subtype="csv", filename="様式.csv")
    return m.as_bytes()


def test_eml_body_and_attachment_as_member(tmp_path):
    blob = _eml()
    assert kind_of(blob, "a.eml") == "eml"
    p = _p(tmp_path, "a.eml", blob)
    k, doc, parts = api.window(p)
    t = "\n".join(x for _n, x in parts)
    assert k == "eml" and "申請について" in t and "本文の一行目" in t and "様式.csv" in t and "yamada@example.jp" in t
    assert {f["id"] for f in api.fields(f"{p}!/様式.csv")} >= {"R2.C2"}
    au = api.audit(p)
    assert au["ok"] and any(c["name"] == "読み落とし（本文）" for c in au["checks"])
    g = api.grade(p)
    assert g["write"]["grade"] == "D" and g["grade"] == "B"
    with pytest.raises(api.ApiError, match="メール"):
        api.fill(f"{p}!/様式.csv", {"R2.C2": "1"}, str(tmp_path / "b.eml"))


def test_eml_without_attachment_and_html_only(tmp_path):
    m = EmailMessage()
    m["Subject"] = "件名"
    m.set_content("<html><body><p>HTML の本文</p><script>x()</script></body></html>", subtype="html")
    p = _p(tmp_path, "a.eml", m.as_bytes())
    t = "\n".join(x for _n, x in api.window(p)[2])
    assert "HTML の本文" in t and "x()" not in t and "タグの間の字" in t
    assert api.grade(p)["grade"] == "B"


def _msg() -> bytes:
    from test_embeds import make_cfb

    u16 = lambda s: s.encode("utf-16-le")  # noqa: E731
    return make_cfb({"__properties_version1.0": bytes(32),
                     "__substg1.0_0037001F": u16("件名の字"),
                     "__substg1.0_0C1A001F": u16("山田"),
                     "__substg1.0_0E04001F": u16("宛先の人"),
                     "__substg1.0_1000001F": u16("本文の字\r\n二行目"),
                     "__attach_version1.0_#00000000": {"__substg1.0_3707001F": u16("様式.csv"), "__substg1.0_37010102": CSV}})


def test_msg_body_and_attachment(tmp_path):
    blob = _msg()
    assert kind_of(blob) == "msg"
    p = _p(tmp_path, "a.msg", blob)
    t = "\n".join(x for _n, x in api.window(p)[2])
    assert "件名の字" in t and "本文の字" in t and "様式.csv" in t and "CSV" in t
    assert {f["id"] for f in api.fields(f"{p}!/様式.csv")} >= {"R2.C2"}
    assert api.audit(p)["ok"] and api.grade(p)["write"]["grade"] == "D"
