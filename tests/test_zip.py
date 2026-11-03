"""ZIP：中のファイルを「a.zip!/中の名前」で読む・書く（書いた部品のほかはバイト一致）・段階・名前の文字コード。"""

from __future__ import annotations

import io
import struct
import zipfile

import pytest

from todoku import api, ziparc


def _zip(files: dict[str, bytes | str]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in files.items():
            z.writestr(n, b)
    return buf.getvalue()


def test_member_is_read_and_written_like_the_file_itself(tmp_path):
    src = tmp_path / "a.zip"
    src.write_bytes(_zip({"data.csv": "名前,年\n山田,3\n", "meta.csv": "x\n1\n"}))
    assert api.summary(str(src))["kind"] == "zip" and api.fields(str(src)) == []
    assert {f["id"] for f in api.fields(f"{src}!/data.csv")} >= {"R2.C2"}
    out = tmp_path / "b.zip"
    r = api.fill(f"{src}!/data.csv", {"R2.C2": "4"}, str(out))
    assert r["proof_ok"] and r["read_back_ok"]
    assert zipfile.ZipFile(out).read("data.csv").decode() == "名前,年\n山田,4\n"
    assert ziparc.unchanged_except(src.read_bytes(), out.read_bytes(), "data.csv") == []
    with pytest.raises(api.ApiError):
        api.fill(str(src), {"R2.C2": "4"}, str(tmp_path / "c.zip"))  # ZIP そのものには書かない


def test_zip_grade_is_the_lowest_member():
    from todoku.grade import grade_blob

    g = grade_blob(_zip({"a.csv": "x\n1\n", "b.csv": "y\n2\n"}), "a.zip")
    assert g.grade == "S" and g.write == "A", g.headline()
    g = grade_blob(_zip({"a.csv": "x\n1\n", "readme.bin": b"\x00\x01"}), "a.zip")
    assert g.grade == "F" and any(r.title.startswith("readme.bin：") for r in g.reasons), g.headline()


def test_nested_zip_member(tmp_path):
    src = tmp_path / "a.zip"
    src.write_bytes(_zip({"in.zip": _zip({"d.csv": "a\n1\n"})}))
    out = tmp_path / "b.zip"
    api.fill(f"{src}!/in.zip!/d.csv", {"R2.C1": "2"}, str(out))
    inner = zipfile.ZipFile(io.BytesIO(zipfile.ZipFile(out).read("in.zip"))).read("d.csv")
    assert inner == b"a\n2\n"


def _flags(blob: bytes, set_bits: int = 0, clear_bits: int = 0) -> bytes:
    """ZIP のすべての部品の印（ローカルヘッダと中央ディレクトリ）を変える。"""
    b = bytearray(blob)
    for sig, at in ((b"PK\x03\x04", 6), (b"PK\x01\x02", 8)):
        i = b.find(sig)
        while i >= 0:
            f = struct.unpack_from("<H", b, i + at)[0]
            struct.pack_into("<H", b, i + at, (f | set_bits) & ~clear_bits)
            i = b.find(sig, i + 4)
    return bytes(b)


def test_shift_jis_names_are_read_and_marked_as_guessed():
    # 日本の ZIP によくある、UTF-8 の印のない Shift_JIS の名前（同じバイト数の名前に置き換えて作る）
    sj = "申請書.csv".encode("cp932")  # 10 バイト
    blob = _zip({"abcdefghij": "a\n1\n"}).replace(b"abcdefghij", sj)
    m = ziparc.files(blob)[0]
    assert m.name == "申請書.csv" and "推定" in m.named
    assert ziparc.extract(blob, "申請書.csv") == b"a\n1\n"


def test_encrypted_and_bomb_members_are_not_extracted():
    with pytest.raises(ziparc.ArchiveError, match="暗号化"):
        ziparc.extract(_flags(_zip({"a.csv": "x\n"}), set_bits=0x1), "a.csv")
    blob = _zip({"big.csv": "0" * 5_000_000})
    with pytest.raises(ziparc.ArchiveError, match="圧縮率"):
        ziparc.extract(blob, "big.csv")  # 展開して大きくなる ZIP
    with pytest.raises(ziparc.ArchiveError, match="ありません"):
        ziparc.extract(blob, "none.csv")
