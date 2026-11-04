"""とどくが解かない形式（RAR・HEIC・AVIF）：見分けて、別の道具の手順を構造で返す。とどくは外の道具を呼ばない。"""

from __future__ import annotations

import ast
import struct
import zipfile
from pathlib import Path

import pytest

from todoku import api
from todoku.book import kind_of
from todoku.imaging import formats as F
from todoku.legacy import convert_first as C


def _ftyp(major: bytes, *compat: bytes) -> bytes:
    body = major + b"\0\0\0\0" + b"".join(compat)
    return struct.pack(">I", 8 + len(body)) + b"ftyp" + body + b"\0" * 64


SAMPLES = {"a.rar": ("rar", b"Rar!\x1a\x07\x01\x00" + bytes(40)),
           "r4.rar": ("rar", b"Rar!\x1a\x07\x00" + bytes(40)),
           "a.heic": ("heic", _ftyp(b"heic", b"mif1", b"heic")),
           "a.heif": ("heic", _ftyp(b"mif1", b"mif1")),
           "a.avif": ("avif", _ftyp(b"avif", b"mif1", b"avif"))}


@pytest.mark.parametrize("name", SAMPLES)
def test_open_says_format_why_and_next_step(tmp_path, name):
    kind, blob = SAMPLES[name]
    p = tmp_path / name
    p.write_bytes(blob)
    assert kind_of(blob, name) == kind
    s = api.summary(str(p))
    nt = s["needs_tool"]
    assert nt["format"] == C.FORMAT[kind] and nt["why"] == C.WHY[kind]
    step = nt["next_step"]
    assert step["after"] == "変換・展開してできたファイルを、もう一度 open_document で開いてください"
    assert step["writes"] == "書いたものは元の形式には戻りません（PNG や展開したファイルに対して書きます）。本人に伝えてください"
    assert all({"os", "command", "note"} <= set(c) for c in step["commands"])
    text = " ".join(c["command"] for c in step["commands"])
    if kind == "rar":
        assert nt["message"] == "RAR の中身は別の道具で展開してください（展開したフォルダか ZIP なら読めます）"
        assert all(t in text for t in ("bsdtar -xf", "unar", "unrar x", "7z x"))
        assert text.index("bsdtar") < text.index("unrar")
    else:
        assert all(t in text for t in ("sips -s format png", "magick", "ffmpeg -i", "heif-convert"))
        assert name in text and f"{Path(name).stem}.png" in text
    _k, _doc, parts = api.window(str(p))
    md = "".join(t for _h, t in parts)
    assert step["after"] in md and step["writes"] in md and step["commands"][0]["command"] in md


@pytest.mark.parametrize("name", SAMPLES)
def test_grade_is_read_e_write_d(tmp_path, name):
    p = tmp_path / name
    p.write_bytes(SAMPLES[name][1])
    g = api.grade(str(p))
    assert g["grade"] == "E" and g["write"]["grade"] == "D"
    r = g["reasons"][0]
    assert r["title"] == "別の道具で変換してから読む"
    assert "open_document" in r["detail"]
    assert api.audit(str(p))["ok"]


@pytest.mark.parametrize("name", SAMPLES)
def test_view_and_write_are_refused_with_the_next_step(tmp_path, name):
    p = tmp_path / name
    p.write_bytes(SAMPLES[name][1])
    with pytest.raises(api.ApiError) as e:
        api.view_page(str(p), 1)
    assert "open_document" in str(e.value)
    with pytest.raises(api.ApiError) as e:
        api.to_pdf(str(p), str(tmp_path / "o.pdf"))
    assert "open_document" in str(e.value)
    with pytest.raises(api.ApiError) as e:
        api.fill(str(p), {"x": "1"}, str(tmp_path / ("o" + p.suffix)))
    assert "元の形式には戻りません" in str(e.value)


def test_mcp_open_document_carries_the_structure(tmp_path):
    from todoku import mcp_server

    p = tmp_path / "a.heic"
    p.write_bytes(SAMPLES["a.heic"][1])
    r = mcp_server.call_tool("open_document", {"path": str(p)})
    assert r["structuredContent"]["needs_tool"]["next_step"]["commands"]
    assert "sips -s format png a.heic --out a.png" in r["content"][1]["text"]


def test_member_inside_rar_says_extract_first(tmp_path):
    p = tmp_path / "a.rar"
    p.write_bytes(SAMPLES["a.rar"][1])
    with pytest.raises(api.ApiError) as e:
        api.window(f"{p}!/中.xlsx")
    assert "RAR の中身は別の道具で展開してください" in str(e.value)


def test_member_of_zip_goes_the_same_way(tmp_path):
    p = tmp_path / "a.zip"
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("写真.heic", SAMPLES["a.heic"][1])
    s = api.summary(f"{p}!/写真.heic")
    assert "sips -s format png '写真.heic' --out '写真.png'" in s["needs_tool"]["next_step"]["commands"][0]["command"]


def test_decoding_routes_say_the_next_step_for_embedded_images():
    for blob in (SAMPLES["a.heic"][1], SAMPLES["a.avif"][1]):
        with pytest.raises(F.ImageError) as e:
            F.decode(blob)
        assert "別の道具" in str(e.value)


def test_file_names_are_quoted_in_commands():
    cmd = C.needs_tool("heic", "写真 1;rm -rf x.heic")["next_step"]["commands"][0]["command"]
    assert cmd == "sips -s format png '写真 1;rm -rf x.heic' --out '写真 1;rm -rf x.png'"


def test_todoku_does_not_call_any_conversion_tool_itself():
    """手順を返すだけ：この部品は subprocess も os.system も使わない。"""
    src = Path(C.__file__).read_text(encoding="utf-8")
    names = {a.name.split(".")[0] for n in ast.walk(ast.parse(src)) if isinstance(n, ast.Import) for a in n.names}
    assert not names & {"subprocess", "os"}
