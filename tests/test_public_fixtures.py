"""コミットした架空の書類（tests/fixtures/public/）で、ファイルから開く本物の道筋（api：窓・欄・書き込み・検査・段階）を通す。

書類は scripts/fixtures/make_public.py が作った（個人情報なし・書体のデータなし、CC0）。正解（窓に出るはずの字・欄の名前・
書く値）は、作った側が入れたものを MANIFEST.json に書いてある（とどくの出力に合わせて決めたものではない）。
corpus/ がなくても（CI でも）動く。
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

from todoku import api
from todoku.grade import grade_blob

PUBLIC = Path(__file__).parent / "fixtures" / "public"
MANIFEST = json.loads((PUBLIC / "MANIFEST.json").read_text(encoding="utf-8"))
ENTRIES = MANIFEST["files"]
BY_NAME = {e["file"]: e for e in ENTRIES}
LIMIT = 100 * 1024

# 見つかっている不具合（直ったら XPASS になる。strict ではないので、直した人のテストは落ちない）
KNOWN: dict = {}


def _ids(entries, key):
    return [pytest.param(e, id=e["file"]) for e in entries if e.get(key)]


def _mark(kind: str, e: dict):
    why = KNOWN.get((kind, e["file"]))
    return pytest.param(e, id=e["file"], marks=pytest.mark.xfail(reason=why, strict=False)) if why else pytest.param(e, id=e["file"])


def _path(e: dict) -> str:
    return str(PUBLIC / e["file"])


def _markdown(e: dict) -> tuple[str, dict, str]:
    k, doc, parts = api.window(_path(e), password=e.get("password"), encoding=e.get("encoding"))
    return k, doc, "".join(text for _name, text in parts)


# ====================================================================== ファイルそのもの
def test_manifest_lists_every_file_and_each_is_small():
    on_disk = {p.name for p in PUBLIC.iterdir() if p.is_file()} - {"MANIFEST.json", "README.md"}
    assert on_disk == set(BY_NAME), on_disk ^ set(BY_NAME)
    big = {p.name: p.stat().st_size for p in PUBLIC.iterdir() if p.stat().st_size >= LIMIT}
    assert not big, big


def test_pdf_generators_reproduce_committed_files():
    """PDF は標準ライブラリだけの生成器で作ったので、作り直すとバイトまで同じ（コミットしたものが生成器の出力そのもの）。"""
    gen = Path(__file__).resolve().parent.parent / "scripts" / "fixtures"
    sys.path.insert(0, str(gen))
    try:
        import make_pdfs
    finally:
        sys.path.remove(str(gen))
    for name, fn in make_pdfs.ALL.items():
        assert fn() == (PUBLIC / name).read_bytes(), name


# ====================================================================== 窓（読む）
@pytest.mark.parametrize("e", [pytest.param(e, id=e["file"]) for e in ENTRIES])
def test_window_has_the_written_text(e):
    k, doc, md = _markdown(e)
    assert k == e["kind"]
    missing = [t for t in e.get("texts", []) if t not in md]
    assert not missing, missing
    if e.get("order"):
        at = [md.find(t) for t in e["order"]]
        assert at == sorted(at), list(zip(e["order"], at))
    if e.get("pages"):
        assert len(doc["pages"]) == e["pages"]
    for m in e.get("members", []):
        assert m in md


@pytest.mark.parametrize("e", _ids(ENTRIES, "fields"))
def test_fields_named_in_the_document(e):
    ids = {f["id"] for f in api.fields(_path(e), password=e.get("password"))}
    assert set(e["fields"]) <= ids, set(e["fields"]) - ids


def test_xlsx_list_validation_options():
    e = BY_NAME["sheet.xlsx"]
    by = {f["id"]: f for f in api.fields(_path(e))}
    for fid, opts in e["options"].items():
        assert by[fid].get("options") == opts


def test_image_only_pdf_is_not_read_as_text():
    e = BY_NAME["scanned.pdf"]
    g = grade_blob((PUBLIC / e["file"]).read_bytes(), e["file"]).to_dict()
    assert g["grade"] not in ("S", "A", "B"), g["headline"]
    assert api.fields(_path(e)) == []


def test_encrypted_pdf_needs_the_password():
    e = BY_NAME["encrypted.pdf"]
    with pytest.raises(api.PasswordNeeded):
        api.window(_path(e))
    with pytest.raises(api.ApiError):
        api.window(_path(e), password="wrong")


# ====================================================================== 書き込み（証明と読み戻し）
def _fill_values(e: dict) -> dict:
    values = dict(e.get("fill") or {})
    if e.get("fill_by_label"):
        fields = api.fields(_path(e), password=e.get("password"), encoding=e.get("encoding"))
        for label, value in e["fill_by_label"].items():
            hit = [f["id"] for f in fields if f.get("label") == label and not f.get("agent")]
            assert hit, (label, [(f["id"], f.get("label")) for f in fields])
            values[hit[0]] = value
    return values


def _need_fonts(e: dict) -> None:
    from todoku.pdf import shape

    for fid, script in (e.get("needs") or {}).items():
        if shape.font_for(script) is None:
            pytest.skip(f"このパソコンに {script} を組み立てられる書体がない（欄 {fid}）")


FILLS = [e for e in ENTRIES if e.get("fill") or e.get("fill_by_label")]


def _fill_param(e: dict):
    marks = [pytest.mark.fonts] if e.get("needs") else []
    if KNOWN.get(("fill", e["file"])):
        marks.append(pytest.mark.xfail(reason=KNOWN[("fill", e["file"])], strict=False))
    return pytest.param(e, id=e["file"], marks=marks)


@pytest.mark.parametrize("e", [_fill_param(e) for e in FILLS])
def test_fill_proves_and_reads_back(e, tmp_path):
    _need_fonts(e)
    src = PUBLIC / e["file"]
    before = src.read_bytes()
    values = _fill_values(e)
    out = tmp_path / ("out" + src.suffix)
    r = api.fill(str(src), values, str(out), password=e.get("password"), encoding=e.get("encoding"))
    assert r["proof_ok"] and r["read_back_ok"], r
    assert src.read_bytes() == before  # 元のファイルは変えない
    # 書き出したファイルを開き直すと、書いた字がそのまま窓に出る（推定した空のマスは、書くと空でなくなるので欄の一覧からは消える）
    _k, _doc, parts = api.window(str(out), password=e.get("password"), encoding=e.get("encoding"))
    md = "".join(text for _name, text in parts)
    for fid, v in values.items():
        if isinstance(v, str):
            assert v in md, (fid, v)
    if e.get("password"):
        with pytest.raises(api.PasswordNeeded):  # 書き出したファイルも、同じパスワードで保護されたまま
            api.window(str(out))


def test_fill_member_of_zip(tmp_path):
    e = BY_NAME["bundle.zip"]
    for member, values in e["member_fill"].items():
        out = tmp_path / "out.zip"
        r = api.fill(f"{_path(e)}!/{member}", values, str(out))
        assert r["proof_ok"] and r["read_back_ok"], r
        got = {f["id"]: f.get("value") for f in api.fields(f"{out}!/{member}")}
        assert all(got.get(k) == v for k, v in values.items()), got


# ====================================================================== 検査・段階
@pytest.mark.parametrize("e", [_mark("audit", e) for e in ENTRIES])
def test_audit_passes(e):
    au = api.audit(_path(e), password=e.get("password"), encoding=e.get("encoding"))
    assert au["ok"], [(c["name"], c.get("detail"), c.get("items", [])[:3]) for c in au["checks"] if not c["ok"]]


@pytest.mark.parametrize("e", [pytest.param(e, id=e["file"]) for e in ENTRIES])
def test_grade_is_on_the_scale(e):
    g = grade_blob((PUBLIC / e["file"]).read_bytes(), e["file"], password=e.get("password"), encoding=e.get("encoding")).to_dict()
    assert g["grade"] in g["scale"], g
    assert g["write"]["grade"] in g["write_scale"], g
