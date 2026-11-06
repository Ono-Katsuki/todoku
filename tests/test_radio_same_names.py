"""同じシートに同じ名前のオプションボタンの組があるとき：組の位置で id を分け、窓・欄・書き込み・読み戻し・証明・検査でそろう。"""

from __future__ import annotations

import io
import re
import zipfile

import pytest

import test_controls as tc
from todoku import api
from todoku.xlsx.audit import run as audit
from todoku.xlsx.edit import EditError, fill
from todoku.xlsx.proof import prove
from todoku.xlsx.window import build
from todoku.xlsx.workbook import Workbook

TWO = [
    ("GBox", "Group Box 1", (5, 11, 9, 13), 'objectType="GBox" noThreeD="1"', ""),
    ("Radio", "rdoPlan", (6, 11, 7, 12), 'objectType="Radio" lockText="1"', ""),
    ("Radio", "rdoModify", (7, 11, 8, 12), 'objectType="Radio" lockText="1"', ""),
    ("GBox", "Group Box 1", (5, 28, 9, 30), 'objectType="GBox" noThreeD="1"', ""),
    ("Radio", "rdoPlan", (6, 28, 7, 29), 'objectType="Radio" lockText="1"', ""),
    ("Radio", "rdoModify", (7, 28, 8, 29), 'objectType="Radio" lockText="1"', ""),
    ("GBox", "Group Box 9", (1, 40, 4, 42), 'objectType="GBox" noThreeD="1"', ""),
    ("Radio", "rdoOther", (2, 40, 3, 41), 'objectType="Radio" lockText="1"', ""),
    ("Radio", "rdoOther2", (3, 40, 4, 41), 'objectType="Radio" lockText="1"', ""),
]
TEXT = {"rdoPlan": "新築", "rdoModify": "増築", "rdoOther": "はい", "rdoOther2": "いいえ", "Group Box 1": "計画", "Group Box 9": "ほか"}


@pytest.fixture
def blob(monkeypatch):
    monkeypatch.setattr(tc, "PARTS", TWO)
    monkeypatch.setattr(tc, "TEXT", TEXT)
    monkeypatch.setitem(tc.VML_TYPE, "Radio", "Radio")
    return tc.make_book()


def _radios(blob) -> dict:
    return {f["id"]: f for sh in build(Workbook(blob))["sheets"] for f in sh["fields"] if f["kind"] == "radio"}


def test_only_clashing_groups_get_a_position(blob):
    f = _radios(blob)
    assert set(f) == {"申込!rdoPlan@G12", "申込!rdoPlan@G29", "申込!rdoOther"}
    assert f["申込!rdoPlan@G12"]["controls"] == ["rdoPlan@G12", "rdoModify@G12"]
    assert f["申込!rdoOther"]["controls"] == ["rdoOther", "rdoOther2"]  # ぶつからない組の id は変えない
    assert f["申込!rdoPlan@G29"]["options"] == ["新築", "増築"]


def test_window_shows_the_same_ids(blob):
    from todoku.xlsx.window import to_markdown

    text = to_markdown(build(Workbook(blob)))
    assert "［rdoPlan@G12］" in text and "［rdoPlan@G29］" in text and "［rdoModify@G29］" in text
    assert "［rdoOther］" in text and "［rdoOther2］" in text  # ぶつからない組は、いままでの名前


def test_write_chooses_only_that_group_and_reads_back(blob):
    res = fill(Workbook(blob), {"申込!rdoPlan@G29": "増築"})
    f = _radios(res.blob)
    assert f["申込!rdoPlan@G29"]["value"] == "増築"
    assert f["申込!rdoPlan@G12"]["value"] is None  # もう一つの組は変わらない
    pr = prove(blob, res.blob, res.allowed_inputs, {(p.sheet, p.address): p.after for p in res.patches if p.kind == "control"})
    assert pr.ok, pr.unexpected
    assert [c["control"] for c in pr.control_changes] == ["rdoModify@G29"]


def test_both_groups_hold_their_own_choice(blob):
    res = fill(Workbook(blob), {"申込!rdoPlan@G12": "新築", "申込!rdoPlan@G29": "増築"})
    f = _radios(res.blob)
    assert f["申込!rdoPlan@G12"]["value"] == "新築" and f["申込!rdoPlan@G29"]["value"] == "増築"


def test_a_button_of_a_group_can_be_named_with_its_position(blob):
    res = fill(Workbook(blob), {"申込!rdoModify@G12": True})
    f = _radios(res.blob)
    assert f["申込!rdoPlan@G12"]["value"] == "増築" and f["申込!rdoPlan@G29"]["value"] is None


def test_the_bare_name_is_refused_with_the_candidates(blob):
    with pytest.raises(EditError, match="rdoPlan@G12.*rdoPlan@G29"):
        fill(Workbook(blob), {"申込!rdoPlan": "増築"})


def test_proof_catches_a_change_in_the_other_group_of_the_same_name(blob):
    """名前が同じでも、指示していない組が変わったら不合格。"""
    res = fill(Workbook(blob), {"申込!rdoPlan@G12": "新築"})
    wrong = fill(Workbook(res.blob), {"申込!rdoPlan@G29": "新築"})  # もう一つの組の同じ名前のボタン
    allowed = {(p.sheet, p.address): p.after for p in res.patches if p.kind == "control"}
    pr = prove(res.blob, wrong.blob, wrong.allowed_inputs, allowed)
    assert not pr.ok and any("rdoPlan@G29" in x for x in pr.unexpected), pr.unexpected


def test_fill_tool_reads_back_each_group(blob, tmp_path):
    p = tmp_path / "a.xlsx"
    p.write_bytes(blob)
    assert {"申込!rdoPlan@G12", "申込!rdoPlan@G29"} <= {f["id"] for f in api.fields(str(p))}
    res = api.fill(str(p), {"申込!rdoPlan@G12": "増築", "申込!rdoPlan@G29": "新築"}, str(tmp_path / "o.xlsx"))
    assert res["read_back_ok"] and res["proof_ok"], res
    assert "増築" in res["say_to_user"] and "新築" in res["say_to_user"]
    f = {x["id"]: x for x in api.fields(res["out_path"])}
    assert f["申込!rdoPlan@G12"]["value"] == "増築" and f["申込!rdoPlan@G29"]["value"] == "新築"


def test_audit_passes_on_both_groups(blob):
    au = audit(blob, "a.xlsx").to_dict()
    bad = [c for c in au["checks"] if not c["ok"]]
    assert not [c for c in bad if "読み戻し" in c["name"]], bad
    assert not any("rdoPlan" in str(c) for c in bad), bad


def test_only_the_state_changes_in_the_file(blob):
    """組の動きを決める部品の作り（グループボックス・図形の位置）は変えない。変わるのはチェックの印だけ。"""
    res = fill(Workbook(blob), {"申込!rdoPlan@G29": "増築"})
    with zipfile.ZipFile(io.BytesIO(blob)) as a, zipfile.ZipFile(io.BytesIO(res.blob)) as b:
        va, vb = a.read("xl/drawings/vmlDrawing1.vml").decode(), b.read("xl/drawings/vmlDrawing1.vml").decode()
    assert re.sub(r"[ \t]*<x:Checked>1</x:Checked>\r?\n?", "", vb) == va
