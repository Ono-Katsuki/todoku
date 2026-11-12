"""同じシートに同じ名前のチェックボックス・ドロップダウン・スピンボタンがあるとき：部品の位置で id を分け、欄・書き込み・読み戻し・証明・検査でそろう。

前は、同じ名前の部品は同じ id になっていた：名前で書くと、何も伝えずに 1 つ目に書き、読み戻しは別の部品と照らして「違い」と伝え、
2 つ目から後ろの部品は名指しできなかった。いまは、名前がぶつかる部品だけ id に自分の位置（左上のセル）を付ける（オプションボタンの組と同じ決まり）。
"""

from __future__ import annotations

import io
import zipfile

import pytest

import test_controls as tc
from todoku import api
from todoku.xlsx.audit import run as audit
from todoku.xlsx.edit import EditError, fill
from todoku.xlsx.proof import prove
from todoku.xlsx.window import build, to_markdown
from todoku.xlsx.workbook import Workbook

SAME = [
    ("CheckBox", "Check Box 1", (1, 5, 2, 6), 'objectType="CheckBox" lockText="1"', ""),
    ("CheckBox", "Check Box 1", (1, 8, 2, 9), 'objectType="CheckBox" lockText="1"', ""),
    ("CheckBox", "Check Box 1", (1, 8, 2, 9), 'objectType="CheckBox" lockText="1"', ""),  # 同じ名前で、同じセルにある
    ("CheckBox", "Check Box 3", (1, 11, 2, 12), 'objectType="CheckBox" lockText="1"', ""),
    ("Drop", "Drop Down 4", (4, 5, 6, 6), 'objectType="Drop" dropStyle="combo" dx="16" fmlaRange="$A$1:$A$3" sel="1" val="0"',
     "<x:FmlaRange>$A$1:$A$3</x:FmlaRange><x:Sel>1</x:Sel><x:Val>0</x:Val>"),
    ("Drop", "Drop Down 4", (4, 8, 6, 9), 'objectType="Drop" dropStyle="combo" dx="16" fmlaRange="$A$1:$A$3" sel="1" val="0"',
     "<x:FmlaRange>$A$1:$A$3</x:FmlaRange><x:Sel>1</x:Sel><x:Val>0</x:Val>"),
    ("Spin", "Spinner 6", (8, 5, 9, 6), 'objectType="Spin" dx="16" max="10" min="0" page="10" val="3"',
     "<x:Val>3</x:Val><x:Min>0</x:Min><x:Max>10</x:Max><x:Page>10</x:Page>"),
    ("Spin", "Spinner 6", (8, 8, 9, 9), 'objectType="Spin" dx="16" max="10" min="0" page="10" val="3"',
     "<x:Val>3</x:Val><x:Min>0</x:Min><x:Max>10</x:Max><x:Page>10</x:Page>"),
]
TEXT = {"Check Box 1": "はい", "Check Box 3": "ほか"}


@pytest.fixture
def blob(monkeypatch):
    monkeypatch.setattr(tc, "PARTS", SAME)
    monkeypatch.setattr(tc, "TEXT", TEXT)
    monkeypatch.setitem(tc.VML_TYPE, "CheckBox", "Checkbox")
    return tc.make_book()


def _fields(blob) -> dict:
    out: dict = {}
    for sh in build(Workbook(blob))["sheets"]:
        for f in sh["fields"]:
            if f["kind"] in ("checkbox", "select", "spin"):
                assert f["id"] not in out, f["id"]  # どの欄の id も、1 つの部品だけを指す
                out[f["id"]] = f
    return out


def _allowed(res) -> dict:
    return {(p.sheet, p.address): p.after for p in res.patches if p.kind == "control"}


def test_only_clashing_controls_get_a_position(blob):
    f = _fields(blob)
    assert set(f) == {"申込!Check Box 1@B6", "申込!Check Box 1@B9", "申込!Check Box 1@B9.1027", "申込!Check Box 3",
                      "申込!Drop Down 4@E6", "申込!Drop Down 4@E9", "申込!Spinner 6@I6", "申込!Spinner 6@I9"}
    assert f["申込!Check Box 3"]["control"] == "Check Box 3"  # ぶつからない部品の id は変えない
    assert f["申込!Check Box 1@B9"]["control"] == "Check Box 1@B9" and f["申込!Check Box 1@B9"]["at"] == "B9"
    text = to_markdown(build(Workbook(blob)))
    assert "［Check Box 1@B6］" in text and "［Check Box 1@B9］" in text and "［Check Box 3］" in text


def test_write_changes_only_the_named_checkbox(blob):
    res = fill(Workbook(blob), {"申込!Check Box 1@B9": True})
    f = _fields(res.blob)
    assert f["申込!Check Box 1@B9"]["value"] is True
    assert f["申込!Check Box 1@B6"]["value"] is False and f["申込!Check Box 1@B9.1027"]["value"] is False
    pr = prove(blob, res.blob, res.allowed_inputs, _allowed(res))
    assert pr.ok, pr.unexpected
    assert [c["control"] for c in pr.control_changes] == ["Check Box 1@B9"]


def test_write_to_same_named_dropdowns_and_spinners(blob):
    res = fill(Workbook(blob), {"申込!Drop Down 4@E9": "大阪", "申込!Spinner 6@I6": 7})
    f = _fields(res.blob)
    assert f["申込!Drop Down 4@E9"]["value"] == "大阪" and f["申込!Drop Down 4@E6"]["value"] == "東京"
    assert f["申込!Spinner 6@I6"]["value"] == 7 and f["申込!Spinner 6@I9"]["value"] == 3
    pr = prove(blob, res.blob, res.allowed_inputs, _allowed(res))
    assert pr.ok, pr.unexpected
    assert sorted(c["control"] for c in pr.control_changes) == ["Drop Down 4@E9", "Spinner 6@I6"]


def test_the_bare_name_is_refused_with_the_candidates(blob):
    with pytest.raises(EditError, match="Check Box 1@B6.*Check Box 1@B9.*Check Box 1@B9.1027"):
        fill(Workbook(blob), {"申込!Check Box 1": True})
    with pytest.raises(EditError, match="Spinner 6@I6.*Spinner 6@I9"):
        fill(Workbook(blob), {"申込!Spinner 6": 5})
    assert fill(Workbook(blob), {"申込!Check Box 3": True}).patches  # ぶつからない名前は、そのまま書ける


def test_proof_catches_a_change_in_the_other_checkbox_of_the_same_name(blob):
    """名前が同じでも、指示していない部品が変わったら不合格。"""
    res = fill(Workbook(blob), {"申込!Check Box 1@B6": True})
    wrong = fill(Workbook(blob), {"申込!Check Box 1@B9": True})
    pr = prove(blob, wrong.blob, wrong.allowed_inputs, _allowed(res))
    assert not pr.ok and any("Check Box 1@B9" in x for x in pr.unexpected), pr.unexpected


def test_fill_tool_reads_back_each_control(blob, tmp_path):
    p = tmp_path / "a.xlsx"
    p.write_bytes(blob)
    res = api.fill(str(p), {"申込!Check Box 1@B6": True, "申込!Check Box 1@B9.1027": True, "申込!Drop Down 4@E6": "名古屋"},
                   str(tmp_path / "o.xlsx"))
    assert res["read_back_ok"] and res["proof_ok"] and ".unverified" not in res["out_path"], res
    f = {x["id"]: x for x in api.fields(res["out_path"])}
    assert f["申込!Check Box 1@B6"]["value"] is True and f["申込!Check Box 1@B9"]["value"] is False
    assert f["申込!Check Box 1@B9.1027"]["value"] is True and f["申込!Drop Down 4@E6"]["value"] == "名古屋"
    assert f["申込!Drop Down 4@E9"]["value"] == "東京"


def test_audit_passes_on_same_named_controls(blob):
    au = audit(blob, "a.xlsx").to_dict()
    bad = [c for c in au["checks"] if not c["ok"]]
    assert not [c for c in bad if "読み戻し" in c["name"] or "書き込み" in c["name"]], bad
    assert not any("Check Box 1" in str(c) or "Spinner 6" in str(c) or "Drop Down 4" in str(c) for c in bad), bad


def test_only_the_state_changes_in_the_file(blob):
    """部品の名前・位置は変えない（id のためだけに区別する）。変わるのは、書いた部品の状態だけ。"""
    res = fill(Workbook(blob), {"申込!Check Box 1@B9": True})
    with zipfile.ZipFile(io.BytesIO(blob)) as a, zipfile.ZipFile(io.BytesIO(res.blob)) as b:
        changed = [n for n in a.namelist() if a.read(n) != b.read(n)]
        assert a.read("xl/worksheets/sheet1.xml") == b.read("xl/worksheets/sheet1.xml")
    assert sorted(changed) == ["xl/ctrlProps/ctrlProp2.xml", "xl/drawings/vmlDrawing1.vml"]
