"""もとからあった字や値を変えた所は、結果と「本人に伝えること」に出る（空欄に書いた所と同じ形で伝えない）。
決め方は、どの形式でも同じ：書いた所の、書く前の字と、書いた後の字を比べる（頼み方では決めない）。
証明の及ぶ範囲（とどくが書き出したファイルだけ）も、結果に添える。"""

from __future__ import annotations

import pytest

from todoku import api

openpyxl = pytest.importorskip("openpyxl")


def _xlsx(tmp_path, build, name="a.xlsx"):
    wb = openpyxl.Workbook()
    build(wb.active)
    p = tmp_path / name
    wb.save(p)
    return p


def test_overwrite_text_that_loses_printed_text_is_told(tmp_path):
    p = _xlsx(tmp_path, lambda ws: ws.__setitem__("B3", "（仮称）協議会"))
    r = api.fill(str(p), {"Sheet!B3": {"value": "西新宿介護ネット", "overwrite_text": True}}, str(tmp_path / "o.xlsx"))
    assert r["proof_ok"] and r["verified"]
    assert r["replaced_existing"] == [{"id": "Sheet!B3", "removed": "（仮称）協議会", "wrote": "西新宿介護ネット", "how": "removed"}]
    told = [x for x in r["say_to_user"].split("\n") if x.startswith("本人に伝えること：")]
    assert len(told) == 1 and "もとからあった字を消して書き換えた所が 1 か所" in told[0]
    assert "Sheet!B3「（仮称）協議会」→「西新宿介護ネット」" in told[0]


def test_overwrite_text_that_keeps_printed_text_is_not_a_replacement(tmp_path):
    # 元の字が、書いた字の中に同じ順で残っている（書き足しただけ）
    p = _xlsx(tmp_path, lambda ws: ws.__setitem__("B3", "新宿区"))
    r = api.fill(str(p), {"Sheet!B3": {"value": "東京都新宿区西新宿2-8-1", "overwrite_text": True}}, str(tmp_path / "o.xlsx"))
    assert r["proof_ok"] and "replaced_existing" not in r
    assert "もとからあった字" not in r["say_to_user"]


def test_anchor_write_that_removes_text_is_told_and_insertion_is_not(tmp_path):
    p = tmp_path / "a.txt"
    p.write_text("（仮称）協議会 設立趣意書\n代表者：\n", encoding="utf-8")
    r = api.fill(str(p), {"@1": {"old": "（仮称）協議会 設立", "new": "西新宿介護ネット 設立"},
                          "@2": {"old": "代表者：", "new": "代表者：山田花子"}}, str(tmp_path / "o.txt"))
    assert r["proof_ok"] and r["read_back_ok"]
    assert [(d["removed"], d["wrote"]) for d in r["replaced_existing"]] == [("（仮称）協議会", "西新宿介護ネット")]
    assert "「（仮称）協議会」→「西新宿介護ネット」" in r["say_to_user"]
    assert "山田花子」" not in r["say_to_user"].split("本人に伝えること：", 1)[1].split("／")[0]
    # 字を消しただけ
    r = api.fill(str(p), {"@1": {"old": "（仮称）協議会", "new": "協議会"}}, str(tmp_path / "o2.txt"))
    assert r["replaced_existing"][0]["removed"] == "（仮称）" and r["replaced_existing"][0]["wrote"] == ""
    assert "「（仮称）」→（字を消しただけ）" in r["say_to_user"]


def test_proof_scope_is_stated_in_every_fill_result(tmp_path):
    p = _xlsx(tmp_path, lambda ws: ws.__setitem__("A1", "氏名"))
    r = api.fill(str(p), {"Sheet!B1": "山田"}, str(tmp_path / "o.xlsx"))
    assert "ほかの道具で書き直したファイル" in r["proof_applies_to"] and "本人に必ず伝えてください" in r["proof_applies_to"]
    assert "replaced_existing" not in r


@pytest.mark.parametrize("old, new, how", [
    ("申請者", "届出者", "removed"),
    ("1000", "10000", "removed"),  # 数は、桁を足しても別の数
    ("１０００", "1000.0", None),  # 同じ数
    ("1,5", "15", "removed"),  # コンマは、桁区切りか小数点か決められない：字として比べる
    ("1,000", "1000", "removed"),
    ("山田", "山本と田中", "inside"),  # 字は残っているが、字と字のあいだに入った
    ("年月日", "8年11月10日", "inside"),
    ("令和　年　月　日", "令和8年11月10日", None),  # 空白の所に書いた
    ("氏名（　　　）", "氏名（山田花子）", None),
    ("____年__月", "2026年11月", None),
    ("□有　□無", "■有　□無", None),  # 箱の字の印
    ("氏名：", "氏名：山田", None),  # 後ろに足した
    ("新宿区", "東京都新宿区西新宿2-8-1", None),  # 前と後ろに足した
    ("", "山田", None),
    ("山田", "山田", None),
])
def test_what_happened_to_the_text_that_was_there(old, new, how):
    assert api._changed_existing(old, new) == how


def test_a_csv_cell_that_had_a_value_is_told(tmp_path):
    p = tmp_path / "a.csv"
    p.write_text("氏名,住所,電話\n山田太郎,大阪市北区1-1,06-1111-2222\n,,\n", encoding="utf-8")
    r = api.fill(str(p), {"R2.C2": "東京都新宿区西新宿2-8-1", "R3.C1": "佐藤花子", "R2.C3": "06-1111-2222（内線 12）"}, str(tmp_path / "o.csv"))
    assert r["proof_ok"] and r["read_back_ok"]
    # 値のあったマスだけ。空のマスに書いた所・後ろに足しただけの所は数えない
    assert r["replaced_existing"] == [{"id": "R2.C2", "removed": "大阪市北区1-1", "wrote": "東京都新宿区西新宿2-8-1", "how": "removed"}]
    assert "R2.C2「大阪市北区1-1」→「東京都新宿区西新宿2-8-1」" in r["say_to_user"]


def test_a_json_value_that_was_there_is_told(tmp_path):
    p = tmp_path / "a.json"
    p.write_text('{"name": "山田太郎", "addr": "", "n": 3}', encoding="utf-8")
    r = api.fill(str(p), {"/name": "佐藤花子", "/addr": "東京", "/n": 5}, str(tmp_path / "o.json"))
    assert r["proof_ok"] and r["read_back_ok"]
    assert [(d["id"], d["removed"], d["wrote"]) for d in r["replaced_existing"]] == [("/name", "山田太郎", "佐藤花子"), ("/n", "3", "5")]


def test_an_excel_number_written_over_by_address_is_told(tmp_path):
    # 欄の一覧にないセル（番地で書く）の数：桁を足しただけに見えても、別の数
    def build(ws):
        ws["A3"], ws["B3"], ws["A4"], ws["B4"] = "金額", 1000, "人数", 3

    p = _xlsx(tmp_path, build)
    r = api.fill(str(p), {"Sheet!B3": 10000, "Sheet!B4": 3}, str(tmp_path / "o.xlsx"))
    assert r["proof_ok"] and r["read_back_ok"]
    assert [(d["id"], d["removed"], d["wrote"], d["how"]) for d in r["replaced_existing"]] == [("Sheet!B3", "1000", "10000", "removed")]


def test_text_put_between_the_characters_is_told_separately(tmp_path):
    p = tmp_path / "a.txt"
    p.write_text("届出者 山田\n", encoding="utf-8")
    r = api.fill(str(p), {"@1": {"old": "山田", "new": "山本と田中"}}, str(tmp_path / "o.txt"))
    assert r["proof_ok"] and r["read_back_ok"]
    assert r["replaced_existing"] == [{"id": r["replaced_existing"][0]["id"], "removed": "山田", "wrote": "山本と田中", "how": "inside"}]
    assert "もとからあった字のあいだに字を足した所が 1 か所" in r["say_to_user"]
    assert "消して書き換えた所" not in r["say_to_user"]
