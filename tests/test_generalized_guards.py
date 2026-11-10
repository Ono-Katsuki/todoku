"""書き込みの守りは、形式や道によらず同じ：書類に起きたことで決める。
ふつうの道と大きな書類の道（Excel・CSV・JSON Lines）、OpenDocument の表・.xlsb、ZIP の中のファイル、コマンド（todoku fill）、
Word の古いフォームの欄（.docx・.doc・RTF）を、同じ書き込みでくらべる。"""

from __future__ import annotations

import json
import struct
import zipfile

import pytest
import xlsb_maker
from test_docx import make_docx, p
from test_openformats import SHEET, _w, make_ods, make_rtf, sj

from todoku import api, cli, lazy, limits, pkgsig

openpyxl = pytest.importorskip("openpyxl")

OVERWRITE = "もとからあった字を消して書き換えた所"


@pytest.fixture(autouse=True)
def clean(monkeypatch, tmp_path):
    monkeypatch.setenv("TODOKU_JOBS_DIR", str(tmp_path / "jobs"))
    lazy.clear()
    api._WINDOWS.clear()
    api._HAND_ONLY.clear()
    yield
    lazy.clear()
    api._WINDOWS.clear()
    api._HAND_ONLY.clear()


def _roads(monkeypatch, path, values, **kw) -> tuple[dict, dict]:
    """同じ書き込みを、ふつうの道と大きな書類の道で（断られたら {"refused": 理由}）。"""
    got = []
    for budget in (None, 300):
        with monkeypatch.context() as m:
            if budget:
                m.setattr(limits, "CONTEXT_BUDGET", budget)
            lazy.clear()
            api._WINDOWS.clear()
            try:
                r = api.fill(str(path), values, str(path.parent / f"out{budget or 0}{path.suffix}"), overwrite=True, **kw)
            except api.ApiError as e:
                r = {"refused": str(e)}
            got.append(r)
    lazy.clear()
    api._WINDOWS.clear()
    assert "大きな書類" not in (got[0].get("say_to_user") or "")
    assert "refused" in got[1] or "大きな書類" in got[1]["say_to_user"]  # 2 つ目は、大きな書類の道を通った
    return got[0], got[1]


def _book(tmp_path, name="a.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "申請"
    ws["A1"], ws["A2"], ws["A3"] = "氏名", "□同意する", "東京都"
    for r in range(5, 80):
        ws.cell(r, 1, f"項目{r}")
        ws.cell(r, 2, r)
    path = tmp_path / name
    wb.save(path)
    return path


def _signed(src, dst):
    """パッケージの電子署名のある .xlsx（署名の部品と、その関係を足す）。"""
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for i in zi.infolist():
            data = zi.read(i.filename)
            if i.filename == "_rels/.rels":
                data = data.replace(b"</Relationships>", b'<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/package/2006/'
                                    b'relationships/digital-signature/origin" Target="_xmlsignatures/origin.sigs"/></Relationships>')
            zo.writestr(i, data)
        zo.writestr("_xmlsignatures/origin.sigs", "")
        zo.writestr("_xmlsignatures/sig1.xml", '<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>')
    return dst


# ====================================================================== 大きな書類の道も、ふつうの道と同じ守り
def test_printed_text_overwrite_fails_proof_on_both_roads(tmp_path, monkeypatch):
    path = _book(tmp_path)
    for r in _roads(monkeypatch, path, {"申請!A1": "山田"}):
        assert r["proof_ok"] is False and r["verified"] is False
        assert any("印刷された字「氏名」" in u for u in r["unexpected"])
        assert r["replaced_existing"] == [{"id": "申請!A1", "removed": "氏名", "wrote": "山田", "how": "removed"}]
        assert ".unverified" in r["out_path"]
    for r in _roads(monkeypatch, path, {"申請!A1": {"value": "山田", "overwrite_text": True}}):
        assert r["proof_ok"] and r["verified"] and not r["unexpected"]
        assert r["replaced_existing"][0]["removed"] == "氏名" and OVERWRITE in r["say_to_user"]  # 明示した書き換えも、消えた字は伝える


def test_true_on_a_text_cell_is_refused_on_both_roads(tmp_path, monkeypatch):
    path = _book(tmp_path)
    for r in _roads(monkeypatch, path, {"申請!A3": True}):
        assert "チェックの箱の字" in r["refused"] and "東京都" in r["refused"]
    for r in _roads(monkeypatch, path, {"申請!A2": True}):  # 箱の字のあるセル：箱だけを変える
        assert r["proof_ok"] and r["read_back_ok"] and "replaced_existing" not in r
        assert any("「□」を「■」" in w for w in r["warnings"])


def test_changed_number_is_told_on_both_roads(tmp_path, monkeypatch):
    path = _book(tmp_path)
    for r in _roads(monkeypatch, path, {"申請!B7": 99}):
        assert r["proof_ok"] and r["replaced_existing"] == [{"id": "申請!B7", "removed": "7", "wrote": "99", "how": "removed"}]
        assert "近くの字「項目7」" in r["say_to_user"]  # 同じ行の左の字は、どちらの道でも事実として出す
    for r in _roads(monkeypatch, path, {"申請!B7": 7}):  # 同じ数：変わっていない
        assert r["proof_ok"] and "replaced_existing" not in r
    for r in _roads(monkeypatch, path, {"申請!B7": True}):  # 数のセルを true にした：数が消えた
        assert r["replaced_existing"][0]["removed"] == "7" and OVERWRITE in r["say_to_user"]


def test_large_road_does_not_say_no_nearby_text_for_rows_it_did_not_read(tmp_path, monkeypatch):
    path = _book(tmp_path)
    normal, large = _roads(monkeypatch, path, {"申請!A90": "追記"})
    assert "近くの字「項目79」" in normal["say_to_user"]  # ふつうの道：同じ列を上へたどる
    assert "近くの字なし" not in large["say_to_user"] and "同じ行の左に字なし（上の行は見ていません）" in large["say_to_user"]


def test_package_signature_is_told_on_both_roads(tmp_path, monkeypatch):
    path = _signed(_book(tmp_path), tmp_path / "s.xlsx")
    for r in _roads(monkeypatch, path, {"申請!C7": "メモ"}):
        assert r["digital_signature"]["invalidated"] and r["digital_signature"]["kind"] == "opc"
        assert r["warnings"][0].startswith("この書類には電子署名があります") and "電子署名" in r["say_to_user"]
    with zipfile.ZipFile(path) as z:
        assert pkgsig.find_in(z.namelist(), z.read).to_dict() == pkgsig.find(path.read_bytes()).to_dict()
    assert pkgsig.find(_book(tmp_path, "plain.xlsx").read_bytes()) is None


def test_hand_only_places_reach_the_result_on_both_roads_and_in_a_zip(tmp_path, monkeypatch):
    path = _book(tmp_path)
    finding = [{"check": "cannot_do", "title": "自署と押印", "detail": "申請者の署名欄に自署・押印", "where": ["A5"]}]
    for r in _roads(monkeypatch, path, {"申請!C8": "メモ"}):
        assert "hand_only" not in r and "まだ確かめていません" in r["hand_only_unchecked"]
    api.grade(str(path), findings=finding)
    for r in _roads(monkeypatch, path, {"申請!C8": "メモ"}):
        assert r["hand_only"][0]["title"] == "自署と押印" and "hand_only_unchecked" not in r
        assert "手で行う所" in r["say_to_user"] and "自署と押印" in r["say_to_user"]
    api._HAND_ONLY.clear()
    pack = tmp_path / "pack.zip"
    with zipfile.ZipFile(pack, "w") as z:
        z.write(path, "中/a.xlsx")
    member = f"{pack}!/中/a.xlsx"
    api.grade(member, findings=finding)
    r = api.fill(member, {"申請!B8": 5}, str(tmp_path / "from_zip.xlsx"))
    assert r["hand_only"][0]["title"] == "自署と押印"
    assert r["say_to_user"].count(OVERWRITE) == 1 and len(r["replaced_existing"]) == 1  # 同じことを 2 度言わない


def test_csv_and_jsonl_tell_replaced_values_on_both_roads(tmp_path, monkeypatch):
    c = tmp_path / "a.csv"
    c.write_text("氏名,住所\n" + "".join(f"人{i},町{i}\n" for i in range(60)), encoding="utf-8")
    normal, large = _roads(monkeypatch, c, {"R3.C1": "山田"})
    assert normal["proof_ok"] and large["proof_ok"]
    assert normal["replaced_existing"] == large["replaced_existing"] == [{"id": "R3.C1", "removed": "人1", "wrote": "山田", "how": "removed"}]
    j = tmp_path / "a.jsonl"
    j.write_text("".join(json.dumps({"name": f"人{i}", "n": i, "ok": False, "memo": None}, ensure_ascii=False) + "\n" for i in range(80)),
                 encoding="utf-8")
    for values, gone in (({"L3/name": "山田"}, ("人2", "山田")), ({"L3/n": 99}, ("2", "99")), ({"L3/memo": "x"}, None)):
        normal, large = _roads(monkeypatch, j, values)
        assert normal["proof_ok"] and large["proof_ok"]
        assert normal.get("replaced_existing") == large.get("replaced_existing")
        if gone is None:
            assert "replaced_existing" not in normal  # null に書いた：消えた字はない
        else:
            assert (normal["replaced_existing"][0]["removed"], normal["replaced_existing"][0]["wrote"]) == gone


# ====================================================================== OpenDocument の表・.xlsb も、Excel と同じ決まり
def _ods(tmp_path):
    cell = '<table:table-cell office:value-type="string" calcext:value-type="string"><text:p>{}</text:p></table:table-cell>'
    rows = SHEET.replace('<table:table-row table:number-rows-repeated="3">',
                         "<table:table-row>" + cell.format("□ 同意する") + cell.format("男") + '<table:table-cell table:number-columns-repeated="2"/>'
                         '</table:table-row><table:table-row table:number-rows-repeated="2">')
    assert rows != SHEET
    path = tmp_path / "a.ods"
    path.write_bytes(make_ods(rows))
    return path


def _try(path, values, **kw):
    try:
        return api.fill(str(path), values, str(path.parent / f"out{path.suffix}"), overwrite=True, **kw)
    except api.ApiError as e:
        return {"refused": str(e)}


@pytest.mark.parametrize("kind", ["ods", "xlsb"])
def test_sheet_formats_share_the_printed_text_rule(tmp_path, kind):
    if kind == "ods":
        path, label, text, box, twice = _ods(tmp_path), "Sheet1!A1", "氏名", "Sheet1!A5", "Sheet1!B1"
    else:
        path, label, text, box, twice = tmp_path / "a.xlsb", "申請!C1", "直接の字", None, "申請!B2"
        path.write_bytes(xlsb_maker.sample())
    name = "山田太郎"  # .xlsb の字のセルは、同じ字数の字だけ書ける（形式の決まり）
    r = _try(path, {label: name})
    assert r["proof_ok"] is False and any(f"印刷された字「{text}」" in u for u in r["unexpected"]) and ".unverified" in r["out_path"]
    assert r["replaced_existing"] == [{"id": label, "removed": text, "wrote": name, "how": "removed"}]
    r = _try(path, {label: {"value": name, "overwrite_text": True}})
    assert r["proof_ok"] and r["verified"] and r["replaced_existing"][0]["removed"] == text
    assert "チェックの箱の字" in _try(path, {label: True})["refused"]
    sheet, cell = twice.split("!")
    assert "同じセル" in _try(path, {twice: 1, f"{sheet}!{cell.lower()}": 2})["refused"]
    if box:
        r = _try(path, {box: True})
        assert r["proof_ok"] and "replaced_existing" not in r and any("「□」を「■」" in w for w in r["warnings"])


def test_xlsb_number_before_and_after_are_both_the_shown_text(tmp_path):
    path = tmp_path / "a.xlsb"
    path.write_bytes(xlsb_maker.sample())
    r = _try(path, {"申請!B2": 99})
    assert r["proof_ok"] and r["replaced_existing"] == [{"id": "申請!B2", "removed": "1,235円", "wrote": "99円", "how": "removed"}]
    r = _try(path, {"申請!B2": 1234.5})  # 同じ数：見えている字ではなく、セルの値でくらべる
    assert r["proof_ok"] and "replaced_existing" not in r


# ====================================================================== コマンドも、道具と同じ道を通る
def test_cli_fill_exit_codes_follow_proof_and_read_back(tmp_path, capsys):
    path = _book(tmp_path)
    out = tmp_path / "o.xlsx"
    assert cli.main(["fill", str(path), "--set", "申請!C7=メモ", "-o", str(out)]) == 0
    assert out.is_file() and out.with_suffix(".proof.json").is_file() and out.with_suffix(".proof.html").is_file()
    proof = json.loads(out.with_suffix(".proof.json").read_text(encoding="utf-8"))
    assert proof["proof_ok"] and proof["read_back_ok"]
    bad = tmp_path / "bad.xlsx"
    assert cli.main(["fill", str(path), "--set", "申請!A1=山田", "-o", str(bad)]) == 1  # 印刷された字を消した：証明が通らない
    assert not bad.exists() and (tmp_path / "bad.unverified.xlsx").is_file()
    said = capsys.readouterr().out
    assert "印刷された字「氏名」" in said and OVERWRITE in said
    wb = openpyxl.load_workbook(path)
    wb.active["C5"] = "=B5*2"
    formula = tmp_path / "f.xlsx"
    wb.save(formula)
    assert cli.main(["fill", str(formula), "--set", "申請!C5=1", "-o", str(tmp_path / "no.xlsx")]) == 2  # 書けない：書き出さない
    assert not (tmp_path / "no.xlsx").exists() and "書き込めません" in capsys.readouterr().err
    c = tmp_path / "a.csv"
    c.write_text("氏名,住所\n,\n", encoding="utf-8")
    assert cli.main(["fill", str(c), "--set", "R2.C1=山田", "-o", str(tmp_path / "o.csv")]) == 0  # Excel でない形式も同じ道
    assert (tmp_path / "o.csv").read_text(encoding="utf-8") == "氏名,住所\n山田,\n"


# ====================================================================== Word の古いフォームの欄：設定を読み、入力する所でない欄には書かない
def _docx_field(name, inner, text="     "):
    return ('<w:p><w:r><w:t>' + name + '：</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="' + name + '"/>' + inner
            + '</w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>' + text + '</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')


def _rtf_field(label, name, props, result, extra=""):
    return (_w("pard") + _w("f1 ") + sj(label) + "{" + _w("field") + "{" + _w("*") + _w("fldinst {FORMTEXT {") + _w("*") + _w("formfield{")
            + props + "{" + _w("*") + _w("ffname " + name + "}") + extra + "}}}}{" + _w("fldrslt {") + result + "}}}" + _w("par"))


def _form(tmp_path, kind):
    """数の欄（出るときに計算し直す）・計算の欄・今日の日付の欄・入力を受け付けない欄・ふつうの欄。欄の id（名前）を返す。"""
    if kind == "docx":
        body = (p("申請書")
                + _docx_field("数量", '<w:enabled/><w:calcOnExit/><w:textInput><w:type w:val="number"/><w:format w:val="#,##0"/></w:textInput>')
                + _docx_field("合計", '<w:enabled w:val="0"/><w:calcOnExit w:val="0"/><w:textInput><w:type w:val="calculated"/>'
                              '<w:default w:val="=数量*2"/></w:textInput>', "0")
                + _docx_field("日付", '<w:enabled/><w:textInput><w:type w:val="currentDate"/></w:textInput>', "2026/01/01")
                + _docx_field("固定", '<w:enabled w:val="false"/><w:textInput/>', "決まった字")
                + _docx_field("電話", '<w:enabled/><w:textInput><w:maxLength w:val="13"/></w:textInput>'))
        path = tmp_path / "ff.docx"
        path.write_bytes(make_docx(body))
        return path, {"number": "数量", "calculated": "合計", "date": "日付", "disabled": "固定", "plain": "電話"}
    blank = _w("u8194") + _w("'20") + _w("u8194") + _w("'20")
    path = tmp_path / "ff.rtf"
    path.write_bytes(make_rtf(
        _w("pard") + _w("f1 ") + sj("申請書") + _w("par"),
        _rtf_field("数量", "Qty", _w("fftype0") + _w("fftypetxt1") + _w("ffrecalc1"), blank, "{" + _w("*") + _w("ffformat #,##0}")),
        _rtf_field("合計", "Sum", _w("fftype0") + _w("fftypetxt5") + _w("ffprot1"), "0", "{" + _w("*") + _w("ffdeftext =Qty*2}")),
        _rtf_field("日付", "Day", _w("fftype0") + _w("fftypetxt3"), "2026/01/01"),
        _rtf_field("固定", "Fix", _w("fftype0") + _w("ffprot"), "abc"),  # 数のない \ffprot は 1
        _rtf_field("電話", "Tel", _w("fftype0") + _w("ffprot0"), blank)))
    return path, {"number": "Qty", "calculated": "Sum", "date": "Day", "disabled": "Fix", "plain": "Tel"}


@pytest.mark.parametrize("kind", ["docx", "rtf"])
def test_legacy_form_field_settings_are_read_and_respected(tmp_path, kind):
    path, ids = _form(tmp_path, kind)
    fl = {f["id"]: f for f in api.fields(str(path))}
    number, calc, day, fixed, plain = (fl[ids[k]] for k in ("number", "calculated", "date", "disabled", "plain"))
    assert number["text_type"] == "number" and number["format"] == "#,##0" and not number.get("readonly")
    assert calc["readonly"] and "計算の欄" in calc["not_for_input"] and calc["expression"].startswith("=")
    assert day["readonly"] and "今日の日付" in day["not_for_input"]
    assert fixed["readonly"] and "入力を受け付けない" in fixed["not_for_input"]
    assert not plain.get("readonly") and "not_for_input" not in plain and "text_type" not in plain
    r = _try(path, {ids["plain"]: "03-1234-5678"})
    assert r["proof_ok"] and r["read_back_ok"] and not r["warnings"]
    r = _try(path, {ids["number"]: "1234"})  # 出るときに計算し直す欄：計算の欄は計算し直していないと伝える
    assert r["proof_ok"] and len(r["warnings"]) == 1 and "計算し直していません" in r["warnings"][0] and ids["calculated"] in r["warnings"][0]
    r = _try(path, {ids["number"]: "たくさん"})  # 数の欄に、数字のない字：書くが、伝える
    assert r["proof_ok"] and any("数字がありません" in w for w in r["warnings"])
    for key, why in (("calculated", "計算の欄"), ("date", "今日の日付"), ("disabled", "入力を受け付けない")):
        assert why in _try(path, {ids[key]: "x"})["refused"]
    r = _try(path, {ids["disabled"]: "x"}, force=True)  # 決めて書くなら書ける：消えた字は伝える
    assert r["proof_ok"] and r["read_back_ok"] and r["replaced_existing"][0]["wrote"] == "x" and OVERWRITE in r["say_to_user"]


def _xstz(s: str) -> bytes:
    return struct.pack("<H", len(s)) + s.encode("utf-16-le") + b"\0\0"


def _ffdata(bits: int, name: str, default: str = "", fmt: str = "") -> bytes:
    body = struct.pack("<IHHH", 0xFFFFFFFF, bits, 0, 0) + _xstz(name) + _xstz(default) + _xstz(fmt) + _xstz("") * 4
    return struct.pack("<IH", 68 + len(body), 68) + bytes(62) + body


def test_doc_ffdata_bits_for_protection_recalculation_and_text_type():
    from todoku.doc.book import parse_ffdata

    for bits, want in ((0, (0, False, False)), (0x0200, (0, True, False)), (0x4000 | (1 << 11), (1, False, True)), ((5 << 11) | 0x0200, (5, True, False))):
        got = parse_ffdata(_ffdata(bits, "N", "=A*2", "#,##0"), 0)
        assert (got["type_txt"], got["prot"], got["recalc"]) == want
        assert (got["name"], got["default"], got["format"]) == ("N", "=A*2", "#,##0")


def test_doc_form_field_data_is_read_only_from_the_character_marked_as_data():
    # Word の書き出した .doc：欄の始めの字（0x13）にも絵の位置 0 が付いているが、欄のデータの印はない。そこを欄のデータとして読まない
    from todoku.doc.book import DocFile

    junk = b"\xff" * 40
    doc = DocFile.__new__(DocFile)
    doc.data = junk + _ffdata((5 << 11) | 0x0200, "Total", "=A*2")
    doc.controls = []
    props = {10: {"data": False, "pic": 0}, 21: {"data": True, "pic": len(junk)}}
    doc.char_props = lambda cp: props.get(cp, {"data": False, "pic": None})
    ctl = doc._form_field({"begin": 10, "instr": " FORMTEXT \x01", "sep": 22}, None)
    assert (ctl.name, ctl.text_type, ctl.expression, ctl.disabled) == ("Total", "calculated", "=A*2", True)
    doc.char_props = lambda cp: {"data": False, "pic": 0}  # 印の付いた字がない：欄のデータは読まない（名前のない欄のまま）
    ctl = doc._form_field({"begin": 10, "instr": " FORMTEXT \x01", "sep": 22}, None)
    assert ctl.name == "" and not ctl.disabled and ctl.text_type == ""


def test_public_doc_form_still_reads_its_field_names():
    from pathlib import Path

    path = Path(__file__).parent / "fixtures" / "public" / "form.doc"
    names = {f["id"] for f in api.fields(str(path))}
    assert {"Tel", "Mail"} <= names
