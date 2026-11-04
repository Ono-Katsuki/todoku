"""第5回の見直しの残り（合成の書類で確かめる）。

1. 書いたことで行が折り返す見込み（目安）を伝える：Word（表のマス・1 行目の字下げ）・.doc（公開の様式）・ODT・Excel・.ods。
   証明と読み戻しは字が正しいことしか見ないので、「○１・２種」の「種」が次の行へ行く・「有・○無」の「無」が離れる、といった
   見た目の変化を、警告と本人に伝えることに入れる（書き込みは止めない）
2. Excel の○・□の位置の見積もり：揃え・字下げ・字の大きさと書体の幅の種類・セルの余白・結合・縦の揃え・字の幅。
   見積もりの誤差が字の幅の半分を超えるときは描かずに断る（「位置を確かめられない」）。列の幅の単位（数字の幅）は、
   ＭＳ Ｐゴシック 11pt で 8 ピクセル（Excel が図形の位置を保存した様式で確かめた値）
3. 表の格子（週間の予定表）のマスの見出し：行の見出し（左）／列の見出し（上）。表の上の別の欄の見出しを拾わない
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import pytest

from todoku import api

MSG = "書いたことで行が折り返す見込み（目安）"


def _save(tmp_path, wb_, name="t.xlsx") -> Path:
    p = tmp_path / name
    wb_.save(p)
    return p


# ====================================================================== 1. 折り返し：Excel
def _wrap_book(tmp_path, text: str):
    openpyxl = pytest.importorskip("openpyxl")
    from openpyxl.styles import Alignment, Font

    wb_ = openpyxl.Workbook()
    ws = wb_.active
    ws.title = "S"
    ws["A1"] = "要介護認定"
    ws["B1"] = text
    ws.merge_cells("B1:C2")
    ws["B1"].alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    ws["B1"].font = Font(name="ＭＳ ゴシック", size=9)
    ws.column_dimensions["B"].width = 3
    ws.column_dimensions["C"].width = 3  # Calibri 11（数字 7 ピクセル）：21 ピクセルずつ。字を置ける幅は約 28 pt
    return _save(tmp_path, wb_)


def test_xlsx_write_that_makes_the_cell_wrap_is_reported(tmp_path):
    p = _wrap_book(tmp_path, "有・無")
    r = api.fill(str(p), {"@1": {"old": "有・無", "new": "有・○無", "cell": "S!B1"}}, str(tmp_path / "o.xlsx"))
    assert r["verified"], (r["unexpected"], r["read_back_differences"])
    hit = [w for w in r["warnings"] if MSG in w]
    assert hit and "S!B1" in hit[0] and "1 行から 2 行" in hit[0] and "B1:C2" in hit[0], r["warnings"]
    assert MSG in r["say_to_user"]
    assert any(x.get("wrap") for x in r["fit"])


def test_xlsx_write_that_still_fits_has_no_wrap_warning(tmp_path):
    p = _wrap_book(tmp_path, "有・無")
    r = api.fill(str(p), {"@1": {"old": "有・無", "new": "○・無", "cell": "S!B1"}}, str(tmp_path / "o.xlsx"))
    assert not any(MSG in w for w in r["warnings"]), r["warnings"]


# ====================================================================== 1. 折り返し：Word
def _docx_cell(text: str, grid_twips: int, size_half_points: int = 24, first_line: int = 0) -> bytes:
    from test_docx import make_docx

    ind = f'<w:ind w:firstLine="{first_line}"/>' if first_line else ""
    run = f'<w:r><w:rPr><w:sz w:val="{size_half_points}"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'
    body = (f'<w:tbl><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="{grid_twips}"/></w:tblGrid><w:tr>'
            '<w:tc><w:p><w:r><w:t>障害等級</w:t></w:r></w:p></w:tc>'
            f'<w:tc><w:tcPr><w:tcW w:w="{grid_twips}" w:type="dxa"/></w:tcPr><w:p><w:pPr>{ind}</w:pPr>{run}</w:p></w:tc></w:tr></w:tbl>'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:left="1701" w:right="1701" w:top="1985" w:bottom="1701"/></w:sectPr>')
    return make_docx(body)


def test_docx_circle_substitute_that_pushes_a_char_to_the_next_line_is_reported(tmp_path):
    # 12pt の字が 8 字（96pt）入るマス（2136 twip − 左右の余白 108 twip ずつ = 96pt）：「１　・　２　種」（7 字）は 1 行、
    # 「○１　・　２　種」（8 字）もちょうど入る。「○１　・　２　　種」（9 字）は「種」が次の行へ
    src = tmp_path / "t.docx"
    src.write_bytes(_docx_cell("１　・　２　種", 2136))
    r = api.fill(str(src), {"@1": {"old": "１　・　２　種", "new": "○１　・　２　　種"}}, str(tmp_path / "o.docx"))
    assert r["verified"], (r["unexpected"], r["read_back_differences"])
    hit = [w for w in r["warnings"] if MSG in w]
    assert hit and "1 行から 2 行" in hit[0] and "表のマスの幅" in hit[0], r["warnings"]
    assert MSG in r["say_to_user"] and r["wrap"]


def test_docx_write_that_fits_has_no_wrap_warning(tmp_path):
    src = tmp_path / "t.docx"
    src.write_bytes(_docx_cell("１　・　２　種", 3000))
    r = api.fill(str(src), {"@1": {"old": "１　・　２　種", "new": "○１　・　２　種"}}, str(tmp_path / "o.docx"))
    assert not any(MSG in w for w in r["warnings"]), r["warnings"]


def test_docx_first_line_indent_counts(tmp_path):
    # 1 行目の字下げ（firstLine 1200 twip = 60pt）：字を置ける幅が 1 行目だけ狭い
    src = tmp_path / "t.docx"
    src.write_bytes(_docx_cell("銀行　信金", 3000, 24, first_line=1200))
    r = api.fill(str(src), {"@1": {"old": "銀行　信金", "new": "架空　銀行　信金"}}, str(tmp_path / "o.docx"))
    assert any(MSG in w for w in r["warnings"]), r["warnings"]
    src.write_bytes(_docx_cell("銀行　信金", 3000, 24))
    r = api.fill(str(src), {"@1": {"old": "銀行　信金", "new": "架空　銀行　信金"}}, str(tmp_path / "o2.docx"))
    assert not any(MSG in w for w in r["warnings"]), r["warnings"]


def test_wrap_line_counts():
    from todoku.wrapcheck import count_lines, wraps

    assert count_lines("あいうえお", 50, 10) == 1 and count_lines("あいうえおか", 50, 10) == 2
    assert count_lines("あいうえお、", 50, 10) == 1  # 句読点は行の終わりにぶら下げる
    assert count_lines("ab cdefgh", 30, 10) == 2  # 英字の語は、語の初めから次の行へ
    assert count_lines("あ\nい", 100, 10) == 2
    assert wraps("有・無", "有・○無", 35, 10, 10) == (1, 2, True)
    assert wraps("有・無", "有・○無", 42, 10, 10) == (1, 2, False)  # 幅いっぱいに近い：少し狭く測るソフトでは折り返す
    assert wraps("有・無", "○有・無", 60, 10, 10) is None


# ====================================================================== 1. 折り返し：.doc（公開の様式）
DOC = Path(__file__).resolve().parent.parent / "corpus" / "pairs" / "www.city.kitakyushu.lg.jp" / "2adce6eea3a7_blank.doc"


@pytest.mark.corpus
def test_doc_write_that_wraps_in_a_table_cell_is_reported(tmp_path):
    if not DOC.exists():
        pytest.skip("corpus なし")
    src = tmp_path / "k.doc"
    src.write_bytes(DOC.read_bytes())
    r = api.fill(str(src), {"@1": {"old": "有　・　無", "new": "有（週三回、午前と午後に通所しています）　・　無",
                                   "occurrence": 1}}, str(tmp_path / "o.doc"))
    assert r["verified"], (r["unexpected"], r["read_back_differences"])
    hit = [w for w in r["warnings"] if MSG in w]
    assert hit and "表のマスの幅" in hit[0] and "字の大きさ 10.5" in hit[0], r["warnings"]
    r = api.fill(str(src), {"@1": {"old": "有　・　無", "new": "○有　・　無", "occurrence": 1}}, str(tmp_path / "o2.doc"))
    assert not any(MSG in w for w in r["warnings"]), r["warnings"]


# ====================================================================== 1. 折り返し：OpenDocument
def test_odt_write_that_wraps_in_a_table_cell_is_reported(tmp_path):
    from test_openformats import make_odt

    auto = '<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="3cm"/></style:style>'
    body = ('<table:table table:name="t1"><table:table-column table:style-name="co1" table:number-columns-repeated="2"/>'
            '<table:table-row><table:table-cell><text:p>外出</text:p></table:table-cell>'
            '<table:table-cell><text:p>有・無</text:p></table:table-cell></table:table-row></table:table>')
    src = tmp_path / "t.odt"
    src.write_bytes(make_odt(body, auto))
    r = api.fill(str(src), {"@1": {"old": "有・無", "new": "有（毎日）・無"}}, str(tmp_path / "o.odt"))
    assert r["verified"], (r["unexpected"], r.get("read_back_differences"))
    assert any(MSG in w for w in r["warnings"]), r["warnings"]
    r = api.fill(str(src), {"@1": {"old": "有・無", "new": "○有・無"}}, str(tmp_path / "o2.odt"))
    assert not any(MSG in w for w in r["warnings"]), r["warnings"]


def test_ods_write_that_wraps_a_cell_is_reported(tmp_path):
    from test_openformats import ODS_AUTO, make_ods

    auto = ODS_AUTO + ('<style:style style:name="cw" style:family="table-cell"><style:table-cell-properties fo:wrap-option="wrap"/>'
                       '<style:text-properties fo:font-size="10pt"/></style:style>'
                       '<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="2cm"/></style:style>')
    table = ('<table:table table:name="Sheet1"><table:table-column table:style-name="co1" table:number-columns-repeated="2"/>'
             '<table:table-row><table:table-cell office:value-type="string" calcext:value-type="string"><text:p>住所</text:p></table:table-cell>'
             '<table:table-cell table:style-name="cw"/></table:table-row></table:table>')
    src = tmp_path / "t.ods"
    src.write_bytes(make_ods(table, auto, ""))
    r = api.fill(str(src), {"Sheet1!B1": "東京都千代田区一丁目"}, str(tmp_path / "o.ods"))
    assert any(MSG in w and "Sheet1!B1" in w for w in r["warnings"]), r["warnings"]
    r = api.fill(str(src), {"Sheet1!B1": "東京"}, str(tmp_path / "o2.ods"))
    assert not any(MSG in w for w in r["warnings"]), r["warnings"]


# ====================================================================== 2. Excel の○の位置
def _circle_book(tmp_path, text: str, font: str, horizontal: str = "left", indent: int = 0, default_font: str | None = None,
                 width: float = 40, vertical: str = "center", wrap: bool = False, name: str = "t.xlsx"):
    openpyxl = pytest.importorskip("openpyxl")
    from openpyxl.styles import Alignment, Font

    wb_ = openpyxl.Workbook()
    ws = wb_.active
    ws.title = "S"
    ws["A1"] = "選択"
    ws["B1"] = text
    ws["B1"].font = Font(name=font, size=11)
    ws["B1"].alignment = Alignment(horizontal=horizontal, indent=indent, vertical=vertical, wrap_text=wrap)
    ws.column_dimensions["B"].width = width
    ws.row_dimensions[1].height = 30
    p = _save(tmp_path, wb_, name)
    if default_font:
        import re
        import zipfile

        data = {}
        with zipfile.ZipFile(p) as z:
            for n in z.namelist():
                data[n] = z.read(n)
        data["xl/styles.xml"] = re.sub(rb'<name val="Calibri" ?/>', f'<name val="{default_font}"/>'.encode(), data["xl/styles.xml"], count=1)
        with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
            for n, d in data.items():
                z.writestr(n, d)
    return p


def test_xlsx_circle_refuses_when_the_estimate_error_exceeds_half_a_char(tmp_path):
    # 「Ｐ」の書体の左揃えで、囲む字の前に幅の変わるかなが長く並ぶ：誤差が字の幅の半分を超えるので描かない
    text = "あいうえおかきくけこさしすせそ・有・無"
    p = _circle_book(tmp_path, text, "ＭＳ Ｐゴシック")
    with pytest.raises(api.ApiError, match="位置を確かめられない"):
        api.fill(str(p), {"@c": {"circle": "無", "cell": "S!B1"}}, str(tmp_path / "o.xlsx"))
    # 等幅の書体（ＭＳ ゴシック）なら、同じ字でも位置が決まる。結果に誤差の目安（mm）を添える
    q = _circle_book(tmp_path, text, "ＭＳ ゴシック", name="m.xlsx")
    r = api.fill(str(q), {"@c": {"circle": "無", "cell": "S!B1"}}, str(tmp_path / "o2.xlsx"))
    assert r["verified"] and r["circled"][0]["error_mm"] < 1.0 and "mm 以内" in r["circled"][0]["uncertain"]
    # 右揃えなら、後ろの字だけが効く（「無」の後ろに字はない）
    s = _circle_book(tmp_path, text, "ＭＳ Ｐゴシック", horizontal="right", name="r.xlsx")
    r = api.fill(str(s), {"@c": {"circle": "無", "cell": "S!B1"}}, str(tmp_path / "o3.xlsx"))
    assert r["verified"], r


def test_xlsx_circle_refuses_distributed_alignment(tmp_path):
    p = _circle_book(tmp_path, "有・無", "ＭＳ ゴシック", horizontal="distributed")
    with pytest.raises(api.ApiError, match="distributed"):
        api.fill(str(p), {"@c": {"circle": "無", "cell": "S!B1"}}, str(tmp_path / "o.xlsx"))


def test_xlsx_column_unit_and_text_box_estimate(tmp_path):
    from todoku.book import open_book
    from todoku.xlsx import circle as C

    # ＭＳ Ｐゴシック 11pt：数字の幅 8 ピクセル（既定の列の幅 9 = 72 ピクセル）。Calibri 11pt：7 ピクセル
    assert C.mdw_for("ＭＳ Ｐゴシック", 11) == (8, True) and C.mdw_for("ＭＳ Ｐゴシック", 10) == (7, True)
    assert C.mdw_for("Calibri", 11) == (7, True) and C.mdw_for("游ゴシック", 11)[0] == 8
    assert C.width_px(9, 8) == 72 and C.width_px(None, 8) == 72 and C.width_px(None, 7) == 64
    fam = C.family("HG丸ｺﾞｼｯｸM-PRO")
    assert not fam.cjk_prop and fam.latin_mono  # 「PRO」は「Ｐ」の書体ではない
    assert C.family("ＭＳ Ｐ明朝").cjk_prop and C.family("HGPｺﾞｼｯｸE").cjk_prop and C.family("Meiryo UI").cjk_prop
    assert C.family("游ゴシック").pitch > 1.6  # 游ゴシックは行が高い（既定の行の高さ 18.75pt）
    # 字下げ：左揃えの字は、字下げの段ごとに数字 3 つ分右へ
    p = _circle_book(tmp_path, "有・無", "ＭＳ ゴシック", default_font="ＭＳ Ｐゴシック")
    q = _circle_book(tmp_path, "有・無", "ＭＳ ゴシック", indent=2, default_font="ＭＳ Ｐゴシック", name="i.xlsx")
    wa, wb2 = open_book(p.read_bytes(), "a.xlsx"), open_book(q.read_bytes(), "b.xlsx")
    ga, _ = C.glyphs(wa, wa.sheets[0], 1, 2, "有・無")
    gb, _ = C.glyphs(wb2, wb2.sheets[0], 1, 2, "有・無")
    assert abs((gb[0].x0 - ga[0].x0) - 2 * 3 * 8 * 0.75) < 1e-6
    # 中央揃え：字の並びがセルの中央
    s = _circle_book(tmp_path, "有・無", "ＭＳ ゴシック", horizontal="center", default_font="ＭＳ Ｐゴシック", width=20, name="c.xlsx")
    ws_ = open_book(s.read_bytes(), "c.xlsx")
    g, _ = C.glyphs(ws_, ws_.sheets[0], 1, 2, "有・無")
    x0 = C.col_pt(ws_, ws_.sheets[0], 1)
    w = C.col_pt(ws_, ws_.sheets[0], 2)
    assert abs((g[0].x0 + g[-1].x1) / 2 - (x0 + w / 2)) < 1e-6


def test_xlsx_circle_refuses_when_wrap_is_too_close_to_call(tmp_path):
    # 折り返すセルで、行の字の幅がセルの幅に近く（かなの幅の見積もりの範囲に入る）、自動で改行されるかが分からない
    p = _circle_book(tmp_path, "あいうえお・有・無", "ＭＳ Ｐゴシック", wrap=True, width=20)
    with pytest.raises(api.ApiError, match="位置を確かめられない"):
        api.fill(str(p), {"@c": {"circle": "有", "cell": "S!B1"}}, str(tmp_path / "o.xlsx"))


# ====================================================================== 3. 表の格子の見出し
def _weekly(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    from openpyxl.styles import Border, Side

    thin = Side(style="thin")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    wb_ = openpyxl.Workbook()
    ws = wb_.active
    ws.title = "S"
    # 表の上の別の欄：受給者証番号・障害支援区分（同じ列の上にある）
    ws["A1"] = "受給者証番号"
    ws.merge_cells("A1:D1")
    ws["E1"] = "障害支援区分"
    ws.merge_cells("E1:F1")
    ws["G1"] = "計画作成日"
    for c in "ABCDEFGH":
        ws[f"{c}1"].border = box
    # 格子：見出しの行（月・火・水・木）と、行の見出しの列（時刻。数の値で、表示は h:mm）
    for c, day in zip("EFGH", ["月", "火", "水", "木"]):
        ws[f"{c}2"] = day
        ws[f"{c}2"].border = box
    for r, t in ((3, dt.time(6, 0)), (7, dt.time(8, 0))):  # 5・6 行目は、行の見出しのない行（7:00 の行）
        ws[f"A{r}"] = t
        ws[f"A{r}"].number_format = "h:mm"
        ws.merge_cells(f"A{r}:D{r + 1}")
    for r in range(2, 9):
        for c in "ABCD":
            ws[f"{c}{r}"].border = box
        for c in "EFGH":
            ws[f"{c}{r}"].border = box
    return _save(tmp_path, wb_, "w.xlsx")


def test_weekly_grid_labels_are_nearest_text_with_both_facts(tmp_path):
    # 罫線の格子（週間の予定表）のマス：数の決まり（見出しの行の仲間の数・下の行の見出しの数）で表の作りを決めず、
    # 左と上のいちばん近い字（あいだの空の欄をこえる）を、離れ方つきで near に出す（数の値の時刻は number の事実）。
    # 見出しは、ほかの欄と同じ決まり（左のすぐ隣の字、なければ上の見出し）。字でない数の値は見出しにしない
    p = _weekly(tmp_path)
    _k, doc, _parts = api.window(str(p))
    raw = {f["id"]: f for f in doc["sheets"][0]["fields"]}
    for fid, left, above in (("S!E7", "8:00", "月"), ("S!H3", "6:00", "木"), ("S!F8", "8:00", "火"), ("S!F4", "6:00", "火")):
        near = raw[fid]["near"]
        assert near["left"]["label"] == left and near["above"]["label"] == above, (fid, near)
        assert near["left"]["mm"] >= 0 and near["above"]["mm"] >= 0 and "cols" in near["left"] and "rows" in near["above"]
    assert raw["S!E7"]["label"] == "月" and raw["S!E7"]["near"]["left"]["number"]  # 左のすぐ隣は時刻（数の値）
    assert raw["S!H3"]["label"] == "木"  # すぐ上（0 mm）。左の時刻は 3 列先
    assert raw["S!E5"]["label"] == "月" and "left" not in raw["S!E5"]["near"]  # 行の見出しのない行：上の字だけ
    by = {f["id"]: f for f in api.fields(str(p))}
    for f in by.values():
        if f["id"].startswith("S!") and f.get("label") and f["id"][2:3] in "EFGH" and f["id"][3:].isdigit() and int(f["id"][3:]) >= 3:
            assert "受給者証番号" not in f["label"] and "障害支援区分" not in f["label"] and "計画作成日" not in f["label"], f
