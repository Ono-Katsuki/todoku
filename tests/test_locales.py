"""日本語に限らない読み書きの規則：組み込みの表示形式・Word の日付の欄の言語と暦・ODF の暦・段落番号・字の幅・日付の入力。"""

import datetime as dt
import io

import pytest

from todoku.docx.document import number_text
from todoku.docx.fill import FillError, format_word_date
from todoku.formula.numfmt import BUILTIN, JA_JP_BUILTIN, LocaleDependent, builtin_id, format_value, is_date_format
from todoku.formula.values import Unsupported
from todoku.textwidth import cells, text_cells

NOV3 = 46329  # 2026-11-03（1900 年の数え方）


# ====================================================================== 1. 組み込みの表示形式（番号 14・22・27〜36・50〜58）
@pytest.mark.parametrize("fid", [14, 22, 27, 31, 32, 36, 50, 58])
def test_locale_dependent_builtin_formats_are_not_shown_as_japanese(fid):
    fmt = BUILTIN[fid]
    assert builtin_id(fmt) == fid and is_date_format(fmt)
    with pytest.raises(LocaleDependent) as e:
        format_value(NOV3, fmt)
    assert e.value.shown == "2026-11-03" and isinstance(e.value, Unsupported)
    assert f"番号 {fid}" in str(e.value) and "開く PC の地域の設定で見え方が変わる" in str(e.value)


def test_builtin_iso_rendering_with_time_and_1904():
    with pytest.raises(LocaleDependent) as e:
        format_value(NOV3 + 0.5, BUILTIN[22])
    assert e.value.shown == "2026-11-03T12:00:00"
    with pytest.raises(LocaleDependent) as e:
        format_value(0.75, BUILTIN[32])
    assert e.value.shown == "18:00:00"
    with pytest.raises(LocaleDependent) as e:
        format_value(NOV3 - 1462, BUILTIN[14], True)
    assert e.value.shown == "2026-11-03"


def test_fixed_builtin_formats_and_the_japanese_table():
    assert format_value(NOV3, BUILTIN[15]) == "3-Nov-26"  # 15 は地域で変わらない（18.8.30）
    assert format_value(NOV3, JA_JP_BUILTIN[14]) == "2026/11/3"
    assert format_value(NOV3, JA_JP_BUILTIN[58]) == "令和8年11月3日"


def test_workbook_with_builtin_14_is_told_but_keeps_read_grade(tmp_path):
    """値（日付）は正確に分かるので読む段階は下げない。見え方が PC の地域で変わることだけ伝える。"""
    openpyxl = pytest.importorskip("openpyxl")
    from todoku import api

    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Date"
    ws["B1"] = dt.date(2026, 11, 3)
    ws["B1"].number_format = "mm-dd-yy"  # openpyxl は組み込みの番号 14 で保存する
    path = tmp_path / "d.xlsx"
    wb.save(path)
    g = api.grade(str(path))
    r = next(r for r in g["reasons"] if r["title"].startswith("組み込みの表示形式（番号 14）"))
    assert r["grade"] == "A" and r["where"] == ["Sheet!B1"]
    assert g["grade"] in ("S", "A")
    assert "再現できない表示形式" not in [x["title"] for x in g["reasons"]]


# ====================================================================== 2. Word の日付の欄（w:lid・w:calendar）
D = dt.date(2026, 11, 3)


@pytest.mark.parametrize("fmt,lid,cal,want", [
    ("ggge年M月d日", "ja-JP", "japan", "令和8年11月3日"),
    ("ggge年M月d日（aaa）", "ja-JP", None, "令和8年11月3日（火）"),
    ("dddd d MMMM yyyy", "fr-FR", "gregorian", "mardi 3 novembre 2026"),
    ("d MMMM yyyy", "th-TH", "thai", "3 พฤศจิกายน 2569"),
    ("e年M月d日", "zh-TW", "taiwan", "115年11月3日"),
    ("yyyy.MM.dd", "ko-KR", "korea", "4359.11.03"),
    ("d MMMM yyyy", "ru-RU", None, "3 ноября 2026"),  # 日と並ぶ月は生格
    ("MMMM yyyy", "ru-RU", None, "ноябрь 2026"),
    ("dddd, MMMM d, yyyy", "0409", None, "Tuesday, November 3, 2026"),  # 16進の LCID
    ("d' de 'MMMM' de 'yyyy", "es-ES", None, "3 de noviembre de 2026"),  # '…' は字のまま
    ("yyyy/MM/dd", "ar-SA", "hijri", "1448/05/22"),
    ("yyyy/M/d H:mm", "ja-JP", None, "2026/11/3 0:00"),
])
def test_word_date_follows_lid_and_calendar(fmt, lid, cal, want):
    assert format_word_date(D, fmt, lid, cal) == want


@pytest.mark.parametrize("fmt,lid,cal", [
    ("d MMMM yyyy", "he-IL", "hebrew"),  # ヘブライ暦は作らない
    ("d MMMM yyyy", "hi-IN", "saka"),
    ("dddd", None, None),  # 言語の指定がなければ、名前を出さない
    ("ggge年", None, None),  # 言語も暦も和暦でなければ、元号を出さない
    ("ggge年", "en-US", "gregorian"),
    ("MMMM", "ar-SA", "hijri"),  # ヒジュラ暦の月の名前
    ("yyyy年", "ja-JP", "japan"),  # 和暦の y
    ("d MMMM yyyy", "nl-NL", None),  # 名前の表のない言語
    ("h:mm tt", "en-US", None),  # 知らない記号
])
def test_word_date_refuses_what_it_cannot_reproduce(fmt, lid, cal):
    with pytest.raises(FillError):
        format_word_date(D, fmt, lid, cal)


def test_meiji_starts_on_the_same_day_as_excel():
    from todoku.formula.dates import ERAS

    assert format_word_date(dt.date(1868, 1, 10), "ggge年", "ja-JP", "japan") == "明治元年"
    assert [e for e in ERAS if e[0] == "明治"][0][2] == dt.date(1868, 1, 1)


def test_date_control_lid_and_calendar_are_read(tmp_path):
    from test_word_findings import docx, p, r

    from todoku import api

    body = p(r("Date : "), '<w:sdt><w:sdtPr><w:tag w:val="date"/><w:id w:val="4"/><w:showingPlcHdr/><w:date>'
             '<w:dateFormat w:val="dddd d MMMM yyyy"/><w:lid w:val="fr-FR"/><w:calendar w:val="gregorian"/></w:date></w:sdtPr>'
             '<w:sdtContent><w:r><w:t>Choisir</w:t></w:r></w:sdtContent></w:sdt>')
    path = tmp_path / "f.docx"
    path.write_bytes(docx(body))
    res = api.fill(str(path), {"date": "2026-11-03"}, str(tmp_path / "o.docx"))
    assert res["written_as"] == {"date": "mardi 3 novembre 2026"}, res


# ====================================================================== 3. ODF の暦
def _odf_date(body: str, attrs: str = "") -> str:
    from todoku.odf.numstyle import excel_format
    from todoku.odf.xml import parse, top

    ns = ('xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" '
          'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"')
    data = f'<number:date-style {ns} style:name="N1" {attrs}>{body}</number:date-style>'.encode()
    return excel_format(top(parse(data)), data, lambda n: None)[0]


def test_odf_calendars():
    f = _odf_date('<number:year number:style="long" number:calendar="buddhist"/><number:text>-</number:text><number:month/>',
                  'number:language="th" number:country="TH"')
    assert format_value(NOV3, f) == "2569-11"
    f = _odf_date('<number:year number:calendar="ROC"/><number:text>年</number:text><number:month number:calendar="ROC"/>'
                  '<number:text>月</number:text><number:day number:calendar="ROC"/><number:text>日</number:text>')
    assert format_value(NOV3, f) == "115年11月3日"
    f = _odf_date('<number:era number:style="long" number:calendar="gengou"/><number:year number:calendar="gengou"/>'
                  '<number:text>年</number:text><number:month/><number:text>月</number:text>', 'number:language="ja" number:country="JP"')
    assert format_value(NOV3, f) == "令和8年11月"
    f = _odf_date('<number:year number:style="long" number:calendar="hijri"/><number:text>/</number:text>'
                  '<number:month number:calendar="hijri"/>', 'number:language="ar" number:country="SA"')
    assert format_value(NOV3, f) == "1448/5"
    # ヘブライ暦・混じった暦・民国の元号の名前は作らない（グレゴリオ暦の年を黙って出さない）
    assert _odf_date('<number:year number:calendar="jewish"/>') == ""
    assert _odf_date('<number:year number:calendar="ROC"/><number:month number:calendar="buddhist"/>') == ""
    assert _odf_date('<number:era number:calendar="ROC"/><number:year number:calendar="ROC"/>') == ""


def test_excel_era_calendar_months_are_gregorian():
    # 和暦（暦の番号 03）の指定でも、月と日はグレゴリオ暦と同じ
    assert format_value(NOV3, '[$-30411]ggge"年"m"月"d"日"') == "令和8年11月3日"


# ====================================================================== 4. 段落番号（ST_NumberFormat）
@pytest.mark.parametrize("fmt,n,want", [
    ("aiueo", 1, "ｱ"), ("aiueo", 47, "ｱｱ"), ("aiueoFullWidth", 3, "ウ"), ("iroha", 3, "ﾊ"), ("irohaFullWidth", 3, "ハ"),
    ("japaneseCounting", 21, "二十一"), ("japaneseCounting", 110, "百十"), ("japaneseLegal", 3, "参"),
    ("ideographDigital", 10, "一〇"), ("chineseCounting", 12, "十二"), ("chineseLegalSimplified", 3, "叁"),
    ("ideographLegalTraditional", 3, "參"), ("ideographZodiacTraditional", 1, "甲子"), ("ideographEnclosedCircle", 2, "㊁"),
    ("ganada", 3, "다"), ("chosung", 3, "ㄷ"), ("koreanCounting", 23, "이십삼"), ("koreanLegal", 3, "셋"), ("koreanDigital", 3, "삼"),
    ("russianLower", 3, "в"), ("russianUpper", 10, "К"), ("hebrew1", 15, "טו"), ("hebrew1", 123, "קכג"), ("hebrew2", 3, "ג"),
    ("hindiNumbers", 12, "१२"), ("thaiNumbers", 12, "๑๒"), ("thaiCounting", 21, "ยี่สิบเอ็ด"), ("thaiCounting", 11, "สิบเอ็ด"),
    ("numberInDash", 4, "- 4 -"), ("upperLetter", 27, "AA"),
])
def test_list_numbers(fmt, n, want):
    assert number_text(n, fmt) == want


@pytest.mark.parametrize("fmt,n", [
    ("arabicAbjad", 2), ("arabicAlpha", 2), ("hindiVowels", 2), ("hindiConsonants", 2), ("thaiLetters", 2),
    ("vietnameseCounting", 2), ("cardinalText", 2), ("custom", 2), ("chineseCounting", 100), ("japaneseLegal", 10),
    ("russianLower", 29), ("ganada", 15), ("iroha", 25), ("someFutureFormat", 1),
])
def test_unknown_list_numbers_are_marked_not_decimal(fmt, n):
    t = number_text(n, fmt)
    assert t == f"〔{n}：番号の書式「{fmt}」は再現できません〕"


def test_word_grade_flags_unreproducible_numbering():
    from test_word_findings import docx

    from todoku.grade import grade_blob

    num = ('<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="arabicAbjad"/>'
           '<w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>')
    import test_word_findings as twf

    old = twf.NUMBERING
    twf.NUMBERING = num
    try:
        blob = docx('<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
                    '<w:r><w:t>الاسم</w:t></w:r></w:p>', numbering=True)
    finally:
        twf.NUMBERING = old
    g = grade_blob(blob, "n.docx")
    r = next(r for r in g.reasons if r.title == "再現できない段落番号")
    assert r.grade == "B" and "arabicAbjad" in r.detail


# ====================================================================== 5. 字の幅
def test_widths_follow_east_asian_width():
    assert cells("あ") == 2 and cells("Ａ") == 2 and cells("A") == 1
    assert cells("Ж") == 1 and cells("Ж", True) == 2  # キリル文字は曖昧（A）：東アジアの文脈だけ全角
    assert cells("ب") == 1 and cells("क") == 1 and cells("́") == 0
    assert text_cells("Имя") == 3 and text_cells("名前") == 4 and text_cells("○有") == 4
    assert text_cells("الاسم", False) == 5 and text_cells("नाम", False) == 3  # アラビア文字・デーヴァナーガリーは 2 ではない


def test_xlsx_circle_marks_proportional_scripts(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    from todoku.book import open_book
    from todoku.xlsx.circle import glyphs, width

    assert width("Д", 10) == 5 and width("Д", 10, True) == 10
    wb_ = openpyxl.Workbook()
    ws = wb_.active
    ws["A1"] = "Да / Нет"
    ws["A2"] = "有・無"
    for c in ("A1", "A2"):
        ws[c].font = openpyxl.styles.Font(name="ＭＳ ゴシック")
    path = tmp_path / "c.xlsx"
    wb_.save(path)
    wb = open_book(path.read_bytes(), "c.xlsx")
    sh = wb.sheets[0]
    _g, notes = glyphs(wb, sh, 1, 1, "Да / Нет")
    assert any("キリル文字" in n for n in notes)
    _g, notes = glyphs(wb, sh, 2, 1, "有・無")
    assert not notes


def test_docx_circle_marks_proportional_scripts():
    from test_docx import make_docx

    from todoku.docx import circle as dc
    from todoku.docx.document import DocxDocument

    doc = DocxDocument(make_docx("<w:p><w:r><w:t>Пол: мужской / женский</w:t></w:r></w:p>"), "c.docx")
    sd = dc.resolve(doc, {"circle": "мужской"}, {})
    assert sd["notes"] and "幅の変わる字" in sd["notes"][0]
    doc = DocxDocument(make_docx("<w:p><w:r><w:t>性別　男・女</w:t></w:r></w:p>"), "c.docx")
    assert dc.resolve(doc, {"circle": "男"}, {})["notes"] == []


# ====================================================================== 6. 日付の入力
def test_ambiguous_day_month_is_refused(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    from todoku.book import open_book
    from todoku.xlsx.edit import EditError, interpret_input

    wb_ = openpyxl.Workbook()
    ws = wb_.active
    ws["A1"].number_format = "dd/mm/yyyy"
    ws["A2"].number_format = "General"
    ws["A1"] = None
    path = tmp_path / "e.xlsx"
    wb_.save(path)
    buf = io.BytesIO(path.read_bytes())
    wb = open_book(buf.getvalue(), "e.xlsx")
    name = wb.sheets[0].name
    if (1, 1) not in wb.sheets[0].cells:
        pytest.skip("空のセルの書式が保存されていません")
    with pytest.raises(EditError, match="ISO"):
        interpret_input(wb, name, 1, 1, "03/11/2026")
    v, notes = interpret_input(wb, name, 1, 1, "03/11/2026", refuse_ambiguous=False)
    assert v == "03/11/2026" and notes
    v, _n = interpret_input(wb, name, 1, 1, "2026-11-03")  # ISO は、セルの表示形式で表示する日付になる
    assert v == float(NOV3)
    v, notes = interpret_input(wb, name, 1, 1, "3 Nov 2026")
    assert v == "3 Nov 2026" and any("ISO" in n for n in notes)


# ====================================================================== 7. 段階の問い
def test_cannot_do_question_is_language_neutral():
    from todoku.grade import AGENT_CHECKS

    q = next(c for c in AGENT_CHECKS if c["id"] == "cannot_do")["question"]
    assert "署名・印・拇印・公証など、手で行う所" in q
