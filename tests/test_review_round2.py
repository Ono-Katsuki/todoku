"""2 回目の見直しの指摘：地域で見え方が変わる日付の試算・Word の日付の欄の言語・Word の□・丸の伝え方など。"""

import datetime as dt

import pytest

from todoku.formula.numfmt import BUILTIN


def _date_book(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "開始日"
    ws["B1"] = dt.date(2026, 11, 1)
    ws["B1"].number_format = "mm-dd-yy"  # openpyxl は組み込みの番号 14 で保存する
    ws["A2"] = "翌日"
    ws["B2"] = "=B1+1"
    ws["B2"].number_format = "mm-dd-yy"
    path = tmp_path / "d.xlsx"
    wb.save(path)
    return path


def test_calculate_shows_locale_dates_as_iso_with_a_note(tmp_path):
    from todoku import api

    r = api.calculate(str(_date_book(tmp_path)), {"Sheet!B1": "2026-11-03"})
    ch = {c["cell"]: c for c in r["changed"]}
    assert ch["Sheet!B2"]["after"] == "2026-11-04"
    assert "番号 14" in ch["Sheet!B2"]["note"] and "地域" in ch["Sheet!B2"]["note"]


def test_explain_shows_locale_dates_as_iso_with_a_note(tmp_path):
    from todoku import api

    from todoku.book import open_book
    from todoku.xlsx import model as xm

    path = _date_book(tmp_path)
    wb = open_book(path.read_bytes(), path.name)
    wb.sheet("Sheet").cells[(2, 2)].value = 46328.0  # Excel が保存する計算済みの値（openpyxl は書かない）
    node = xm.explain(xm.model(wb, api.window(str(path))[1]), "Sheet!B2")
    assert node["value"] == "2026-11-02" and "番号 14" in node["value_note"]


def test_chart_cache_locale_date_is_iso():
    from todoku.embeds.chart import number_text

    assert number_text("46329", BUILTIN[14]) == "2026-11-03"


def test_vba_range_text_of_locale_date_is_iso_and_told():
    from test_vbarun import book, out
    from todoku.xlsx.workbook import Style

    # A1：ファイルの値。A2：マクロが書いた値（計算した値の側の道）
    wb = book('Sub Test()\n    Range("B1").Value = Range("A1").Text\n    Range("A2").Value = 46330\n'
              '    Range("B2").Value = Range("A2").Text\nEnd Sub\n', {"A1": 46329, "A2": 0})
    wb.styles.append(Style(BUILTIN[14], True, False))
    for a in ((1, 1), (2, 1)):
        wb.sheets[0].cells[a].style = len(wb.styles) - 1
    got, r = out(wb)
    assert got["入力!B1"] == "2026-11-03" and got["入力!B2"] == "2026-11-04"
    assert any("入力!A1" in n and "番号 14" in n for n in r.notes)
    assert any("入力!A2" in n and "番号 14" in n for n in r.notes)


# ====================================================================== Word の旧来の日付の欄の言語
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _date_field_docx(fmt: str, *, rpr: str = "", ppr: str = "", styles: str = "", settings: str = "") -> bytes:
    import io
    import zipfile

    body = (f'<w:p>{ppr}<w:r><w:t>日付：</w:t></w:r><w:r>{rpr}<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="D"/><w:enabled/>'
            f'<w:textInput><w:type w:val="date"/><w:format w:val="{fmt}"/></w:textInput></w:ffData></w:fldChar></w:r>'
            '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>     </w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/document.xml", f'<?xml version="1.0" encoding="UTF-8"?><w:document {W}><w:body>{body}</w:body></w:document>')
        if styles:
            z.writestr("word/styles.xml", f'<?xml version="1.0" encoding="UTF-8"?><w:styles {W}>{styles}</w:styles>')
        if settings:
            z.writestr("word/settings.xml", f'<?xml version="1.0" encoding="UTF-8"?><w:settings {W}>{settings}</w:settings>')
    return buf.getvalue()


def _fill_date(blob: bytes):
    from todoku.docx.document import DocxDocument
    from todoku.docx.fill import fill

    doc = DocxDocument(blob)
    return doc.controls[0], fill(doc, {"D": "2026-11-03"})


DEFAULTS_JA = ('<w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr>'
               '</w:rPrDefault></w:docDefaults>')


@pytest.mark.parametrize("kw,where", [
    ({"styles": DEFAULTS_JA}, "docDefaults"),
    ({"styles": '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:rPr><w:lang w:eastAsia="ja-JP"/></w:rPr></w:style>'},
     "段落のスタイル「a」"),
    ({"styles": '<w:style w:type="paragraph" w:styleId="base"><w:rPr><w:lang w:eastAsia="ja-JP"/></w:rPr></w:style>'
                '<w:style w:type="paragraph" w:styleId="b"><w:basedOn w:val="base"/></w:style>',
      "ppr": '<w:pPr><w:pStyle w:val="b"/></w:pPr>'}, "段落のスタイル「b」"),
    ({"styles": '<w:style w:type="character" w:styleId="c"><w:rPr><w:lang w:eastAsia="ja-JP"/></w:rPr></w:style>',
      "rpr": '<w:rPr><w:rStyle w:val="c"/></w:rPr>'}, "字のスタイル「c」"),
    ({"settings": '<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/>'}, "themeFontLang"),
])
def test_legacy_date_language_follows_the_style_hierarchy(kw, where):
    ctl, res = _fill_date(_date_field_docx("ggge年M月d日", **kw))
    assert ctl.date_lid == "ja-JP" and where in ctl.date_lid_from
    assert [p.value for p in res.patches] == ["令和8年11月3日"]
    told = [w for w in res.warnings if "themeFontLang" in w]
    assert bool(told) == (where == "themeFontLang")  # 文書の設定から決めたときは、そう伝える


def test_direct_run_language_wins_over_defaults():
    from todoku.docx.document import DocxDocument

    blob = _date_field_docx("ggge年M月d日", rpr='<w:rPr><w:lang w:eastAsia="zh-TW"/></w:rPr>', styles=DEFAULTS_JA)
    ctl = DocxDocument(blob).controls[0]
    assert ctl.date_lid == "zh-TW" and ctl.date_lid_from == "字の要素の w:lang"


def test_unknown_language_numeric_format_is_written_and_era_is_refused():
    from todoku.docx.fill import FillError

    ctl, res = _fill_date(_date_field_docx("yyyy/MM/dd"))
    assert ctl.date_lid == "" and [p.value for p in res.patches] == ["2026/11/03"]
    # ECMA-376 第1部 17.16.4.1 の日付の記号に g・e はない：元号の記号だけで和暦とは決めない
    with pytest.raises(FillError):
        _fill_date(_date_field_docx("ggge年M月d日"))



# ====================================================================== Word の影つきの白い四角
def test_docx_shadowed_squares_are_question_not_rule_checkboxes():
    from test_docx import make_docx, p
    from todoku.docx.document import DocxDocument
    from todoku.docx.window import AGENT, LAYOUT, Spots

    sp = Spots(DocxDocument(make_docx(p("❑はい　　❑いいえ") + p("□新規　　□変更"))))
    boxes = [it for it in sp.items.values() if it["kind"] == "checkbox"]
    shadow = [it for it in boxes if it["label"] in ("はい", "いいえ")]
    plain = [it for it in boxes if it["label"] in ("新規", "変更")]
    assert len(shadow) == 2 and all(it["provenance"] == AGENT and it["why_key"] == "bullet_box" for it in shadow)
    assert len(plain) == 2 and all(it["provenance"] == LAYOUT for it in plain)


# ====================================================================== 囲んだ所は、いつも伝える
def test_circles_are_always_told_even_when_the_key_was_written():
    from todoku.api import _to_tell

    written = {"S!B1": "有"}
    result = {"circled": [{"key": "S!B1", "sheet": "S", "label": "外出の有無", "text": "有"}]}
    say = _to_tell([{"id": "S!B1", "sheet": "S", "label": "外出の有無", "value": "有　・　無"}], written, result)["say_to_user"]
    assert "○で囲んだ：S・外出の有無：「有」" in say


# ====================================================================== サーバーの立ち上げ：ps の確かめ・UTF-8 の標準出力
def test_missing_ps_is_in_the_instructions_from_the_start(monkeypatch):
    import os
    import subprocess

    from todoku import isolate, mcp_server

    if os.name == "nt":
        pytest.skip("Windows では働き手を分けない")
    monkeypatch.setenv("TODOKU_ISOLATE", "1")
    monkeypatch.setattr(isolate, "_PS_OK", None)
    monkeypatch.setattr(isolate, "_TOLD", set())

    def no_ps(*a, **k):
        raise FileNotFoundError("ps")

    monkeypatch.setattr(subprocess, "run", no_ps)
    text = mcp_server._instructions()  # まだ一度も道具を呼んでいない
    assert "ps が使えないので" in text


def test_server_writes_utf8_lines(monkeypatch):
    import io
    import json
    import sys

    from todoku import mcp_server

    raw = io.BytesIO()
    out = io.TextIOWrapper(raw, encoding="cp1252", newline="\r\n")  # Windows の ANSI のコードページの標準出力
    req = json.dumps({"jsonrpc": "2.0", "id": 1, "method": "tools/list"}).encode() + b"\n"
    monkeypatch.setattr(sys, "stdout", out)
    monkeypatch.setattr(sys, "stdin", io.TextIOWrapper(io.BytesIO(req), encoding="cp1252"))
    monkeypatch.setenv("TODOKU_ISOLATE", "0")
    mcp_server.main()
    out.flush()
    data = raw.getvalue()
    assert data.endswith(b"\n") and not data.endswith(b"\r\n")
    body = data.decode("utf-8")  # cp1252 では書けない日本語の説明も、UTF-8 で出ている
    assert json.loads(body)["result"]["tools"] and "書類" in body
