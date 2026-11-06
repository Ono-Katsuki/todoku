"""ふりがな：ファイルに保存された読みの単位をそのまま出す。非表示のシートのふりがなも窓に出す。"""

from __future__ import annotations

import io
import zipfile

from todoku import api

_SST = ('<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<si><t>サービス等の利用</t><rPh sb="0" eb="8"><t>さーびすとうのりよう</t></rPh></si>'
        '<si><t>年号を選択</t><rPh sb="0" eb="2"><t>ネンゴウ</t></rPh><rPh sb="3" eb="5"><t>センタク</t></rPh></si></sst>')


def _xlsx(hidden: bool) -> bytes:
    st = ' state="hidden"' if hidden else ""
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
                   '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                   '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml", '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="A" sheetId="1" r:id="rId1"/>'
                   f'<sheet name="B" sheetId="2" r:id="rId2"{st}/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
                   '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>')
        z.writestr("xl/sharedStrings.xml", _SST)
        z.writestr("xl/worksheets/sheet1.xml", '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                   '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>')
        z.writestr("xl/worksheets/sheet2.xml", '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                   '<sheetData><row r="1"><c r="A1" t="s"><v>1</v></c></row></sheetData></worksheet>')
    return b.getvalue()


def _window(tmp_path, hidden: bool) -> str:
    src = tmp_path / "a.xlsx"
    src.write_bytes(_xlsx(hidden))
    return "\n".join(s for _n, s in api.window(str(src))[2])


def test_stored_reading_is_not_cut_where_it_shares_kana_with_the_text(tmp_path):
    assert "さーびすとうのりよう" in _window(tmp_path, False)


def test_each_stored_unit_is_shown(tmp_path):
    assert "ネンゴウ・センタク" in _window(tmp_path, False)


def test_furigana_of_a_hidden_sheet_is_in_the_window(tmp_path):
    assert "ネンゴウ・センタク" in _window(tmp_path, True)
