"""字のない図形を、種類・位置・大きさの事実として窓に出し、その上に字を重ねて書ける（チェックだとは決めつけない）。"""

from __future__ import annotations

import io
import zipfile

import pytest

from todoku import api
from todoku.xlsx.workbook import Workbook

DRAWING = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
           '<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>161925</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>104775</xdr:rowOff></xdr:from>'
           '<xdr:to><xdr:col>0</xdr:col><xdr:colOff>285750</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>47625</xdr:rowOff></xdr:to>'
           '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2129" name="Rectangle 4"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="1" y="2"/>'
           '<a:ext cx="123825" cy="133350"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
           '<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>')


def _xlsx(drawing: str = DRAWING) -> bytes:
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                   '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml", '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        z.writestr("xl/worksheets/sheet1.xml", '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="2"><c r="A2" t="inlineStr"><is><t>　　移行を希望します。</t></is></c></row></sheetData>'
                   '<drawing r:id="rId1"/></worksheet>')
        z.writestr("xl/worksheets/_rels/sheet1.xml.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>')
        z.writestr("xl/drawings/drawing1.xml", drawing)
    return b.getvalue()


def _window(path) -> str:
    _k, _d, parts = api.window(str(path))
    return "\n".join(s for _n, s in parts)


def test_bare_shape_facts_are_in_the_workbook_and_window(tmp_path):
    wb = Workbook(_xlsx(), "a.xlsx")
    s = wb.sheets[0].shapes[0]
    assert (s.sid, s.geom, s.ext, s.fill, s.line) == ("2129", "rect", (123825, 133350), "FFFFFF", "000000")
    src = tmp_path / "a.xlsx"
    src.write_bytes(_xlsx())
    text = _window(src)
    line = next(x for x in text.splitlines() if "字のない図形 2129" in x)
    assert "Rectangle 4" in line and "A2〜A3" in line and "種類 rect" in line and "3.4×3.7 mm" in line
    assert "左から 4.5 mm" in line and "塗り #FFFFFF" in line and "線 #000000" in line
    assert "移行を希望します。" in line and "A2" in line
    assert "チェック" not in line  # 決めつけない
    assert "on_shape" in text


def test_mark_is_written_on_top_without_changing_the_shape(tmp_path):
    src = tmp_path / "a.xlsx"
    src.write_bytes(_xlsx())
    out = tmp_path / "b.xlsx"
    r = api.fill(str(src), {"@m1": {"on_shape": "2129", "sheet": "S", "mark": "✓"}}, str(out), overwrite=True)
    assert r["proof_ok"] and r["read_back_ok"], r
    assert r["circled"][0]["style"] == "mark" and "uncertain" not in r["circled"][0]
    assert "図形の上に字を書いた" in r["say_to_user"]
    wb = Workbook(out.read_bytes(), "b.xlsx")
    base, mine = wb.sheets[0].shapes
    assert (mine.text, mine.frm, mine.to) == ("✓", base.frm, base.to) and base.geom == "rect"
    a, b = zipfile.ZipFile(src), zipfile.ZipFile(out)
    assert a.read("xl/worksheets/sheet1.xml") == b.read("xl/worksheets/sheet1.xml")
    assert b.read("xl/drawings/drawing1.xml").startswith(a.read("xl/drawings/drawing1.xml")[:-len(b"</xdr:wsDr>")])
    assert "✓" in _window(out)
    with pytest.raises(api.ApiError) as e:  # 同じ図形に 2 度は書かない
        api.fill(str(out), {"@m1": {"on_shape": "2129", "mark": "×"}}, str(tmp_path / "c.xlsx"), overwrite=True)
    assert "もうあります" in str(e.value)


@pytest.mark.parametrize("spec,word", [({"on_shape": "9", "mark": "✓"}, "ありません"), ({"on_shape": "2129"}, "mark"),
                                       ({"on_shape": "2129", "mark": "✓", "x": 1}, "分からない")])
def test_bad_specs_are_refused(tmp_path, spec, word):
    src = tmp_path / "a.xlsx"
    src.write_bytes(_xlsx())
    with pytest.raises(api.ApiError) as e:
        api.fill(str(src), {"@m1": spec}, str(tmp_path / "b.xlsx"), overwrite=True)
    assert word in str(e.value)


def test_checkbox_caption_is_in_the_reading_order():
    from types import SimpleNamespace

    from todoku.xlsx.window import Box, Layout

    lay = Layout.__new__(Layout)
    ctl = SimpleNamespace(text="A.\\r\\n   Nombre comercial".replace("\\r\\n", "\r\n"), link_cell=None, link=None, name="Check Box 1")
    b = Box(1, 1, 1, 1, "check", None, "☐", ctl)
    assert lay._token(b).endswith("A. Nombre comercial") and lay._token(b).startswith("☐［")
    assert "Nombre" not in lay._token(Box(1, 1, 1, 1, "check", None, "☐", SimpleNamespace(text="", link_cell=None, link=None, name="C")))
