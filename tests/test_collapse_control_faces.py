"""読み順で「同じ形の行」にまとめても、チェックボックスの面の字が違う行の字を窓から落とさない。"""

from __future__ import annotations

import io
import zipfile

from todoku import api

_CT = ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
       '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
       '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
       '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
       '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')


def _shape(i: int, row: int, face: str) -> str:
    return (f'<v:shape id="_x0000_s{1000 + i}" type="#_x0000_t201"><v:textbox><div>{face}</div></v:textbox>'
            f'<x:ClientData ObjectType="Checkbox"><x:Anchor>0, 5, {row - 1}, 2, 2, 5, {row - 1}, 15</x:Anchor></x:ClientData></v:shape>')


def _xlsx(faces: list[str]) -> bytes:
    rows = "".join(f'<row r="{r}"><c r="C{r}" t="inlineStr"><is><t>同じ字</t></is></c></row>' for r in range(1, len(faces) + 1))
    vml = ('<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">'
           + "".join(_shape(i, i + 1, f) for i, f in enumerate(faces)) + "</xml>")
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _CT)
        z.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml", '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        z.writestr("xl/worksheets/sheet1.xml", '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                   f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>{rows}</sheetData><legacyDrawing r:id="rId1"/></worksheet>')
        z.writestr("xl/worksheets/_rels/sheet1.xml.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/></Relationships>')
        z.writestr("xl/drawings/vmlDrawing1.vml", vml)
    return b.getvalue()


def _window(path) -> str:
    return "\n".join(s for _n, s in api.window(str(path))[2])


def test_rows_with_different_faces_stay_in_the_window(tmp_path):
    faces = ["工場 A", "工場 B", "工場 C", "工場 D", "工場 E"]
    src = tmp_path / "a.xlsx"
    src.write_bytes(_xlsx(faces))
    text = _window(src)
    for f in faces:
        assert f in text


def test_rows_with_the_same_face_may_still_collapse(tmp_path):
    src = tmp_path / "a.xlsx"
    src.write_bytes(_xlsx(["同じ面"] * 5))
    assert "同じ形の行が" in _window(src)


def test_button_face_is_in_the_window(tmp_path):
    blob = _xlsx(["同じ面"])
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(blob)) as zi, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            b = zi.read(it.filename)
            if it.filename.endswith(".vml"):
                b = b.replace(b'ObjectType="Checkbox"', b'ObjectType="Button"').replace("同じ面".encode(), "印刷する".encode())
            zo.writestr(it, b)
    src = tmp_path / "a.xlsx"
    src.write_bytes(out.getvalue())
    assert "印刷する" in _window(src)
