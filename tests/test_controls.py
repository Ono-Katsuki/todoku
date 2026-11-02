"""Excel のフォームの部品（オプションボタン・ドロップダウン・リストボックス・スピンボタン）を読み書きする。
ファイルは MS-XLSX 2.6.65 CT_FormControlPr と VML の x:ClientData のとおりに組み立てる。"""

from __future__ import annotations

import io
import zipfile

import pytest

from todoku.xlsx.audit import run as audit
from todoku.xlsx.calc import verify
from todoku.xlsx.edit import EditError, fill
from todoku.xlsx.proof import prove
from todoku.xlsx.window import build
from todoku.xlsx.workbook import Workbook

NS = ('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
      'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
      'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"')
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"

# (形, 名前, 左上・右下（列, 行。0から）, ctrlProp の属性, VML の中身)
PARTS = [
    ("GBox", "Group Box 1", (5, 1, 9, 6), 'objectType="GBox" noThreeD="1"', ""),
    ("Radio", "Option Button 2", (6, 2, 8, 3), 'objectType="Radio" checked="Checked" firstButton="1" fmlaLink="$C$1" lockText="1"',
     "<x:Checked>1</x:Checked><x:FmlaLink>$C$1</x:FmlaLink><x:FirstButton/>"),
    ("Radio", "Option Button 3", (6, 4, 8, 5), 'objectType="Radio" fmlaLink="$C$1" lockText="1"', "<x:FmlaLink>$C$1</x:FmlaLink>"),
    ("Drop", "Drop Down 4", (1, 8, 3, 9), 'objectType="Drop" dropStyle="combo" dx="16" fmlaLink="$D$1" fmlaRange="$A$1:$A$3" sel="1" val="0"',
     "<x:FmlaLink>$D$1</x:FmlaLink><x:FmlaRange>$A$1:$A$3</x:FmlaRange><x:Sel>1</x:Sel><x:Val>0</x:Val>"),
    ("List", "List Box 5", (1, 11, 3, 14), 'objectType="List" dx="16" fmlaRange="$A$1:$A$3" selType="multi" multiSel="1"',
     "<x:FmlaRange>$A$1:$A$3</x:FmlaRange><x:SelType>Multi</x:SelType><x:MultiSel>1</x:MultiSel>"),
    ("Spin", "Spinner 6", (1, 16, 2, 18), 'objectType="Spin" dx="16" fmlaLink="$E$1" max="10" min="0" page="10" val="3"',
     "<x:FmlaLink>$E$1</x:FmlaLink><x:Val>3</x:Val><x:Min>0</x:Min><x:Max>10</x:Max><x:Page>10</x:Page>"),
]
VML_TYPE = {"GBox": "GBox", "Radio": "Radio", "Drop": "Drop", "List": "List", "Spin": "Spin"}
TEXT = {"Group Box 1": "性別", "Option Button 2": "男", "Option Button 3": "女"}


def make_book(radio_link_value: float = 1.0) -> bytes:
    cells = ('<row r="1"><c r="A1" t="inlineStr"><is><t>東京</t></is></c><c r="C1"><v>{rv}</v></c><c r="D1"><v>1</v></c>'
             '<c r="E1"><v>3</v></c><c r="F1"><f>C1*100+D1*10+E1</f><v>{fv}</v></c></row>'
             '<row r="2"><c r="A2" t="inlineStr"><is><t>大阪</t></is></c></row>'
             '<row r="3"><c r="A3" t="inlineStr"><is><t>名古屋</t></is></c></row>').format(
        rv=int(radio_link_value), fv=int(radio_link_value * 100 + 10 + 3))
    ctl_xml = ""
    rels = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" '
            'Target="../drawings/vmlDrawing1.vml"/>']
    shapes = []
    for i, (kind, name, (c1, r1, c2, r2), attrs, inner) in enumerate(PARTS, 1):
        sid = 1024 + i
        rid = f"rId{i + 1}"
        rels.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" '
                    f'Target="../ctrlProps/ctrlProp{i}.xml"/>')
        ctl_xml += (f'<mc:AlternateContent><mc:Choice Requires="x14"><control shapeId="{sid}" r:id="{rid}" name="{name}">'
                    f'<controlPr defaultSize="0" autoFill="0" autoLine="0" autoPict="0"><anchor moveWithCells="1">'
                    f'<from><xdr:col>{c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r1}</xdr:row><xdr:rowOff>0</xdr:rowOff></from>'
                    f'<to><xdr:col>{c2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r2}</xdr:row><xdr:rowOff>0</xdr:rowOff></to>'
                    f'</anchor></controlPr></control></mc:Choice></mc:AlternateContent>')
        text = TEXT.get(name, "")
        tb = f'<v:textbox><div><font>{text}</font></div></v:textbox>' if text else ""
        shapes.append(f'<v:shape id="_x0000_s{sid}" type="#_x0000_t201" style="position:absolute">{tb}\n'
                      f'  <x:ClientData ObjectType="{VML_TYPE[kind]}">\n   <x:Anchor>{c1}, 0, {r1}, 0, {c2}, 0, {r2}, 0</x:Anchor>\n'
                      + "".join(f"   {x}\n" for x in inner.replace("><", ">\n<").split("\n") if x) +
                      '  </x:ClientData>\n </v:shape>')
    sheet = (f'<?xml version="1.0" encoding="UTF-8"?><worksheet {NS}><sheetData>{cells}</sheetData>'
             f'<legacyDrawing r:id="rId1"/><mc:AlternateContent><mc:Choice Requires="x14"><controls>{ctl_xml}</controls>'
             f'</mc:Choice></mc:AlternateContent></worksheet>')
    vml = ('<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
           'xmlns:x="urn:schemas-microsoft-com:office:excel">\n' + "\n".join(shapes) + "\n</xml>")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        over = "".join(f'<Override PartName="/xl/ctrlProps/ctrlProp{i}.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>'
                       for i in range(1, len(PARTS) + 1))
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                   + over + '</Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                   '</Relationships>')
        z.writestr("xl/workbook.xml", f'<?xml version="1.0" encoding="UTF-8"?><workbook {NS}><sheets>'
                   '<sheet name="申込" sheetId="1" r:id="rId1"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                   '</Relationships>')
        z.writestr("xl/worksheets/sheet1.xml", sheet)
        z.writestr("xl/worksheets/_rels/sheet1.xml.rels", '<?xml version="1.0" encoding="UTF-8"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "".join(rels) + '</Relationships>')
        z.writestr("xl/drawings/vmlDrawing1.vml", vml)
        for i, (_k, _n, _a, attrs, _v) in enumerate(PARTS, 1):
            z.writestr(f"xl/ctrlProps/ctrlProp{i}.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                       f'<formControlPr xmlns="{X14}" {attrs}/>')
    return buf.getvalue()


def _fields(blob):
    return {f["id"]: f for sh in build(Workbook(blob))["sheets"] for f in sh["fields"]}


def test_controls_are_fields_with_options():
    f = _fields(make_book())
    assert f["申込!C1"]["kind"] == "radio" and f["申込!C1"]["options"] == ["男", "女"] and f["申込!C1"]["value"] == "男"
    assert f["申込!C1"]["label"] == "性別"  # グループボックスの見出し
    assert f["申込!D1"]["kind"] == "select" and f["申込!D1"]["options"] == ["東京", "大阪", "名古屋"] and f["申込!D1"]["value"] == "東京"
    assert f["申込!List Box 5"]["kind"] == "multiselect" and f["申込!List Box 5"]["value"] == ["東京"]
    assert f["申込!E1"]["kind"] == "spin" and (f["申込!E1"]["min"], f["申込!E1"]["max"], f["申込!E1"]["value"]) == (0, 10, 3)


def test_write_controls_updates_cells_states_and_formulas():
    blob = make_book()
    res = fill(Workbook(blob), {"申込!C1": "女", "申込!D1": "名古屋", "申込!List Box 5": ["大阪", "名古屋"], "申込!E1": 7})
    pr = prove(blob, res.blob, {(p.sheet, p.address) for p in res.patches if p.kind == "input"},
               {(p.sheet, p.address) for p in res.patches if p.kind == "control"})
    assert pr.ok, (pr.unexpected, pr.lint_errors)
    after = Workbook(res.blob)
    sh = after.sheet("申込")
    assert sh.cells[(1, 3)].value == 2 and sh.cells[(1, 4)].value == 3 and sh.cells[(1, 5)].value == 7
    assert sh.cells[(1, 6)].value == 2 * 100 + 3 * 10 + 7  # つながったセルを使う数式も計算し直す
    assert verify(after).summary()["mismatch"] == 0
    f = _fields(res.blob)
    assert f["申込!C1"]["value"] == "女" and f["申込!D1"]["value"] == "名古屋" and f["申込!E1"]["value"] == 7
    assert f["申込!List Box 5"]["value"] == ["大阪", "名古屋"]
    st = {c.name: c for c in sh.controls}
    assert st["Option Button 2"].checked is None and st["Option Button 3"].checked == "Checked"


def test_write_refuses_values_outside_the_control():
    with pytest.raises(EditError):
        fill(Workbook(make_book()), {"申込!D1": "福岡"})
    with pytest.raises(EditError):
        fill(Workbook(make_book()), {"申込!E1": 11})


def test_inconsistent_radio_group_is_left_undecided():
    # つながったセルの値（2）と、選ばれたボタン（1番目）が食い違う：組の順を規則では決められないので、書かない
    blob = make_book(radio_link_value=2)
    f = _fields(blob)
    assert "undecided" in f["申込!C1"]
    with pytest.raises(EditError):
        fill(Workbook(blob), {"申込!C1": "女"})


def test_audit_writes_and_reads_back_every_control():
    au = audit(make_book(), "controls.xlsx")
    bad = [(c.name, c.detail, c.items[:5]) for c in au.checks if not c.ok]
    assert not bad, bad
