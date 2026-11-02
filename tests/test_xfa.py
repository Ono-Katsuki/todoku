"""XFA のフォーム（Adobe XML Forms Architecture 3.3）の読み書きのテスト。

外部のファイルは使わず、テストの中で小さな XFA の PDF（動的・静的・datasets のないもの・1つのストリームのもの・暗号化したもの）を
組み立てて確かめる。集めた公開の様式（corpus/wide_xfa）があれば、それでも確かめる。
"""

from __future__ import annotations

import zlib
from pathlib import Path

import pytest

from todoku.pdf import xfa
from todoku.pdf.audit import run as audit
from todoku.pdf.doc import PdfDocument
from todoku.pdf.fill import FillError, Update, fill
from todoku.pdf.form import read_fields
from todoku.pdf.objects import Name, PdfString, Ref, Stream, dumps
from todoku.pdf.proof import prove
from todoku.pdf.window import build, markdown_parts
from todoku.pdf.xfapicture import Locale, conforms, format_value, parse_value

CORPUS = Path(__file__).resolve().parent.parent / "corpus"
WIDE = sorted((CORPUS / "wide_xfa").rglob("*.pdf")) if (CORPUS / "wide_xfa").exists() else []

TEMPLATE = """<template xmlns="http://www.xfa.org/schema/xfa-template/3.3/">
<subform name="form1" layout="tb" locale="en_US">
 <pageSet><pageArea name="Page1"><contentArea w="8in" h="10in"/><medium stock="letter" short="8.5in" long="11in"/>
   <draw name="Footer"><value><text>Form T-1 (2026)</text></value></draw></pageArea></pageSet>
 <proto><draw name="Style" id="sty"><font typeface="Arial"/><value><text>Unused proto text</text></value></draw></proto>
 <subform name="Page1" layout="tb">
  <draw name="Title" use="#sty"><value><text>Application for a library card</text></value></draw>
  <draw name="Intro"><value><exData contentType="text/html"><body xmlns="http://www.w3.org/1999/xhtml"><p>Please write in
   <span style="font-weight:bold">block letters</span>.</p><p>Second paragraph</p></body></exData></value></draw>
  <field name="LastName"><ui><textEdit/></ui><caption><value><text>Family name</text></value></caption>
   <value><text maxChars="10"/></value><assist><toolTip>Your family name</toolTip><speak>Family name as on passport</speak></assist>
   <validate nullTest="error"/></field>
  <field name="Age"><ui><numericEdit/></ui><caption><value><text>Age</text></value></caption><value><integer/></value></field>
  <field name="Birth"><ui><dateTimeEdit/><picture>date{MM/DD/YYYY}</picture></ui><caption><value><text>Date of birth</text></value></caption>
   <value><date/></value><format><picture>date{MM/DD/YYYY}</picture></format></field>
  <field name="Member"><ui><checkButton/></ui><caption><value><text>Already a member</text></value></caption>
   <items><integer>1</integer><integer>0</integer></items></field>
  <exclGroup name="Gender"><caption><value><text>Gender</text></value></caption>
   <field name="M"><ui><checkButton shape="round"/></ui><caption><value><text>Male</text></value></caption><items><text>M</text></items></field>
   <field name="F"><ui><checkButton shape="round"/></ui><caption><value><text>Female</text></value></caption><items><text>F</text></items></field>
  </exclGroup>
  <field name="Color"><ui><choiceList/></ui><caption><value><text>Favourite colour</text></value></caption>
   <items><text>Red</text><text>Blue</text></items><items save="1" presence="hidden"><text>R</text><text>B</text></items></field>
  <field name="City"><ui><choiceList/></ui><caption><value><text>City</text></value></caption>
   <bindItems ref="!Lists.City[*]" labelRef="$" valueRef="code"/></field>
  <field name="Zip"><ui><textEdit/></ui><caption><value><text>Postal code</text></value></caption>
   <validate formatTest="error"><picture>text{99999}</picture></validate></field>
  <field name="Total"><ui><numericEdit/></ui><caption><value><text>Total</text></value></caption>
   <calculate><script contentType="application/x-javascript">this.rawValue = Age.rawValue * 2;</script></calculate></field>
  <field name="Note"><ui><textEdit/></ui><caption><value><text>Note</text></value></caption><bind match="none"/></field>
  <field name="Secret" presence="hidden"><ui><textEdit/></ui></field>
  <subform><field name="Phone"><ui><textEdit/></ui><caption><value><text>Phone</text></value></caption></field></subform>
  <subform name="Row" layout="tb"><occur min="1" max="3"/>
   <field name="Item"><ui><textEdit/></ui><caption><value><text>Item</text></value></caption></field></subform>
  <field name="Print"><ui><button/></ui><caption><value><text>Print</text></value></caption>
   <event activity="click"><script contentType="application/x-javascript">xfa.host.print(1, "0", "0", 0, 0, 0, 0, 0);</script></event></field>
  <field name="Sig"><ui><signature/></ui></field>
 </subform>
 <event activity="initialize"><script contentType="application/x-javascript">// starts</script></event>
</subform>
</template>"""

DATASETS = """<xfa:datasets xmlns:xfa="http://www.xfa.org/schema/xfa-data/1.0/"><xfa:data><form1><Page1><LastName>Suzuki</LastName><Age/>
<Row><Item>apple</Item></Row><Row><Item>banana &amp; pear</Item></Row><Extra>kept value</Extra></Page1></form1></xfa:data>
<Lists><City code="T">Tokyo</City><City code="O">Osaka</City></Lists></xfa:datasets>"""

CONFIG_DYN = """<config xmlns="http://www.xfa.org/schema/xci/3.1/"><present><pdf><version>1.7</version></pdf></present>
<acrobat><acrobat7><dynamicRender>required</dynamicRender></acrobat7></acrobat></config>"""
CONFIG_STATIC = """<config xmlns="http://www.xfa.org/schema/xci/3.1/"><present><pdf><version>1.7</version></pdf></present></config>"""
PRE = '<?xml version="1.0" encoding="UTF-8"?>\n<xdp:xdp xmlns:xdp="http://ns.adobe.com/xdp/">'
POST = "</xdp:xdp>"
ID = "form1[0].Page1[0]."


def make_xfa_pdf(template=TEMPLATE, datasets=DATASETS, *, dynamic=True, single=False, acro=None, enc=None) -> bytes:
    """XFA の PDF：部品は名前・ストリームの配列（single なら XDP 全体の1つのストリーム）。acro は PDF の欄（名前 → 種類）。"""
    packets = [("preamble", PRE), ("config", CONFIG_DYN if dynamic else CONFIG_STATIC), ("template", template)]
    if datasets is not None:
        packets.append(("datasets", datasets))
    packets.append(("postamble", POST))
    objs: dict[int, object] = {}
    content = b"BT /F1 12 Tf 72 720 Td (Please wait...) Tj ET" if dynamic else b"BT /F1 12 Tf 72 720 Td (Library card) Tj ET"
    objs[1] = {"Type": Name("Catalog"), "Pages": Ref(2), "AcroForm": Ref(5)}
    if dynamic:
        objs[1]["NeedsRendering"] = True
    objs[2] = {"Type": Name("Pages"), "Kids": [Ref(3)], "Count": 1}
    objs[3] = {"Type": Name("Page"), "Parent": Ref(2), "MediaBox": [0, 0, 612, 792], "Contents": Ref(4),
               "Resources": {"Font": {"F1": Ref(6)}}}
    objs[4] = Stream({}, content)
    objs[6] = {"Type": Name("Font"), "Subtype": Name("Type1"), "BaseFont": Name("Helvetica"), "Encoding": Name("WinAnsiEncoding")}
    n = 7
    if single:
        objs[n] = Stream({"Filter": Name("FlateDecode")}, zlib.compress("".join(t for _, t in packets).encode()))
        xfa_entry = Ref(n)
        n += 1
    else:
        xfa_entry = []
        for name, text in packets:
            objs[n] = Stream({"Filter": Name("FlateDecode")}, zlib.compress(text.encode()))
            xfa_entry += [PdfString(name.encode()), Ref(n)]
            n += 1
    fields = []
    annots = []
    for i, (fname, kind) in enumerate((acro or {}).items()):
        y = 600 - 30 * i
        w = {"Type": Name("Annot"), "Subtype": Name("Widget"), "T": PdfString(fname.encode()), "Rect": [150, y, 400, y + 20],
             "P": Ref(3), "F": 4}
        if kind == "text":
            w.update({"FT": Name("Tx"), "DA": PdfString(b"/Helv 10 Tf 0 g")})
        else:
            w.update({"FT": Name("Btn"), "V": Name("Off"), "AS": Name("Off"),
                      "AP": {"N": {"1": Ref(n + 1), "Off": Ref(n + 1)}}})
        objs[n] = w
        if kind != "text":
            objs[n + 1] = Stream({"Type": Name("XObject"), "Subtype": Name("Form"), "BBox": [0, 0, 20, 20]}, b"")
        fields.append(Ref(n))
        annots.append(Ref(n))
        n += 2
    if annots:
        objs[3]["Annots"] = annots
    objs[5] = {"Fields": fields, "XFA": xfa_entry, "DA": PdfString(b"/Helv 0 Tf 0 g"),
               "DR": {"Font": {"Helv": Ref(6)}}}
    out = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
    offs = {}
    for num in sorted(objs):
        o = objs[num]
        offs[num] = len(out)
        if isinstance(o, Stream):
            data, d = o.raw, dict(o.dict)
            if enc is not None:
                data, d = enc.crypt(enc.stm, data, num, 0), enc.strings(d, num)
            d["Length"] = len(data)
            body = dumps(d) + b"\nstream\n" + data + b"\nendstream"
        else:
            body = dumps(enc.strings(o, num) if enc is not None else o)
        out += f"{num} 0 obj\n".encode() + body + b"\nendobj\n"
    size = max(objs) + 1
    trailer = {"Size": size, "Root": Ref(1)}
    if enc is not None:
        from test_pdf_password import FILE_ID

        offs[size] = len(out)
        out += f"{size} 0 obj\n".encode() + dumps(enc.dict) + b"\nendobj\n"
        trailer["Encrypt"] = Ref(size)
        trailer["Size"] = size + 1
        ids = [PdfString(FILE_ID), PdfString(FILE_ID)]
        for s_ in ids:
            s_.hexform = True
        trailer["ID"] = ids
    xref = len(out)
    out += f"xref\n0 {trailer['Size']}\n0000000000 65535 f \n".encode()
    for num in range(1, trailer["Size"]):
        out += (f"{offs[num]:010d} 00000 n \n" if num in offs else "0000000000 00000 f \n").encode()
    out += b"trailer\n" + dumps(trailer) + f"\nstartxref\n{xref}\n%%EOF\n".encode()
    return bytes(out)


def fields_of(blob: bytes, password=None) -> dict:
    win = build(PdfDocument(blob, "x.pdf", password=password) if password else PdfDocument(blob, "x.pdf"))
    return {f["id"]: f for f in win["xfa"]["fields"]}


def packets(blob: bytes, password=None) -> dict[str, bytes]:
    doc = PdfDocument(blob, password=password) if password else PdfDocument(blob)
    return {p.name: p.data for p in xfa.Xdp(doc).packets}


# ====================================================================== 読む
def test_read_dynamic_form():
    blob = make_xfa_pdf()
    win = build(PdfDocument(blob, "lib.pdf"))
    xw = win["xfa"]
    assert xw["readable"] and xw["kind"] == "dynamic" and win["form"]["xfa_kind"] == "dynamic"
    f = {d["id"]: d for d in xw["fields"]}
    assert f[ID + "LastName[0]"]["value"] == "Suzuki" and f[ID + "LastName[0]"]["label"] == "Family name"
    assert f[ID + "LastName[0]"]["maxlen"] == 10 and f[ID + "LastName[0]"]["required"]
    assert f[ID + "Row[1].Item[0]"]["value"] == "banana & pear"
    assert ID + "Row[2].Item[0]" not in f  # データの数だけ（min 1・max 3）
    assert f[ID + "#subform[0].Phone[0]"]["kind"] == "text"  # 名前のない subform は #subform[番号]
    assert f[ID + "Gender[0]"]["kind"] == "radio" and f[ID + "Gender[0]"]["options"] == ["M", "F"]
    assert f[ID + "Gender[0]"]["option_labels"] == ["Male", "Female"]
    assert f[ID + "Color[0]"]["options"] == ["Red", "Blue"] and f[ID + "Color[0]"]["export_values"] == ["R", "B"]
    assert f[ID + "City[0]"]["options"] == ["Tokyo", "Osaka"] and f[ID + "City[0]"]["export_values"] == ["T", "O"]
    assert f[ID + "Member[0]"]["value"] is False and f[ID + "Member[0]"]["on_value"] == "1"
    assert f[ID + "Note[0]"]["readonly"] and f[ID + "Note[0]"]["no_data"]
    assert f[ID + "Total[0]"]["readonly"] and f[ID + "Total[0]"]["calculated"] == "error"
    assert f[ID + "Secret[0]"]["hidden"]
    assert f[ID + "Sig[0]"]["readonly"] and f[ID + "Print[0]"]["readonly"]
    reading = "\n".join(xw["reading"])
    assert "Application for a library card" in reading and "Please write in block letters." in reading
    assert "Second paragraph" in reading and "Family name ［form1[0].Page1[0].LastName[0]＝Suzuki］" in reading
    assert any("Form T-1 (2026)" in s for s in xw["master"])
    assert any("Extra[0]：kept value" in s for s in xw["unbound_data"])
    assert any("Tokyo" in s and "code=T" in s for s in xw["lists"])
    assert any("Unused proto text" in s for s in xw["unshown"])
    assert {s["language"] for s in xw["scripts"]} == {"JavaScript"} and len(xw["scripts"]) == 3
    md = "\n".join(t for _, t in markdown_parts(win))
    assert "XFA のフォーム（動的なフォーム" in md and "this.rawValue = Age.rawValue * 2;" in md
    # 動的なフォームでは、ページの描画（代わりの表示）から入力欄を推定しない
    assert all(f_.get("source") != "layout" for pg in win["pages"] for f_ in pg["fields"])


def test_audit_dynamic_form():
    au = audit(make_xfa_pdf(), "lib.pdf")
    checks = {c.name: c for c in au.checks}
    assert checks["読み落とし（XFA の字）"].ok, checks["読み落とし（XFA の字）"].items
    assert checks["全入力欄への書き込み"].ok, checks["全入力欄への書き込み"].items
    assert checks["読み戻し"].ok, checks["読み戻し"].items
    assert checks["フォームの仕掛け"].ok and checks["フォームの仕掛け"].agent  # 自動で動くスクリプトがある


def test_reading_check_finds_missing_text():
    from todoku.pdf.audit import check_xfa

    doc = PdfDocument(make_xfa_pdf())
    win = build(doc)
    win["xfa"]["reading"] = [s for s in win["xfa"]["reading"] if "Second paragraph" not in s]
    win["xfa"]["_plain"] = [s for s in win["xfa"]["_plain"] if "Second paragraph" not in s]
    c = check_xfa(doc, win)[0]
    assert not c.ok and any("Second paragraph" in i for i in c.items)


# ====================================================================== 書く
def test_write_dynamic_form_and_proof():
    blob = make_xfa_pdf()
    values = {ID + "LastName[0]": "Tanaka", ID + "Age[0]": "42", ID + "Birth[0]": "04/01/1990", ID + "Member[0]": True,
              ID + "Gender[0]": "Female", ID + "Color[0]": "Blue", ID + "City[0]": "Osaka", ID + "Zip[0]": "12345",
              ID + "#subform[0].Phone[0]": "03-1234", ID + "Row[1].Item[0]": "cherry <ripe>"}
    res = fill(PdfDocument(blob), values)
    assert res.blob[:len(blob)] == blob
    pr = prove(blob, res.blob, set(values))
    assert pr.ok, (pr.unexpected, pr.lint_errors)
    assert pr.xfa["identical_packets"] == 4 and pr.xfa["data_changes"] >= len(values)
    a, b = packets(blob), packets(res.blob)
    assert a["template"] == b["template"] and a["config"] == b["config"]
    got = fields_of(res.blob)
    assert got[ID + "LastName[0]"]["value"] == "Tanaka" and got[ID + "Age[0]"]["value"] == "42"
    assert got[ID + "Birth[0]"]["value"] == "1990-04-01" and got[ID + "Birth[0]"]["display"] == "04/01/1990"
    assert got[ID + "Member[0]"]["value"] is True and got[ID + "Gender[0]"]["value"] == "F"
    assert got[ID + "Color[0]"]["value"] == "B" and got[ID + "City[0]"]["value"] == "O"
    assert got[ID + "Row[1].Item[0]"]["value"] == "cherry <ripe>" and got[ID + "Row[0].Item[0]"]["value"] == "apple"
    ds = b["datasets"]
    assert b"<Birth>1990-04-01</Birth>" in ds and b"<Item>cherry &lt;ripe&gt;</Item>" in ds
    assert b"<Extra>kept value</Extra>" in ds and b'<City code="T">Tokyo</City>' in ds
    # 書いた所のほかは、元のバイト列のまま
    assert ds.startswith(a["datasets"][:a["datasets"].index(b"Suzuki")])
    written = {p.id: p.value for p in res.patches}
    assert written[ID + "Birth[0]"] == "1990-04-01" and written[ID + "Gender[0]"] == "F"


@pytest.mark.parametrize("key, value, why", [
    ("LastName[0]", "Abcdefghijk", "10 字まで"),
    ("Age[0]", "4.5", "整数"),
    ("Birth[0]", "next week", "日付"),
    ("Color[0]", "Green", "選択肢にありません"),
    ("Zip[0]", "12a45", "書式 text{99999}"),
    ("Note[0]", "x", "データに結びつかない"),
    ("Total[0]", "3", "計算"),
    ("Member[0]", "maybe", "はい／いいえ"),
    ("Gender[0]", "X", "選択肢にありません"),
])
def test_write_refuses_constraint_violations(key, value, why):
    with pytest.raises(FillError) as e:
        fill(PdfDocument(make_xfa_pdf()), {ID + key: value})
    assert why in str(e.value)


def test_write_refuses_places_outside_xfa():
    with pytest.raises(FillError) as e:
        fill(PdfDocument(make_xfa_pdf()), {"@1": {"after": "Please", "value": "x"}})
    assert "XFA のフォーム" in str(e.value)


def test_proof_detects_other_data_and_template_changes():
    blob = make_xfa_pdf()
    res = fill(PdfDocument(blob), {ID + "LastName[0]": "Tanaka"})
    doc = PdfDocument(res.blob)
    x = xfa.Xdp(doc)
    ds = next(p for p in x.packets if p.name == "datasets")
    tp = next(p for p in x.packets if p.name == "template")
    for ref, data, words in ((ds.ref, ds.data.replace(b"kept value", b"changed!!"), "書くと言っていないデータ"),
                             (tp.ref, tp.data.replace(b"library card", b"library pass"), "template が変わっています")):
        up = Update(doc)
        up.put(ref, Stream({"Filter": Name("FlateDecode")}, zlib.compress(data)))
        pr = prove(blob, up.build(), {ID + "LastName[0]"})
        assert not pr.ok and any(words in u for u in pr.unexpected), pr.unexpected
    # 書くと言っていない欄
    pr = prove(blob, res.blob, {ID + "Age[0]"})
    assert not pr.ok


def test_write_creates_datasets_when_missing():
    blob = make_xfa_pdf(datasets=None)
    f = fields_of(blob)
    assert f[ID + "LastName[0]"]["value"] is None
    res = fill(PdfDocument(blob), {ID + "LastName[0]": "Kato", ID + "Row[0].Item[0]": "pen"})
    pr = prove(blob, res.blob, {ID + "LastName[0]", ID + "Row[0].Item[0]"})
    assert pr.ok, pr.unexpected
    p = packets(res.blob)
    assert list(p) == ["preamble", "config", "template", "datasets", "postamble"]
    got = fields_of(res.blob)
    assert got[ID + "LastName[0]"]["value"] == "Kato" and got[ID + "Row[0].Item[0]"]["value"] == "pen"
    assert b"<form1><Page1>" in p["datasets"]


def test_single_stream_xfa():
    blob = make_xfa_pdf(single=True)
    assert fields_of(blob)[ID + "LastName[0]"]["value"] == "Suzuki"
    res = fill(PdfDocument(blob), {ID + "LastName[0]": "Mori"})
    pr = prove(blob, res.blob, {ID + "LastName[0]"})
    assert pr.ok, pr.unexpected
    assert fields_of(res.blob)[ID + "LastName[0]"]["value"] == "Mori"


def test_static_form_writes_data_and_pdf_field():
    acro = {ID + "LastName[0]": "text", ID + "Birth[0]": "text", ID + "Member[0]": "check"}
    blob = make_xfa_pdf(dynamic=False, acro=acro)
    win = build(PdfDocument(blob))
    assert win["xfa"]["kind"] == "static"
    page = {f["id"]: f for f in win["pages"][0]["fields"]}
    assert page[ID + "LastName[0]"]["value"] == "Suzuki" and page[ID + "LastName[0]"]["source"] == "xfa"
    vals = {ID + "LastName[0]": "Ito", ID + "Birth[0]": "1990-04-01", ID + "Member[0]": True}
    res = fill(PdfDocument(blob), vals)
    pr = prove(blob, res.blob, set(vals))
    assert pr.ok, pr.unexpected
    after = PdfDocument(res.blob)
    acro_v = {f.name: f.value_text() for f in read_fields(after)}
    assert acro_v[ID + "LastName[0]"] == "Ito" and acro_v[ID + "Birth[0]"] == "1990-04-01"
    birth = next(f for f in read_fields(after) if f.name == ID + "Birth[0]")
    ap = after.resolve(after.resolve(after.resolve(birth.widgets[0].ref)["AP"])["N"])
    assert b"04/01/1990".hex().upper().encode() in after.decode(ap)  # 見た目は、XFA の見た目の書式で整えた字
    page = {f["id"]: f for f in build(after)["pages"][0]["fields"]}
    assert page[ID + "Member[0]"]["value"] is True
    au = audit(blob, "static.pdf")
    names = next(c for c in au.checks if c.name == "XFA の名前と PDF の欄")
    assert names.ok


def test_encrypted_xfa_form():
    from test_pdf_password import OWNER, USER, Enc

    enc = Enc(rev=4, user=USER.encode(), owner=OWNER.encode(), cfm="AESV2")
    blob = make_xfa_pdf(enc=enc)
    assert fields_of(blob, USER)[ID + "LastName[0]"]["value"] == "Suzuki"
    res = fill(PdfDocument(blob, password=USER), {ID + "LastName[0]": "Endo"})
    pr = prove(blob, res.blob, {ID + "LastName[0]"}, password=USER)
    assert pr.ok, pr.unexpected
    assert fields_of(res.blob, USER)[ID + "LastName[0]"]["value"] == "Endo"
    assert b"Endo" not in res.blob  # 暗号化されたまま


def test_global_fields_share_data():
    t = TEMPLATE.replace('<field name="Phone"><ui><textEdit/></ui>', '<field name="LastName"><ui><textEdit/></ui><bind match="global"/>')
    t = t.replace('<field name="LastName"><ui><textEdit/></ui><caption><value><text>Family name</text></value></caption>',
                  '<field name="LastName"><ui><textEdit/></ui><bind match="global"/><caption><value><text>Family name</text></value></caption>', 1)
    blob = make_xfa_pdf(template=t)
    f = fields_of(blob)
    other = ID + "#subform[0].LastName[1]"  # 名前のない subform は透明：番号は Page1 の中で数える
    assert f[other]["value"] == "Suzuki" and ID + "LastName[0]" in f[other]["shared_with"]
    with pytest.raises(FillError):
        fill(PdfDocument(blob), {ID + "LastName[0]": "A", other: "B"})
    res = fill(PdfDocument(blob), {ID + "LastName[0]": "Abe"})
    assert fields_of(res.blob)[other]["value"] == "Abe"


def test_bare_ampersand_and_unnamed_subform_numbering():
    root = xfa.parse_xml(b"<a><b>x & y</b><c/></a>")
    a = root.elements[0]
    assert a.elements[0].text() == "x & y" and a.elements[1].s == len(b"<a><b>x & y</b>")


# ====================================================================== 書式（picture clause）
@pytest.mark.parametrize("pic, hint, canon, shown", [
    ("date{MM/DD/YYYY}", "date", "2026-04-01", "04/01/2026"),
    ("date{D MMMM YYYY}", "date", "2026-04-01", "1 April 2026"),
    ("date{DD/MM/YY}", "date", "2026-11-03", "03/11/26"),
    ("text{999-99-9999}", "text", "123456789", "123-45-6789"),
    ("text{'('999')' 999-9999}", "text", "5551234567", "(555) 123-4567"),
    ("num{z,zz9.99}", "num", "2157.5", "2,157.50"),
    ("num{($zzzz9.99)}", "num", "-12.5", "($12.50)"),
])
def test_picture_format_and_parse(pic, hint, canon, shown):
    loc = Locale()
    assert format_value(canon, pic, hint, loc) == shown
    assert parse_value(shown, pic, hint, loc) in (canon, canon.lstrip("+"))
    assert conforms(canon, pic, hint, loc)


def test_picture_alternatives_and_predefined():
    loc = Locale()
    assert conforms("12345", "text{99999-9999}|text{99999}", "text", loc)
    assert not conforms("1234", "text{99999-9999}|text{99999}", "text", loc)
    assert format_value("2026-07-04", "date.long{}", "date", loc) == "July 4, 2026"
    assert format_value("", "null{}|text{999}", "text", loc) == ""
    assert conforms("x", "date{YYYY} time{HH}", "date", loc) is None  # 決められない書式


# ====================================================================== 集めた様式
@pytest.mark.skipif(not WIDE, reason="corpus/wide_xfa がありません")
@pytest.mark.parametrize("path", WIDE, ids=[p.name for p in WIDE])
def test_corpus_reading(path):
    from todoku.pdf.audit import check_xfa

    doc = PdfDocument(path.read_bytes(), path.name)
    win = build(doc)
    assert win["xfa"]["readable"]
    checks = check_xfa(doc, win)
    for c in checks:
        assert c.ok, (c.name, c.items[:5])


SMALL = [p for p in WIDE if p.stat().st_size < 400_000]


@pytest.mark.skipif(not SMALL, reason="corpus/wide_xfa がありません")
@pytest.mark.parametrize("path", SMALL, ids=[p.name for p in SMALL])
def test_corpus_writing(path):
    from todoku.pdf.audit import check_writing

    blob = path.read_bytes()
    doc = PdfDocument(blob, path.name)
    for c in check_writing(doc, build(doc)):
        assert c.ok, (c.name, c.items[:5])
