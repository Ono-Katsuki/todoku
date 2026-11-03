"""PDF のフォーム：ひとつ選ぶ欄の組の見出しと選択肢ごとの字（options_labels）、画面にも紙にも出ない欄の印（合成した書類で確かめる）。"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from pdf_forms_maker import pdf, stream  # noqa: E402

from todoku import api  # noqa: E402
from todoku.pdf.doc import PdfDocument  # noqa: E402
from todoku.pdf.window import build, to_markdown  # noqa: E402


def _build(texts: list[tuple[float, float, str]], boxes: list[tuple[float, str]], y: float = 730,
           name_flags: int = 4) -> bytes:
    """texts：(x, y, 字)。boxes：(x, 入の状態の名前) の、Radio の印のないチェックの部品（gender の組）。"""
    objs: list[bytes] = []

    def add(b):
        objs.append(b if isinstance(b, bytes) else b.encode("latin-1"))
        return len(objs)

    def set_(i, b):
        objs[i - 1] = b if isinstance(b, bytes) else b.encode("latin-1")

    cat, pages = add(""), add("")
    font = add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
    zadb = add("<< /Type /Font /Subtype /Type1 /BaseFont /ZapfDingbats >>")
    dr = f"<< /Font << /Helv {font} 0 R /ZaDb {zadb} 0 R >> >>"

    def apbox(on):
        return add(stream(f"/Type /XObject /Subtype /Form /BBox [0 0 12 12] /Resources {dr}",
                          b"q BT /ZaDb 10 Tf 2 2 Td (4) Tj ET Q" if on else b""))

    body = b" ".join(f"BT /Helv 10 Tf {x} {yy} Td ({t}) Tj ET".encode() for x, yy, t in
                     texts + [(50, 760, "Name")])
    cs = add(stream("", body))
    page = add("")
    name = add(f"<< /Type /Annot /Subtype /Widget /FT /Tx /T (name) /Rect [100 755 180 771] /DA (/Helv 12 Tf 0 g) "
               f"/F {name_flags} /P {page} 0 R >>")
    par = add("")
    kids = []
    for x, st in boxes:
        on, off = apbox(True), apbox(False)
        kids.append(add(f"<< /Type /Annot /Subtype /Widget /Parent {par} 0 R /Rect [{x} {y - 3} {x + 12} {y + 9}] /AS /Off "
                        f"/AP << /N << /{st} {on} 0 R /Off {off} 0 R >> >> /MK << /CA (4) >> /DA (/ZaDb 0 Tf 0 g) "
                        f"/F 4 /P {page} 0 R >>"))
    set_(par, f"<< /FT /Btn /T (gender) /V /Off /Kids [{' '.join(f'{k} 0 R' for k in kids)}] >>")
    annots = " ".join(f"{k} 0 R" for k in [name] + kids)
    set_(page, f"<< /Type /Page /Parent {pages} 0 R /MediaBox [0 0 612 792] /Contents {cs} 0 R "
               f"/Resources << /Font << /Helv {font} 0 R >> >> /Annots [{annots}] >>")
    set_(pages, f"<< /Type /Pages /Kids [{page} 0 R] /Count 1 >>")
    set_(cat, f"<< /Type /Catalog /Pages {pages} 0 R /AcroForm << /Fields [{name} 0 R {par} 0 R] "
              f"/DR {dr} /DA (/Helv 0 Tf 0 g) >> >>")
    return pdf(objs, cat)


# 「Gender: Male □　Female □」：選択肢の字は部品の左
BEFORE = dict(texts=[(50, 730, "Gender: Male"), (160, 730, "Female")], boxes=[(120, "M"), (200, "F")])
# 「Sex: □ Man　□ Woman」：選択肢の字は部品の右
AFTER = dict(texts=[(50, 730, "Sex:"), (90, 730, "Man"), (150, 730, "Woman")], boxes=[(75, "M"), (135, "W")])


def _gender(blob: bytes) -> dict:
    w = build(PdfDocument(blob))
    return next(f for pg in w["pages"] for f in pg["fields"] if f["id"] == "gender")


def test_group_label_before_options():
    f = _gender(_build(**BEFORE))
    assert f["label"] == "Gender:"
    assert f["options"] == ["M", "F"] and f["options_labels"] == ["Male", "Female"]


def test_group_label_after_options():
    f = _gender(_build(**AFTER))
    assert f["label"] == "Sex:"
    assert f["options_labels"] == ["Man", "Woman"]


def test_markdown_shows_option_labels():
    md = to_markdown(build(PdfDocument(_build(**BEFORE))))
    assert "Male（M）／Female（F）" in md


def test_say_to_user_shows_the_chosen_option_label(tmp_path):
    src = tmp_path / "a.pdf"
    src.write_bytes(_build(**BEFORE))
    r = api.fill(str(src), {"gender": "F"}, str(tmp_path / "b.pdf"), overwrite=True)
    assert r["proof_ok"] and r["read_back_ok"], r
    assert "Gender：Female" in r["say_to_user"], r["say_to_user"]
    assert "Male" not in r["say_to_user"].split("記入した内容")[1].split("\n")[2]
    after = _gender((tmp_path / "b.pdf").read_bytes())
    assert after["value"] == "Female" and after["export_value"] == "F"


def test_option_can_be_chosen_by_its_printed_label(tmp_path):
    src = tmp_path / "a.pdf"
    src.write_bytes(_build(**AFTER))
    r = api.fill(str(src), {"gender": "Woman"}, str(tmp_path / "b.pdf"), overwrite=True)
    assert r["proof_ok"] and r["read_back_ok"], r
    from todoku.pdf.form import read_fields

    g = {f.name: f for f in read_fields(PdfDocument((tmp_path / "b.pdf").read_bytes()))}["gender"]
    assert g.value.name == "W" and [w.state for w in g.widgets] == ["Off", "W"]


def test_undecidable_labels_keep_the_old_label():
    # 選択肢の字が片側にそろわない（2つめの部品の左右に字がない）：組の見出しを作らない
    f = _gender(_build(texts=[(50, 730, "Choice")], boxes=[(100, "A"), (300, "B")]))
    assert "options_labels" not in f


def test_hidden_and_unprinted_fields_are_marked():
    w = build(PdfDocument(_build(**BEFORE, name_flags=2)))
    pg = w["pages"][0]
    name = next(f for f in pg["fields"] if f["id"] == "name")
    assert name["hidden"] and "Hidden" in name["hidden_why"]
    assert any("［name：画面にも紙にも出ない］" in ln for ln in pg["reading"]), pg["reading"]
    w = build(PdfDocument(_build(**BEFORE, name_flags=0)))
    pg = w["pages"][0]
    name = next(f for f in pg["fields"] if f["id"] == "name")
    assert name["not_printed"] and "hidden" not in name
    assert any("［name：紙に印刷されない］" in ln for ln in pg["reading"])
    assert "紙に印刷されない欄" in to_markdown(w)
    plain = next(f for f in build(PdfDocument(_build(**BEFORE)))["pages"][0]["fields"] if f["id"] == "name")
    assert "hidden" not in plain and "not_printed" not in plain
