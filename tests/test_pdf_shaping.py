"""字の組み立て（シェーピング）をして PDF に書く：アラビア文字（ペルシャ語・ウルドゥー語）・ヘブライ文字・インド系の文字・タイ文字。

書体は、このパソコンにある書体を使う（なければ、そのテストは飛ばす）。テストの PDF は、テストの中で組み立てる。
hb-shape（HarfBuzz）があれば、同じ書体（morx などを除いて GSUB・GPOS だけにした写し）で、字形の番号と位置が同じかを照らす。
"""

from __future__ import annotations

import re
import shutil
import subprocess

import pytest

from todoku import api
from todoku.pdf import glyphs, shape
from todoku.pdf.doc import PdfDocument
from todoku.pdf.fill import appearance_readback, fill
from todoku.pdf.form import read_fields

from pdf_forms_maker import build_all, pdf, stream


def _font(script: str):
    sf = shape.font_for(script)
    if sf is None:
        pytest.skip(f"このパソコンに、{script} を組み立てられる書体がない")
    return sf


def _write(tmp_path, blob: bytes, name: str = "in.pdf") -> str:
    p = tmp_path / name
    p.write_bytes(blob)
    return str(p)


def _field_ap(path: str, name: str = "family"):
    doc = PdfDocument(open(path, "rb").read())
    f = next(x for x in read_fields(doc) if x.name == name)
    st = glyphs.widget_appearance(doc, doc.resolve(f.ref))
    return doc, st


def wide_form(n: int, w: int = 520, h: int = 30, comb: bool = False) -> bytes:
    """幅の広い一行の欄 w0..w{n-1}（comb なら 1 字ずつの枠 8 字）。"""
    objs: list[bytes] = []

    def add(b):
        objs.append(b if isinstance(b, bytes) else b.encode("latin-1"))
        return len(objs)

    cat, pages, page = add(""), add(""), add("")
    font = add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
    fields = []
    for i in range(n):
        y = 800 - 36 * (i + 1)
        ap = add(stream(f"/Type /XObject /Subtype /Form /BBox [0 0 {w} {h}]", b"/Tx BMC EMC"))
        extra = " /Ff 16777216 /MaxLen 8" if comb else ""
        fields.append(add(f"<< /Type /Annot /Subtype /Widget /FT /Tx /T (w{i}) /Rect [40 {y} {40 + w} {y + h}] /P {page} 0 R "
                          f"/DA (/Helv 12 Tf 0 g) /F 4{extra} /AP << /N {ap} 0 R >> >>"))
    cs = add(stream("", b""))
    refs = " ".join(f"{f} 0 R" for f in fields)
    objs[page - 1] = (f"<< /Type /Page /Parent {pages} 0 R /MediaBox [0 0 600 842] /Contents {cs} 0 R "
                      f"/Resources << /Font << /F1 {font} 0 R >> >> /Annots [{refs}] >>").encode()
    objs[pages - 1] = f"<< /Type /Pages /Kids [{page} 0 R] /Count 1 >>".encode()
    objs[cat - 1] = (f"<< /Type /Catalog /Pages {pages} 0 R /AcroForm << /Fields [{refs}] "
                     f"/DR << /Font << /Helv {font} 0 R >> >> /DA (/Helv 0 Tf 0 g) >> >>").encode()
    return pdf(objs, cat)


# ====================================================================== Unicode の決まり
def test_bidi_levels_and_order():
    lv, para = shape.bidi_levels("Ali محمد 1990")
    assert para == 0
    assert lv[:4] == [0, 0, 0, 0] and lv[4:8] == [1, 1, 1, 1] and lv[9:] == [2, 2, 2, 2]
    lv, para = shape.bidi_levels("محمد 1990 Ali")
    assert para == 1 and lv[5:9] == [2, 2, 2, 2] and lv[10:] == [2, 2, 2]
    assert shape.reorder([0, 1, 1, 1, 0]) == [0, 3, 2, 1, 4]
    lv, _ = shape.bidi_levels("שלום (abc) עולם")  # 対の括弧（N0）：中が左から右でも、外の向きに合わせる
    assert lv[5] == lv[9] == 1
    with pytest.raises(glyphs.GlyphError, match="双方向"):
        shape.bidi_levels("abc‮def")


def test_joining_types_from_unicode_data():
    assert shape.joining_type(0x0628) == "D"  # ب
    assert shape.joining_type(0x0627) == "R"  # ا
    assert shape.joining_type(0x0621) == "U"  # ء
    assert shape.joining_type(0x064E) == "T"  # َ（一般カテゴリ Mn）
    assert shape.joining_type(0x0640) == "C"  # ـ
    assert shape.joining_type(0x06C1) == "D" and shape.joining_type(0x06D2) == "R"  # ہ ے（ウルドゥー語）


def test_indic_syllables():
    cats = "".join(shape.indic_props(ord(c), "Devanagari")[0] for c in "क्षेत्र")
    assert [k for k, _a, _b in shape.syllables(cats)] == ["consonant", "consonant"]
    cats = "".join(shape.indic_props(ord(c), "Bengali")[0] for c in "ি")
    assert shape.syllables(cats)[0][0] == "broken"
    assert shape.units("आवेदक का नाम") == ["आ", "वे", "द", "क", " ", "का", " ", "ना", "म"]
    assert shape.units("क्षेत्र") == ["क्षे", "त्र"]
    assert shape.units("abc") == ["a", "b", "c"]


# ====================================================================== 組み立て
def test_arabic_joining_and_lam_alef():
    sf = _font("Arabic")
    cm = sf.tt.cmap
    gl, _w = shape.shape_run(sf, "بب", True)
    gids = [g[0] for g in gl]
    assert len(gids) == 2 and cm[0x0628] not in gids  # 語頭形と語末形（独立形でない）
    assert gl[0][1] < gl[1][1]  # 見た目の順（左が語末）
    gl, _w = shape.shape_run(sf, "لا", True)
    assert len(gl) == 1 and gl[0][0] not in (cm[0x0644], cm[0x0627])  # ラーム・アリフの合字


def test_arabic_marks_positioned_on_base():
    sf = _font("Arabic")
    gl, _w = shape.shape_run(sf, "بَ", True)
    assert len(gl) == 2
    base, mark = gl if gl[0][0] != sf.tt.cmap.get(0x064E) else gl[::-1]
    assert mark[2] != 0 or mark[1] != base[1] + base[3]  # 記号は基の字の上に動かしてある


def test_devanagari_reph_and_pre_base_matra():
    sf = _font("Devanagari")
    cm = sf.tt.cmap
    gl, _w = shape.shape_run(sf, "र्क", False)
    assert gl[0][0] == cm[0x0915] and len(gl) == 2 and gl[1][0] != cm[0x0930]  # レーフは基の子音の後ろ
    gl, _w = shape.shape_run(sf, "कि", False)
    assert gl[1][0] == cm[0x0915] and gl[0][1] < gl[1][1]  # ि は前に


def test_tamil_two_part_vowel():
    sf = _font("Tamil")
    cm = sf.tt.cmap
    gl, _w = shape.shape_run(sf, "கொ", False)
    assert [g[0] for g in gl] == [cm[0x0BC6], cm[0x0B95], cm[0x0BBE]]  # ொ = ெ（前）+ ா（後ろ）


def test_thai_sara_am():
    sf = _font("Thai")
    gl, _w = shape.shape_run(sf, "ทำ", False)
    assert len(gl) == 3  # ำ をニッカヒットとサラ・アーに分ける


HB_WORDS = {
    "Arabic": ["الاسم الكامل", "تاريخ الميلاد", "لله", "نام خانوادگی", "درخواست", "مُحَمَّد", "بِسْمِ", "می‌خواهم", "۱۴۰۲"],
    "Hebrew": ["שלום", "שָׁלוֹם", "תאריך לידה"],
    "Devanagari": ["आवेदक का नाम", "क्षेत्र", "श्री", "र्क", "र्कि", "हस्ताक्षर", "क्‍ष", "क्‌ष", "ज़िला", "दृष्टि", "ॐ"],
    "Bengali": ["আবেদনকারীর নাম", "ক্ষ", "র্য", "কোথায়", "র‍্যাব", "উৎস", "স্ত্রী"],
    "Gurmukhi": ["ਪੰਜਾਬ", "ਪ੍ਰਧਾਨ", "ਕ੍ਰਿਪਾ", "ਦਸਤਖ਼ਤ", "ਸਿੱਖਿਆ"],
    "Gujarati": ["અરજદાર", "ક્ષ", "શ્રી", "કર્મ", "પ્રમાણપત્ર"],
    "Oriya": ["ଓଡ଼ିଆ", "କର୍ମ", "ଶ୍ରୀ", "ସ୍ୱାକ୍ଷର", "କୌ"],
    "Tamil": ["விண்ணப்பதாரர் பெயர்", "கொ", "க்ஷ", "கௌ", "ஸ்ரீ"],
    "Telugu": ["దరఖాస్తుదారు", "క్ష", "శ్రీ", "కర్మ", "ర్క", "క్షేత్రం"],
    "Kannada": ["ಅರ್ಜಿದಾರ", "ಕ್ಷ", "ಶ್ರೀ", "ಕರ್ಮ", "ಕೊ", "ಕೋ", "ಸರ್ಕಾರ"],
    "Thai": ["ชื่อ-นามสกุล", "ที่อยู่", "น้ำ", "ผู้ยื่นคำขอ", "ญี่ปุ่น", "พ.ศ. ๒๕๖๗"],
}
_HB_SCRIPT = {"Arabic": "arab", "Hebrew": "hebr", "Devanagari": "deva", "Bengali": "beng", "Gurmukhi": "guru",
              "Gujarati": "gujr", "Oriya": "orya", "Tamil": "taml", "Telugu": "telu", "Kannada": "knda", "Thai": "thai"}


@pytest.mark.parametrize("script", sorted(HB_WORDS))
def test_same_glyphs_and_positions_as_harfbuzz(script, tmp_path):
    """HarfBuzz（hb-shape）と、字形の番号・位置（字形の単位で ±1）・幅が同じ。書体は、morx・kerx を除いた写し
    （HarfBuzz は morx があればそちらを使うので）。"""
    hb = shutil.which("hb-shape")
    if hb is None:
        pytest.skip("hb-shape（HarfBuzz）がない")
    sf = _font(script)
    from todoku.pdf.embedfont import _sfnt

    fp = tmp_path / "f.ttf"
    fp.write_bytes(_sfnt({k: v for k, v in sf.tt.t.items() if k not in ("morx", "mort", "kerx", "kern")}))
    k = sf.tt.upem / 1000.0
    for w in HB_WORDS[script]:
        rtl = sf.script.rtl
        gl, width = shape.shape_run(sf, w, rtl)
        mine = [(g, round(x * k), round(y * k)) for g, x, y, _a in gl]
        r = subprocess.run([hb, f"--font-file={fp}", "--no-glyph-names", f"--script={_HB_SCRIPT[script]}",
                            f"--direction={'rtl' if rtl else 'ltr'}", w], capture_output=True, text=True, timeout=60)
        ref, pen = [], 0
        for item in r.stdout.strip().strip("[]").split("|"):
            m = re.match(r"(\d+)=\d+(?:@(-?\d+),(-?\d+))?\+(-?\d+)", item)
            g, xo, yo, xa = int(m[1]), int(m[2] or 0), int(m[3] or 0), int(m[4])
            if not (xa == 0 and g == sf.gid(0x20) and ("‌" in w or "‍" in w)):  # 幅のない ZWJ・ZWNJ
                ref.append((g, pen + xo, yo))
            pen += xa
        assert [x[0] for x in mine] == [x[0] for x in ref], (w, mine, ref)
        assert all(abs(a[1] - b[1]) <= 1 and abs(a[2] - b[2]) <= 1 for a, b in zip(mine, ref)), (w, mine, ref)
        assert abs(round(width * k) - pen) <= 1, w


# ====================================================================== 書き込みと読み戻し
FORM_WORDS = {
    "Arabic": ["الاسم الكامل", "تاريخ الميلاد", "لله", "نام خانوادگی", "درخواست"],
    "Hebrew": ["מספר זהות"],
    "Devanagari": ["आवेदक का नाम", "क्षेत्र", "श्री", "र्क"],
    "Bengali": ["আবেদনকারীর নাম", "ক্ষ", "র্য"],
    "Tamil": ["விண்ணப்பதாரர் பெயர்", "கொ", "க்ஷ"],
    "Thai": ["ชื่อ-นามสกุล"],
    "Gujarati": ["અરજદાર"], "Gurmukhi": ["ਪੰਜਾਬ"], "Oriya": ["ଆବେଦନ"], "Telugu": ["దరఖాస్తుదారు"], "Kannada": ["ಅರ್ಜಿದಾರ"],
}


@pytest.mark.parametrize("script", sorted(FORM_WORDS))
def test_fill_shaped_text_with_proof_and_read_back(script, tmp_path):
    _font(script)
    words = FORM_WORDS[script]
    src = _write(tmp_path, wide_form(len(words)))
    out = str(tmp_path / "o.pdf")
    r = api.fill(src, {f"w{i}": w for i, w in enumerate(words)}, out)
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    assert any("シェーピング" in w for w in r["warnings"])
    for i, w in enumerate(words):
        doc, st = _field_ap(out, f"w{i}")
        raw = doc.decode(st) if st.num else st.raw
        assert b"/ActualText <FEFF" + w.encode("utf-16-be").hex().upper().encode() in raw  # 行の元の字（論理の順）
        text, probs = glyphs.appearance_text(doc, st)
        assert probs == [] and text == w
    new = open(out, "rb").read()
    assert b"/ToUnicode" in new and b"/FontFile2" in new


def test_mixed_direction_line_and_wrap(tmp_path):
    _font("Arabic")
    _font("Devanagari")
    src = _write(tmp_path, build_all())
    value = "Ali محمد 1990 (test)"
    r = api.fill(src, {"family": value}, str(tmp_path / "o.pdf"))
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    doc, st = _field_ap(str(tmp_path / "o.pdf"))
    assert glyphs.appearance_text(doc, st) == (value, [])
    # 複数行の欄：単語の切れ目で折り返し、音節を切り離さない
    long = "आवेदक का नाम और पता हस्ताक्षर क्षेत्र " * 3
    from todoku.pdf.fill import wrap

    lines = wrap(long.strip(), 10, 150)
    assert len(lines) > 1 and all(not ln.startswith(("्", "ा", "ि", "े")) for ln in lines)
    assert " ".join(lines).split() == long.split()


def test_read_back_catches_a_changed_glyph(tmp_path):
    sf = _font("Devanagari")
    doc = PdfDocument(build_all())
    res = fill(doc, {"family": "र्क"})
    assert appearance_readback(PdfDocument(res.blob), res.patches) == []
    # 見た目の字形を 1 つ別の字形に変える（元の字のままの ActualText でも、組み立てた並びと違うので見つける）
    blob = res.blob
    tdoc = PdfDocument(blob)
    f = next(x for x in read_fields(tdoc) if x.name == "family")
    st = glyphs.widget_appearance(tdoc, tdoc.resolve(f.ref))
    raw = tdoc.decode(st) if st.num else st.raw
    first = re.search(rb"\[<([0-9A-F]{4})>", raw)[1]
    other = b"%04X" % sf.tt.cmap[0x0930]
    assert other != first
    from todoku.pdf.objects import Stream

    stream_obj = Stream(dict(st.dict), raw.replace(b"[<" + first + b">", b"[<" + other + b">", 1))
    text, probs = glyphs.appearance_text(tdoc, stream_obj)
    assert text == "र्क" and any("組み立てた字形と違います" in p for p in probs)


def test_read_back_catches_dotted_circle():
    sf = _font("Bengali")
    dot = sf.gid(0x25CC)
    if dot is None:
        pytest.skip("書体に点線の丸がない")

    class Info:
        name = "ABCDEF+" + sf.name

    probs, ok = glyphs._check_span("ক", [(Info(), dot)])
    assert not ok and any("点線の丸" in p for p in probs)


# ====================================================================== 断る
def test_refuses_broken_cluster(tmp_path):
    _font("Bengali")
    src = _write(tmp_path, build_all())
    with pytest.raises(api.ApiError, match=r"「ি」.*音節の形になっていません"):
        api.fill(src, {"family": "ি"}, str(tmp_path / "o.pdf"))


@pytest.mark.parametrize("value, script", [("മലയാളം", "マラヤーラム文字"), ("ශ්‍රී", "シンハラ文字"), ("ក្មេង", "クメール文字"),
                                           ("မြန်မာ", "ミャンマー文字")])
def test_refuses_unverified_scripts(tmp_path, value, script):
    src = _write(tmp_path, build_all())
    with pytest.raises(api.ApiError, match=f"{script}.*確かめていない"):
        api.fill(src, {"family": value}, str(tmp_path / "o.pdf"))


def test_refuses_without_a_shaping_font(tmp_path, monkeypatch):
    monkeypatch.setenv("TODOKU_PDF_FONT_ARABIC", "")
    src = _write(tmp_path, build_all())
    with pytest.raises(api.ApiError, match="TODOKU_PDF_FONT_ARABIC"):
        api.fill(src, {"family": "محمد"}, str(tmp_path / "o.pdf"))


def test_refuses_a_missing_joining_form(monkeypatch):
    """書体に語末形のない字（Tahoma のウルドゥー語の ہ）は、つながらない形で描かずに断る。"""
    import os

    p = "/System/Library/Fonts/Supplemental/Tahoma.ttf"
    if not os.path.exists(p):
        pytest.skip("Tahoma がない")
    monkeypatch.setenv("TODOKU_PDF_FONT_ARABIC", p)
    with pytest.raises(glyphs.GlyphError, match=r"「ہ」（U\+06C1）の語末形"):
        shape.line("پتہ")


def test_falls_back_to_the_next_font_for_a_missing_form():
    import os

    if not (os.path.exists("/System/Library/Fonts/Supplemental/Tahoma.ttf")
            and os.path.exists("/System/Library/Fonts/NotoNastaliq.ttc")):
        pytest.skip("Tahoma・Noto Nastaliq Urdu がない")
    ln = shape.line("پتہ")
    assert ln.segs[0].font.name == "NotoNastaliqUrdu"


def test_refuses_comb_field_and_bidi_controls(tmp_path):
    _font("Arabic")
    src = _write(tmp_path, wide_form(1, comb=True))
    with pytest.raises(api.ApiError, match="comb"):
        api.fill(src, {"w0": "محمد"}, str(tmp_path / "o.pdf"))
    src = _write(tmp_path, build_all(), "b.pdf")
    with pytest.raises(api.ApiError, match="双方向"):
        api.fill(src, {"family": "محمد‮abc"}, str(tmp_path / "o2.pdf"))


def test_latin_combining_marks_still_refused():
    with pytest.raises(glyphs.GlyphError, match="シェーピング"):
        glyphs.choose(chr(0x301))
