"""OpenDocument の文書・RTF・HWP の監査で見つかったことの確かめ（合成の文書）。"""

from __future__ import annotations

import io
import struct
import zipfile

import pytest

from todoku import api
from todoku.book import kind_of

# ---------------------------------------------------------------------- .odt：LibreOffice のコンテンツコントロール
ODT_NS = ('xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
          'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
          'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
          'xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" '
          'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" office:version="1.3"')


def odt(body: str) -> bytes:
    content = (f'<?xml version="1.0" encoding="UTF-8"?><office:document-content {ODT_NS}><office:automatic-styles>'
               '<style:style style:name="T1" style:family="text"/></office:automatic-styles>'
               f'<office:body><office:text>{body}</office:text></office:body></office:document-content>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr(zipfile.ZipInfo("mimetype"), "application/vnd.oasis.opendocument.text")
        z.writestr("content.xml", content, zipfile.ZIP_DEFLATED)
        z.writestr("META-INF/manifest.xml", '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest '
                   'xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">'
                   '<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>'
                   '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>',
                   zipfile.ZIP_DEFLATED)
    return buf.getvalue()


CC = ('<text:p>氏名：<loext:content-control loext:showing-place-holder="true" loext:plain-text="true" loext:alias="氏名" '
      'loext:tag="name" loext:id="1">ここをクリックして入力</loext:content-control></text:p>'
      '<text:p>同意する：<loext:content-control loext:checkbox="true" loext:checked-state="☒" loext:unchecked-state="☐" '
      'loext:tag="agree" loext:id="2"><text:span text:style-name="T1">☐</text:span></loext:content-control></text:p>'
      '<text:p>区分：<loext:content-control loext:showing-place-holder="true" loext:dropdown="true" loext:tag="kubun" loext:id="3">'
      '<loext:list-item loext:display-text="新規" loext:value="1"/><loext:list-item loext:display-text="変更" loext:value="2"/>'
      '選んでください</loext:content-control></text:p>'
      '<text:p>申請日：<loext:content-control loext:showing-place-holder="true" loext:date="true" loext:date-format="yyyy年M月d日" '
      'loext:tag="date" loext:id="4">日付を選択</loext:content-control></text:p>'
      '<text:p>人数：<field:fieldmark-start text:name="ninzu" field:type="vnd.oasis.opendocument.field.FORMTEXT"/>'
      '<field:fieldmark-separator/><field:fieldmark-end/></text:p>')


def test_odt_content_controls_read(tmp_path):
    path = tmp_path / "a.odt"
    path.write_bytes(odt(CC))
    fl = {f["id"]: f for f in api.fields(str(path))}
    assert fl["氏名"]["kind"] == "text" and fl["氏名"]["value"] is None and fl["氏名"]["label"] == "氏名"
    assert fl["agree"]["kind"] == "checkbox" and fl["agree"]["value"] is False
    assert fl["kubun"]["kind"] == "dropdown" and fl["kubun"]["options"] == ["新規", "変更"]
    assert fl["date"]["kind"] == "date"
    assert "ninzu" in fl  # field:fieldmark の名前（text:name）


def test_odt_content_controls_fill_proof_readback(tmp_path):
    path = tmp_path / "a.odt"
    path.write_bytes(odt(CC))
    out = str(tmp_path / "o.odt")
    r = api.fill(str(path), {"氏名": "山田 花子", "agree": True, "kubun": "変更", "date": "2026-11-03", "ninzu": "12"}, out)
    assert r["proof_ok"], r["unexpected"]
    assert r["read_back_ok"], r["read_back_differences"]
    assert r["written_as"] == {"date": "2026年11月3日"}
    xml = zipfile.ZipFile(out).read("content.xml").decode()
    assert "showing-place-holder" not in xml and 'loext:checked="true"' in xml and "☒" in xml
    assert 'loext:current-date="2026-11-03T00:00:00Z"' in xml
    say = r["say_to_user"]
    assert "人数：12" in say and "：：" not in say and "申請日：2026年11月3日" in say
    with pytest.raises(api.ApiError, match="選択肢"):
        api.fill(str(path), {"kubun": "廃止"}, str(tmp_path / "x.odt"))


# ---------------------------------------------------------------------- .rtf
B = chr(92)  # 逆斜線


def u(text: str) -> str:
    """RTF の Unicode の字の書き方（逆斜線 u と番号、代わりの字は「?」）。"""
    return "".join(c if ord(c) < 0x80 else B + "u%d" % (ord(c) if ord(c) < 0x8000 else ord(c) - 0x10000) + B + "'3f" for c in text)


def g(*words: str) -> str:
    """制御語を並べる：g("pard") → 逆斜線 pard。"""
    return "".join(B + w for w in words)


RTF = ("{" + g("rtf1", "ansi", "ansicpg1252", "uc1", "deff0") + "{" + g("fonttbl") + "{" + g("f0") + " Arial;}}\n"
       + g("pard") + " " + u("人数：") + "{" + g("*", "bkmkstart") + " ninzu}{" + g("*", "bkmkend") + " ninzu}"
       + "{" + g("field") + "{" + g("*", "fldinst") + " { FORMTEXT }}{" + g("fldrslt") + " {}}}" + g("par") + "\n"
       + g("pard") + " " + u("郵送：") + "{" + g("*", "bkmkstart") + " yuso}{" + g("*", "bkmkend") + " yuso}"
       + "{" + g("field") + "{" + g("*", "fldinst") + " { FORMCHECKBOX }}{" + g("fldrslt") + " {}}}" + g("par") + "\n"
       + g("pard") + " " + u("性別：") + "{" + g("field") + "{" + g("*", "fldinst") + " { FORMDROPDOWN {" + g("*", "formfield")
       + "{" + g("fftype2", "ffres0") + "{" + g("*", "ffname") + " sex}{" + g("*", "ffl") + " " + u("男") + "}{" + g("*", "ffl")
       + " " + u("女") + "}}}}}{" + g("fldrslt") + " {}}}" + g("par") + "\n"
       + g("pard") + " " + u("備考：") + "{" + g("deleted") + " " + u("旧い") + "}{" + u("新しい") + "}" + g("par") + "\n"
       + g("pard") + " " + u("宛名：") + "{" + g("field") + "{" + g("*", "fldinst") + " MERGEFIELD " + u("宛") + "}{" + g("fldrslt")
       + " " + u("«宛»") + "}}" + g("par") + "}")


def test_rtf_findings(tmp_path):
    path = tmp_path / "a.rtf"
    path.write_bytes(RTF.encode("ascii"))
    fl = {f["id"]: f for f in api.fields(str(path))}
    assert fl["sex"]["options"] == ["男", "女"]  # \ffl の中の \uN
    assert fl["ninzu"]["label"] == "人数"  # 名前はしおりから、見出しは前の字から
    assert fl["yuso"]["readonly"] and "formfield" in fl["yuso"]["why"]
    _k, _d, parts = api.window(str(path))
    t = parts[0][1]
    assert "備考：新しい〔変更履歴で削除された字（いまの本文ではない）: 旧い〕" in t
    out = str(tmp_path / "o.rtf")
    r = api.fill(str(path), {"ninzu": "12", "sex": "女"}, out)
    assert r["proof_ok"] and r["read_back_ok"]
    with pytest.raises(api.ApiError, match="yuso: この欄は、状態を入れる所"):
        api.fill(str(path), {"yuso": True}, str(tmp_path / "x.rtf"))
    with pytest.raises(api.ApiError, match="フィールド"):
        api.fill(str(path), {"@1": {"old": "«宛»", "new": "区長"}}, str(tmp_path / "y.rtf"))
    g = api.grade(str(path))
    titles = {x["title"] for x in g["reasons"]}
    assert "変更履歴" in titles
    assert g["write"]["grade"] == "D"  # 書けない入力欄
    assert api.audit(str(path))["ok"]


# ---------------------------------------------------------------------- HWP・HWPX
def _hwpx() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr(zipfile.ZipInfo("mimetype"), "application/hwp+zip")
        z.writestr("version.xml", "<v/>")
        z.writestr("Contents/section0.xml", "<s/>")
        z.writestr("META-INF/manifest.xml", "<m/>")
    return buf.getvalue()


def _cfb(streams: dict[str, bytes]) -> bytes:
    """最小の複合ファイル（512 バイトのセクター、ストリームは 4096 バイト以上にして、FAT だけで置く）。"""
    sec = 512
    names = list(streams)
    data = [streams[n].ljust(4096, b"\0") for n in names]
    n_data = sum(len(d) // sec for d in data)
    n_dir = 1
    fat_count = 1
    fat = []
    start = {}
    at = 0
    for n, d in zip(names, data):
        k = len(d) // sec
        start[n] = at
        fat += [at + i + 1 for i in range(k - 1)] + [0xFFFFFFFE]
        at += k
    dir_sec = at
    fat += [0xFFFFFFFE]
    fat_sec = at + n_dir
    fat += [0xFFFFFFFD]
    fat += [0xFFFFFFFF] * (sec // 4 - len(fat))

    def entry(name: str, typ: int, child: int, start_: int, size: int, left=0xFFFFFFFF, right=0xFFFFFFFF) -> bytes:
        nm = name.encode("utf-16-le") + b"\0\0"
        return (nm.ljust(64, b"\0") + struct.pack("<HBB", len(nm), typ, 1) + struct.pack("<III", left, right, child)
                + b"\0" * 16 + b"\0" * 4 + b"\0" * 16 + struct.pack("<IQ", start_, size)[:4] + struct.pack("<Q", size))

    ents = [entry("Root Entry", 5, 1, 0xFFFFFFFE, 0)]
    for i, n in enumerate(names):
        ents.append(entry(n, 2, 0xFFFFFFFF, start[n], len(data[i]), right=(i + 2) if i + 1 < len(names) else 0xFFFFFFFF))
    dirb = b"".join(ents).ljust(sec, b"\0")
    hdr = (b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 16 + struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + b"\0" * 6
           + struct.pack("<IIIIIIIII", 0, fat_count, dir_sec, 0, 4096, 0xFFFFFFFE, 0, 0xFFFFFFFE, 0)
           + struct.pack("<I", fat_sec) + struct.pack("<I", 0xFFFFFFFF) * 108)
    del n_data
    return hdr.ljust(sec, b"\0") + b"".join(data) + dirb + struct.pack(f"<{sec // 4}I", *fat)


@pytest.mark.parametrize("kind", ["hwpx", "hwp"])
def test_hwp_is_detected_and_broken_one_refused_clearly(tmp_path, kind):
    """HWP・HWPX は見分けて hwpx/ で読む（tests/test_hwpx.py）。中身の壊れたものは、推測で読まずに理由を返し、段階は F。"""
    blob = _hwpx() if kind == "hwpx" else _cfb({"FileHeader": b"HWP Document File" + b"\0" * 15, "BodyText": b"x"})
    assert kind_of(blob) == kind
    path = tmp_path / f"a.{kind}"
    path.write_bytes(blob)
    with pytest.raises(api.ApiError, match="本文の節"):
        api.window(str(path))
    g = api.grade(str(path))
    assert g["grade"] == "F" and "読めませんでした" in g["reasons"][0]["detail"]
