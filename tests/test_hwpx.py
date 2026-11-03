"""한글の文書：HWPX（読む・書く・証明・検査・段階）と HWP 5（読むだけ）。

HWPX は、OWPML（KS X 6101）の形のとおりに、テストの中で ZIP を組み立てる（mimetype を先頭に圧縮しないで置き、
Contents/content.hpf・header.xml・section0.xml）。HWP 5 は、CFB の中に FileHeader・DocInfo・BodyText/Section0 の
レコード（raw deflate で圧縮）を組み立てる。
"""

from __future__ import annotations

import io
import json
import re
import struct
import zipfile
import zlib

import pytest

from todoku import api
from todoku.book import kind_of
from todoku.cfb import build_cfb
from todoku.docx.fill import FillError
from todoku.grade import grade_blob
from todoku.hwpx.audit import run, run5
from todoku.hwpx.document import HwpxDocument
from todoku.hwpx.fill import fill, written_of
from todoku.hwpx.proof import prove
from todoku.hwpx.window import build, markdown_parts
from todoku.zipraw import RawZip

NS = ('xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
      'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" '
      'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:opf="http://www.idpf.org/2007/opf/" '
      'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"')
XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
LSA = '<hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="600" horzpos="0" horzsize="42520" flags="393216"/></hp:linesegarray>'

HEADER = (XML + f'<hh:head {NS} version="1.4" secCnt="1"><hh:refList>'
          '<hh:charProperties itemCnt="2">'
          '<hh:charPr id="0" height="1000" borderFillIDRef="1"><hh:underline type="NONE" shape="SOLID" color="#000000"/></hh:charPr>'
          '<hh:charPr id="1" height="1000" borderFillIDRef="1"><hh:underline type="BOTTOM" shape="SOLID" color="#000000"/></hh:charPr>'
          '</hh:charProperties>'
          '<hh:numberings itemCnt="1"><hh:numbering id="1" start="0">'
          '<hh:paraHead start="1" level="1" numFormat="DIGIT">^1.</hh:paraHead>'
          '<hh:paraHead start="1" level="2" numFormat="HANGUL_SYLLABLE">^2)</hh:paraHead>'
          '</hh:numbering></hh:numberings>'
          '<hh:paraProperties itemCnt="3">'
          '<hh:paraPr id="0"><hh:heading type="NONE" idRef="0" level="0"/></hh:paraPr>'
          '<hh:paraPr id="1"><hh:heading type="NUMBER" idRef="1" level="0"/></hh:paraPr>'
          '<hh:paraPr id="2"><hh:heading type="NUMBER" idRef="1" level="1"/></hh:paraPr>'
          '</hh:paraProperties></hh:refList></hh:head>')
HPF = (XML + f'<opf:package {NS} version="" unique-identifier="" id=""><opf:metadata><opf:language>ko</opf:language></opf:metadata>'
       '<opf:manifest><opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>'
       '<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/></opf:manifest>'
       '<opf:spine><opf:itemref idref="header"/><opf:itemref idref="section0" linear="yes"/></opf:spine></opf:package>')
CONTAINER = (XML + '<ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" '
             'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"><ocf:rootfiles>'
             '<ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/></ocf:rootfiles></ocf:container>')
MANIFEST = XML + '<odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>'
VERSION = (XML + '<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" tagetApplication="WORDPROCESSOR" '
           'major="5" minor="1" micro="0" buildNumber="1" os="1" xmlVersion="1.4" application="Hancom Office Hangul"/>')


def para(*runs: str, pr: str = "0") -> str:
    return f'<hp:p id="0" paraPrIDRef="{pr}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">{"".join(runs)}{LSA}</hp:p>'


def run_(inner: str = "", cp: str = "0") -> str:
    return f'<hp:run charPrIDRef="{cp}">{inner}</hp:run>' if inner else f'<hp:run charPrIDRef="{cp}"/>'


def t(s: str) -> str:
    return f"<hp:t>{s}</hp:t>"


def tc(col: int, row: int, inner: str = "", cs: int = 1, rs: int = 1, protect: int = 0) -> str:
    body = para(run_(t(inner))) if inner else para(run_())
    return (f'<hp:tc name="" header="0" hasMargin="0" protect="{protect}" editable="0" dirty="0" borderFillIDRef="1">'
            f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER">{body}</hp:subList>'
            f'<hp:cellAddr colAddr="{col}" rowAddr="{row}"/><hp:cellSpan colSpan="{cs}" rowSpan="{rs}"/>'
            '<hp:cellSz width="10000" height="1000"/><hp:cellMargin left="0" right="0" top="0" bottom="0"/></hp:tc>')


def click(fid: str, name: str, direction: str, content: str = "", dirty: str = "0") -> str:
    cmd = f"Clickhere:set:66:Direction:wstring:{len(direction)}:{direction} HelpState:wstring:0: "
    begin = (f'<hp:ctrl><hp:fieldBegin id="{fid}" type="CLICK_HERE" name="{name}" editable="1" dirty="{dirty}" zorder="-1" '
             f'fieldid="627272811"><hp:parameters cnt="2" name=""><hp:integerParam name="Prop">9</hp:integerParam>'
             f'<hp:stringParam name="Command" xml:space="preserve">{cmd}</hp:stringParam></hp:parameters></hp:fieldBegin></hp:ctrl>')
    end = f'<hp:ctrl><hp:fieldEnd beginIDRef="{fid}" fieldid="627272811"/></hp:ctrl>'
    if content:  # 한글が書くように、値は別の字の並び（hp:run）に置く
        return run_(begin) + run_(t(content), "1") + run_(end + "<hp:t/>")
    return run_(begin + end + "<hp:t/>")


FORM_ATTRS = ('foreColor="#000000" backColor="#FFFFFF" groupName="" tabStop="1" editable="1" tabOrder="1" enabled="1" '
              'borderTypeIDRef="0" drawFrame="1" printable="1" command=""')
SZ = ('<hp:sz width="7000" widthRelTo="ABSOLUTE" height="1500" heightRelTo="ABSOLUTE" protect="0"/>'
      '<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" '
      'horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>')


def section() -> str:
    body = [
        para(run_('<hp:secPr id="" textDirection="HORIZONTAL" outlineShapeIDRef="1"/>'
                  '<hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/></hp:ctrl>'
                  + t("보조금 지급 신청서"))),
        para(run_(t("신청인")), pr="1"),
        para(run_(t("연락처")), pr="2"),
        para(run_(t("신청 내용")), pr="1"),
        para(run_('<hp:tbl id="1" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" rowCnt="3" colCnt="3" cellSpacing="0" '
                  'borderFillIDRef="1" noAdjust="0">' + SZ
                  + "<hp:tr>" + tc(0, 0, "성명") + tc(1, 0) + tc(2, 0, "사진", rs=2) + "</hp:tr>"
                  + "<hp:tr>" + tc(0, 1, "주소") + tc(1, 1) + "</hp:tr>"
                  + "<hp:tr>" + tc(0, 2, "담당자 확인") + tc(1, 2, cs=2, protect=1) + "</hp:tr>"
                  + "</hp:tbl><hp:t/>")),
        para(run_(t("전화번호: ")) + click("11", "전화번호", "전화번호를 입력하세요")),
        para(run_(t("이메일: ")) + click("12", "", "이메일", content="이메일")),
        para(run_(t("생년월일 ")) + run_(t("            "), "1")),
        para(run_(t("□ 예 □ 아니오"))),
        para(run_(f'<hp:checkBtn caption="개인정보 수집에 동의" value="UNCHECKED" radioGroupName="" triState="0" backStyle="OPAQUE" '
                  f'name="동의" {FORM_ATTRS}><hp:formCharPr charPrIDRef="0" followContext="0" autoSz="0" wordWrap="0"/>{SZ}</hp:checkBtn>'
                  + t(" 성별 ")
                  + f'<hp:radioBtn caption="남" value="UNCHECKED" radioGroupName="성별" triState="0" backStyle="OPAQUE" name="r1" {FORM_ATTRS}>'
                    f'<hp:formCharPr charPrIDRef="0"/>{SZ}</hp:radioBtn>'
                  + f'<hp:radioBtn caption="여" value="CHECKED" radioGroupName="성별" triState="0" backStyle="OPAQUE" name="r2" {FORM_ATTRS}>'
                    f'<hp:formCharPr charPrIDRef="0"/>{SZ}</hp:radioBtn><hp:t/>')),
        para(run_(t("지역 ")
                  + f'<hp:comboBox listBoxRows="10" listBoxWidth="0" editEnable="0" selectedValue="" name="지역" {FORM_ATTRS}>'
                    '<hp:formCharPr charPrIDRef="0"/><hp:listItem displayText="서울" value="11"/><hp:listItem displayText="부산" value="26"/>'
                    f'{SZ}</hp:comboBox>'
                  + t(" 나이 ")
                  + f'<hp:edit multiLine="0" passwordChar="" maxLength="3" scrollBars="NONE" tabKeyBehavior="NEXT_OBJECT" numOnly="1" '
                    f'readOnly="0" alignText="LEFT" name="나이" {FORM_ATTRS}><hp:formCharPr charPrIDRef="0"/>{SZ}<hp:text/></hp:edit><hp:t/>')),
        para(run_('<hp:ctrl><hp:header id="1" applyPageType="BOTH"><hp:subList id="">' + para(run_(t("서식 제1호")))
                  + '</hp:subList></hp:header></hp:ctrl><hp:ctrl><hp:footer id="2" applyPageType="BOTH"><hp:subList id="">'
                  + para(run_(t("행정안전부"))) + '</hp:subList></hp:footer></hp:ctrl>'
                  + t("유의 사항") + '<hp:ctrl><hp:footNote number="1" suffixChar="41" instId="1"><hp:subList id="">'
                  + para(run_(t("사실과 다르면 환수합니다"))) + "</hp:subList></hp:footNote></hp:ctrl><hp:t/>")),
        para(run_('<hp:rect id="5" zOrder="1" numberingType="PICTURE" textWrap="IN_FRONT_OF_TEXT"><hp:drawText lastWidth="1000" name="" editable="0">'
                  '<hp:subList id="">' + para(run_(t("접수 번호"))) + '</hp:subList></hp:drawText><hp:shapeComment>사각형입니다.</hp:shapeComment></hp:rect>'
                  '<hp:pic id="6" zOrder="2" numberingType="PICTURE"><hc:img binaryItemIDRef="image1"/><hp:shapeComment>기관 로고</hp:shapeComment></hp:pic>'
                  '<hp:pic id="7" zOrder="3" numberingType="PICTURE"><hc:img binaryItemIDRef="image2"/></hp:pic>'
                  + t("변경<hp:deleteBegin Id=\"1\" TcId=\"1\"/>삭제한 글<hp:deleteEnd Id=\"1\" TcId=\"1\" paraend=\"0\"/> 끝"))),
        para(run_()),
    ]
    return XML + f"<hs:sec {NS}>" + "".join(body) + "</hs:sec>"


def make_hwpx(sec: str | None = None, extra: dict | None = None, manifest: str = MANIFEST) -> bytes:
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, "w") as z:
        z.writestr(zipfile.ZipInfo("mimetype"), "application/hwp+zip", compress_type=zipfile.ZIP_STORED)
        z.writestr("version.xml", VERSION, compress_type=zipfile.ZIP_STORED)
        z.writestr("Contents/header.xml", HEADER, compress_type=zipfile.ZIP_DEFLATED)
        z.writestr("Contents/section0.xml", sec or section(), compress_type=zipfile.ZIP_DEFLATED)
        z.writestr("Preview/PrvText.txt", "보조금 지급 신청서", compress_type=zipfile.ZIP_DEFLATED)
        z.writestr("settings.xml", XML + '<ha:HWPApplicationSetting xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app"/>',
                   compress_type=zipfile.ZIP_DEFLATED)
        z.writestr("META-INF/container.xml", CONTAINER, compress_type=zipfile.ZIP_DEFLATED)
        z.writestr("Contents/content.hpf", HPF, compress_type=zipfile.ZIP_DEFLATED)
        z.writestr("META-INF/manifest.xml", manifest, compress_type=zipfile.ZIP_DEFLATED)
        for k, v in (extra or {}).items():
            z.writestr(k, v, compress_type=zipfile.ZIP_DEFLATED)
    return bio.getvalue()


@pytest.fixture(scope="module")
def blob() -> bytes:
    return make_hwpx()


def _fields(b: bytes) -> dict:
    return {f["id"]: f for f in build(HwpxDocument(b, "a.hwpx"))["fields"]}


def _lines(b: bytes) -> list[str]:
    return [ln for st in build(HwpxDocument(b, "a.hwpx"))["stories"] for ln in st["reading"]]


# ====================================================================== 見分ける・読む
def test_kind(blob):
    assert kind_of(blob, "a.hwpx") == "hwpx"
    assert kind_of(blob) == "hwpx"


def test_window_reads_everything(blob):
    win = build(HwpxDocument(blob, "a.hwpx"))
    lines = [re.sub(r"［P\d+\.a\d+？］", "", ln) for st in win["stories"] for ln in st["reading"]]  # 規則では決められない所の印は除いて見る
    text = "\n".join(lines)
    assert "보조금 지급 신청서" in lines[0]
    assert "1. 신청인" in text and "가) 연락처" in text and "2. 신청 내용" in text  # 段落の番号（numbering の書式）
    assert "성명 ｜ ［T1.R1.C2］ ｜ 사진" in text  # 縦に結合したマス：下の行には出さない
    assert "주소 ｜ ［T1.R2.C2］" in text
    assert "전화번호: ［전화번호］" in text
    assert "（テキストボックス）접수 번호" in text
    assert "〔図: 기관 로고〕" in text and "〔図2（説明なし）〕" in text
    assert "변경 끝〔非表示の文字: 삭제한 글〕" in text  # 変更の記録で消した字は、非表示の字
    kinds = {st["kind"] for st in win["stories"]}
    assert {"body", "header", "footer", "footnotes"} <= kinds
    f = {x["id"]: x for x in win["fields"]}
    assert f["전화번호"]["placeholder"] == "전화번호를 입력하세요" and f["전화번호"]["value"] is None
    # 手を入れていない（dirty="0"）欄の字が案内の字と同じ：한글は案内の字として扱う（値ではない）
    unnamed = next(x for x in win["fields"] if x.get("placeholder") == "이메일")
    assert unnamed["value"] is None
    assert f["동의"]["kind"] == "checkbox" and f["동의"]["value"] is False and f["동의"]["label"] == "개인정보 수집에 동의"
    assert f["성별"]["options"] == ["남", "여"] and f["성별"]["value"] == "여"
    assert f["지역"]["options"] == ["서울", "부산"] and not f["지역"]["value"]
    assert f["나이"]["maxlen"] == 3 and f["나이"]["format"] == "数字だけ"
    assert f["T1.R3.C2"]["readonly"] and f["T1.R3.C2"]["protected"]  # 保護したマス
    boxes = [x for x in win["fields"] if x["kind"] == "checkbox" and x["id"].startswith("P")]
    assert [x["label"] for x in boxes] == ["예", "아니오"]
    under = [x for x in win["fields"] if x["kind"] == "inline" and not x.get("agent") and "생년월일" in x["label"]]
    assert under, "下線を引いた空白の空欄"
    assert win["pictures"][1]["alt"] is None
    assert markdown_parts(win)[0][1].startswith("# a.hwpx")


def test_audit_all_ok(blob):
    au = run(blob, "a.hwpx")
    bad = [(c.name, c.items[:5]) for c in au.checks if not c.ok]
    assert not bad, bad


# ====================================================================== 書く
def _write(blob: bytes, values: dict):
    res = fill(HwpxDocument(blob, "a.hwpx"), values)
    pr = prove(blob, res.blob, written_of(res.patches))
    assert pr.ok, (pr.unexpected, pr.lint_errors)
    return res


def test_fill_proof_and_readback(blob):
    f0 = _fields(blob)
    under = next(x["id"] for x in f0.values() if x["kind"] == "inline" and not x.get("agent") and "생년월일" in x["label"])
    box = next(x["id"] for x in f0.values() if x["kind"] == "checkbox" and x["label"] == "예")
    unnamed = next(x["id"] for x in f0.values() if x.get("placeholder") == "이메일")
    values = {"T1.R1.C2": "홍길동", "T1.R2.C2": "서울특별시 종로구 세종대로 209", "전화번호": "010-1234-5678",
              unnamed: "hong@example.kr", "동의": True, "성별": "남", "지역": "부산", "나이": "42", under: "1990-01-02",
              box: True, "@1": {"old": "유의 사항", "new": "유의 사항(필독)"}}
    res = _write(blob, values)
    f1 = _fields(res.blob)
    for k, v in values.items():
        if k.startswith("@"):
            continue
        assert f1[k]["value"] == v, (k, f1[k]["value"])
    anchored = next(p.id for p in res.patches if p.anchor)
    assert f1[anchored]["value"] == "(필독)"
    # 書き換えたのは本文の節だけ：ほかの部品は、ヘッダーと圧縮されたバイト列まで同じ
    za, zb = RawZip(blob), RawZip(res.blob)
    assert za.names() == zb.names()
    for n in za.names():
        same = za.get(n).local == zb.get(n).local and za.get(n).data == zb.get(n).data
        assert same == (n != "Contents/section0.xml"), n
    sec = zb.read("Contents/section0.xml").decode()
    assert 'name="전화번호" editable="1" dirty="1"' in sec  # 한글で書いた欄と同じ印
    assert 'value="CHECKED" radioGroupName="" triState="0" backStyle="OPAQUE" name="동의"' in sec
    assert 'selectedValue="26"' in sec and "<hp:text>42</hp:text>" in sec
    assert '<hp:bookmark name="_todoku_T1.R1.C2"/>' in sec
    # 行の配置の控え：字を変えた段落だけ除く（ほかの段落は残る）
    before = za.read("Contents/section0.xml").decode().count("<hp:linesegarray>")
    after = sec.count("<hp:linesegarray>")
    assert 0 < before - after < before
    lines = "\n".join(_lines(res.blob))
    assert f"☑［{box}］ 예 ☐［" in lines and "<hp:t>■</hp:t>" in sec  # □ は ■ にする
    # 書いた欄の id は、書いた後も同じ：もう一度は書かない
    with pytest.raises(FillError, match="すでに書き込んだ"):
        fill(HwpxDocument(res.blob), {"T1.R1.C2": "다시"})


def test_fill_refuses(blob):
    with pytest.raises(FillError, match="保護したマス"):
        fill(HwpxDocument(blob), {"T1.R3.C2": "x"})
    with pytest.raises(FillError, match="選択肢にありません"):
        fill(HwpxDocument(blob), {"지역": "대구"})
    with pytest.raises(FillError, match="数字だけ"):
        fill(HwpxDocument(blob), {"나이": "마흔"})
    with pytest.raises(FillError, match="字まで"):
        fill(HwpxDocument(blob), {"나이": "1234"})
    with pytest.raises(FillError, match="○で囲む"):
        fill(HwpxDocument(blob), {"@c": {"circle": "예"}})


def test_proof_catches_tampering(blob):
    res = fill(HwpxDocument(blob), {"T1.R1.C2": "홍길동"})
    z = RawZip(res.blob)
    sec = z.read("Contents/section0.xml").replace("주소".encode(), "住所".encode())
    z.replace("Contents/section0.xml", sec)
    pr = prove(blob, z.build(), written_of(res.patches))
    assert not pr.ok and any("書いていない段落" in u or "外が変わって" in u for u in pr.unexpected)


def test_api_roundtrip(tmp_path, blob):
    src = tmp_path / "신청서.hwpx"
    src.write_bytes(blob)
    s = api.summary(str(src))
    assert s["kind"] == "hwpx" and s["kind_ja"].startswith("한글")
    fl = {f["id"]: f for f in api.fields(str(src))}
    assert fl["전화번호"]["placeholder"] == "전화번호를 입력하세요"
    out = tmp_path / "신청서_記入.hwpx"
    r = api.fill(str(src), {"T1.R1.C2": "홍길동", "전화번호": "010-1234-5678", "성별": "남"}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], json.dumps(r, ensure_ascii=False)
    assert "홍길동" in r["say_to_user"]
    g = api.grade(str(src))
    assert g["grade"] == "A" and g["write"]["grade"] == "B"  # 空のマス・空欄は配置からの推定
    d = api.audit(str(out))
    assert d["ok"] or all(c["ok"] for c in d["checks"] if c["name"] != "全入力欄への書き込み"), d


def test_grade_marked_fields_only():
    sec = XML + f"<hs:sec {NS}>" + para(run_(t("성명: ")) + click("1", "성명", "이름을 쓰세요")) + "</hs:sec>"
    g = grade_blob(make_hwpx(sec), "a.hwpx")
    assert g.grade == "A" and g.write == "A", g.headline()


def test_scripts_are_macros():
    js = "function OnDocument_Open()\r\n{\r\n\tComboBox1.InsertString(\"서울\", 0);\r\n}\r\nfunction OnDocument_New()\r\n{\r\n}\r\n"
    g = grade_blob(make_hwpx(extra={"Scripts/sourceScripts": js.encode("utf-16-le")}), "a.hwpx")
    assert g.grade == "C" and g.write == "C" and any(r.title == "한글のマクロ" for r in g.reasons)
    empty = "function OnDocument_New()\r\n{\r\n}\r\n"
    g2 = grade_blob(make_hwpx(extra={"Scripts/sourceScripts": empty.encode("utf-16-le")}), "a.hwpx")
    assert not any(r.title == "한글のマクロ" for r in g2.reasons)


def test_encrypted_is_f():
    man = (XML + '<odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">'
           '<odf:file-entry odf:full-path="Contents/section0.xml" odf:media-type="application/xml" odf:size="10">'
           '<odf:encryption-data odf:checksum-type="SHA1" odf:checksum="x"/></odf:file-entry></odf:manifest>')
    g = grade_blob(make_hwpx(manifest=man), "a.hwpx")
    assert g.grade == "F" and g.write == "D"


# ====================================================================== HWP 5（読むだけ）
def _rec(tag: int, level: int, data: bytes) -> bytes:
    n = len(data)
    if n >= 0xFFF:
        return struct.pack("<II", tag | (level << 10) | (0xFFF << 20), n) + data
    return struct.pack("<I", tag | (level << 10) | (n << 20)) + data


def _ext(code: int, cid: str = "    ") -> list[int]:
    v = sum(ord(ch) << s for ch, s in zip(cid, (24, 16, 8, 0)))
    return [code, v & 0xFFFF, v >> 16, 0, 0, 0, 0, code]


def _ptext(units: list[int]) -> bytes:
    return struct.pack(f"<{len(units)}H", *units)


def _w(s: str) -> list[int]:
    return list(struct.unpack(f"<{len(s)}H", s.encode("utf-16-le")))


def _ctrl(cid: str, rest: bytes = b"") -> bytes:
    v = sum(ord(ch) << s for ch, s in zip(cid, (24, 16, 8, 0)))
    return struct.pack("<I", v) + rest


def _para(units: list[int], level: int = 0, ctrls: bytes = b"") -> bytes:
    head = struct.pack("<IIHBBHHHI", len(units), 0, 0, 0, 0, 1, 0, 0, 0)
    return _rec(0x42, level, head) + _rec(0x43, level + 1, _ptext(units)) + ctrls


def make_hwp5(props: int = 1) -> bytes:
    direction = "이름을 쓰세요"
    cmd = f"Clickhere:set:40:Direction:wstring:{len(direction)}:{direction} HelpState:wstring:0: "
    clk = _ctrl("%clk", struct.pack("<IB", 0, 0) + struct.pack("<H", len(cmd)) + cmd.encode("utf-16-le") + struct.pack("<I", 7))
    p1 = _para(_w("복지 급여 신청서") + [13])
    p2 = _para(_w("성명: ") + _ext(3, "%clk") + _ext(4) + [13], ctrls=_rec(0x47, 1, clk))

    def cell(col: int, row: int, text: str) -> bytes:
        lh = struct.pack("<HHIHHHHII4HH", 1, 0, 0, col, row, 1, 1, 1000, 1000, 0, 0, 0, 0, 1)
        return _rec(0x48, 2, lh) + _para(_w(text) + [13], level=2)

    tbl = (_rec(0x47, 1, _ctrl("tbl ", b"\0" * 40)) + _rec(0x4D, 2, b"\0" * 20)
           + cell(0, 0, "주소") + cell(1, 0, ""))
    p3 = _para(_ext(11, "tbl ") + [13], ctrls=tbl)
    body = p1 + p2 + p3
    head = (b"HWP Document File".ljust(32, b"\0") + struct.pack("<II", 0x05010000, props)).ljust(256, b"\0")

    def z(b: bytes) -> bytes:
        if not props & 1:
            return b
        c = zlib.compressobj(9, zlib.DEFLATED, -15)
        return c.compress(b) + c.flush()

    streams = {"FileHeader": head, "DocInfo": z(b""), "BodyText/Section0": z(body)}
    if props & 4:
        streams = {"FileHeader": head, "DocInfo": z(b""), "ViewText/Section0": b"\x01" * 300}
    return build_cfb(streams)


def test_hwp5_reads(tmp_path):
    b = make_hwp5()
    assert kind_of(b, "a.hwp") == "hwp"
    p = tmp_path / "a.hwp"
    p.write_bytes(b)
    k, win, parts = api.window(str(p))
    lines = [re.sub(r"［P\d+\.a\d+？］", "", ln) for st in win["stories"] for ln in st["reading"]]
    assert lines[0] == "복지 급여 신청서"
    assert "성명: ［ctl1］" in lines[1]
    assert "주소 ｜ ［T1.R1.C2］" in lines
    fl = {f["id"]: f for f in api.fields(str(p))}
    assert fl["ctl1"]["placeholder"] == "이름을 쓰세요" and fl["ctl1"]["readonly"]
    with pytest.raises(api.ApiError, match="HWPX で保存し直せば"):
        api.fill(str(p), {"ctl1": "홍길동"}, str(tmp_path / "b.hwp"))
    g = api.grade(str(p))
    assert g["write"]["grade"] == "D" and "HWPX で保存し直せば書ける" in json.dumps(g, ensure_ascii=False)
    au = run5(b, "a.hwp")
    assert au.ok, [(c.name, c.items) for c in au.checks if not c.ok]


def test_hwp5_distribution_is_f():
    g = grade_blob(make_hwp5(props=1 | 4), "a.hwp")
    assert g.grade == "F" and g.write == "D" and g.reasons[0].title == "配布用の文書"
    g = grade_blob(make_hwp5(props=1 | 2), "a.hwp")
    assert g.grade == "F" and g.reasons[0].title == "パスワード"
