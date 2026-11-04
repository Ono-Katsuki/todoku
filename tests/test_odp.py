"""OpenDocument の発表資料（.odp）と図形描画（.odg）：スライドの枠の字・表・ノートを読み、字を目印に書く。"""

from __future__ import annotations

import io
import zipfile

from todoku import api

NS = ('xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
      'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
      'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"')


def _pkg(body: str, mime: str = "application/vnd.oasis.opendocument.presentation") -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("mimetype", mime, compress_type=zipfile.ZIP_STORED)
        z.writestr("META-INF/manifest.xml", '<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">'
                   f'<manifest:file-entry manifest:full-path="/" manifest:media-type="{mime}"/>'
                   '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>')
        z.writestr("content.xml", f'<?xml version="1.0" encoding="UTF-8"?><office:document-content {NS} office:version="1.3"><office:body>{body}</office:body></office:document-content>')
    return buf.getvalue()


PAGES = ('<office:presentation>'
         '<draw:page draw:name="表紙"><draw:frame><draw:text-box><text:p>申請の説明会</text:p></draw:text-box></draw:frame></draw:page>'
         '<draw:page draw:name="手順"><draw:frame><draw:text-box><text:p>氏名：　　　　</text:p></draw:text-box></draw:frame>'
         '<draw:frame><table:table><table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell><text:p>窓口</text:p></table:table-cell>'
         '<table:table-cell><text:p>区役所</text:p></table:table-cell></table:table-row></table:table></draw:frame>'
         '<presentation:notes><draw:frame><draw:text-box><text:p>話す人のメモ</text:p></draw:text-box></draw:frame></presentation:notes></draw:page>'
         '</office:presentation>')


def _win(path) -> str:
    return "\n".join(t for _n, t in api.window(str(path))[2])


def test_odp_reads_every_slide_table_and_notes(tmp_path):
    p = tmp_path / "a.odp"
    p.write_bytes(_pkg(PAGES))
    k, doc, _parts = api.window(str(p))
    win = _win(p)
    assert k == "odp" and doc["kind_ja"].startswith("OpenDocument の発表資料")
    for x in ("申請の説明会", "氏名", "窓口", "区役所", "話す人のメモ", "スライド 1：表紙", "スライド 2：手順", "（ノート）"):
        assert x in win, x


def test_odg_reads_pages(tmp_path):
    p = tmp_path / "a.odg"
    p.write_bytes(_pkg('<office:drawing><draw:page draw:name="図"><draw:custom-shape><text:p>図形の字</text:p></draw:custom-shape></draw:page></office:drawing>',
                       "application/vnd.oasis.opendocument.graphics"))
    assert "図形の字" in _win(p)


def test_odp_grade_and_audit_do_not_refuse(tmp_path):
    p = tmp_path / "a.odp"
    p.write_bytes(_pkg(PAGES))
    g = api.grade(str(p))
    assert g["grade"] not in ("F",) and g["write"]["grade"] != "D"
    au = api.audit_whole(str(p))
    assert all(c["ok"] for c in au["checks"]), [c for c in au["checks"] if not c["ok"]]


def test_odp_write_changes_only_the_text(tmp_path):
    p = tmp_path / "a.odp"
    p.write_bytes(_pkg(PAGES))
    r = api.fill(str(p), {"@1": {"after": "氏名：", "value": "山田"}}, str(tmp_path / "o.odp"))
    assert r["proof_ok"] and r["read_back_ok"]
    assert "氏名：［" in _win(tmp_path / "o.odp") or "氏名：山田" in _win(tmp_path / "o.odp")
    a, b = zipfile.ZipFile(p), zipfile.ZipFile(tmp_path / "o.odp")
    assert [n for n in a.namelist() if a.read(n) != b.read(n)] == ["content.xml"]
