"""世界調査の、Common Crawl の記録の読み方（ネットにつながずに確かめる）。"""

from __future__ import annotations

import gzip
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

import world_frame  # noqa: E402
import world_survey as ws  # noqa: E402


def _record(payload: bytes, truncated: bool = False) -> bytes:
    http = b"HTTP/1.1 200 OK\r\nContent-Type: application/pdf\r\n\r\n" + payload
    head = (b"WARC/1.0\r\nWARC-Type: response\r\nWARC-Date: 2026-10-10T08:14:01Z\r\n"
            + (b"WARC-Truncated: length\r\n" if truncated else b"")
            + b"Content-Length: " + str(len(http)).encode() + b"\r\n\r\n")
    return gzip.compress(head + http + b"\r\n\r\n")


def test_parse_warc_slices_by_content_length():
    body = b"%PDF-1.5\n...\n%%EOF\r\n\r\n"  # 中身の終わりの CRLF も中身のうち。記録の区切りの CRLF CRLF は除く
    w = ws.parse_warc(_record(body))
    assert w["payload"] == body
    assert w["http"]["content-type"] == "application/pdf"
    assert "warc-truncated" not in w["warc"]


def test_obtain_checks_digest_and_truncation(monkeypatch):
    body = b"%PDF-1.4 hello"
    raw = _record(body)
    monkeypatch.setattr(ws, "cc_get", lambda *a: raw)
    s = {"url": "https://example.go.jp/a.pdf", "digest": ws.sha1_b32(body), "filename": "f", "offset": 0, "length": len(raw)}
    info, got = ws.obtain(s)
    assert info["reach"] == "ok" and info["source"] == "cc" and got == body
    info, got = ws.obtain({**s, "digest": "X" * 32})
    assert info["reach"] == "digest_mismatch" and got is None
    # 切れた記録：robots.txt が許さなければ取り直さず truncated
    monkeypatch.setattr(ws, "cc_get", lambda *a: _record(body, truncated=True))
    monkeypatch.setattr(ws, "allowed", lambda url: False)
    info, got = ws.obtain(s)
    assert info["reach"] == "truncated" and info["live"] == "robots" and got is None


def test_suffix_for_uses_detected_type():
    assert ws.suffix_for({"url": "https://x.gov.br/a.xlsm", "mime_detected": "application/vnd.ms-excel.sheet.macroenabled.12"}) == ".xlsm"
    assert ws.suffix_for({"url": "https://x.gov.br/download?id=3", "mime_detected": "application/msword"}) == ".doc"
    assert ws.suffix_for({"url": "https://x.gov.br/a.pdf", "mime_detected": "application/pdf"}) == ".pdf"


def test_host_and_surt():
    assert world_frame.host_of("https://www.Bmas.de/DE/start.html") == "bmas.de"
    assert world_frame.host_of("(www.x.gov)") is None
    assert world_frame.surt("city.sapporo.lg.jp") == "jp,lg,sapporo,city"
