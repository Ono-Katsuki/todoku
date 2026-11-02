"""MCP サーバー（標準入出力の JSON-RPC）と、形式を問わない入口（api）のテスト。"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

from todoku import api

from test_docx import SAMPLE_BODY, make_docx
from test_pdf import sample_form

FIX = Path(__file__).parent / "fixtures"


def _session(msgs: list[dict]) -> list[dict]:
    data = "".join(json.dumps(m, ensure_ascii=False) + "\n" for m in msgs).encode()
    out = subprocess.run([sys.executable, "-m", "todoku.mcp_server"], input=data, capture_output=True, timeout=120)
    return [json.loads(x) for x in out.stdout.decode().splitlines() if x.strip()]


def test_mcp_protocol_and_fill(tmp_path):
    src = tmp_path / "s.docx"
    src.write_bytes(make_docx(SAMPLE_BODY))
    dst = tmp_path / "out.docx"
    replies = _session([
        {"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {"protocolVersion": "2025-06-18", "capabilities": {}}},
        {"jsonrpc": "2.0", "method": "notifications/initialized"},
        {"jsonrpc": "2.0", "id": 2, "method": "tools/list"},
        {"jsonrpc": "2.0", "id": 3, "method": "tools/call", "params": {"name": "list_fields", "arguments": {"path": str(src)}}},
        {"jsonrpc": "2.0", "id": 4, "method": "tools/call", "params": {"name": "fill_document", "arguments": {
            "path": str(src), "out_path": str(dst), "values": {"Tel": "03-1234-5678"}}}},
        {"jsonrpc": "2.0", "id": 5, "method": "tools/call", "params": {"name": "fill_document", "arguments": {
            "path": str(src), "out_path": str(src), "values": {"Tel": "1"}}}},
    ])
    by = {r["id"]: r for r in replies}
    assert by[1]["result"]["serverInfo"]["name"] == "todoku"
    assert {t["name"] for t in by[2]["result"]["tools"]} >= {"open_document", "read_window", "list_fields", "fill_document", "audit_document"}
    assert any(f["id"] == "Tel" for f in by[3]["result"]["structuredContent"]["fields"])
    res = by[4]["result"]["structuredContent"]
    assert res["proof_ok"] and res["read_back_ok"] and dst.exists()
    assert by[5]["result"]["isError"]  # 元のファイルは上書きしない
    assert src.read_bytes() == make_docx(SAMPLE_BODY)


def test_api_all_formats(tmp_path):
    pdf = tmp_path / "f.pdf"
    pdf.write_bytes(sample_form())
    xlsx = FIX / "kami_meisai.xlsx"
    for path in (str(pdf), str(xlsx)):
        s = api.summary(path)
        assert s["fields"] > 0
        fl = [f for f in api.fields(path) if not f.get("readonly") and f["kind"] in ("text", "number")]
        f = fl[0]
        val = 3 if f["kind"] == "number" else "とどく"
        res = api.fill(path, {f["id"]: val}, str(tmp_path / ("out" + Path(path).suffix)))
        assert res["proof_ok"] and res["read_back_ok"], res


def test_mcp_grade_document_combines_agent_findings(tmp_path):
    src = tmp_path / "s.docx"
    src.write_bytes(make_docx(SAMPLE_BODY))
    replies = _session([
        {"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {"protocolVersion": "2025-06-18", "capabilities": {}}},
        {"jsonrpc": "2.0", "id": 2, "method": "tools/call", "params": {"name": "grade_document", "arguments": {"path": str(src)}}},
        {"jsonrpc": "2.0", "id": 3, "method": "tools/call", "params": {"name": "grade_document", "arguments": {
            "path": str(src), "findings": [{"grade": "F", "title": "押印", "detail": "代表者の印", "where": "P9"}],
            "answered": ["cannot_do", "vague_value"]}}},
    ])
    by = {r["id"]: r for r in replies}
    rule = by[2]["result"]["structuredContent"]
    assert rule["grade"] == "B" and {c["id"] for c in rule["agent_checks"]} == {"cannot_do", "vague_value"}
    final = by[3]["result"]["structuredContent"]
    assert final["grade"] == "F" and final["without_barriers"] == "B" and not final["agent_checks"]
    assert any(r["by"] == "agent" for r in final["reasons"])
