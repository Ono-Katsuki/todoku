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
            "path": str(src), "findings": [{"check": "cannot_do", "title": "押印", "detail": "代表者の印", "where": "P9"}],
            "answered": ["cannot_do", "ui_only", "vague_value"]}}},
    ])
    by = {r["id"]: r for r in replies}
    rule = by[2]["result"]["structuredContent"]
    # 規則の読む段階は A（字は正確に読める）。表題「申請書」の後ろ（規則では決められない所）は、書く段階の B
    assert rule["grade"] == "A" and {c["id"] for c in rule["agent_checks"]} == {"cannot_do", "ui_only", "vague_value"}
    assert rule["write"]["grade"] == "B"
    final = by[3]["result"]["structuredContent"]
    assert final["grade"] == "A" and final["write"]["grade"] == "D" and final["without_barriers"] == "B" and not final["agent_checks"]
    assert any(r["by"] == "agent" for r in final["write"]["reasons"])
    assert "書く段階 D" in by[3]["result"]["content"][0]["text"]


def test_fill_result_can_be_told_in_dialogue(tmp_path):
    # 本人は書類を開いて確かめない：書いた所（見出しと、読み直した値）・まだ空いている欄を、エージェントが対話で伝えられる
    from todoku import api

    src = tmp_path / "s.docx"
    src.write_bytes(make_docx(SAMPLE_BODY))
    r = api.fill(str(src), {"T1.R1.C2": "とどく花子", "P2.1": "8"}, str(tmp_path / "out.docx"))
    assert r["proof_ok"] and r["read_back_ok"]
    assert {d["id"]: d["value"] for d in r["written"]} == {"T1.R1.C2": "とどく花子", "P2.1": "8"}
    assert "- 氏名：とどく花子" in r["say_to_user"] and "- 令和［8］年 月 日" in r["say_to_user"], r["say_to_user"]
    assert "まだ空いている欄" in r["say_to_user"]
    assert r["still_empty_count"] >= 1 and all(not e["id"].startswith(("T1.R1.C2", "P2.1")) for e in r["still_empty"])


def test_say_to_user_lists_everything_written_by_place_and_joins_split_fields():
    # 一つずつ聞かずにまとめて書いた後、記入した内容を全部、シートごとに一覧で伝える。分かれた欄（〒の前と後ろ）は1行に
    from todoku.api import _to_tell

    after = [{"id": f"申請!B{i}", "label": "番号", "value": str(i), "sheet": "申請"} for i in range(10, 80)]
    after += [{"id": "申請!B3", "label": "〒［▶］-［　］", "value": "123", "sheet": "申請"},
              {"id": "申請!D3", "label": "〒［　］-［▶］", "value": "4567", "sheet": "申請"},
              {"id": "別紙!C2", "label": "氏名", "value": "山田花子", "sheet": "別紙"}]
    written = {f["id"]: f["value"] for f in after}
    say = _to_tell(after, written, {})["say_to_user"]
    assert "記入した内容（73 か所" in say and "ほか" not in say  # 全部を伝える
    assert "- 〒［123］-［4567］" in say and say.index("【申請】") < say.index("【別紙】")
    assert "- 番号：79" in say
