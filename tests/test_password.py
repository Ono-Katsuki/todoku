"""パスワードで保護された書類：本人が対話で伝えたパスワードで開き、書いた後も同じパスワードで保護したまま書き出す。"""

from __future__ import annotations

from pathlib import Path

import pytest

from todoku import api
from todoku.mcp_server import call_tool, handle

CRYPTO = Path(__file__).resolve().parent.parent / "corpus" / "crypto"
PASSWORD = "Password1234_"  # 公開の試験用ファイル（Office で保存したもの）のパスワード


def _need(name: str) -> Path:
    path = CRYPTO / "inputs" / name
    if not path.exists():
        pytest.skip("corpus/crypto なし")
    return path


@pytest.mark.parametrize("name", ["example_password.xlsx", "example_password.docx", "ecma376standard_password.docx"])
def test_needs_and_checks_the_password(name):
    path = str(_need(name))
    with pytest.raises(api.PasswordNeeded):
        api.window(path)
    with pytest.raises(api.PasswordWrong):
        api.window(path, "違うパスワード")
    k, doc, parts = api.window(path, PASSWORD)
    assert k in ("xlsx", "docx") and parts and parts[0][1]


def test_fill_keeps_the_same_password(tmp_path):
    from todoku.officecrypto import decrypt_ooxml, verify_integrity

    src = _need("example_password.xlsx")
    fl = [f for f in api.fields(str(src), PASSWORD) if not f.get("readonly")]
    if not fl:
        pytest.skip("書ける欄のない試験用ファイル")
    key = fl[0]["id"]
    out = tmp_path / "out.xlsx"
    r = api.fill(str(src), {key: "とどく"}, str(out), password=PASSWORD)
    assert r["proof_ok"] and r["read_back_ok"], r
    assert "同じパスワード" in r["password_protected"] and PASSWORD not in str(r)
    blob = out.read_bytes()
    assert blob[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # 暗号の入れ物のまま
    inner, k2 = decrypt_ooxml(blob, PASSWORD)
    assert inner[:2] == b"PK" and verify_integrity(blob, k2)
    with pytest.raises(api.PasswordNeeded):
        api.fields(str(out))
    assert {f["id"]: f["value"] for f in api.fields(str(out), PASSWORD)}[key] == "とどく"


def test_mcp_asks_for_the_password_in_dialogue():
    path = str(_need("example_password.docx"))
    r = handle({"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": "open_document", "arguments": {"path": path}}})
    res = r["result"]
    assert not res["isError"] and res["structuredContent"]["needs_password"]
    assert "本人にパスワードを聞いて" in res["content"][0]["text"]
    r = call_tool("open_document", {"path": path, "password": PASSWORD})
    assert "Word" in r["content"][0]["text"]
