"""MCP サーバーが、Claude Code と同じ起動のしかた（.mcp.json の command・args）で立ち上がり、標準入出力で答えるかを確かめる。

initialize → notifications/initialized → tools/list を送り、サーバーの名前と道具の一覧を確かめる。標準ライブラリだけ。
使い方:  python .github/scripts/mcp_smoke.py            （.mcp.json の todoku を起動）
         python .github/scripts/mcp_smoke.py --plugin   （.claude-plugin/plugin.json の起動のしかた）
         python .github/scripts/mcp_smoke.py -- cmd args （起動するコマンドを指定）
"""

from __future__ import annotations

import json
import queue
import shutil
import subprocess
import sys
import threading
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
NEED = {"open_document", "read_window", "list_fields", "fill_document", "audit_document", "grade_document"}
TIMEOUT = 180  # 初回は uv が環境を作るので長め


def command() -> list[str]:
    if "--" in sys.argv:
        return sys.argv[sys.argv.index("--") + 1:]
    if "--plugin" in sys.argv:  # Claude Code のプラグインとして入れたときの起動（${CLAUDE_PLUGIN_ROOT} はこのリポジトリ）
        cfg = json.loads((ROOT / ".claude-plugin" / "plugin.json").read_text(encoding="utf-8"))["mcpServers"]["todoku"]
    else:
        cfg = json.loads((ROOT / ".mcp.json").read_text(encoding="utf-8"))["mcpServers"]["todoku"]
    exe = shutil.which(cfg["command"]) or cfg["command"]  # Windows では uv.exe を探す
    return [exe, *(a.replace("${CLAUDE_PLUGIN_ROOT}", str(ROOT)) for a in cfg.get("args", []))]


def main() -> int:
    cmd = command()
    print("起動:", " ".join(cmd), flush=True)
    p = subprocess.Popen(cmd, cwd=ROOT, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    lines: queue.Queue = queue.Queue()

    def pump():
        for raw in p.stdout:
            lines.put(raw)
        lines.put(None)

    threading.Thread(target=pump, daemon=True).start()

    def send(msg: dict) -> None:
        p.stdin.write((json.dumps(msg, ensure_ascii=False) + "\n").encode("utf-8"))
        p.stdin.flush()

    def reply(want_id: int) -> dict:
        while True:
            try:
                raw = lines.get(timeout=TIMEOUT)
            except queue.Empty:
                raise SystemExit(f"{TIMEOUT} 秒待っても id {want_id} の返事がない")
            if raw is None:
                err = p.stderr.read().decode("utf-8", "replace")
                raise SystemExit(f"サーバーが終わった（返事なし）。stderr:\n{err}")
            text = raw.decode("utf-8").strip()
            if not text:
                continue
            msg = json.loads(text)  # 標準出力には JSON-RPC の行だけが出るはず（ほかの字が混じれば、ここで落ちる）
            if msg.get("id") == want_id:
                return msg

    try:
        send({"jsonrpc": "2.0", "id": 1, "method": "initialize",
              "params": {"protocolVersion": "2025-06-18", "capabilities": {}, "clientInfo": {"name": "ci-smoke", "version": "0"}}})
        init = reply(1)
        assert init["result"]["serverInfo"]["name"] == "todoku", init
        send({"jsonrpc": "2.0", "method": "notifications/initialized"})
        send({"jsonrpc": "2.0", "id": 2, "method": "tools/list"})
        tools = {t["name"] for t in reply(2)["result"]["tools"]}
        missing = NEED - tools
        assert not missing, f"道具が足りない: {sorted(missing)}"
        sample = ROOT / "tests" / "fixtures" / "public" / "form.docx"
        send({"jsonrpc": "2.0", "id": 3, "method": "tools/call",
              "params": {"name": "list_fields", "arguments": {"path": str(sample)}}})
        res = reply(3)["result"]
        assert not res.get("isError"), res
        ids = {f["id"] for f in res["structuredContent"]["fields"]}
        assert "Tel" in ids, ids
        print(f"MCP: {init['result']['serverInfo']['name']} {init['result']['serverInfo'].get('version', '')}・"
              f"道具 {len(tools)} 個・list_fields（{sample.name}）{len(ids)} 欄", flush=True)
    finally:
        try:
            p.stdin.close()
        except OSError:
            pass
        try:
            p.wait(timeout=30)
        except subprocess.TimeoutExpired:
            p.kill()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
