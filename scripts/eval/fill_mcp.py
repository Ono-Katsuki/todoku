"""記入のエージェントに渡す todoku の MCP サーバー（本物の todoku.mcp_server に、評価のための囲いをかぶせる）：
- 作業のフォルダの外のパスを開かせない（前の記入の記録・ほかの書類が見えないように。見えない状態を道具の側でも守る）
- 重い処理は一度に一つ（評価全体で同じ錠）
- 呼び出しの記録を機械で残す（proof_ok・read_back_ok・断った理由は、エージェントの申告でなく、ここから数える）
- eval_render_office：Office の書類を PDF にする（本物のエージェントが手元の LibreOffice で確かめるのと同じこと）。todoku の機能ではない
使い方（python -I）: fill_mcp.py <todoku の src> <作業のフォルダ> <記録の JSONL> <LibreOffice の設定のフォルダ>"""

import json
import os
import subprocess
import sys
import time
from pathlib import Path

src, sandbox, log_path, lo_profile = sys.argv[1:5]
sys.path.insert(0, src)
sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import Heavy  # noqa: E402

if os.nice(0) < 10:
    os.nice(10 - os.nice(0))
from todoku import mcp_server as ms  # noqa: E402

SANDBOX = os.path.realpath(sandbox)
RENDER = {"name": "eval_render_office",
          "description": "Excel・Word の書類（作業のフォルダの中）を、LibreOffice で PDF にする。できた PDF を view_page で見て、書いた見た目を確かめる。描き手は LibreOffice で、Excel・Word の実物とは少し違うことがある。",
          "inputSchema": {"type": "object", "properties": {"path": {"type": "string"}}, "required": ["path"]}}
ms.TOOLS.append(RENDER)
orig = ms.handle


def outside(args) -> str | None:
    def walk(v, key=""):
        if isinstance(v, dict):
            for k, x in v.items():
                got = walk(x, k)
                if got:
                    return got
        elif isinstance(v, list):
            for x in v:
                got = walk(x, key)
                if got:
                    return got
        elif isinstance(v, str) and key in ("path", "out_path", "file", "files", "image_path", "seal", "seal_path", "signature_image"):
            p = os.path.realpath(v.split("!/")[0])
            if p != SANDBOX and not p.startswith(SANDBOX + os.sep):
                return v
        return None

    return walk(args)


def render(path: str) -> dict:
    out = os.path.join(SANDBOX, "render")
    os.makedirs(out, exist_ok=True)
    r = subprocess.run(["soffice", f"-env:UserInstallation=file://{lo_profile}", "--headless", "--norestore", "--convert-to", "pdf", "--outdir", out, path],
                       capture_output=True, text=True, timeout=180)
    pdf = os.path.join(out, Path(path).stem + ".pdf")
    if not os.path.exists(pdf):
        return {"content": [{"type": "text", "text": f"PDF にできませんでした：{r.stderr[-300:]}"}], "isError": True}
    return {"content": [{"type": "text", "text": f"PDF にしました：{pdf}"}], "structuredContent": {"pdf": pdf}}


def result_summary(name: str, args: dict, res: dict) -> dict:
    sc = res.get("structuredContent") or {}
    text = " ".join(c.get("text", "") for c in res.get("content") or [] if c.get("type") == "text")
    d = {"is_error": bool(res.get("isError"))}
    if name == "fill_document":
        d["out_path"] = args.get("out_path")
        d["values"] = args.get("values")
        for k in ("proof_ok", "read_back_ok", "warnings", "uncertain", "say_to_user", "circled", "written_rects", "read_back_differences", "unexpected"):
            if k in sc:
                d[k] = sc[k]
        if res.get("isError"):
            d["refused"] = text[:1500]
    elif name == "audit_document":
        d["summary"] = text[:600]
        for k in ("ok", "score", "problems"):
            if k in sc:
                d[k] = sc[k] if k != "problems" else str(sc[k])[:600]
    elif res.get("isError"):
        d["error"] = text[:600]
    return d


def handle(msg: dict):
    if msg.get("method") != "tools/call":
        return orig(msg)
    params = msg.get("params") or {}
    name, args = params.get("name", ""), params.get("arguments") or {}
    t = time.time()
    bad = outside(args)
    if bad:
        reply = {"jsonrpc": "2.0", "id": msg.get("id"), "result": {"content": [{"type": "text", "text": f"この評価では、作業のフォルダ（{SANDBOX}）の外は開けません：{bad}"}], "isError": True}}
        entry = {"tool": name, "denied_path": bad}
    elif name == "eval_render_office":
        with Heavy():
            res = render(args["path"])
        reply = {"jsonrpc": "2.0", "id": msg.get("id"), "result": res}
        entry = {"tool": name, "is_error": bool(res.get("isError"))}
    else:
        with Heavy():
            reply = orig(msg)
        res = (reply or {}).get("result") or {}
        entry = {"tool": name, **result_summary(name, args, res)}
    entry["seconds"] = round(time.time() - t, 2)
    entry["at"] = round(t, 1)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False, default=str) + "\n")
    return reply


ms.handle = handle
raise SystemExit(ms.main())
