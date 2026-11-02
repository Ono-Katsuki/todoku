"""とどくの MCP サーバー：どの AI エージェントからも、書類を「窓」で読み、証明つきで書き込めるようにする。

標準入出力で JSON-RPC 2.0 を1行ずつやりとりする（Model Context Protocol の stdio）。外部のライブラリは使わない。
道具：
- open_document   書類を開き、形式・入力欄の数・窓の最初の部分を返す
- read_window     窓（Markdown）の n 番目の部分を返す（大きな書類は分けてある）
- list_fields     入力欄の一覧（名前・見出し・種類・選択肢・いまの値・書けるか）
- fill_document   入力欄に書き込み、別の名前で書き出す。元のファイルは変えない。証明と読み戻しの結果を返す
- audit_document  この書類を「全て正確に読み書きできるか」を確かめる（読み落とし・全欄への書き込み・読み戻し）
- grade_document  書類の段階（S〜F）。規則で決めた段階に、エージェントが窓を読んで確かめたことを合わせる
"""

from __future__ import annotations

import json
import sys
import traceback

from . import api

PROTOCOL = "2025-06-18"
INSTRUCTIONS = (
    "とどく（todoku）は、Excel・PDF・Word の申請書類を、内容を一字も変えずに読める「窓」にし、入力欄に証明つきで書き込む道具です。"
    "まず open_document で書類を開き、read_window で窓を読み、list_fields で入力欄の名前を確かめてから、"
    "fill_document で書き込んでください。書き込みは元のファイルを変えず、別の名前で書き出します。"
    "窓の［T1.R2.C3］［p1.5］［シート!B3］などが入力欄の名前です。推定した見出しは context（その行）と合わせて確かめてください。"
)

TOOLS = [
    {"name": "open_document",
     "description": "書類（Excel .xlsx/.xlsm/.xls・PDF・Word .docx/.doc）を開き、形式・入力欄の数・窓の分け方と、窓（Markdown）の最初の部分を返す。",
     "inputSchema": {"type": "object", "properties": {"path": {"type": "string", "description": "書類のファイルのパス"}},
                     "required": ["path"]}},
    {"name": "read_window",
     "description": "窓（Markdown）の part 番目の部分を返す（1から。大きな書類は目次と部分に分けてある）。",
     "inputSchema": {"type": "object", "properties": {"path": {"type": "string"},
                                                      "part": {"type": "integer", "minimum": 1, "default": 1}},
                     "required": ["path"]}},
    {"name": "list_fields",
     "description": "入力欄の一覧を JSON で返す：id（書き込むときの名前）・label（見出し）・kind（種類）・options（選択肢）・value（いまの値）・readonly（書けない欄）・context（その行）。",
     "inputSchema": {"type": "object", "properties": {"path": {"type": "string"},
                                                      "only_empty": {"type": "boolean", "default": False}},
                     "required": ["path"]}},
    {"name": "fill_document",
     "description": "入力欄に書き込み、out_path に書き出す（元のファイルは変えない）。values は {欄の id: 値}。チェックは true/false。"
                    "結果として、変わったのが書いた欄（とその計算結果・見た目）だけであることの証明と、書き込み後のファイルを読み直した結果を返す。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"}, "out_path": {"type": "string", "description": "書き出すファイルのパス（元と別の名前）"},
         "values": {"type": "object", "description": "{欄の id: 値}"},
         "force": {"type": "boolean", "default": False, "description": "保護された欄にも書く（本人が了承したときだけ）"}},
         "required": ["path", "out_path", "values"]}},
    {"name": "audit_document",
     "description": "この書類を、エージェントが全て正確に読み書きできるかを確かめる（窓とは別の読み方で読み落としがないか、全入力欄に書けるか、書いた値が読み戻せるか）。"
                    "段階（S 素で問題ない／A 規則だけで読み書きできる／B エージェントが読めば正確に入れられる／C エージェントが工夫して変換できる／"
                    "D エージェントの解釈が必要／E エージェントが得意でない形式／F アクセスできない）と、その理由も返す。",
     "inputSchema": {"type": "object", "properties": {"path": {"type": "string"}}, "required": ["path"]}},
    {"name": "grade_document",
     "description": "書類の段階（S〜F）を返す。規則で決めた段階に、エージェントが窓を読んで確かめたこと（findings）を合わせる。"
                    "規則では決めないこと（押印・自署のようなエージェントにはできない記入＝F、書く値をぼかす書き方＝D）は、"
                    "agent_checks の問いとして返すので、窓を読んで答えてください。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"},
         "findings": {"type": "array", "description": "見つけたこと：[{grade: 'F'|'D'|…, title, detail, where}]",
                      "items": {"type": "object", "properties": {"grade": {"type": "string"}, "title": {"type": "string"},
                                                                 "detail": {"type": "string"},
                                                                 "where": {"type": ["string", "array"]}}}},
         "answered": {"type": "array", "items": {"type": "string"},
                      "description": "確かめ終えた問いの id（見つからなかったものも含む。例 ['cannot_do', 'vague_value']）"}},
         "required": ["path"]}},
]


def _text(obj) -> dict:
    return {"type": "text", "text": obj if isinstance(obj, str) else json.dumps(obj, ensure_ascii=False, indent=1, default=str)}


def call_tool(name: str, args: dict) -> dict:
    if name == "open_document":
        path = args["path"]
        s = api.summary(path)
        _k, _doc, parts = api.window(path)
        head = (f"{s['kind_ja']}・入力欄 {s['fields']} か所（書ける欄 {s['writable_fields']}）・窓は {s['window_parts']} 部分"
                + ("（read_window で part を指定して続きを読めます）" if s["window_parts"] > 1 else ""))
        return {"content": [_text(head), _text(parts[0][1])], "structuredContent": s}
    if name == "read_window":
        _k, _doc, parts = api.window(args["path"])
        n = int(args.get("part", 1))
        if not 1 <= n <= len(parts):
            raise api.ApiError(f"part は 1〜{len(parts)} です")
        return {"content": [_text(f"部分 {n}/{len(parts)}：{parts[n - 1][0]}"), _text(parts[n - 1][1])]}
    if name == "list_fields":
        fl = api.fields(args["path"])
        if args.get("only_empty"):
            fl = [f for f in fl if f.get("value") in (None, "", False)]
        return {"content": [_text(fl)], "structuredContent": {"fields": fl}}
    if name == "fill_document":
        res = api.fill(args["path"], args.get("values") or {}, args["out_path"], bool(args.get("force", False)))
        ok = res.get("proof_ok") and res.get("read_back_ok")
        head = ("書き込みました。" if ok else "書き込みましたが、確認が必要な点があります。") + f"{res.get('proof', '')}"
        return {"content": [_text(head), _text(res)], "structuredContent": res, "isError": not ok}
    if name == "grade_document":
        g = api.grade(args["path"], args.get("findings") or [], args.get("answered") or [])
        lines = [g["headline"]] + [f"  - [{r['grade']}]{'（エージェントの判断）' if r.get('by') == 'agent' else ''} {r['title']}: {r['detail']}"
                                   for r in g["reasons"]]
        lines += [f"  - ？（あれば {c['grade']}）{c['question']}" for c in g.get("agent_checks", [])]
        return {"content": [_text("\n".join(lines)), _text(g)], "structuredContent": g}
    if name == "audit_document":
        au = api.audit(args["path"])
        lines = [("[OK] " if c["ok"] else "[要確認] ") + f"{c['name']}: {c['detail']}" for c in au["checks"]]
        lines.append("結論: " + ("全て正確に読み書きできます" if au["ok"] else "読み書きに確認が必要な点があります"))
        g = au.get("grade")
        if g:
            lines.append(g["headline"])
            lines += [f"  - [{r['grade']}] {r['title']}: {r['detail']}" for r in g["reasons"] if r["grade"] not in ("S", "A")]
            lines += [f"  - ？（あれば {c['grade']}）{c['question']}" for c in g.get("agent_checks", [])]
            if g.get("agent_checks"):
                lines.append("  窓を読んで確かめたら、grade_document に findings（見つけたこと）と answered（確かめた問いの id）を渡してください。")
        return {"content": [_text("\n".join(lines)), _text(au)], "structuredContent": au}
    raise api.ApiError(f"そのような道具はありません: {name}")


def handle(msg: dict) -> dict | None:
    mid = msg.get("id")
    method = msg.get("method")
    params = msg.get("params") or {}
    if method == "initialize":
        ver = params.get("protocolVersion") or PROTOCOL
        return {"jsonrpc": "2.0", "id": mid, "result": {
            "protocolVersion": ver, "capabilities": {"tools": {"listChanged": False}},
            "serverInfo": {"name": "todoku", "title": "とどく", "version": "0.2.0"}, "instructions": INSTRUCTIONS}}
    if method in ("notifications/initialized", "notifications/cancelled", "initialized"):
        return None
    if method == "ping":
        return {"jsonrpc": "2.0", "id": mid, "result": {}}
    if method == "tools/list":
        return {"jsonrpc": "2.0", "id": mid, "result": {"tools": TOOLS}}
    if method == "tools/call":
        name = params.get("name", "")
        args = params.get("arguments") or {}
        try:
            result = call_tool(name, args)
        except (api.ApiError, KeyError, ValueError) as e:
            result = {"content": [_text(f"できませんでした：{e}")], "isError": True}
        except Exception as e:  # noqa: BLE001
            result = {"content": [_text(f"処理の途中で問題が起きました：{type(e).__name__}: {e}\n{traceback.format_exc()[-1500:]}")],
                      "isError": True}
        return {"jsonrpc": "2.0", "id": mid, "result": result}
    if mid is None:
        return None
    return {"jsonrpc": "2.0", "id": mid, "error": {"code": -32601, "message": f"知らない命令です: {method}"}}


def main() -> int:
    stdin = sys.stdin.buffer
    out = sys.stdout
    for raw in stdin:
        line = raw.decode("utf-8", errors="replace").strip()
        if not line:
            continue
        try:
            msg = json.loads(line)
        except json.JSONDecodeError:
            out.write(json.dumps({"jsonrpc": "2.0", "id": None, "error": {"code": -32700, "message": "JSON として読めません"}}) + "\n")
            out.flush()
            continue
        msgs = msg if isinstance(msg, list) else [msg]
        replies = [r for r in (handle(m) for m in msgs) if r is not None]
        for r in replies:
            out.write(json.dumps(r, ensure_ascii=False, default=str) + "\n")
        out.flush()
    return 0


if __name__ == "__main__":
    sys.exit(main())
