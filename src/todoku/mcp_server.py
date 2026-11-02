"""とどくの MCP サーバー：どの AI エージェントからも、書類を「窓」で読み、証明つきで書き込めるようにする。

標準入出力で JSON-RPC 2.0 を1行ずつやりとりする（Model Context Protocol の stdio）。外部のライブラリは使わない。
道具：
- open_document   書類を開き、形式・入力欄の数・窓の最初の部分を返す
- read_window     窓（Markdown）の n 番目の部分を返す（大きな書類は分けてある）
- list_fields     入力欄の一覧（名前・見出し・種類・選択肢・いまの値・書けるか）
- fill_document   入力欄に書き込み、別の名前で書き出す。元のファイルは変えない。証明と読み戻しの結果を返す
- audit_document  この書類を「全て正確に読み書きできるか」を確かめる（読み落とし・全欄への書き込み・読み戻し）
- run_macro       Excel のマクロ（ボタン）を動かし、マクロが変えたセルを証明つきで書き出す
- grade_document  書類の段階（読む S〜F・書く A〜D）。規則で決めた段階に、エージェントが窓を読んで確かめたことを合わせる
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
    "［P12.a1？］のように「？」のついた所は、規則では欄と決められない所です。窓を読んで判断し、書く所なら同じように id で書きます。"
    "窓に欄として示していない所にも、字を目印にして書けます（fill_document の説明を見てください）。"
    "本人は書類を Word・Excel・Acrobat で開いて確かめません。対話だけで確かめられるように、書き込みの結果の say_to_user"
    "（書いた所の見出しと値・まだ空いている欄・本人に伝えること）を、そのまま本人に伝えてください。"
    "本人に欄を一つずつ聞かないでください。住所・名前・事業所の情報など、本人がこれまでに伝えたこと・ほかの書類・記憶から分かる値は、"
    "先にまとめて fill_document で書き、書いた後の say_to_user（記入した内容の一覧）を本人に伝えて、直す所がないかだけを確かめてください。"
    "分からない値・本人が決めること（選ぶもの・マクロの「はい／いいえ」など）は、勝手に決めず、全部をまとめて一度に、選びやすい選択肢の形で聞いてください。"
    "パスワードで保護された書類は「パスワードが要る」と返るので、本人にパスワードを聞いて、どの道具にも password で渡してください"
    "（パスワードは結果にも記録にも残しません。書き出したファイルも、元と同じパスワードで保護されたままです）。"
    "窓の「画像」の節で「説明（代わりの字）がありません」とある画像は、中身を字として読めません。申請に関わりそうな位置なら、"
    "その位置（ページ・上下左右・近くの字）を本人に伝え、中身を推測で埋めないでください。段階の alerts（アクセシビリティの注意）は、"
    "段階とは別に伝えてください。"
)

TOOLS = [
    {"name": "open_document",
     "description": "書類（Excel .xlsx/.xlsm/.xls・PDF・Word .docx/.doc）を開き、形式・入力欄の数・窓の分け方と、窓（Markdown）の最初の部分を返す。",
     "inputSchema": {"type": "object", "properties": {"path": {"type": "string", "description": "書類のファイルのパス"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"}},
                     "required": ["path"]}},
    {"name": "read_window",
     "description": "窓（Markdown）の part 番目の部分を返す（1から。大きな書類は目次と部分に分けてある）。",
     "inputSchema": {"type": "object", "properties": {"path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"},
                                                      "part": {"type": "integer", "minimum": 1, "default": 1}},
                     "required": ["path"]}},
    {"name": "list_fields",
     "description": "入力欄の一覧を JSON で返す：id（書き込むときの名前）・label（見出し）・kind（種類）・options（選択肢）・value（いまの値）・readonly（書けない欄）・context（その行）。",
     "inputSchema": {"type": "object", "properties": {"path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"},
                                                      "only_empty": {"type": "boolean", "default": False}},
                     "required": ["path"]}},
    {"name": "fill_document",
     "description": "入力欄に書き込み、out_path に書き出す（元のファイルは変えない）。values は {欄の id: 値}。チェックは true/false。"
                    "窓で［P12.a1？］と「？」のついた所は、規則では欄と決められない所：窓を読んで書く所と判断したときだけ、同じように id で書く。"
                    "Word で、窓に欄として示していない所に書くときは、{\"@1\": {\"old\": 段落の中の元の字, \"new\": 書き換えた後の字}} の形で、"
                    "字を目印にして書ける（old が文書に1つしかないときはそのまま。いくつもあるときは near に近くの欄の id か、"
                    "occurrence に何番目か を足す）。old と new の違う所だけが変わり、書いた所は id「P段落.w番号」になる。"
                    "PDF で、窓に欄として示していない所に書くときは、{\"@1\": {\"after\": 見出しの字, \"value\": 書く字, \"page\": ページ}} の形で、"
                    "見出しの右（\"where\": \"below\" なら下）の空いている所に書ける（書いた所は id「pページ.w番号」）。"
                    "PDF の伏せ字（「○○」）の上には {\"@1\": {\"cover\": \"○○\", \"value\": 書く字}}（白い下地の欄で覆って書く。元の字は消さない）。"
                    "結果として、変わったのが書いた欄（とその計算結果・見た目）だけであることの証明と、書き込み後のファイルを読み直した結果を返す。"
                    "本人は書類を開いて確かめないので、結果の say_to_user（書いた所の見出しと値・まだ空いている欄・本人に伝えること）を、"
                    "対話でそのまま本人に伝えて確かめる。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"}, "out_path": {"type": "string", "description": "書き出すファイルのパス（元と別の名前）"},
         "values": {"type": "object", "description": "{欄の id: 値}。Word で欄のない所は {\"@1\": {\"old\": …, \"new\": …}}、"
                                                      "PDF で欄のない所は {\"@1\": {\"after\": …, \"value\": …}}"},
         "force": {"type": "boolean", "default": False, "description": "保護された欄にも書く（本人が了承したときだけ）"}},
         "required": ["path", "out_path", "values"]}},
    {"name": "audit_document",
     "description": "この書類を、エージェントが全て正確に読み書きできるかを確かめる（窓とは別の読み方で読み落としがないか、全入力欄に書けるか、書いた値が読み戻せるか）。"
                    "読む段階（S 素で問題ない／A 規則だけで正確に読める／B エージェントが読めば正確／C エージェントが工夫して変換できる／"
                    "D エージェントの解釈が必要／E エージェントが得意でない形式／F アクセスできない）と、"
                    "書く段階（A エージェントがそのまま書ける／B 解析しないと書けない／C 画面（UI）から入力する／D 書けない（アクセスできない・手書き））、"
                    "その理由も返す。",
     "inputSchema": {"type": "object", "properties": {"path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"}}, "required": ["path"]}},
    {"name": "run_macro",
     "description": "Excel のマクロ（ボタンの処理など）を動かし、マクロが変えたセルを証明つきで out_path に書き出す（元のファイルは変えない）。"
                    "マクロの名前は窓の「マクロ」の節にある（例「Module1.印刷」「Sheet2.CommandButton1_Click」）。"
                    "MsgBox の「はい／いいえ」や InputBox、マクロが読むファイルが要るときは、勝手に決めずにその問いを返すので、"
                    "本人に確かめてから answers・files に入れて、もう一度呼ぶ。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"}, "macro": {"type": "string", "description": "動かすマクロ（モジュール名.手続き名）"},
         "out_path": {"type": "string"},
         "answers": {"type": "array", "description": "MsgBox・InputBox の答え（順に。「はい」「いいえ」、入力する値）"},
         "files": {"type": "object", "description": "マクロが読むファイル：{マクロの中のパス: 手元のファイルのパス}"}},
         "required": ["path", "macro", "out_path"]}},
    {"name": "grade_document",
     "description": "書類の段階を返す：読む段階（S〜F）と書く段階（A〜D：そのまま書ける／解析しないと書けない／画面から入力する／書けない）。"
                    "規則で決めた段階に、エージェントが窓を読んで確かめたこと（findings）を合わせる。"
                    "規則では決めないこと（押印・自署のようなエージェントにはできない記入＝書く D、電子申請などの画面からの入力＝書く C、"
                    "書く値をぼかす書き方＝読む D）は、agent_checks の問いとして返すので、窓を読んで答えてください。"
                    "見つけたら findings に {check: 問いの id, title, detail, where} で渡す（段階は問いで決まる）。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"},
         "findings": {"type": "array", "description": "見つけたこと：[{check: 'cannot_do'|'ui_only'|'vague_value', title, detail, where}]。"
                                                      "問いにないことは {score: 'read'|'write', grade, title, detail, where}",
                      "items": {"type": "object", "properties": {"check": {"type": "string"}, "score": {"type": "string"},
                                                                 "grade": {"type": "string"}, "title": {"type": "string"},
                                                                 "detail": {"type": "string"},
                                                                 "where": {"type": ["string", "array"]}}}},
         "answered": {"type": "array", "items": {"type": "string"},
                      "description": "確かめ終えた問いの id（見つからなかったものも含む。例 ['cannot_do', 'ui_only', 'vague_value']）"}},
         "required": ["path"]}},
]


# ---------------------------------------------------------------------- 追加の形式（CSV・TSV・JSON・JSON Lines・RTF・OpenDocument）
_MORE = ("（ほかに CSV・TSV・JSON・JSON Lines・RTF・OpenDocument の文書 .odt と表計算 .ods も開けます）")
for _t in TOOLS:
    if _t["name"] == "open_document":
        _t["description"] += _MORE
    elif _t["name"] == "fill_document":
        _t["description"] += (" RTF と OpenDocument の文書（.odt）は Word と同じ書き方（［P12.1］・字を目印にした {\"@1\": {\"old\": …, \"new\": …}}）、"
                              "OpenDocument の表計算（.ods）は Excel と同じ書き方（シート!B3）。CSV・TSV は「R行.C列」（R3.C2・R3.氏名。"
                              "今ある行の次の行に書くと、行を足す）、JSON は JSON Pointer（/applicant/name。ない名前は足し、配列の終わりは /items/-）、"
                              "JSON Lines は「L行/名前」。")

def _text(obj) -> dict:
    return {"type": "text", "text": obj if isinstance(obj, str) else json.dumps(obj, ensure_ascii=False, indent=1, default=str)}


def _score(check: dict) -> str:
    return "書く段階" if check.get("score") == "write" else "読む段階"


def call_tool(name: str, args: dict) -> dict:
    if name == "open_document":
        path = args["path"]
        pw = args.get("password")
        s = api.summary(path, pw)
        _k, _doc, parts = api.window(path, pw)
        head = (f"{s['kind_ja']}・入力欄 {s['fields']} か所（書ける欄 {s['writable_fields']}）・窓は {s['window_parts']} 部分"
                + ("（read_window で part を指定して続きを読めます）" if s["window_parts"] > 1 else ""))
        return {"content": [_text(head), _text(parts[0][1])], "structuredContent": s}
    if name == "read_window":
        _k, _doc, parts = api.window(args["path"], args.get("password"))
        n = int(args.get("part", 1))
        if not 1 <= n <= len(parts):
            raise api.ApiError(f"part は 1〜{len(parts)} です")
        return {"content": [_text(f"部分 {n}/{len(parts)}：{parts[n - 1][0]}"), _text(parts[n - 1][1])]}
    if name == "list_fields":
        fl = api.fields(args["path"], args.get("password"))
        if args.get("only_empty"):
            fl = [f for f in fl if f.get("value") in (None, "", False)]
        return {"content": [_text(fl)], "structuredContent": {"fields": fl}}
    if name == "fill_document":
        res = api.fill(args["path"], args.get("values") or {}, args["out_path"], bool(args.get("force", False)),
                       password=args.get("password"))
        ok = res.get("proof_ok") and res.get("read_back_ok")
        head = ("書き込みました。" if ok else "書き込みましたが、確認が必要な点があります。") + f"{res.get('proof', '')}"
        # 本人は書類を開いて確かめないので、say_to_user（書いた所の見出しと値・まだ空いている欄・本人に伝えること）を、対話で伝える
        return {"content": [_text(head), _text(res.get("say_to_user", "")), _text(res)], "structuredContent": res,
                "isError": not ok}
    if name == "run_macro":
        res = api.run_macro(args["path"], args["macro"], args["out_path"], args.get("answers") or [], args.get("files") or {},
                            password=args.get("password"))
        if not res.get("done"):
            head = res.get("note", "")
            if res.get("prompt"):
                head = f"マクロが答えを求めています：「{res['prompt']}」。{head}"
            return {"content": [_text(head), _text(res)], "structuredContent": res, "isError": False}
        head = (f"マクロ {res['macro']} を動かしました。" + (f"{res.get('proof', '')}" if res.get("out_path") else res.get("note", ""))
                + (f"　エラー：{res['error']}" if res.get("error") else ""))
        return {"content": [_text(head), _text(res)], "structuredContent": res, "isError": bool(res.get("error"))}
    if name == "grade_document":
        g = api.grade(args["path"], args.get("findings") or [], args.get("answered") or [], password=args.get("password"))
        mark = lambda r: "（エージェントの判断）" if r.get("by") == "agent" else ""  # noqa: E731
        lines = [g["headline"]] + [f"  - 読む[{r['grade']}]{mark(r)} {r['title']}: {r['detail']}" for r in g["reasons"]]
        lines += [f"  - 書く[{r['grade']}]{mark(r)} {r['title']}: {r['detail']}" for r in g["write"]["reasons"]]
        lines += [f"  - ？（あれば{_score(c)} {c['grade']}）{c['question']}" for c in g.get("agent_checks", [])]
        return {"content": [_text("\n".join(lines)), _text(g)], "structuredContent": g}
    if name == "audit_document":
        au = api.audit(args["path"], args.get("password"))
        lines = [("[OK] " if c["ok"] else "[要確認] ") + f"{c['name']}: {c['detail']}" for c in au["checks"]]
        lines.append("結論: " + ("全て正確に読み書きできます" if au["ok"] else "読み書きに確認が必要な点があります"))
        g = au.get("grade")
        if g:
            lines.append(g["headline"])
            lines += [f"  - 読む[{r['grade']}] {r['title']}: {r['detail']}" for r in g["reasons"] if r["grade"] not in ("S", "A")]
            lines += [f"  - 書く[{r['grade']}] {r['title']}: {r['detail']}" for r in g["write"]["reasons"] if r["grade"] != "A"]
            lines += [f"  - ？（あれば{_score(c)} {c['grade']}）{c['question']}" for c in g.get("agent_checks", [])]
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
        except (api.PasswordNeeded, api.PasswordWrong) as e:
            result = {"content": [_text(str(e))], "isError": False,
                      "structuredContent": {"needs_password": True, "wrong_password": isinstance(e, api.PasswordWrong)}}
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
