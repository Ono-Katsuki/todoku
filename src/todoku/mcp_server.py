"""とどくの MCP サーバー：どの AI エージェントからも、書類を「窓」で読み、証明つきで書き込めるようにする。

標準入出力で JSON-RPC 2.0 を1行ずつやりとりする（Model Context Protocol の stdio）。外部のライブラリは使わない。
道具：
- open_document   書類を開き、形式・入力欄の数・窓の最初の部分を返す
- read_window     窓（Markdown）の n 番目の部分を返す（大きな書類は分けてある）
- list_fields     入力欄の一覧（名前・見出し・種類・選択肢・いまの値・書けるか）
- fill_document   入力欄に書き込み、別の名前で書き出す。元のファイルは変えない。証明と読み戻しの結果を返す
- audit_document  この書類を「全て正確に読み書きできるか」を確かめる（読み落とし・全欄への書き込み・読み戻し）
- run_macro       Excel のマクロ（ボタン）を動かし、マクロが変えたセルを証明つきで書き出す
- extract_processing  書類の計算の処理を、規則・構造・マクロとして取り出す（format="python" で、確かめつきの calculate(inputs) のモジュール）
- grade_document  書類の段階（読む S〜F・書く A〜D）。規則で決めた段階に、エージェントが窓を読んで確かめたことを合わせる
"""

from __future__ import annotations

import json
import os
import re
import sys
import traceback
from pathlib import Path

from . import api, limits

PROTOCOL = "2025-06-18"
INSTRUCTIONS = (
    "とどく（todoku）は、申請書類（Excel .xlsx/.xlsm/.xls・Word .docx/.docm/.doc・PDF・一太郎 .jtd・OpenDocument .odt/.ods・"
    "RTF・PowerPoint .pptx・文字だけのファイル .txt/.csv/.tsv/.json・それらをまとめた ZIP）を、内容を一字も変えずに読める「窓」にし、"
    "入力欄に証明つきで書き込む道具です。韓国のワープロ（한글）の HWPX（.hwpx）も読んで書け、HWP 5（.hwp）は読むだけです"
    "（HWP 5 に書くときは、本人に HWPX で保存し直したファイルを頼みます）。"
    "大きな書類は、読む量を分けて読みます（読めない所はありません）：open_document が large の概要を返したら、要る所を read_window に範囲で頼みます"
    "（PDF は pages、Excel・OpenDocument の表は sheet と rows、CSV・TSV は rows、JSON Lines は records、JSON は path_in（JSON Pointer）、.txt・Word・一太郎・한글・PowerPoint・OpenDocument の文書・RTF は paragraphs。結果の next をそのまま渡すと続きが読め、"
    "たどり終えると全部を読んだことになります）。list_fields も同じ範囲で頼みます。段階の deferred（未確認）は、まだ確かめていないことなので、確かめたことのように伝えないでください。"
    "文書に埋め込んだ大きな表は、開くときには解析しません：窓の案内のとおり、read_window・list_fields に sheet: \"E1!シート\" と rows で頼みます。"
    "open_document などが「長い仕事にしました」と返したら、失敗ではありません。もう一度同じ呼び出しをすると、進み具合か結果が返ります。"
    "検査 audit_document は長い仕事として動き、呼ぶたびに進み具合を返します。呼び出しが上限で止まったら、結果の smaller_request の頼み方で頼み直します。"
    "まず open_document で書類を開き、read_window で窓を読み、list_fields で入力欄の名前を確かめてから、"
    "fill_document で書き込んでください。書き込みは元のファイルを変えず、別の名前で書き出します。"
    "窓の［T1.R2.C3］［p1.t45-112］［シート!B3］などが入力欄の名前です。推定した見出しは context（その行）と合わせて確かめてください。"
    "PDF の所の id は位置で決まる：p{ページ}.{作りの頭の字}{左}-{上}（ページの左上からの mm。t 空のマス・m 空のマスの並び・g 行の中の空き・b 箱の字・u 下線の罫線。例 p1.t45-112）。前の版の番号の id（p1.3・p1.a12）は、同じ所と決められないので断る。所の facts は mm（rect_mm・width_mm・前後の字・囲むマスの字と離れ方）、字の大きさは size と size_from（どこから決めたか）。格子に並ぶ空のマスは「［p1.t{x}-{y}］ 空のマス n 個：横 a × 縦 b」とまとめ、x・y の mm の並びと、列の上・行の左の字を示すので、そこから p1.t{x}-{y} の id を作って書く。窓の「［↓余白 12.0 mm］」は、欄にしていない、字のあるマスの字の下の空き（↑ は上）：書く所と判断したら after と where: below（上は box）で書く。"
    "［P12.c5？］のように「？」のついた所は、規則では欄と決められない所です。窓を読んで判断し、書く所なら同じように id で書きます。"
    "とどくは推測で決めず、事実を出します：見出しは印刷された字のとおり（左・上のマスの字と、その離れ方：row_label・column_label・near）、"
    "所の字・幅・前後の字（facts）。どれを見出しとして読むか、書く所か、どこに寄せるかは、エージェントが事実を読んで決めます。"
    "段落・セルの中の所の id は位置で決まり（Word などは P段落.c前の字数、Excel は シート!B3@字の位置）、ほかの所が増えても変わりません。"
    "窓に欄として示していない所にも、字を目印にして書けます（fill_document の説明を見てください）。"
    "本人は書類を Word・Excel・Acrobat で開いて確かめません。書き込みの結果の say_to_user（記入した内容の一覧・まだ空いている欄・"
    "本人に伝えること）は、対話で確かめるための材料です。決まった型で読み上げるより、会話の流れに合わせて自然な言葉で伝えてください"
    "（ただし「本人に伝えること」と、書いた値そのものは省かない）。"
    "本人に分かりきったことを一つずつ聞かないでください。住所・名前・事業所の情報など、本人がこれまでに伝えたこと・ほかの書類・記憶から"
    "分かる値は、先にまとめて書いて、記入した内容を伝え、直す所がないかを確かめるのが基本です。分からない値・本人が決めること"
    "（選ぶもの・マクロの「はい／いいえ」など）は勝手に決めず、なるべくまとめて、選びやすい形で聞いてください。"
    "ただし手順の決まりではありません。本人が会話しながら一緒に埋めたいときは、その流れに合わせ、言われた所から書いてかまいません。"
    "パスワードで保護された書類は「パスワードが要る」と返るので、本人にパスワードを聞いて、どの道具にも password で渡してください"
    "（パスワードは結果にも記録にも残しません。書き出したファイルも、元と同じパスワードで保護されます）。"
    "文字だけのファイル（CSV・.txt など）・古い表（.dbf・.wk1 など）や書庫・メールの名前で、窓に「文字コードを規則では決められません」と候補が並んだら、字を読んで意味の通る方を"
    "選び、そのファイルのどの道具にも encoding で渡してください。"
    "窓の「画像」の節で「説明（代わりの字）がありません」とある画像は、中身を字として読めません。申請に関わりそうなら view_image（PDF の"
    "ページ全体は view_page）で画像を見て読み、読めなければ、その位置（ページ・上下左右・近くの字）を本人に伝えてください。中身を推測で"
    "埋めないでください。画像そのもののファイル（JPEG・PNG・TIFF・SVG など）も open_document で開き、view_page で見ます。RAR・HEIC・HEIF・AVIF と LZH の -lh1-〜-lh3- は、とどくは解きません：open_document の needs_tool（format・why・next_step）に、別の道具のコマンドの例があります。入っている道具を確かめて使い、できたファイルを open_document で開いてください。変換・展開したことは本人に伝えてください（書いたものは元の形式には戻りません）。計算の決まりを知りたいときは、値から推測せず extract_processing で規則として受け取ります。段階の alerts（アクセシビリティの注意）は、"
    "段階とは別に伝えてください。"
    "fill_document の結果の read_back_ok が false のときは、ファイルは書き出し済みですが、読み直した値が"
    "書いた値と違う欄があります（read_back_differences）。確かめずに本人に渡さず、違いを読んで直すか、そのことを本人に伝えてください。"
    "Excel の proof_scope は確かめた範囲の数です（連動しないセルと一致数・照らした計算結果と一致数・連動を決められないセル・計算できないセル）。"
    "結果に deferred（未確認）があれば、ブック全体の計算し直しは済んでいません：書き出したファイルを audit_document に渡すと、確かめて written_proof に返します。"
    "済むまで、確かめたと伝えないでください。written_proof の problems があれば、そのファイルを本人に渡さないでください。"
    "結果に check_visually（書く B）があるときは、その欄を view_page の画像で見て、字の形が正しいかを確かめてください"
    "（字を組み立てて描いた文字は、字の並びは確かめ、見た目は確かめていません）。"
    "結果に replaced_existing があるときは、空欄に書いたのではなく、書いた所にもとからあった字や値を変えています（how が removed：字が"
    "消えた・数が変わった。inside：字は残り、字のあいだに字が入った）。どの形式でも、書く前と後の字を比べて出します。"
    "その 2 つの字を本人に伝え、本人が頼んだ書き換えかを確かめてください。"
    "証明（書いた所のほかは元のまま）と読み直しは、とどくが書き出したファイルについて確かめたことです。本人がほかの道具（Python の"
    "ライブラリなど）での書き直しを求めたときは、そのファイルにはこの確かめがなくなること（書いた所のほかが元のままかは分からなくなること）を、"
    "本人に必ず伝えてください：切り替える前に伝えて本人に決めてもらうのが基本で、すぐに作り直すよう求められたときも、作り直した報告で省かないでください。"
)

from .fillhelp import TOPICS as FILL_TOPICS, help_text as fill_help  # noqa: E402

READ_TABLE_LIMIT = 200  # read_table：out_path がないとき、結果に入れる行の数の既定の上限（limit で変えられる）
_OVERWRITE = {"type": "boolean", "default": False,
              "description": "out_path にすでに別の中身のファイルがあるとき、置き換える（本人が了承したときだけ。元のファイルは決して置き換えない）。"
                             "書き込みを証明か読み戻しで確かめられなかったときは、out_path でなく「名前.unverified.拡張子」に書き出す"
                             "（そのファイルは提出しない）"}

_RANGE_SCHEMA = {"type": ["string", "integer"], "description": "1 から数えた範囲（\"1-20\"・\"500-\"・5）"}
TOOLS = [
    {"name": "open_document",
     "description": "書類（Excel .xlsx/.xlsm/.xls・PDF・Word .docx/.docm/.doc・一太郎 .jtd）を開き、形式・入力欄の数・窓の分け方と、"
                    "窓（Markdown）の最初の部分を返す。한글の HWPX（.hwpx）は読む・書く、HWP 5（.hwp）は読むだけ"
                    "（配布用・パスワードの HWP は読めない。본 제품은 한글과컴퓨터의 한/글 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.）。"
                    "ZIP は中のファイルの一覧を返す。中のファイルは、どの道具でも path に「a.zip!/中の名前」と指定して開く・書く"
                    "（書き出す out_path は ZIP。書いたファイルの部品だけを入れ替える）。",
     "inputSchema": {"type": "object", "properties": {"path": {"type": "string", "description": "書類のファイルのパス"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"}},
                     "required": ["path"]}},
    {"name": "read_window",
     "description": "窓（Markdown）を返す。ふつうの書類は part 番目の部分（1から）。大きな書類は範囲で頼む：PDF は pages（\"1-20\"）、"
                    "Excel・OpenDocument の表は sheet と rows（\"1-500\"）、CSV・TSV は rows、JSON Lines は records、JSON は path_in（JSON Pointer）、.txt・Word・一太郎・한글・PowerPoint・OpenDocument の文書・RTF は paragraphs（行・段落の範囲）。"
                    "どの書類でも範囲で読める。結果には、読んだ範囲・残り・続きの頼み方（next）が付く。next をそのまま渡していけば、"
                    "全部を読める（読む量を分けるだけで、読めない所はない）。max_chars で 1 回の大きさ（字）を変えられる（結果の上限 TODOKU_RESULT_CHARS を超える分は上限に収める）。"
                    "部分の番号は変わらない。1 つの部分が結果の上限に入らないときは、部分の中を分けて返す（next の cursor で同じ部分の続きを読む）。",
     "inputSchema": {"type": "object", "properties": {"path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"},
                                                      "part": {"type": "integer", "minimum": 1, "default": 1},
                                                      "pages": _RANGE_SCHEMA, "rows": _RANGE_SCHEMA, "paragraphs": _RANGE_SCHEMA, "records": _RANGE_SCHEMA,
                                                      "path_in": {"type": "string", "description": "大きな JSON の中の値の位置（JSON Pointer。例 /items/0）"},
                                                      "sheet": {"type": "string", "description": "Excel のシート・Access の表の名前（rows と一緒に）。文書に埋め込んだ表は「E1!シート」"},
                                                      "cursor": {"type": "string", "description": "前の結果の next にある、続きの印"},
                                                      "max_chars": {"type": "integer", "minimum": 200, "description": "1 回に返す大きさ（字）"},
                                                      "raw": {"type": "boolean", "description": "窓を作らずに、ファイルの字を順に読む（欄は見つけない）。窓を作るのが上限で止まった書類のための道"}},
                     "required": ["path"]}},
    {"name": "list_fields",
     "description": "入力欄の一覧を JSON で返す：id（書き込むときの名前）・label（見出し）・kind（種類）・options（選択肢）・value（いまの値）・readonly（書けない欄）・context（その行）。"
                    "同じ行の欄が続くときは、行を 1 度だけ返す：2 つ目からは context の代わりに context_same_as（その行を context に持つ、前の欄の id）。"
                    "行の中のその欄の所は［その欄の id…］（表計算は［番地…］。▶ は、行を持つ欄の所の印）。"
                    "大きな書類は範囲で頼む（read_window と同じ pages・sheet・rows・paragraphs・records。空のマスは only_empty）。結果の next で続きを頼める。"
                    "kind が range の項目は「範囲の所」：入力規則や数式の参照が広いとき、空のセルをまとめて 1 つで示す（id は S!C5:C40000。count が空のセルの数）。"
                    "中のどのマスも番地（S!C7）で書ける。sheet と rows を渡すと、その行の分を 1 マスずつの欄で返す（文書に埋め込んだ表は sheet に「E1!シート」）。"
                    "結果が大きいときは、書くのに要る事実を残して、根拠の事実（near・evidence・why など）を省き、省いたものと頼み方を omitted に書く。"
                    "ids（id のリスト）を渡すと、その欄の事実をすべて返す（context も欄ごとに）。detail: true なら、どの欄の事実も省かず、分けて返す（next で続く）。",
     "inputSchema": {"type": "object", "properties": {"path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"},
                                                      "only_empty": {"type": "boolean", "default": False},
                                                      "pages": _RANGE_SCHEMA, "rows": _RANGE_SCHEMA, "paragraphs": _RANGE_SCHEMA, "records": _RANGE_SCHEMA,
                                                      "path_in": {"type": "string", "description": "大きな JSON の中の値の位置（JSON Pointer。例 /items/0）"},
                                                      "sheet": {"type": "string"}, "cursor": {"type": "string"}, "limit": {"type": "integer", "minimum": 1},
                                                      "ids": {"type": "array", "items": {"type": "string"}, "description": "この欄の事実（near・evidence・why なども）をすべて返す"},
                                                      "detail": {"type": "boolean", "default": False, "description": "どの欄の事実も省かず、大きければ分けて返す"},
                                                      "skip": {"type": "integer", "minimum": 0, "description": "同じ頼み方の結果の、はじめのこの数の欄を飛ばす（結果の next に入っているのを、そのまま渡す）"}},
                     "required": ["path"]}},
    {"name": "fill_document",
     "description": "入力欄に書き込み、out_path に書き出す（元のファイルは変えない）。values は {欄の id: 値}。"
                    "結果は、変わったのが書いた欄（とその計算結果・見た目）だけであることの証明と、書き込み後のファイルを読み直した結果。"
                    "本人は書類を開いて確かめないので、結果の say_to_user（書いた所の見出しと値・まだ空いている欄・本人に伝えること）を、対話でそのまま本人に伝えて確かめる。"
                    "values_not_written_now は今回書いていないのに値の入っている欄（様式の既定の値・前の書き込み）。viewers_differ は、見た目を使わず描き直す表示（Mac のプレビュー・iPhone）で字の終わりが欠け、違う値に見える欄。どちらも本人に伝える。fonts_not_embedded は書体を埋め込まずに描いた字（同梱の書体に字形がない字だけ。cjk_predefined は開くソフトの書体に頼る）。"
                    "欄のない所にも書ける：印刷された選択肢の字を○で囲める（Excel・PDF・Word。詳しくは help の circle）、印刷された字を書き換えられる（old/new。詳しくは help の text）。印を付けられないと本人に伝える前に、これを確かめる。"
                    "\n【値】チェックは true/false（JSON の真偽値）か、その欄が宣言した状態の値（PDF の状態の名前・XFA の値・箱の字）だけ："
                    "「はい」「あり」「有」などの言葉は受けないので、本人の答えを true/false に直して書く。"
                    "表のセルの数は JSON の数か「1234.56」の形の字（「1.234」は 1.234、「1,234」は 1234 と読む）、日付は年が先の形（2026-11-03）で書く。"
                    "本人が確かめる書類ではいつも意味 as を添える：{欄の id: {\"value\": 値, \"as\": \"生年月日\"}}（書類には書かれない。@ の書き込み・○にも。as のなかった所は written_without_as）。"
                    "\n【欄の id】窓の id をそのまま使う。Word などの「P12.c5」は段落の中の所（位置で決まる。古い P12.1・P12.a1 は 1 つに決まるときだけ受ける）、Excel は「シート!B3」"
                    "（セルの字の中の空欄は、空欄ごとの値のリストか、所の id「シート!B3@5」）。［…？］は規則では欄と決められない所：facts と窓を読んで、書く所と判断したときだけ書く。"
                    "alternative_to のある欄は同じ所の別の読み方で、どちらか1つにだけ書く（両方は断る）。Word の箇条書きの □（☐［P段落.b］）は true で書く。"
                    "kind_candidates のある PDF の欄は {\"kind\": \"text\", \"value\": 値} で種類を決められる。"
                    "\n【欄のない所】values に \"@1\" などの鍵で書く："
                    "字を書き換える（Word・Excel など）{\"old\": 元の字, \"new\": 後の字}（old と new の違う所だけが変わる。同じ字がいくつもあれば near か occurrence、Excel は cell・sheet。Word のフィールドの結果の字は断る）。"
                    "字の前・後ろに足す {\"before\" か \"after\": 字, \"value\": 値}（PDF の after は見出しの右の空き、\"where\": \"below\" で下。\"page\"）。"
                    "PDF の枠 {\"page\": ページ, \"box\": [左, 上, 右, 下]（mm）, \"value\": 値}。PDF の伏せ字（「○○」）の上 {\"cover\": \"○○\", \"value\": 値}（元の字は消さない）。"
                    "\n【○で囲む】{\"@c1\": {\"circle\": 囲む字, \"near\": 近くの見出し}}：窓は印刷された選択肢（有・無など）を欄として並べないので、字を読んで選ぶ。"
                    "同じ字がいくつもあれば near か occurrence（Excel は cell でも）で 1 つに決める。決まらなければ候補を示して断る。PDF は楕円（字に触れるなら断るので style・box を決める）、Excel（.xlsx）は楕円の図形、"
                    "Word（.docx）は既定で字の四角い囲み線。古い .xls・.doc には描けない。Excel・Word の位置は見積もりなので uncertain を本人に伝える。"
                    "\n【本人の印影・署名の画像】本人に、どの書類のどこにどの画像を置くかを伝えて同意を得たときだけ {\"@1\": {\"seal\" か \"signature\": 画像のパス, \"consent\": true, \"page\": ページ, \"box\": [左, 上, 右, 下]（mm）}}"
                    "（PDF のみ。consent: true がなければ置かない。字を目印に cover・after でも置ける）。実印・自署（手書き）が求められる所では受け付けられないことがある：様式の注意書きを読んで本人に伝える。"
                    "\n【明示しないと断る】印刷された字のあるセルをまるごと書き換える {\"value\": 値, \"overwrite_text\": true}。PDF の画面にも紙にも出ない欄 {\"value\": 値, \"allow_hidden\": true}（隠れた欄と分かったうえでだけ）。"
                    "\n【字の置き方】既定は左で、字の意味からは決めない。PDF の空き・枠と Word・Excel の字の中の空欄は {\"value\": 値, \"align\": \"left\"|\"center\"|\"right\"}（PDF は pad_mm も）。書いた枠は結果の written_rects。"
                    "\n【書いた後】PDF の重なりは、はっきりしたものだけ断り、きわどい所は warnings に所と位置（mm）を添えて知らせる。書いた PDF を view_page で見て、warnings の所を一つずつ確かめ、読みにくければ box・style を変えて書き直す。"
                    "Excel・Word などは描いて確かめられないので、say_to_user の見出しと値・折り返し／はみ出し／「####」の警告（wrap・fit）・○の uncertain を確かめる。"
                    "\n【ほかの形式】.odt・.rtf・.txt・.pptx・.hwpx は Word と同じ、.ods は Excel と同じ。CSV・TSV は「R3.C2」、JSON は JSON Pointer、JSON Lines は「L行/名前」。HWP 5 には書けない。"
                    "\n【詳しい書き方】help に話題（id：欄の id の読み方・text：字の書き換えと差し込み・pdf：PDF の欄のない所・align：字の置き方・as：値の意味・seal：印影と署名・circle：○で囲む・formats：ほかの形式・all：すべて）を渡すと、書き込まずに説明だけを返す（そのときの path・out_path は空の字、values は {} でよい）。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"}, "out_path": {"type": "string", "description": "書き出すファイルのパス（元と別の名前）"},
         "values": {"type": "object", "description": "{欄の id: 値}。本人が確かめる書類では、値に意味を添える {欄の id: {\"value\": 値, \"as\": \"生年月日\"}}。"
                                                      "Word・Excel で字の前・後ろに足すときは {\"@1\": {\"before\" か \"after\": 字, \"value\": …}}、"
                                                      "Word で欄のない所は {\"@1\": {\"old\": …, \"new\": …}}、"
                                                      "PDF で欄のない所は {\"@1\": {\"after\": …, \"value\": …}}、"
                                                      "本人が同意した印影・署名の画像は {\"@1\": {\"seal\": 画像のパス, \"consent\": true, …}}、"
                                                      "印刷された字を○で囲むときは {\"@c1\": {\"circle\": 囲む字, \"near\": 近くの見出し}}"},
         "force": {"type": "boolean", "default": False, "description": "保護された欄にも書く（本人が了承したときだけ）"},
         "overwrite": _OVERWRITE,
         "help": {"type": "string", "enum": [*FILL_TOPICS, "all"], "description": "書き込まずに、この話題の詳しい書き方だけを返す（circle は○で囲む、text は字の書き換え）"}},
         "required": ["path", "out_path", "values"]}},
    {"name": "audit_document",
     "description": "大きな書類では長い仕事として動く（すぐ返り、呼ぶたびに進み具合・ここまでの結果を返す。止まっても控えの続きから進む。job: true で大きさによらず長い仕事に、restart: true で始めから）。この書類を、エージェントが全て正確に読み書きできるかを確かめる（窓とは別の読み方で読み落としがないか、全入力欄に書けるか、書いた値が読み戻せるか）。"
                    "読む段階（中身を読むことだけで決める。S 素で問題ない／A 規則だけで正確に読める（地域で見え方が変わる組み込みの日付の表示形式は ISO 8601 の字で示して伝える）／B エージェントが読めば正確"
                    "（中身は窓にすべて出ていて、決まらない所は窓の根拠でエージェントが決める：文字コードの候補・再現できない段落番号の書式など。説明のない画像は段階を変えず、注意 alerts に示す）／"
                    "C エージェントが工夫して変換できる（マクロ・確かめられない数式・窓の読み落とし）／"
                    "D エージェントの解釈が必要／E エージェントが得意でない形式（スキャン・中身が画像）／F アクセスできない）と、"
                    "書く段階（A エージェントがそのまま書ける／B 解析しないと書けない／C 画面（UI）から入力する／D 書けない（アクセスできない・署名・印・拇印・公証など、手で行う所））、"
                    "その理由も返す。",
     "inputSchema": {"type": "object", "properties": {"path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"}}, "required": ["path"]}},
    {"name": "run_macro",
     "description": "Excel のマクロ（ボタンの処理など）を動かし、マクロが変えたセルを証明つきで out_path に書き出す（元のファイルは変えない）。"
                    "マクロの名前は窓の「マクロ」の節にある（例「Module1.印刷」「Sheet2.CommandButton1_Click」）。"
                    "MsgBox の「はい／いいえ」や InputBox、マクロが読むファイルが要るときは、勝手に決めずにその問いを返すので、"
                    "本人に確かめてから answers・files に入れて、もう一度呼ぶ。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"}, "macro": {"type": "string", "description": "動かすマクロ（モジュール名.手続き名）"},
         "out_path": {"type": "string"}, "overwrite": _OVERWRITE,
         "answers": {"type": "array", "description": "MsgBox・InputBox の答え（順に。「はい」「いいえ」、入力する値）"},
         "files": {"type": "object", "description": "マクロが読むファイル：{マクロの中のパス: 手元のファイルのパス}"}},
         "required": ["path", "macro", "out_path"]}},
    {"name": "read_table",
     "description": "Excel の表を、見出しの列で行ごとに取り出す（エージェントが表を読む処理を書かなくてよいように）。"
                    "header は表の見出しにある字（例「合成」「サービスコード」）。その字の行（縦に結合した見出しはその範囲）を見出しとする。"
                    "その下の行が2段目の見出しかどうかは規則で決めず、rows_below_header として返すので、読んで、続きなら header_rows を"
                    "指定してもう一度呼ぶ。列の名前は見出しを上から「／」でつなぐ。結合セルは画面どおりに範囲へ広げ、値は表示形式どおりの字、値ごとに元のセルを添える。"
                    "sheet を省くと、その見出しのあるすべてのシートから。行が多いときは out_path（.csv か .json）に書き出す"
                    "（結果には列の名前・行数・はじめの 20 行）。表の後ろの注の行もそのまま返すので、どこまでが表かは読んで決める。"
                    "返さなかった行（使っている範囲の中の値のない行など）は omitted_rows に数と行の番号の範囲で書く。"
                    "Access の表も読める：sheet に表の名前、header に列の名前。rows で行の範囲を頼むと、その範囲だけを返す（1 行目が列の名前、2 行目からレコード）。"
                    "すべての列が NULL のレコードも行として返す。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"},
         "header": {"type": "string", "description": "表の見出しにある字"},
         "sheet": {"type": "string", "description": "シートの名前（省くと、見出しのあるすべてのシート）"},
         "header_rows": {"type": "integer", "minimum": 1, "description": "見出しの行数（省くと自動）"},
         "out_path": {"type": "string", "description": "すべての行を書き出すファイル（.csv か .json）"},
         "rows": {**_RANGE_SCHEMA, "description": "sheet の行の番号の範囲（例「2-501」）。その範囲の行だけを返す（sheet といっしょに）"},
         "cursor": {"type": "string", "description": "大きなブックで、前の結果の next の続きを読むとき"},
         "limit": {"type": "integer", "minimum": 1, "description": f"out_path がないとき、結果に入れる行の数の上限（既定 {READ_TABLE_LIMIT}）。"
                   "超えた行は返さず、行の数（rows）と返した行の数（rows_returned）を返す"}},
         "required": ["path", "header"]}},
    {"name": "explain_calculation",
     "description": "1 つの値の説明：cell を指すと、そのセルの計算を、番地を見出しに置き換えた式（例「＝「報酬総額」-「処遇改善加算の総額」」）で入力欄までさかのぼる。"
                    "cell を省くと、入力欄と結果のまとまりの一覧。本人に「この金額はどう決まるの？」と説明するときに使う。"
                    "計算の決まりそのものを取り出すのは extract_processing。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"},
         "cell": {"type": "string", "description": "説明するセル（シート!番地）。省くと、結果のまとまりの一覧"}},
         "required": ["path"]}},
    {"name": "try_calculation",
     "description": "Excel の様式で、値の組から結果を試算する（ファイルは書かない）。values は fill_document と同じ {欄の id: 値}。"
                    "「区分Ⅱにしたらいくら？」のような問いに、書類を書き換えずに答えるため。入力規則に合わない値は断る。"
                    "結果の say_to_user（変わる所の見出しと、前 → 後）を、会話で伝える。"
                    "cells（[\"シート!A5\", …]）を渡すと、そのセルの計算結果を返す（動的配列の広がった先のマスも、そのマスだけ計算する）。"
                    "とても大きな配列の集計が時間の中で終わらないときは、進んだ所を伝え、別のプロセスで続ける（同じ呼び出しをもう一度すると、進み具合か結果を返す）。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"},
         "values": {"type": "object", "description": "{欄の id: 値}"},
         "cells": {"type": "array", "items": {"type": "string"}, "description": "結果を見たいセル（シート!番地）"}},
         "required": ["path"]}},
    {"name": "extract_processing",
     "description": "書類の計算の処理を、処理として取り出す（値ではなく、計算の決まりそのものを受け取る道具）。Excel の仲間（.xlsx・.xls・.xlsb・.ods など）と PDF のフォーム。"
                    "同じ形の数式は 1 つの規則にまとめる（例「D3:D100000 = B{row}*C{row}」。A1 と R1C1 の正規形、読む範囲、例外のセルつき。10 万行でも規則は数個）。"
                    "あわせて、入力（見出しの事実・読む規則・入力規則）、パラメータ、参照表（中身はデータとして。read_table で引ける）、名前、出力（どこからも読まれないセル）、"
                    "規則の順番と循環、IF・IFS・SWITCH の分岐、配列数式（展開しない）、外部参照、TODAY・OFFSET・INDIRECT（静的に分からない所）、隠れたシート・行、シートのあいだの流れ、"
                    "マクロ（ソース・読み書きするセル・いつ動くか）、条件付き書式・入力規則、ピボット・Power Query（セルの数式で書けない処理としての定義）、PDF の計算の順と欄のあいだの規則（独自の JavaScript は計算せずそのまま）。"
                    "format=\"python\" で、標準ライブラリだけで動く calculate(inputs) -> dict のモジュールを out_path に書き出す。書き出す前に、別のプロセスで動かして保存された値・とどくの計算と全セルを比べ、"
                    "一致の割合・すべての不一致・規則ごとの確かめの結果（verification）を返す。計算の規則を知りたいときはこの道具を使う。1 つの値の説明は explain_calculation、値を変えた試算は try_calculation。"
                    "結果が上限（TODOKU_RESULT_CHARS）に入らないときは、判断の土台（summary・verification・python・say_to_user・limits など）と、規則の id・sheet・range・text を残す。"
                    "入らない規則は next をそのまま渡すと続く（全規則をたどれる。続きは規則だけで、書き出しも確かめもしない）。省いたものと頼み方は omitted に書く。"
                    "規則の詳細（reads・a1・r1c1 など）は ids に規則の id、入力・参照表などの節は section に節の名前（どちらも入らなければ next で続く）、シートごとは sheet、全部を 1 つのファイルでは out_path の JSON。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"},
         "format": {"type": "string", "enum": ["json", "python"], "description": "json（既定）か python"},
         "out_path": {"type": "string", "description": "python：書き出すモジュール（.py。必須）。json：全部を書き出す .json（大きなブックのとき）"},
         "sheet": {"type": "string", "description": "このシートの規則・入力だけを返す"},
         "ids": {"type": "array", "items": {"type": "string"}, "description": "この規則（R1 など）の事実をすべて返す。結果が大きくて省いた詳細を取るとき。format・out_path とは一緒に使わない"},
         "section": {"type": "string", "description": "結果の 1 つの節（inputs・lookup_tables・rules など。omitted.sections の名前）を省かずに返す。format・out_path とは一緒に使わない"},
         "skip": {"type": "integer", "minimum": 0, "description": "続きの位置（前の結果の next をそのまま渡す）"},
         "overwrite": _OVERWRITE},
         "required": ["path"]}},
    {"name": "view_page",
     "description": "PDF・XPS のページ、または画像のファイル（JPEG・PNG・GIF・BMP・TIFF・WebP・JPEG 2000 など。page は何枚目か）を画像で見る（スキャンした紙のように、字の情報がない画像だけのページを読むため）。"
                    "上と左の帯に、ページの左上からの mm の目盛りが付く。小さな字は region（[左, 上, 右, 下] mm）で拡大して読む。"
                    "書く所は、画像を見て決めた枠を fill_document に {\"@1\": {\"page\": ページ, \"box\": [左, 上, 右, 下]（mm）, \"value\": 書く字}} で渡す。"
                    "書く前に marks（[{\"box\": [...], \"label\": \"1\"}]）を渡すと、その枠を赤で描くので、位置が合っているかを目で確かめる。"
                    "ページは、字（埋め込みのフォントの字形）・線・画像・欄の見た目（/AP）を描く。書いた後のファイルを見ると、値の入った欄が青い枠で描かれ、値は文で返る。"
                    "書いた字が欄からはみ出す・つぶれる・横倒し・印刷された字に重なるなら、書いた見た目が正しくないので、書き直す（書いた後は、いつも見て確かめる）。"
                    "フォントが埋め込まれていない字は代わりのフォントで描き（結果の substituted）、使える字形がない字は灰色の四角になる。"
                    "JBIG2・JPEG 2000 の画像も描く（JPEG 2000 は、画面に要る解像度で開き、そのことを結果に書く。細かな字は region で拡大：範囲に掛かる所だけを開くので、どの部分でも元の解像度まで開ける。）。開けない画像は灰色の枠と対角線で示して結果に書くので、その所は描けたと思わず、本人に伝える。そのとおりに描けなかったものは、結果の not_drawn に短く並ぶ。置いた印影・署名の画像も、その位置に描く。画像の中の字は推測で埋めず、読めない所は本人に伝える。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"},
         "page": {"type": "integer", "minimum": 1},
         "region": {"type": "array", "items": {"type": "number"}, "description": "[左, 上, 右, 下]（mm）。省くとページ全体"},
         "marks": {"type": "array", "items": {"type": "object"}, "description": "赤で描く枠：[{box: [左, 上, 右, 下]（mm）, label: 番号}]"},
         "dpi": {"type": "number", "minimum": 10, "maximum": 1200,
                 "description": "解像度。省くとページ全体は 100 dpi、region は長い辺 1600 画素。画像の画素の数には上限があり、超えるときは小さくする"}},
         "required": ["path", "page"]}},
    {"name": "view_image",
     "description": "書類の中の画像を見る（Word・Excel・PowerPoint・OpenDocument・HWPX・.doc・.xls・RTF の絵と、PDF のページの画像・注釈の画像）。"
                    "窓の「書類の中の画像」の id（図1・シート!図1・p1.i1・p1.a1 など）か、部品の名前で指す。PNG・JPEG はそのまま渡し、GIF・BMP・TIFF・WebP・JPEG 2000 は PNG に、"
                    "EMF・WMF は描いて PNG にする（描けない記録は not_drawn。絵の中の字は text にも出す）。読めない所は推測で埋めず、本人に伝える。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"}, "image": {"type": "string", "description": "画像の id か部品の名前"},
         "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"},
         "region": {"type": "array", "items": {"type": "number"}, "description": "[左, 上, 右, 下]（画素）。大きな画像の一部を元の細かさで見るとき"}},
         "required": ["path", "image"]}},
    {"name": "to_pdf",
     "description": "画像のファイル（JPEG・PNG・TIFF など）を PDF に包む。画像には直接書けないので、できた PDF に fill_document で枠（mm）を指定して書く（証明はその PDF について行う）。"
                    "元の画像は変えない。JPEG はそのまま入れる。解像度の記録のない画像は A4 に収まる大きさで包み、そのことを結果に書く。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"}, "out_path": {"type": "string", "description": "書き出す PDF の場所（元のファイルとは別の名前）"},
         "overwrite": {"type": "boolean"}}, "required": ["path", "out_path"]}},
    {"name": "grade_document",
     "description": "書類の段階を返す：読む段階（S〜F）と書く段階（A〜D：そのまま書ける／解析しないと書けない／画面から入力する／書けない）。"
                    "規則で決めた段階に、エージェントが窓を読んで確かめたこと（findings）を合わせる。"
                    "規則では決めないこと（署名・印・拇印・公証など、手で行う所＝書く D、電子申請などの画面からの入力＝書く C、"
                    "書く値をぼかす書き方＝読む D）は、agent_checks の問いとして返すので、窓を読んで答えてください。"
                    "見つけたら findings に {check: 問いの id, title, detail, where} で渡す（段階は問いで決まる）。"
                    "所の多い理由の where は、はじめの所と「ほか N か所」（すべての所は where_all: true）。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"},
         "findings": {"type": "array", "description": "見つけたこと：[{check: 'cannot_do'|'ui_only'|'vague_value', title, detail, where}]。"
                                                      "問いにないことは {score: 'read'|'write', grade, title, detail, where}",
                      "items": {"type": "object", "properties": {"check": {"type": "string"}, "score": {"type": "string"},
                                                                 "grade": {"type": "string"}, "title": {"type": "string"},
                                                                 "detail": {"type": "string"},
                                                                 "where": {"type": ["string", "array"]}}}},
         "answered": {"type": "array", "items": {"type": "string"},
                      "description": "確かめ終えた問いの id（見つからなかったものも含む。例 ['cannot_do', 'ui_only', 'vague_value']）"},
         "where_all": {"type": "boolean", "default": False, "description": "所の多い理由の、すべての所（where_all）も返す"},
         "skip": {"type": "integer", "minimum": 0, "description": "すべての所が一度に入らないときの、続きの位置（前の結果の next をそのまま渡す）"}},
         "required": ["path"]}},
]


# ---------------------------------------------------------------------- 追加の形式（CSV・TSV・JSON・JSON Lines・RTF・OpenDocument・.txt・PowerPoint）
_MORE = ("（ほかに CSV・TSV・JSON・JSON Lines・RTF・OpenDocument の文書 .odt と表計算 .ods・文字だけのファイル .txt・"
         "PowerPoint .pptx・한글 .hwpx（読む・書く）と .hwp（HWP 5。読むだけ）も開けます。"
         "Excel .xlsb（読む。値のあるセルは同じ大きさの上書き、空のセルはレコードを足して書く）・dBase .dbf（読む。項目の幅の中の上書きを書ける）・古い PowerPoint .ppt・XPS .xps/.oxps・EPUB・メール .eml/.msg・Excel 2〜4・Word 1.x/2.0・"
         "Lotus .wk1・Access .mdb/.accdb・OpenOffice.org 1.x .sxw/.sxc・iWork .pages/.numbers/.key は読むだけ。"
         "tar・.tar.gz・.gz（読む・書く）と .tar.bz2・.tar.xz・.bz2・.xz・LZH・7z（読むだけ。パスワードは password）は ZIP と同じく「a.tar!/中の名前」で中のファイルを指す。RAR・HEIC・HEIF・AVIF は別の道具で展開・変換してから開く（結果の needs_tool に手順）。"
         "DocuWorks は読めないので、次の一手を返す）")
for _t in TOOLS:
    if _t["name"] == "open_document":
        _t["description"] += _MORE

# 文字だけのファイル（CSV・TSV・JSON・JSON Lines・.txt）の文字コード：規則で決められず、窓に候補が並んだとき、エージェントが選んだもの
_ENCODING = {"type": "string",
             "description": "文字だけのファイル（CSV・TSV・JSON・JSON Lines・.txt）で、窓に「文字コードを規則では決められません」と"
                            "候補が並んだとき、字を読んで選んだ文字コード（候補の encoding の値。例 \"cp932\"・\"gbk\"）。"
                            "その文字コードで厳密に読めることを確かめて、窓・欄・書き込み・段階をその文字コードで読み直す"
                            "（書き込みは、書いた所だけをその文字コードのバイトにする）。同じファイルには、どの道具にも同じ値を渡す。"
                            "書庫・メール（ZIP・tar・gzip・LZH・.msg）で、窓に「名前の文字コードを規則では決められません」と候補が並んだときは、"
                            "その書庫の path に、選んだ文字コードを渡すと、一覧を読み直す（中のファイルは、選んだ読み方の名前でそのまま指せる）。"
                            "8 ビットの字の表（dBASE .dbf・Lotus 1-2-3 .wk1・Excel 2〜4 の .xls）で、窓に「値の文字コードを規則では決められません」と"
                            "候補が並んだときも、選んだ文字コードを渡す（ファイルの ASCII でない字の全部を、その 1 つの文字コードで読み直す）"}
for _t in TOOLS:
    if _t["name"] == "audit_document":
        _t["inputSchema"]["properties"]["job"] = {"type": "boolean", "description": "大きさによらず、長い仕事として動かす（呼び出しの時間の上限に当たったとき）"}
        _t["inputSchema"]["properties"]["restart"] = {"type": "boolean", "description": "控えを捨てて、始めから確かめ直す"}
    if _t["name"] == "open_document":
        _t["inputSchema"]["properties"]["encoding"] = _ENCODING
    if _t["name"] in ("read_window", "list_fields", "fill_document", "audit_document", "grade_document"):
        # 開いたあとの道具には短く（選び方は open_document の説明と、サーバーの説明にある）
        _t["inputSchema"]["properties"]["encoding"] = {"type": "string", "description": "文字だけのファイル（CSV・.txt など）、書庫・メールの名前、8 ビットの字の表（.dbf・.wk1 など）で、窓に文字コードの候補が並んだときに選んだ"
                                                                                       "文字コード（候補の encoding の値。例 \"cp932\"）。同じファイルには、どの道具にも同じ値を渡す"}
    if _t["name"] in ("read_table", "explain_calculation", "try_calculation", "extract_processing"):
        _t["inputSchema"]["properties"]["encoding"] = {"type": "string", "description": "8 ビットの字の表（.dbf・.wk1 など）で、窓の候補から選んだ文字コード（ほかの道具に渡したものと同じ値）"}


def _text(obj) -> dict:
    """結果の字：JSON は字下げなしで返す（字下げの空白と改行は、読む側の字の数を増やすだけで、事実を足さない）。"""
    return {"type": "text", "text": obj if isinstance(obj, str) else json.dumps(obj, ensure_ascii=False, separators=(",", ":"), default=str)}


def _said(res: dict, *head: str) -> list[dict]:
    """content：head・say_to_user・残りの JSON（say_to_user は 1 度だけ。structuredContent には、そのまま入っている）。"""
    return [*map(_text, head), _text(res.get("say_to_user", "")), _text({k: v for k, v in res.items() if k != "say_to_user"})]


def _where_shown(g: dict) -> dict:
    """段階の結果（MCP の形だけ。api.grade の中身は変えない）：理由ごとの where_all（すべての所）は、頼まれたときだけ返す。
    where（はじめの所と「ほか N か所」）は、いつも返す。外したことと頼み方は omitted に書く。"""
    cut = 0

    def slim(reasons: list[dict]) -> list[dict]:
        nonlocal cut
        cut += sum(1 for r in reasons if "where_all" in r)
        return [{k: v for k, v in r.items() if k != "where_all"} for r in reasons]

    shown = {**g, "reasons": slim(g.get("reasons") or []), "write": {**g["write"], "reasons": slim(g["write"].get("reasons") or [])},
             "alerts": slim(g.get("alerts") or [])}
    if cut:
        shown["omitted"] = {"where_all": "所の多い理由の where は、はじめの所と「ほか N か所」です。すべての所は、grade_document に where_all: true を渡すと、"
                                         "理由ごとの where_all に返ります（findings・answered を渡したときは、同じものも渡す）"}
    return shown


def _score(check: dict) -> str:
    return "書く段階" if check.get("score") == "write" else "読む段階"


# list_fields の結果が limits.RESULT_CHARS に入らないときに省く事実：どれを見出しとして読むか・書く所かを決める根拠。書くのに要る事実
# （id・label・kind・options・value・context・readonly・rules・boxes など）は、ここに挙げないので、いつも残る。省いた事実は ids・detail で取れる
_FIELD_DETAIL = ("near", "evidence", "why", "facts", "neighbours", "reason_codes", "input_reason_codes", "why_code", "label_method", "label_key")
_FIELD_RANGE = ("pages", "sheet", "rows", "paragraphs", "records", "path_in", "cursor")


def _size(obj) -> int:
    """結果の大きさ（limits.RESULT_CHARS と比べる字の数）：字下げつきの JSON の字数。返す字は字下げなしなので、これより少ない
    （上限に入る量の決まりは、字下げなしで返すようにしても変えない）。"""
    return len(obj) if isinstance(obj, str) else len(json.dumps(obj, ensure_ascii=False, indent=1, default=str))


_MARK = "▶"  # 表計算の窓の行で、その欄の所に付く印（［▶B3］）


def _line_of(head: str, fid: str) -> str | None:
    """行を持つ欄の context（head）から、同じ行の欄 fid の context を作る：▶ のない行はそのまま。▶ のある行は、▶ を
    fid の番地の所（［番地…）に移す。番地の所が 1 つに決まらなければ None。"""
    if _MARK not in head:
        return head
    line = head.replace(_MARK, "")
    key = "［" + fid.rsplit("!", 1)[-1]
    at, found = line.find(key), []
    while at >= 0:
        nxt = line[at + len(key):at + len(key) + 1]
        if not (nxt.isascii() and nxt.isalnum()):  # ［B3 は ［B30 の頭ではない
            found.append(at)
        at = line.find(key, at + 1)
    if len(found) != 1:
        return None
    return line[:found[0] + 1] + _MARK + line[found[0] + 1:]


def _shared_lines(fl: list[dict]):
    """同じ行の欄が続くとき、行（context）を 1 度だけ返す：2 つ目からは context の代わりに context_same_as（その行を context に
    持つ欄の id）。_line_of で元の context とちょうど同じ字に戻せる欄だけを置き換えるので、事実は減らない（api.fields の中身は変えない）。
    欄を 1 つずつ返す（何万もの欄のうち、結果に入る分だけを作る）。はじめの n 個の形は、はじめの n 個の欄だけで決まる。"""
    head_id = head = None
    for f in fl:
        c, fid = f.get("context"), f.get("id")
        if isinstance(c, str) and c and isinstance(fid, str):
            if head is not None and _line_of(head, fid) == c:
                yield {("context_same_as" if k == "context" else k): (head_id if k == "context" else v) for k, v in f.items()}
                continue
            head_id, head = fid, c
        yield f


def _share_lines(fl: list[dict]) -> list[dict]:
    return list(_shared_lines(fl))


def _fields_range_args(args: dict) -> dict:
    rng = {k: args[k] for k in _FIELD_RANGE if args.get(k) not in (None, "")}
    if "path_in" in rng:
        rng["path"] = rng.pop("path_in")
    return rng


def _fits(rest: dict, fl, budget: int) -> bool:
    """{**rest, "fields": fl} を字にした大きさ（_size）が budget 以下か。欄を 1 つずつ数えて足し、越えた所でやめる
    （何万もの欄の結果の全体を、大きさを知るためだけに字にしない）。数え方は _size と同じ字数になる：欄は 1 つ目のあと、「,」と改行の分だけ増える。"""
    if not fl:
        return _size({**rest, "fields": fl}) <= budget
    used = _size({**rest, "fields": []}) + 1
    if used > budget:
        return False
    empty = _size({"fields": []})
    for f in fl:
        used += _size({"fields": [f]}) - empty - 1
        if used > budget:
            return False
    return True


def _take_fields(fl, room: int) -> int:
    """room 字に入る、はじめの欄の数（進むように、少なくとも 1 つ）。"""
    empty = _size({"fields": []})
    used = n = 0
    for f in fl:
        used += _size({"fields": [f]}) - empty
        if used > room and n:
            break
        n += 1
    return n


def _ask_again(args: dict, rng: dict, **more) -> dict:
    """同じ頼み方を、次の呼び出しの引数として（next にそのまま入れる）。"""
    ask = {("path_in" if k == "path" else k): v for k, v in rng.items()}
    ask.update({k: args[k] for k in ("only_empty", "limit", "detail") if args.get(k)})
    ask.update(more)
    return ask


def _fit_fields(res: dict, args: dict, rng: dict) -> dict | None:
    """list_fields の結果を limits.RESULT_CHARS に収める（MCP の結果の形だけ。api.fields の中身は変えない）。そのままで入るなら None。
    順に：見た行の空のマスの一覧を切る → 根拠の事実を省く → 分けて返す。どれも、省いたことと頼み方を同じ結果に書く。"""
    skip = int(args.get("skip") or 0)
    detail = bool(args.get("detail"))
    budget = limits.RESULT_CHARS
    limit = int(args.get("limit") or 0)
    fl = res.get("fields") or []
    rest = {k: v for k, v in res.items() if k != "fields"}
    share = iter if detail else _shared_lines  # detail: true は、どの欄も api.fields のとおりに
    if not skip and not (limit and len(fl) > limit) and (_fits(rest, share(fl), budget) if fl else _size(res) <= budget):
        return None
    told: dict = {}
    blanks = rest.get("blank_cells_in_range")
    blanks_size = _size(blanks) if blanks else 0
    if blanks and blanks_size > budget // 4:
        keep = max(1, (budget // 4) * len(blanks) // blanks_size)
        rest["blank_cells_in_range"] = blanks[:keep]
        told["blank_cells_in_range"] = (f"見た行の空のマスの番地のはじめの {keep} 個だけを返しました（全 {rest.get('blank_cell_count', len(blanks))} 個は、どれも番地で書けます）。"
                                        "rows を狭めて頼むと、その行の番地が返ります")
    if not detail and fl and not _fits(rest, share(fl), budget):
        counts: dict[str, int] = {}
        for f in fl:
            for k in _FIELD_DETAIL:
                if k in f:
                    counts[k] = counts.get(k, 0) + 1
        if counts:
            fl = [{k: v for k, v in f.items() if k not in _FIELD_DETAIL} for f in fl]
            first = next((f["id"] for f in fl if "id" in f), "欄の id")
            told["detail"] = {"count_by_name": counts,
                             "how": f"結果が大きいので、根拠の事実（{'・'.join(counts)}）を省きました。ほかの事実は省いていません。list_fields に同じ path と ids: "
                                    f"{json.dumps([first], ensure_ascii=False)} のように id を渡すと、その欄の事実をすべて返します"
                                    + ("（範囲で頼んだ欄は、同じ範囲も渡す）" if rng else "")
                                    + "。detail: true なら、すべての欄の事実を分けて返します（next で続く）"}
    total = len(fl)
    fl = fl[skip:]
    shown = {"omitted": told, **rest}  # 省いたことを先頭に
    n = _take_fields(share(fl), budget - _size(shown) - 200)  # 200：下の、分けたことの知らせの分
    if limit:
        n = min(n, limit)  # 頼まれた数まで（範囲の読みはページ・行の単位で集めるので、ここで欄の数にそろえる）
    shown["fields"] = list(share(fl[:n]))  # 行は、この結果の中の欄を指す（分けた続きでも、行を持つ欄はいつも同じ結果にある）
    shown["fields_total"], shown["fields_returned"] = total, n
    if skip + n < total:
        shown["next"] = _ask_again(args, rng, skip=skip + n)
        told["more_fields"] = f"{'limit の数まで' if limit and n == limit else '一度に入らないので'}、{total} か所のうち {skip + 1}〜{skip + n} か所目を返しました。結果の next をそのまま渡すと続きが読めます"
    elif rest.get("next"):
        ask = {("path_in" if k == "path" else k): v for k, v in rest["next"].items()}
        shown["next"] = {**ask, **{k: args[k] for k in ("only_empty", "limit", "detail") if args.get(k)}}
    if not told:
        del shown["omitted"]
    return shown


def _fields_by_id(args: dict, rng: dict) -> dict:
    """頼まれた id の欄の事実をすべて（省かない）。一度に入らない分は next の ids で続く。"""
    pw, enc = args.get("password"), args.get("encoding")
    want = list(dict.fromkeys(str(x) for x in args["ids"]))
    found: dict = {}
    if rng or api.lazy_open(args["path"], pw, enc) is not None:
        ask: dict | None = rng
        while ask is not None:
            res = api.fields_range(args["path"], ask, False, args.get("limit"), pw, enc)
            found.update({f.get("id"): f for f in res.get("fields") or [] if f.get("id") in want})
            ask = res.get("next") if len(found) < len(want) else None
    else:
        found = {f.get("id"): f for f in api.fields(args["path"], pw, enc) if f.get("id") in want}
    fl = [found[i] for i in want if i in found]
    missing = [i for i in want if i not in found]
    out: dict = {"not_found": missing, "not_found_how": "list_fields の id のとおりに渡してください。範囲で頼んで知った欄は、同じ範囲（sheet・rows・pages など）も渡します"} if missing else {}
    n = _take_fields(fl, limits.RESULT_CHARS - _size(out) - _size({"next": {"ids": want}}))  # next の ids が入る分
    out["fields"] = fl[:n]
    if n < len(fl):
        out["next"] = {**{("path_in" if k == "path" else k): v for k, v in rng.items()}, "ids": [f["id"] for f in fl[n:]]}
    return out


def _list_fields(args: dict) -> dict:
    pw, enc = args.get("password"), args.get("encoding")
    rng = _fields_range_args(args)
    if args.get("ids"):
        res = _fields_by_id(args, rng)
        return {"content": [_text(res)], "structuredContent": res}
    if rng or args.get("limit") or api.lazy_open(args["path"], pw, enc) is not None:
        res = api.fields_range(args["path"], rng, bool(args.get("only_empty")), args.get("limit"), pw, enc)
        shown = _fit_fields(res, args, rng)
        if shown is None and args.get("detail"):
            shown = {**res, "next": {**res["next"], "detail": True}} if res.get("next") else res  # 続きも省かずに
        elif shown is None:
            shown = {**res, "fields": _share_lines(res["fields"])} if res.get("fields") else res
        return {"content": [_text(shown)], "structuredContent": shown}
    fl = api.fields(args["path"], pw, enc)
    if args.get("only_empty"):
        fl = [f for f in fl if f.get("value") in (None, "", False)]
    shown = _fit_fields({"fields": fl}, args, rng)
    if shown is None:
        shown = {"fields": fl if args.get("detail") else _share_lines(fl)}
    return {"content": [_text(shown)], "structuredContent": shown}


# extract_processing の結果が limits.RESULT_CHARS に入らないときも、規則ごとに残す事実。ほかの事実は ids・section・out_path で取れる
_RULE_BRIEF = ("id", "sheet", "range", "text", "exception_count", "unreadable")


def _brief_rule(d: dict) -> dict:
    out = {k: d[k] for k in _RULE_BRIEF if k in d}
    if "kind" in d and d["kind"] != "formula":  # 配列数式・データテーブルなど
        out["kind"] = d["kind"]
    st = (d.get("verification") or {}).get("status")
    if st and st != "proven":
        out["verification"] = st
    return out


def _take_rules(rules: list[dict], room: int) -> list[dict]:
    """room 字に入る、はじめの規則（進むように、少なくとも 1 つ。1 つで入らない text は切り、切ったことを書く）。"""
    empty = _size({"rules": []})
    used = 0
    out: list[dict] = []
    for d in rules:
        b = _brief_rule(d)
        n = _size({"rules": [b]}) - empty
        if used + n > room:
            if out:
                break
            text = str(b.get("text") or "")
            b["text"] = text[:max(0, len(text) - (n - room) - 100)]
            b["text_cut"] = f"長いので切りました（全 {len(text):,} 字）。ids で頼むと全部を返します"
            n = _size({"rules": [b]}) - empty
        out.append(b)
        used += n
    return out


def _shrink(v, room: int):
    """room 字に入るよう一覧をはじめの要素だけにした値と、切った数え。辞書は中の一覧を切る。入らなければ (None, None)。"""
    if isinstance(v, list) and v:
        keep = max(1, room * len(v) // _size(v))
        while keep > 1 and _size(v[:keep]) > room:
            keep = keep * 3 // 4
        return (v[:keep], {"shown": keep, "total": len(v)}) if _size(v[:keep]) <= room else (None, None)
    if isinstance(v, dict):
        out = dict(v)
        cut: dict = {}
        for k in sorted(out, key=lambda k: -_size(out[k])):
            if _size(out) <= room:
                break
            if isinstance(out[k], list) and out[k]:
                part, c = _shrink(out[k], max(100, room - (_size(out) - _size(out[k]))))
                out[k], cut[k] = ([], {"shown": 0, "total": len(v[k])}) if part is None else (part, c)
        return (out, cut) if _size(out) <= room else (None, None)
    return None, None


def _rule_checks(res: dict, room: int) -> dict | None:
    """確かめられなかった規則（format="python" の結果）。続きの呼び出しでは確かめないので、はじめの結果で全部を伝える。入らなければ、書き出しの隣のファイルに。"""
    bad = {d["id"]: {k: v for k, v in d["verification"].items() if k != "counts"} for d in res.get("rules") or []
           if (d.get("verification") or {}).get("status") not in (None, "proven")}
    if not bad:
        return None
    if _size(bad) <= room:
        return {"not_proven": bad}
    target = Path(res["python"]["out_path"])
    side = target.with_name(target.name + ".rules_verification.json")
    side.write_text(json.dumps(bad, ensure_ascii=False, indent=0), encoding="utf-8")
    return {"not_proven_count": len(bad), "file": str(side), "how": "規則ごとの確かめの結果（status・why）は、すべてこのファイルにあります"}


def _fit_processing(res: dict, args: dict, skip: int, say_len: int) -> dict:
    """extract_processing の全部の結果を limits.RESULT_CHARS に収める（MCP の結果の形だけ。api の結果は変えない）。
    順に：節を省く → 規則を 1 行にする → 分けて返す。どれも、省いたことと頼み方を同じ結果に書く。"""
    from .processing import HEAD_KEYS

    budget = limits.RESULT_CHARS - say_len
    rules = res.get("rules") or []
    ask_sheet = {"sheet": args["sheet"]} if args.get("sheet") else {}
    told: dict = {}
    head: dict = {}
    if not skip:
        head = {k: res[k] for k in (*HEAD_KEYS, "xfa", "note", "out_path") if k in res}
        for k in sorted(head, key=lambda k: -_size(head[k])):  # 土台の値でも、大きすぎるものは切るか節に回す（規則の分を空ける）
            if _size(head) <= budget // 3 or k == "say_to_user":
                continue
            small, cut = _shrink(head[k], budget // 6)
            if small is None:
                del head[k]
            else:
                head[k] = small
                told.setdefault("cut", {})[k] = {"lists": cut, "how": f'section: "{k}" で全部を返します'}
        sections = {k: {"chars": _size(v), **({"count": len(v)} if isinstance(v, (list, dict)) else {})}
                    for k, v in res.items() if k != "rules" and k not in head and v not in (None, [], {}, "", False)}
        if sections:
            told["sections"] = {"names": sections,
                                "how": '結果が大きいので、この節は返していません。同じ path と section: "節の名前" を渡すと、省かずに返します（入らなければ next で続く）'}
        told["all"] = (f"全部は JSON に書き出してあります：{res['out_path']}" if res.get("out_path") else
                       "全部を 1 つのファイルにするには、format を省いて（json）out_path に .json を渡します。sheet を渡すと、そのシートの規則・入力だけになります")
        checks = _rule_checks(res, budget // 8)
        if checks:
            told["rule_verification"] = checks
    counts: dict[str, int] = {}
    for d in rules:
        for k in d:
            if k not in _RULE_BRIEF and k != "kind":
                counts[k] = counts.get(k, 0) + 1
    if counts:
        first = rules[skip if skip < len(rules) else 0].get("id", "R1")
        told["rule_detail"] = {"count_by_name": counts,
                               "how": f"規則は 1 行の text と id・シート・範囲だけを返しました（{'・'.join(counts)} を省いています）。同じ path と ids: "
                                      f"{json.dumps([first], ensure_ascii=False)} のように規則の id を渡すと、その規則の事実をすべて返します。"
                                      'すべての規則の詳細は section: "rules"'}
    shown = {"omitted": told, **head, **({"only_sheet": args["sheet"]} if ask_sheet else {})}
    room = budget - _size(shown) - _size({"next": {"skip": len(rules), **ask_sheet}, "rules_total": len(rules), "rules_returned": len(rules)}) - 300
    piece = _take_rules(rules[skip:], max(room, 200))  # 300：下の、分けたことの知らせの分
    shown["rules"] = piece
    shown["rules_total"], shown["rules_returned"] = len(rules), len(piece)
    if skip + len(piece) < len(rules):
        shown["next"] = {"skip": skip + len(piece), **ask_sheet}
        told["more_rules"] = (f"一度に入らないので、{len(rules)} 個のうち {skip + 1}〜{skip + len(piece)} 個目の規則を返しました。"
                              "結果の next をそのまま渡すと続きが読めます（書き出しも確かめも繰り返しません）")
    return shown


def _slice_json(obj: dict, label: str, ask: dict, skip: int) -> dict:
    """省かない値を、JSON の字の位置で分けて返す（つなげると元の JSON。1 つの値がどんなに大きくても入る）。"""
    text = _text(obj)["text"]
    budget = limits.RESULT_CHARS
    if not skip and len(text) <= budget:
        return {"content": [_text(text)], "structuredContent": obj}
    if skip > len(text):
        raise api.ApiError(f"skip は {len(text):,} までです（前の結果の next をそのまま渡してください）")
    room = max(budget - _size(ask) - 400, 200)  # 400：下の、読んだ範囲と続きの知らせの分
    end = min(skip + room, len(text))
    over = len(json.dumps(text[skip:end], ensure_ascii=False)) - 2 - room  # structuredContent では、字の中の " と \ に \ が付く
    if over > 0:
        end = max(end - over, skip + 1)
    if end < len(text):
        nl = text.rfind("\n", skip, end)
        if nl > skip:
            end = nl + 1
    cover = f"{label}の JSON の {skip + 1:,}〜{end:,} 字目（全 {len(text):,} 字。つなげると元の JSON）"
    nxt = {**ask, "skip": end} if end < len(text) else None
    tail = f"続きは next：{json.dumps(nxt, ensure_ascii=False)}" if nxt else "これで終わりです"
    # 字も structuredContent に：structuredContent があると content を見せないクライアントで、covered・next だけになる
    return {"content": [_text(f"{cover}。{tail}"), _text(text[skip:end])],
            "structuredContent": {"text": text[skip:end], "covered": cover, "remaining": f"{len(text) - end:,} 字", "next": nxt, "done": nxt is None}}


def _extract_processing(args: dict) -> dict:
    ids, section, skip = args.get("ids"), args.get("section"), int(args.get("skip") or 0)
    if (ids or section or skip) and (args.get("format") == "python" or args.get("out_path")):
        raise api.ApiError("ids・section・skip は、取り出した結果の続きや詳細を読むためのものです。format と out_path を省いて頼んでください（書き出しと確かめは繰り返しません）")
    if ids and section:
        raise api.ApiError("ids か section のどちらかで頼んでください")
    res, old = api.extract_processing_both(args["path"], args.get("format") or "json", args.get("out_path"), args.get("sheet"),
                                           bool(args.get("overwrite")), args.get("password"), args.get("encoding"))
    ask_sheet = {"sheet": args["sheet"]} if args.get("sheet") else {}
    if ids:
        want = list(dict.fromkeys(str(x) for x in ids))
        ask = {"ids": want, **ask_sheet}
        if _size(ask) > limits.RESULT_CHARS // 2:
            raise api.ApiError('一度に頼む ids が多すぎます。すべての規則の詳細は section: "rules" で頼んでください')
        by = {d.get("id"): d for d in res.get("rules") or []}
        out: dict = {"rules": [by[i] for i in want if i in by]}
        missing = [i for i in want if i not in by]
        if missing:
            out["not_found"] = missing
            out["not_found_how"] = "規則の id（rules の id のとおり）を渡してください。sheet を渡したときは、そのシートの規則だけを探します"
        return _slice_json(out, "規則", ask, skip)
    if section:
        if section not in res:
            raise api.ApiError(f"節「{section}」はありません（ある節：{'・'.join(res)}）")
        return _slice_json({section: res[section]}, f"節「{section}」", {"section": section, **ask_sheet}, skip)
    say = old.get("say_to_user", "")
    if not skip and len(say) + _size(old) <= limits.RESULT_CHARS:
        return {"content": _said(old), "structuredContent": old}
    shown = _fit_processing(res, args, skip, 0 if skip else len(say))
    return {"content": [_text(shown)] if skip else [_text(say), _text(shown)], "structuredContent": shown}


_CUT = re.compile(r"cut:(\d+):(.*)", re.S)  # 1 回に入らなかった所の続き：cut:読んだ字の数:その所を頼んだときの cursor


def _cut_text(head: str, text: str, tail: str, at: int, ask: dict, then: dict | None, prior: str = "") -> dict | None:
    """窓の 1 回の結果を limits.RESULT_CHARS に収める：字の位置で分ける（部分の番号・範囲の決まりは変えない。つなげると元の字のとおり）。
    そのままで入るなら None。ask：同じ所を頼む引数。then：この所を読み終えたあとの next。tail：読み終えたときの知らせ。prior：この所を頼んだときの cursor。"""
    budget = limits.RESULT_CHARS
    if not at and len(head) + len(text) + len(tail) <= budget:
        return None
    if at > len(text):
        raise api.ApiError("cursor がこの所の字の数を超えています（前の結果の next をそのまま渡してください）")
    # 500：下の、読んだ範囲と続きの知らせの分
    end = min(at + max(budget - len(head) - len(tail) - _size(ask) - len(prior) - 500, 200), len(text))
    if end < len(text):
        nl = text.rfind("\n", at, end)
        if nl > at:
            end = nl + 1
    more = end < len(text)
    nxt = {**ask, "cursor": f"cut:{end}:{prior}"} if more else then
    cover = f"この所の {at + 1:,}〜{end:,} 字目（全 {len(text):,} 字）を読みました。"
    foot = (f"（{cover}続きは read_window に {json.dumps(nxt, ensure_ascii=False)} を渡してください）" if more
            else f"（{cover}この所はここまでです）" + tail)
    return {"piece": text[at:end], "foot": foot, "next": nxt, "remaining": f"{len(text) - end:,} 字" if more else ""}


def _read_window(args: dict) -> dict:
    from .lazy import clamp_chars

    at, prior = 0, ""
    m = _CUT.fullmatch(str(args.get("cursor") or ""))
    if m:  # 同じ所を、もとの cursor で読み直して、続きの字から返す
        at, prior = int(m[1]), m[2]
        args = {**args, "cursor": prior}
    rng = {k: args[k] for k in ("pages", "sheet", "rows", "paragraphs", "records", "path_in", "cursor", "raw") if args.get(k) not in (None, "", False)}
    ask = {k: v for k, v in rng.items() if k != "cursor"}  # 分けた続きを頼む引数（read_window の引数の名前のまま）
    ask.update({k: args[k] for k in ("part", "max_chars") if args.get(k) is not None})
    if "path_in" in rng:
        rng["path"] = rng.pop("path_in")
    large = api.lazy_open(args["path"], args.get("password"), args.get("encoding")) if not rng else None
    if rng or large is not None or args.get("max_chars"):
        if args.get("part") is not None and not rng:
            rng["part"] = args["part"]
        # 範囲の 1 回の大きさは結果の上限まで（API の既定 limits.PAGE_CHARS は変えない。続きの cursor は字の位置や行で決まり、大きさによらない）
        most = max(limits.RESULT_CHARS - 1000, 200)
        page, doc = api.read_range(args["path"], rng, min(clamp_chars(args.get("max_chars") or most), most), args.get("password"), args.get("encoding"))
        cut = _cut_text("", page.text, page.footer(), at, ask, page.next, prior)
        # 本文も structuredContent に：structuredContent があると content を見せないクライアントで、covered・next だけになる
        if cut is None:
            return {"content": [_text(page.text), _text(page.footer())], "structuredContent": {"text": page.text, **page.to_dict()}}
        sc = {"text": cut["piece"], **page.to_dict(), "next": cut["next"], "done": cut["next"] is None}
        if cut["remaining"]:
            sc["remaining"] = f"この所の残り {cut['remaining']}" + (f"と {page.remaining}" if page.remaining else "")
        return {"content": [_text(cut["piece"]), _text(cut["foot"])], "structuredContent": sc}
    _k, _doc, parts = api.window(args["path"], args.get("password"), args.get("encoding"))
    n = int(args.get("part", 1))
    if not 1 <= n <= len(parts):
        raise api.ApiError(f"part は 1〜{len(parts)} です")
    head = f"部分 {n}/{len(parts)}：{parts[n - 1][0]}"
    then = {"part": n + 1} if n < len(parts) else None
    cut = _cut_text(head, parts[n - 1][1], f"次は read_window に part: {n + 1} を渡してください" if then else "", at, {"part": n}, then)
    if cut is None:
        return {"content": [_text(head), _text(parts[n - 1][1])]}
    return {"content": [_text(head), _text(cut["piece"]), _text(cut["foot"])],
            "structuredContent": {"text": cut["piece"], "part": n, "parts": len(parts), "remaining": cut["remaining"], "next": cut["next"],
                                  "done": cut["next"] is None}}


def call_tool(name: str, args: dict) -> dict:
    if name == "open_document":
        path = args["path"]
        pw = args.get("password")
        enc = args.get("encoding")
        s = api.summary(path, pw, enc)
        fm = s.get("format")
        if isinstance(fm, dict) and "encoding_notes" in fm:
            # 文字コードの注意の字は、下の窓（概要）の字にそのままある。同じ事実は、format の encoding_basis・encoding_candidates・
            # encoding_chosen_by・control_chars にある
            s = {**s, "format": {k: v for k, v in fm.items() if k != "encoding_notes"}}
        if s.get("large"):
            lz = api.lazy_open(path, pw, enc)
            head = (f"{s['kind_ja']}・大きな書類です（窓の全体は見積もりで {s['estimated_window_chars']:,} 字）。全体を一度には返しません。"
                    "下の概要を見て、要る所を read_window に範囲で頼んでください（読む量を分けるだけで、読めない所はありません）")
            return {"content": [_text(head), _text(lz.overview_text())], "structuredContent": s}
        _k, _doc, parts = api.window(path, pw, enc)
        if s.get("needs_tool"):
            return {"content": [_text(f"{s['kind_ja']}。中身は読んでいません。下の手順で変換・展開してから、できたファイルを開いてください"), _text(parts[0][1])], "structuredContent": s}
        head = (f"{s['kind_ja']}・入力欄 {s['fields']} か所（書ける欄 {s['writable_fields']}）・窓は {s['window_parts']} 部分"
                + ("。下は部分 1 です（続きは read_window に part: 2 から）" if s["window_parts"] > 1 else "。下が窓の全体です"))
        cut = _cut_text(head, parts[0][1], "", 0, {"part": 1}, {"part": 2} if len(parts) > 1 else None)  # 最初の部分が結果の上限に入らないとき
        if cut is not None:
            return {"content": [_text(head), _text(cut["piece"]), _text(cut["foot"])], "structuredContent": {**s, "next": cut["next"]}}
        return {"content": [_text(head), _text(parts[0][1])], "structuredContent": s}
    if name == "read_window":
        return _read_window(args)
    if name == "list_fields":
        return _list_fields(args)
    if name == "fill_document":
        if args.get("help"):
            return {"content": [_text(fill_help(str(args["help"])))]}
        res = api.fill(args["path"], args.get("values") or {}, args["out_path"], bool(args.get("force", False)),
                       password=args.get("password"), encoding=args.get("encoding"), overwrite=bool(args.get("overwrite")))
        ok = res.get("proof_ok") and res.get("read_back_ok")
        head = ("書き込みました。" if ok else "書き込みましたが、確認が必要な点があります。"
                + (f"確かめられなかったので「{res['out_path']}」に書き出しました（提出しないでください）。" if res.get("unverified") else "")
                ) + f"{res.get('proof', '')}"
        if res.get("check_visually"):
            cv = res["check_visually"]
            head += ("　書く[B] 見た目を確かめる欄：" + "・".join(f"{x['id']}（{x['page']}ページ）" if x.get("page") else x["id"]
                                                     for x in cv["fields"]) + "。" + cv["why"])
        if res.get("written_without_as"):
            head += f"　as（書いた値の意味）のない所 {len(res['written_without_as'])} か所：written_without_as。本人が確かめる書類なら as を添えて伝える。"
        if res.get("replaced_existing"):
            head += f"　もとからあった字や値を変えた所 {len(res['replaced_existing'])} か所：replaced_existing。本人が頼んだ書き換えかを、本人に確かめる。"
        # 本人は書類を開いて確かめないので、say_to_user（書いた所の見出しと値・まだ空いている欄・本人に伝えること）を、対話で伝える
        return {"content": _said(res, head), "structuredContent": res, "isError": not ok}
    if name == "run_macro":
        res = api.run_macro(args["path"], args["macro"], args["out_path"], args.get("answers") or [], args.get("files") or {},
                            password=args.get("password"), overwrite=bool(args.get("overwrite")))
        if not res.get("done"):
            head = res.get("note", "")
            if res.get("prompt"):
                head = f"マクロが答えを求めています：「{res['prompt']}」。{head}"
            return {"content": [_text(head), _text(res)], "structuredContent": res, "isError": False}
        head = (f"マクロ {res['macro']} を動かしました。" + (f"{res.get('proof', '')}" if res.get("out_path") else res.get("note", ""))
                + (f"　エラー：{res['error']}" if res.get("error") else ""))
        return {"content": [_text(head), _text(res)], "structuredContent": res, "isError": bool(res.get("error"))}
    if name == "read_table":
        # rows で範囲を頼んだら、limit を省いても、その範囲を切り詰めない。分けて読む道は limit の所で止めて next を返す（読み飛ばす行を作らない）
        limit = args.get("limit", None if args.get("rows") else READ_TABLE_LIMIT)
        res = api.read_table(args["path"], args["header"], args.get("sheet"), args.get("header_rows"), args.get("out_path"),
                             args.get("password"), args.get("cursor"), args.get("rows"), None if args.get("out_path") else limit,
                             args.get("encoding"))
        res.pop("all_rows", None) if args.get("out_path") else None
        if not args.get("out_path") and limit is not None and res["rows"] > limit:
            res["all_rows"] = res["all_rows"][:limit]  # はじめの limit 行（行の数 rows と、返した行の数 rows_returned）
            res["rows_returned"] = limit
            res["note"] = (f"行が {res['rows']} 行あります。はじめの {limit} 行だけを返しました。すべての行は out_path（.csv か .json）"
                           "に書き出すか、limit を大きくして読んでください")
        return {"content": [_text(res)], "structuredContent": res}
    if name == "explain_calculation":
        res = api.calc_model(args["path"], args.get("cell"), args.get("password"), args.get("encoding"))
        return {"content": [_text(res)], "structuredContent": res}
    if name == "try_calculation":
        res = api.calculate(args["path"], args.get("values") or {}, args.get("password"), args.get("cells"), args.get("encoding"))
        return {"content": _said(res), "structuredContent": res}
    if name == "extract_processing":
        return _extract_processing(args)
    if name == "view_page":
        v = api.view_page(args["path"], args["page"], args.get("region"), args.get("marks"), args.get("password"), args.get("dpi"))
        return {"content": [_text(v["text"]), {"type": "image", "data": v["png_base64"], "mimeType": "image/png"}],
                "structuredContent": {"text": v["text"], "width": v["width"], "height": v["height"], "not_drawn": v["not_drawn"]}}
    if name == "view_image":
        v = api.view_image(args["path"], args["image"], args.get("password"), args.get("region"))
        return {"content": [_text(v["text"]), {"type": "image", "data": v["data_base64"], "mimeType": v["mime"]}],
                "structuredContent": {k: v[k] for k in ("id", "mime", "width", "height", "not_drawn", "text")}}
    if name == "to_pdf":
        res = api.to_pdf(args["path"], args["out_path"], bool(args.get("overwrite")))
        return {"content": [_text(res["say_to_user"])], "structuredContent": res}
    if name == "grade_document":
        g = api.grade(args["path"], args.get("findings") or [], args.get("answered") or [], password=args.get("password"),
                      encoding=args.get("encoding"))
        mark = lambda r: ("（エージェントの判断）" if r.get("by") == "agent" else  # noqa: E731
                          "（エージェントの選択を規則で確かめた）" if r.get("by") == "agent_choice" else "")
        lines = [g["headline"]] + [f"  - 読む[{r['grade']}]{mark(r)} {r['title']}: {r['detail']}" for r in g["reasons"]]
        lines += [f"  - 書く[{r['grade']}]{mark(r)} {r['title']}: {r['detail']}" for r in g["write"]["reasons"]]
        lines += [f"  - ？（あれば{_score(c)} {c['grade']}）{c['question']}" for c in g.get("agent_checks", [])]
        shown = g if args.get("where_all") else _where_shown(g)
        skip = int(args.get("skip") or 0)
        if skip or _size(shown) > limits.RESULT_CHARS:  # すべての所を頼んで、一度に入らないとき：分けて返す
            ask = {k: args[k] for k in ("where_all", "findings", "answered") if args.get(k)}
            return _slice_json(shown, "段階", ask, skip)
        return {"content": [_text("\n".join(lines)), _text(shown)], "structuredContent": shown}
    if name == "audit_document":
        au = api.audit(args["path"], args.get("password"), args.get("encoding"), restart=bool(args.get("restart")), job=bool(args.get("job")))
        lines = [("[OK] " if c["ok"] else "[要確認] ") + f"{c['name']}: {c['detail']}" for c in au["checks"]]
        job = au.get("job")
        if job and job.get("failed"):
            lines.insert(0, f"検査は最後まで進みませんでした：{job.get('stopped')}。{au.get('note', '')}")
        elif job and not job["finished"]:
            pr = job.get("progress") or {}
            at = f"{pr.get('text', '始めたところ')}" + (f"（{pr['done']:,}／{pr['total']:,}）" if pr.get("total") else "")
            lines.insert(0, (f"検査は止まりました：{job['stopped']}" if job.get("stopped") else f"検査は別のプロセスで動いています：{at}。") +
                         "もう一度 audit_document を呼ぶと、進み具合と結果が分かります（止まっていれば、控えの続きから進みます）")
        else:
            lines.append("結論: " + (("確かめた所は正確です（未確認の所があります）" if au.get("deferred") else "全て正確に読み書きできます") if au["ok"] else "読み書きに確認が必要な点があります"))
        wp = au.get("written_proof")
        if wp and wp.get("note"):
            lines.insert(0, wp["note"])
        g = au.get("grade")
        if g:
            au = {**au, "grade": _where_shown(g)}
            lines.append(g["headline"])
            lines += [f"  - 読む[{r['grade']}] {r['title']}: {r['detail']}" for r in g["reasons"] if r["grade"] not in ("S", "A")]
            lines += [f"  - 書く[{r['grade']}] {r['title']}: {r['detail']}" for r in g["write"]["reasons"] if r["grade"] != "A"]
            lines += [f"  - ？（あれば{_score(c)} {c['grade']}）{c['question']}" for c in g.get("agent_checks", [])]
            if g.get("agent_checks"):
                lines.append("  窓を読んで確かめたら、grade_document に findings（見つけたこと）と answered（確かめた問いの id）を渡してください。")
        return {"content": [_text("\n".join(lines)), _text(au)], "structuredContent": au}
    raise api.ApiError(f"そのような道具はありません: {name}")


_PATH = re.compile(r"(?:[A-Za-z]:\\|/)[^\s'\"（）()「」、。,:;]*[/\\][^\s'\"（）()「」、。,:;]+")


def _clean(msg: str, args: dict, limit: int = 1500) -> str:
    """結果に出す字：エージェントが渡したパスのほかの絶対パス（とどくの置き場所・一時フォルダなど）は、ファイルの名前だけにする。"""
    given = [str(v) for v in args.values() if isinstance(v, str) and v]

    def keep(m):
        t = m.group(0)
        for g in given:
            gp = os.path.expanduser(g)
            if t in g or t in gp or t.startswith(os.path.dirname(gp) + os.sep):
                return t  # エージェントが渡したパス（と、その隣のファイル：書き出したファイルなど）
        return "…/" + os.path.basename(t.replace("\\", "/").rstrip("/"))

    out = _PATH.sub(keep, str(msg))
    return out if len(out) <= limit else out[:limit] + "…"


def tool_result(name: str, args: dict) -> dict:
    """1 回の道具の呼び出しの結果（MCP の result）。例外は、短い言葉と例外の種類にする（トレースバック・絶対パスは出さない。
    トレースバックは標準エラーの記録にだけ書く）。"""
    from .lazy import smaller_request
    from .limits import STOPPED, LimitExceeded, call_budget

    try:
        with call_budget():
            return call_tool(name, args)
    except (api.PasswordNeeded, api.PasswordWrong) as e:
        return {"content": [_text(str(e))], "isError": False,
                "structuredContent": {"needs_password": True, "wrong_password": isinstance(e, api.PasswordWrong)}}
    except LimitExceeded as e:
        hint = smaller_request(name, args)
        return {"content": [_text(_clean(str(e), args) + "\n" + hint)], "isError": True, "structuredContent": {"stopped": True, "smaller_request": hint}}
    except RecursionError:
        hint = smaller_request(name, args)
        return {"content": [_text(f"{STOPPED}：入れ子や参照が深すぎて、読み終われません（RecursionError）\n{hint}")], "isError": True,
                "structuredContent": {"stopped": True, "smaller_request": hint}}
    except MemoryError:
        hint = smaller_request(name, args)
        return {"content": [_text(f"{STOPPED}：メモリが足りなくなりました（MemoryError）\n{hint}")], "isError": True,
                "structuredContent": {"stopped": True, "smaller_request": hint}}
    except (api.ApiError, KeyError, ValueError) as e:
        kind = "" if isinstance(e, api.ApiError) else f"（{type(e).__name__}）"
        return {"content": [_text(f"できませんでした{kind}：{_clean(str(e), args)}")], "isError": True}
    except Exception as e:  # noqa: BLE001
        traceback.print_exc(file=sys.stderr)  # 記録（標準エラー）にだけ。結果には短く
        return {"content": [_text(f"処理の途中で問題が起きました（{type(e).__name__}）：{_clean(str(e), args, 300)}")],
                "isError": True}


def _instructions() -> str:
    """サーバーの説明：この環境で上限が効かないこと（Windows・ps がない）は、エージェントにも知らせる。"""
    from . import isolate

    isolate.enabled()
    n = isolate.notice()
    return INSTRUCTIONS + (f"（この環境の注意：{n}）" if n else "")


def handle(msg: dict) -> dict | None:
    mid = msg.get("id")
    method = msg.get("method")
    params = msg.get("params") or {}
    if method == "initialize":
        ver = params.get("protocolVersion") or PROTOCOL
        return {"jsonrpc": "2.0", "id": mid, "result": {
            "protocolVersion": ver, "capabilities": {"tools": {"listChanged": False}},
            "serverInfo": {"name": "todoku", "title": "とどく", "version": "0.2.0"}, "instructions": _instructions()}}
    if method in ("notifications/initialized", "notifications/cancelled", "initialized"):
        return None
    if method == "ping":
        return {"jsonrpc": "2.0", "id": mid, "result": {}}
    if method == "tools/list":
        return {"jsonrpc": "2.0", "id": mid, "result": {"tools": TOOLS}}
    if method == "tools/call":
        name = params.get("name", "")
        args = params.get("arguments") or {}
        from . import isolate

        # 別のプロセス（働き手）で、時間とメモリの上限をつけて動かす（壊れた書類でサーバーごと止まらないように）
        result = isolate.window_job_view(name, args)  # 窓を長い仕事で作っている間は、進み具合をすぐ返す
        if result is None:
            result = isolate.worker().call(name, args) if isolate.enabled() else isolate.soft_call(tool_result, name, args)
        return {"jsonrpc": "2.0", "id": mid, "result": result}
    if mid is None:
        return None
    return {"jsonrpc": "2.0", "id": mid, "error": {"code": -32601, "message": f"知らない命令です: {method}"}}


def _utf8_stdio() -> None:
    """MCP の stdio は UTF-8 の JSON の行（改行は \\n）。Windows の標準出力は ANSI のコードページで、改行を \\r\\n に変えるので、
    UTF-8・\\n に設定し直す（設定し直せない流れ（テストの置き換えなど）は、そのまま）。"""
    for stream, kw in ((sys.stdout, {"encoding": "utf-8", "newline": "\n"}), (sys.stdin, {"encoding": "utf-8"})):
        reconf = getattr(stream, "reconfigure", None)
        if reconf is not None:
            try:
                reconf(**kw)
            except (ValueError, OSError):
                pass


def main() -> int:
    from . import isolate

    if isolate.enabled():
        isolate.probe_ps()  # 外からのメモリの見張り（ps）が使えるかを、立ち上げの時に確かめる（説明 instructions に書くため）
        isolate.worker().start()  # 働き手を先に立ち上げておく（はじめの呼び出しを待たせない）
    _utf8_stdio()
    try:
        from . import jobs

        jobs.purge()  # 古い控え（書類の中身が入る）を、立ち上げのときにも消す
    except OSError:
        pass
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
