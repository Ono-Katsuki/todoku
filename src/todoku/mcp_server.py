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
import os
import re
import sys
import traceback

from . import api

PROTOCOL = "2025-06-18"
INSTRUCTIONS = (
    "とどく（todoku）は、申請書類（Excel .xlsx/.xlsm/.xls・Word .docx/.docm/.doc・PDF・一太郎 .jtd・OpenDocument .odt/.ods・"
    "RTF・PowerPoint .pptx・文字だけのファイル .txt/.csv/.tsv/.json・それらをまとめた ZIP）を、内容を一字も変えずに読める「窓」にし、"
    "入力欄に証明つきで書き込む道具です。韓国のワープロ（한글）の HWPX（.hwpx）も読んで書け、HWP 5（.hwp）は読むだけです"
    "（HWP 5 に書くときは、本人に HWPX で保存し直したファイルを頼みます）。"
    "大きな書類は、読む量を分けて読みます（読めない所はありません）：open_document が large の概要を返したら、要る所を read_window に範囲で頼みます"
    "（PDF は pages、Excel は sheet と rows、CSV・TSV は rows、JSON Lines は records、.txt・Word は paragraphs。結果の next をそのまま渡すと続きが読め、"
    "たどり終えると全部を読んだことになります）。list_fields も同じ範囲で頼みます。段階の deferred（未確認）は、まだ確かめていないことなので、確かめたことのように伝えないでください。"
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
    "（パスワードは結果にも記録にも残しません。書き出したファイルも、元と同じパスワードで保護されたままです）。"
    "文字だけのファイル（CSV・.txt など）で、窓に「文字コードを規則では決められません」と候補が並んだときは、字を読んで意味の通る方を"
    "選び、その名前を、そのファイルのどの道具にも encoding で渡してください。"
    "窓の「画像」の節で「説明（代わりの字）がありません」とある画像は、中身を字として読めません。申請に関わりそうなら view_image（PDF の"
    "ページ全体は view_page）で画像を見て読み、読めなければ、その位置（ページ・上下左右・近くの字）を本人に伝えてください。中身を推測で"
    "埋めないでください。画像そのもののファイル（JPEG・PNG・TIFF・HEIC・SVG など）も open_document で開き、view_page で見ます。段階の alerts（アクセシビリティの注意）は、"
    "段階とは別に伝えてください。"
    "fill_document の結果の read_back_ok が false のときは、ファイルはもう書き出されていますが、書いた後のファイルを読み直した値が"
    "書いた値と違う欄があります（read_back_differences）。確かめずに本人に渡さず、違いを読んで直すか、そのことを本人に伝えてください。"
    "結果に check_visually（書く B）があるときは、その欄を view_page でページの画像を見て、字の形が正しいかを確かめてください"
    "（字を組み立てて描いた文字は、字の並びは確かめていますが、見た目そのものは確かめていません）。"
)

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
                    "（配布用・パスワードの HWP は読めない）。"
                    "ZIP は中のファイルの一覧を返す。中のファイルは、どの道具でも path に「a.zip!/中の名前」と指定して開く・書く"
                    "（書き出す out_path は ZIP。書いたファイルの部品だけを入れ替える）。",
     "inputSchema": {"type": "object", "properties": {"path": {"type": "string", "description": "書類のファイルのパス"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"}},
                     "required": ["path"]}},
    {"name": "read_window",
     "description": "窓（Markdown）を返す。ふつうの書類は part 番目の部分（1から）。大きな書類は範囲で頼む：PDF は pages（\"1-20\"）、"
                    "Excel は sheet と rows（\"1-500\"）、CSV・TSV は rows、JSON Lines は records、.txt・Word は paragraphs（行・段落の範囲）。"
                    "どの書類でも範囲で読める。結果には、読んだ範囲・残り・続きの頼み方（next）が付く。next をそのまま渡していけば、"
                    "全部を読める（読む量を分けるだけで、読めない所はない）。max_chars で 1 回の大きさ（字）を変えられる。",
     "inputSchema": {"type": "object", "properties": {"path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"},
                                                      "part": {"type": "integer", "minimum": 1, "default": 1},
                                                      "pages": _RANGE_SCHEMA, "rows": _RANGE_SCHEMA, "paragraphs": _RANGE_SCHEMA, "records": _RANGE_SCHEMA,
                                                      "sheet": {"type": "string", "description": "Excel のシートの名前（rows と一緒に）"},
                                                      "cursor": {"type": "string", "description": "前の結果の next にある、続きの印"},
                                                      "max_chars": {"type": "integer", "minimum": 200, "description": "1 回に返す大きさ（字）"},
                                                      "raw": {"type": "boolean", "description": "窓を作らずに、ファイルの字を順に読む（欄は見つけない）。窓を作るのが上限で止まった書類のための道"}},
                     "required": ["path"]}},
    {"name": "list_fields",
     "description": "入力欄の一覧を JSON で返す：id（書き込むときの名前）・label（見出し）・kind（種類）・options（選択肢）・value（いまの値）・readonly（書けない欄）・context（その行）。"
                    "大きな書類は範囲で頼む（read_window と同じ pages・sheet・rows・paragraphs・records。空のマスは only_empty）。結果の next で続きを頼める。",
     "inputSchema": {"type": "object", "properties": {"path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"},
                                                      "only_empty": {"type": "boolean", "default": False},
                                                      "pages": _RANGE_SCHEMA, "rows": _RANGE_SCHEMA, "paragraphs": _RANGE_SCHEMA, "records": _RANGE_SCHEMA,
                                                      "sheet": {"type": "string"}, "cursor": {"type": "string"}, "limit": {"type": "integer", "minimum": 1}},
                     "required": ["path"]}},
    {"name": "fill_document",
     "description": "入力欄に書き込み、out_path に書き出す（元のファイルは変えない）。values は {欄の id: 値}。"
                    "チェックは true/false（JSON の真偽値）か、その欄が宣言した状態の値（PDF の状態の名前・XFA の値・箱の字）だけ："
                    "「はい」「あり」「有」などの言葉は受けないので、本人の答えを true/false に直して書く。"
                    "窓で［P12.c5？］と「？」のついた所は、規則では欄と決められない所：窓を読んで書く所と判断したときだけ、同じように id で書く。"
                    "Word などの段落の中の所の id「P12.c5」は、12 番目の段落の、前に印刷された字が 5 字ある所（位置で決まる。字をはさまず続く所は .c5.2）。"
                    "古い数え順の id（P12.1・P12.a1）は、その段落にその並びの所が 1 つだけのときだけ受ける（結果の renamed_ids）。"
                    "Excel のセルの中の空欄は、セルの id に空欄ごとの値のリストか、所の id「シート!B3@5」（spans の id）ごとに書ける。"
                    "空白だけの空きの所は、元の空白を印の中に入れて示す（［P1.c2？　］は全角の空白 1 つの空き）。"
                    "Word で、窓に欄として示していない所に書くときは、{\"@1\": {\"old\": 段落の中の元の字, \"new\": 書き換えた後の字}} の形で、"
                    "字を目印にして書ける（old が文書に1つしかないときはそのまま。いくつもあるときは near に近くの欄の id か、"
                    "occurrence に何番目か を足す。near でしぼる範囲は、near の欄が表のマスなら同じ行、ほかは同じ話（本文・ヘッダーなど）。"
                    "near_paras（前後の段落の数）で狭められる）。old と new の違う所だけが変わり、書いた所は id「P段落.w番号」になる。"
                    "差し込み（MERGEFIELD）・日付などのフィールドの結果の字は、Word が更新すると元に戻るので、目印にしても断る。"
                    "Word の箇条書きの印の □（窓の ☐［P段落.b］）は、チェックの欄として true で書く（その段落の印だけが ■・☑ になる）。"
                    "PDF で、窓に欄として示していない所に書くときは、{\"@1\": {\"after\": 見出しの字, \"value\": 書く字, \"page\": ページ}} の形で、"
                    "見出しの右（\"where\": \"below\" なら下）の空いている所に書ける（書いた所は id「pページ.w番号」）。"
                    "PDF の空き・枠（box）・見出しの右（after）に書く字の置き方はエージェントが決める："
                    "{欄の id: {\"value\": 値, \"align\": \"left\"|\"center\"|\"right\", \"pad_mm\": 余白}}"
                    "（@ の書き込みは同じ辞書に align・pad_mm を足す）。既定は左（枠の左の端から 1.5pt ほどの決まった余白）で、"
                    "字の意味（「年」の前など）からは決めない。結果の written_rects に、書いた枠（page・x・y・w・h、ページの左上からの mm）"
                    "・字の大きさ・そろえ方と、view_page の marks に渡せる box_mm が出るので、見て、違えば align か box で書き直す。○は circled の box_mm。"
                    "Word の段落の中の空欄・Excel のセルの字の中の空欄（所の id）も {\"value\": 値, \"align\": \"left\"|\"center\"|\"right\"} で書ける："
                    "空きの幅を保って、前後に半角の空白を足す（字の数で数える。曖昧な幅の字の数え方は \"ambiguous_width\": \"wide\"|\"narrow\"）。"
                    "既定は左（指定がなければ空きを値に置き換え、下線の空欄だけ後ろに空白を足して下線の長さを保つ）。足した空白の字数は結果の aligned に出る。"
                    "字の前・後ろへの差し込み（Word・Excel など、PDF の after と同じ形）：{\"@1\": {\"before\": \"年\", \"occurrence\": 2, \"value\": \"令和8\"}}・{\"@1\": {\"after\": \"氏名\", \"value\": \"山田\"}}"
                    "（字はそのままで、すぐ前・後ろに足す。段落・セルが「年　月　日」のように字で始まる所に：字間をそろえた並びに、先頭の所は足さない。"
                    "1 つに決めるのは occurrence、Word は near・near_paras、Excel は cell・sheet）。PDF の before は字の左の空き。"
                    "PDF の伏せ字（「○○」）の上には {\"@1\": {\"cover\": \"○○\", \"value\": 書く字}}（白い下地の欄で覆って書く。元の字は消さない）。"
                    "PDF の欄で kind_candidates（種類の候補）のあるもの：true/false はチェックの印、字はチェックの候補なら文字として、"
                    "1字ずつ書く欄なら1字ずつ書く。種類を決めて書くときは {欄の id: {\"kind\": \"text\", \"as\": \"郵便番号\", \"value\": 値}}（kind が種類、as は書いた値の意味）。"
                    "書いた値の意味 as：どの値にも {\"value\": 値, \"as\": \"生年月日\"} と添えられる（@ の書き込み・○にも。書類には書かれない）。say_to_user に"
                    "「生年月日：昭和60年4月1日（p1.t45-112／近くの字「住所」）」と出る。as のない所は「p1.t45-112（近くの字「住所」）：…」で、近くの字は事実であって欄の意味ではない。"
                    "本人が確かめる書類ではいつも as を渡す（会話が本人の確かめる手段）。as のなかった所は結果の written_without_as に並ぶ。"
                    "alternative_to のある欄は、その欄と同じ所の別の読み方：どちらか1つにだけ書く（両方に書くと断る）。"
                    "「？」の所の facts（所の字・字数・幅の目安・前後の字・字間を空けた並びの字数 chain）を読んで、書く所か・見出しの字間かを決める。"
                    "PDF の画面にも紙にも出ない欄（欄の hidden・窓の「［名前：画面にも紙にも出ない］」）には書かない：隠れた欄と分かったうえで書くときだけ "
                    "{欄の id: {\"value\": 値, \"allow_hidden\": true}}（出ないことが、本人に伝えることに入る）。"
                    "Excel で、印刷された字のあるセルをまるごと書き換えると決めたときは {セル: {\"value\": 値, \"overwrite_text\": true}}"
                    "（明示がないと、印刷された字を消す書き込みは断るか、証明で想定外の変更として示す）。"
                    "Excel のセルの字の一部を書き換えるときは、Word と同じ {\"@1\": {\"old\": セルの中の元の字, \"new\": 書き換えた後の字}} の形で書ける"
                    "（同じ字がいくつもあれば \"cell\": \"シート!B3\"・\"sheet\"・\"occurrence\" で 1 つに決める）。"
                    "本人の印影・署名の画像（本人は手で押印・自署できない）：本人に、どの書類のどこにどの画像を置くかを伝えて同意を得たときだけ、"
                    "{\"@1\": {\"seal\": 画像（PNG か JPEG）のパス, \"consent\": true, \"page\": ページ, \"box\": [左, 上, 右, 下]（mm）}} か、"
                    "字を目印に {\"@1\": {\"seal\": …, \"consent\": true, \"cover\": \"印\"}}（その字の上。\"after\" なら右、\"where\": \"below\" なら下。"
                    "\"size\" で画像の長い辺の mm）で置く（PDF のみ。署名の画像は \"seal\" の代わりに \"signature\"）。"
                    "consent: true がなければ置かない。画像は縦横の比を保って枠の中に置き、透けない画像がページの字を隠すときは置かない。"
                    "置いた所は id「pページ.s番号」になり、証明は、足したのがその注釈と本人の画像だけであることを確かめる。"
                    "実印・自署（手書き）が求められる所では、画像は受け付けられないことがある：様式の注意書きを読んで確かめ、本人に伝える。"
                    "印刷された字を○で囲む（「有・無」「1.新規 2.更新」「昭和・平成」「普通・当座」などから選ぶ）：{\"@c1\": {\"circle\": 囲む字（窓の字のとおり）, "
                    "\"near\": 近くの見出しの字}}（同じ字がいくつもあるときは near（同じ行→同じ列→近い所の順）か occurrence（読む順の何番目か。near といっしょなら、near に近い候補の中の何番目か：PDF・Word・Excel で同じ）で"
                    "1つに決める。PDF は \"page\"、Excel は \"cell\": \"シート!B3\" でも指せる）。窓は印刷された選択肢を欄として並べない"
                    "（どれが選択肢かは窓の字を読んで決める）。PDF はページに楕円を描き（ほかの字に触れない楕円を、すき間と縦横の比を変えて探す。"
                    "どれも触れるときは断るので、\"style\": \"border\"（字に沿った四角の囲み線。○でなくてよいと本人が確かめたときに）か"
                    "\"box\"（○の外枠 mm）を決める）、Excel（.xlsx）はセルの字の位置を見積もって楕円の図形を足す。Word（.docx）は既定が字の四角い囲み線（w:bdr。□だが字の書式なので位置がずれない）："
                    "Word はページの上の字の位置を保存しないので、字に結び付けた楕円の図形は描き手でずれる（\"style\": \"ellipse\" と明示したときだけ描く）。どちらで書いたかは circled の style・how に出る。"
                    "Excel の○と Word の楕円の位置は見積もりなので、結果の uncertain・本人に伝えることを伝え、確かめてもらう。古い .xls・.doc には○を描けない"
                    "（代わりに字の前に「○」を足す・選ばない字を消すなど、本人に確かめてから字として書く）。"
                    "囲んだ所は結果の circled（id・囲んだ字・見出し・前後の字 around：【】が囲んだ字）に出て、読み戻しで、○がその字だけを囲んでいるかを確かめる。"
                    "near が選択肢の並び（\"near\": \"大・昭・平・令\"）で、囲む字がその項目なら、その所を囲む。"
                    "near の見出しの字の中の字（見出し「有無」の「無」「都県市」の「都」）・ほかの字の並びに埋まった字は選ばず、"
                    "near と別の行の候補へは黙って移らない。決まらなければ候補を前後の字つきで示して断る"
                    "（そのときは occurrence で選ぶ。near の中の字も番号で選べる）。"
                    "結果として、変わったのが書いた欄（とその計算結果・見た目）だけであることの証明と、書き込み後のファイルを読み直した結果を返す。"
                    "PDF の重なりは、はっきりしたもの（印刷された字の上・書いた値どうし・○が字を切る）だけ断り、きわどい所（罫線に触れる・"
                    "字の縁にかかる・印や図の曲線・スキャンのページ）は warnings に所と位置（mm）を添えて知らせる。書いた後は、書いた PDF の"
                    "ページを view_page で見て、warnings の所を一つずつ確かめ、読みにくければ box・style を変えて書き直すこと。"
                    "Excel・Word などは描いて確かめられないので、say_to_user の見出しと値・折り返し／はみ出し／「####」の警告（wrap・fit。"
                    "見積もりの元に「仮に」とあれば、書類に書かれていない値を置いた見積もり）・○の uncertain を確かめる。"
                    "本人は書類を開いて確かめないので、結果の say_to_user（書いた所の見出しと値・まだ空いている欄・本人に伝えること）を、"
                    "対話でそのまま本人に伝えて確かめる。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"}, "out_path": {"type": "string", "description": "書き出すファイルのパス（元と別の名前）"},
         "values": {"type": "object", "description": "{欄の id: 値}。本人が確かめる書類では、値に意味を添える {欄の id: {\"value\": 値, \"as\": \"生年月日\"}}。"
                                                      "Word・Excel で字の前・後ろに足すときは {\"@1\": {\"before\" か \"after\": 字, \"value\": …}}、"
                                                      "Word で欄のない所は {\"@1\": {\"old\": …, \"new\": …}}、"
                                                      "PDF で欄のない所は {\"@1\": {\"after\": …, \"value\": …}}、"
                                                      "本人が同意した印影・署名の画像は {\"@1\": {\"seal\": 画像のパス, \"consent\": true, …}}、"
                                                      "印刷された字を○で囲むときは {\"@c1\": {\"circle\": 囲む字, \"near\": 近くの見出し}}"},
         "force": {"type": "boolean", "default": False, "description": "保護された欄にも書く（本人が了承したときだけ）"},
         "overwrite": _OVERWRITE},
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
                    "（結果には列の名前・行数・はじめの 20 行）。表の後ろの注の行もそのまま返すので、どこまでが表かは読んで決める。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"},
         "header": {"type": "string", "description": "表の見出しにある字"},
         "sheet": {"type": "string", "description": "シートの名前（省くと、見出しのあるすべてのシート）"},
         "header_rows": {"type": "integer", "minimum": 1, "description": "見出しの行数（省くと自動）"},
         "out_path": {"type": "string", "description": "すべての行を書き出すファイル（.csv か .json）"},
         "limit": {"type": "integer", "minimum": 1, "description": f"out_path がないとき、結果に入れる行の数の上限（既定 {READ_TABLE_LIMIT}）。"
                   "超えた行は返さず、行の数（rows）と返した行の数（rows_returned）を返す"}},
         "required": ["path", "header"]}},
    {"name": "explain_calculation",
     "description": "Excel の様式を「データ」（入力欄）と「処理」（入力から結果を出す数式）に分けて返す。式は番地を見出しに置き換えて示す"
                    "（例「＝「報酬総額」-「処遇改善加算の総額」」）。cell を指すと、そのセルの計算を入力欄までさかのぼって説明する。"
                    "様式の計算のきまりを本人に説明するときや、どの入力が結果を左右するかを知るときに使う。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"},
         "cell": {"type": "string", "description": "説明するセル（シート!番地）。省くと、結果のまとまりの一覧"}},
         "required": ["path"]}},
    {"name": "try_calculation",
     "description": "Excel の様式で、値の組から結果を試算する（ファイルは書かない）。values は fill_document と同じ {欄の id: 値}。"
                    "「区分Ⅱにしたらいくら？」のような問いに、書類を書き換えずに答えるため。入力規則に合わない値は断る。"
                    "結果の say_to_user（変わる所の見出しと、前 → 後）を、会話で伝える。",
     "inputSchema": {"type": "object", "properties": {
         "path": {"type": "string"}, "password": {"type": "string", "description": "パスワードで保護された書類のとき、本人に聞いたパスワード（結果にも記録にも残さない）"},
         "values": {"type": "object", "description": "{欄の id: 値}"}},
         "required": ["path", "values"]}},
    {"name": "view_page",
     "description": "PDF のページ、または画像のファイル（JPEG・PNG・GIF・BMP・TIFF・WebP・JPEG 2000 など。page は何枚目か）を画像で見る（スキャンした紙のように、字の情報がない画像だけのページを読むため）。"
                    "上と左の帯に、ページの左上からの mm の目盛りが付く。小さな字は region（[左, 上, 右, 下] mm）で拡大して読む。"
                    "書く所は、画像を見て決めた枠を fill_document に {\"@1\": {\"page\": ページ, \"box\": [左, 上, 右, 下]（mm）, \"value\": 書く字}} で渡す。"
                    "書く前に marks（[{\"box\": [...], \"label\": \"1\"}]）を渡すと、その枠を赤で描くので、位置が合っているかを目で確かめる。"
                    "ページは、字（埋め込みのフォントの字形）・線・画像・欄の見た目（/AP）を描く。書いた後のファイルを見ると、値の入った欄が青い枠で描かれ、値は文で返る。"
                    "書いた字が欄からはみ出す・つぶれる・横倒し・印刷された字に重なるなら、書いた見た目が正しくないので、書き直す（書いた後は、いつも見て確かめる）。"
                    "フォントが埋め込まれていない字は代わりのフォントで描き（結果の substituted）、使える字形がない字は灰色の四角になる。"
                    "JBIG2・JPEG 2000 の画像も描く（JPEG 2000 は、画面に要る解像度で開き、そのことを結果に書く。細かな字は region で拡大：範囲に掛かる所だけを開くので、どの部分でも元の解像度まで開ける。HEIC も範囲のタイルだけを開き、時間が足りないと途中まで返すので、結果のとおり同じ呼び出しをもう一度する）。開けない画像は灰色の枠と対角線で示して結果に書くので、その所は描けたと思わず、本人に伝える。そのとおりに描けなかったものは、結果の not_drawn に短く並ぶ。置いた印影・署名の画像も、その位置に描く。画像の中の字は推測で埋めず、読めない所は本人に伝える。",
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


# ---------------------------------------------------------------------- 追加の形式（CSV・TSV・JSON・JSON Lines・RTF・OpenDocument・.txt・PowerPoint）
_MORE = ("（ほかに CSV・TSV・JSON・JSON Lines・RTF・OpenDocument の文書 .odt と表計算 .ods・文字だけのファイル .txt・"
         "PowerPoint .pptx・한글 .hwpx（読む・書く）と .hwp（HWP 5。読むだけ）も開けます。"
         "Excel .xlsb（読む。値のあるセルの同じ大きさの上書きだけ書ける）・古い PowerPoint .ppt・XPS .xps/.oxps・EPUB・メール .eml/.msg・Excel 2〜4・Word 1.x/2.0・"
         "Lotus .wk1・dBase .dbf・Access .mdb/.accdb・OpenOffice.org 1.x .sxw/.sxc・iWork .pages/.numbers/.key は読むだけ。"
         "tar・.tar.gz・.gz（読む・書く）と LZH（読むだけ）は ZIP と同じく「a.tar!/中の名前」で中のファイルを指す。"
         "7z・RAR・DocuWorks は読めないので、次の一手を返す）")
for _t in TOOLS:
    if _t["name"] == "open_document":
        _t["description"] += _MORE
    elif _t["name"] == "fill_document":
        _t["description"] += (" RTF・OpenDocument の文書（.odt）・文字だけのファイル（.txt。P12 は 12 行目）・PowerPoint（.pptx。表の空のマス T1.R2.C3）は"
                              "Word と同じ書き方（［P12.1］・字を目印にした {\"@1\": {\"old\": …, \"new\": …}}）。.txt と .pptx には書いた所の印を"
                              "置けないので、書いた後のファイルでは残りの空欄の番号が詰まる（続けて書くときは、開き直して id を確かめる）。"
                              "OpenDocument の表計算（.ods）は Excel と同じ書き方（シート!B3）。CSV・TSV は「R行.C列」（R3.C2・R3.氏名。"
                              "今ある行の次の行に書くと、行を足す）、JSON は JSON Pointer（/applicant/name。ない名前は足し、配列の終わりは /items/-）、"
                              "JSON Lines は「L行/名前」。"
                              "한글（.hwpx）は Word と同じ書き方で、누름틀（クリックして書く欄）とフォームの部品は、その名前（名前のない 누름틀は ctl番号）で書く。"
                              "placeholder は値のないときに한글が表示する案内の字で、値ではない。HWP 5（.hwp）には書けない"
                              "（本人に HWPX で保存し直したファイルを頼む）。")

# 文字だけのファイル（CSV・TSV・JSON・JSON Lines・.txt）の文字コード：規則で決められず、窓に候補が並んだとき、エージェントが選んだもの
_ENCODING = {"type": "string",
             "description": "文字だけのファイル（CSV・TSV・JSON・JSON Lines・.txt）で、窓に「文字コードを規則では決められません」と"
                            "候補が並んだとき、字を読んで選んだ文字コード（候補の encoding の値。例 \"cp932\"・\"gbk\"）。"
                            "その文字コードで厳密に読めることを確かめて、窓・欄・書き込み・段階をその文字コードで読み直す"
                            "（書き込みは、書いた所だけをその文字コードのバイトにする）。同じファイルには、どの道具にも同じ値を渡す"}
for _t in TOOLS:
    if _t["name"] == "audit_document":
        _t["inputSchema"]["properties"]["job"] = {"type": "boolean", "description": "大きさによらず、長い仕事として動かす（呼び出しの時間の上限に当たったとき）"}
        _t["inputSchema"]["properties"]["restart"] = {"type": "boolean", "description": "控えを捨てて、始めから確かめ直す"}
    if _t["name"] in ("open_document", "read_window", "list_fields", "fill_document", "audit_document", "grade_document"):
        _t["inputSchema"]["properties"]["encoding"] = _ENCODING


def _text(obj) -> dict:
    return {"type": "text", "text": obj if isinstance(obj, str) else json.dumps(obj, ensure_ascii=False, indent=1, default=str)}


def _score(check: dict) -> str:
    return "書く段階" if check.get("score") == "write" else "読む段階"


def call_tool(name: str, args: dict) -> dict:
    if name == "open_document":
        path = args["path"]
        pw = args.get("password")
        enc = args.get("encoding")
        s = api.summary(path, pw, enc)
        if s.get("large"):
            lz = api.lazy_open(path, pw, enc)
            head = (f"{s['kind_ja']}・大きな書類です（窓の全体は見積もりで {s['estimated_window_chars']:,} 字）。全体を一度には返しません。"
                    "下の概要を見て、要る所を read_window に範囲で頼んでください（読む量を分けるだけで、読めない所はありません）")
            return {"content": [_text(head), _text(lz.overview_text())], "structuredContent": s}
        _k, _doc, parts = api.window(path, pw, enc)
        head = (f"{s['kind_ja']}・入力欄 {s['fields']} か所（書ける欄 {s['writable_fields']}）・窓は {s['window_parts']} 部分"
                + ("（read_window で part を指定して続きを読めます）" if s["window_parts"] > 1 else ""))
        return {"content": [_text(head), _text(parts[0][1])], "structuredContent": s}
    if name == "read_window":
        rng = {k: args[k] for k in ("pages", "sheet", "rows", "paragraphs", "records", "cursor", "raw") if args.get(k) not in (None, "", False)}
        large = api.lazy_open(args["path"], args.get("password"), args.get("encoding")) if not rng else None
        if rng or large is not None or args.get("max_chars"):
            if args.get("part") is not None and not rng:
                rng["part"] = args["part"]
            page, doc = api.read_range(args["path"], rng, args.get("max_chars"), args.get("password"), args.get("encoding"))
            return {"content": [_text(page.text), _text(page.footer())], "structuredContent": page.to_dict()}
        _k, _doc, parts = api.window(args["path"], args.get("password"), args.get("encoding"))
        n = int(args.get("part", 1))
        if not 1 <= n <= len(parts):
            raise api.ApiError(f"part は 1〜{len(parts)} です")
        return {"content": [_text(f"部分 {n}/{len(parts)}：{parts[n - 1][0]}"), _text(parts[n - 1][1])]}
    if name == "list_fields":
        rng = {k: args[k] for k in ("pages", "sheet", "rows", "paragraphs", "records", "cursor") if args.get(k) not in (None, "")}
        if rng or args.get("limit") or api.lazy_open(args["path"], args.get("password"), args.get("encoding")) is not None:
            res = api.fields_range(args["path"], rng, bool(args.get("only_empty")), args.get("limit"), args.get("password"), args.get("encoding"))
            return {"content": [_text(res)], "structuredContent": res}
        fl = api.fields(args["path"], args.get("password"), args.get("encoding"))
        if args.get("only_empty"):
            fl = [f for f in fl if f.get("value") in (None, "", False)]
        return {"content": [_text(fl)], "structuredContent": {"fields": fl}}
    if name == "fill_document":
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
        # 本人は書類を開いて確かめないので、say_to_user（書いた所の見出しと値・まだ空いている欄・本人に伝えること）を、対話で伝える
        return {"content": [_text(head), _text(res.get("say_to_user", "")), _text(res)], "structuredContent": res,
                "isError": not ok}
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
        res = api.read_table(args["path"], args["header"], args.get("sheet"), args.get("header_rows"), args.get("out_path"),
                             args.get("password"))
        res.pop("all_rows", None) if args.get("out_path") else None
        limit = args.get("limit", READ_TABLE_LIMIT)
        if not args.get("out_path") and res["rows"] > limit:
            res["all_rows"] = res["all_rows"][:limit]  # はじめの limit 行（行の数 rows と、返した行の数 rows_returned）
            res["rows_returned"] = limit
            res["note"] = (f"行が {res['rows']} 行あります。はじめの {limit} 行だけを返しました。すべての行は out_path（.csv か .json）"
                           "に書き出すか、limit を大きくして読んでください")
        return {"content": [_text(res)], "structuredContent": res}
    if name == "explain_calculation":
        res = api.calc_model(args["path"], args.get("cell"), args.get("password"))
        return {"content": [_text(res)], "structuredContent": res}
    if name == "try_calculation":
        res = api.calculate(args["path"], args.get("values") or {}, args.get("password"))
        return {"content": [_text(res["say_to_user"]), _text(res)], "structuredContent": res}
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
        return {"content": [_text("\n".join(lines)), _text(g)], "structuredContent": g}
    if name == "audit_document":
        au = api.audit(args["path"], args.get("password"), args.get("encoding"), restart=bool(args.get("restart")), job=bool(args.get("job")))
        lines = [("[OK] " if c["ok"] else "[要確認] ") + f"{c['name']}: {c['detail']}" for c in au["checks"]]
        job = au.get("job")
        if job and not job["finished"]:
            pr = job.get("progress") or {}
            at = f"{pr.get('text', '始めたところ')}" + (f"（{pr['done']:,}／{pr['total']:,}）" if pr.get("total") else "")
            lines.insert(0, (f"検査は止まりました：{job['stopped']}" if job.get("stopped") else f"検査は別のプロセスで動いています：{at}。") +
                         "もう一度 audit_document を呼ぶと、進み具合と結果が分かります（止まっていれば、控えの続きから進みます）")
        else:
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
        result = isolate.worker().call(name, args) if isolate.enabled() else tool_result(name, args)
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
