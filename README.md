# とどく（todoku）

**Claude Code は申請書を書ける。でも、任せられない。とどくは、書いた所のほかは 1 バイトも変えていないと証明する。**

既存の道具では、申請書を任せられませんでした。数か月苦労しました。だから、とどくを作りました。とどくは、書いた値を読み戻し、変えたのが記入した所とその計算結果だけだと、書類ごとに証明します。

行政書類を AI エージェントが正確に読み書きするための MCP サーバー。`dependencies = []`。Office も LibreOffice も不要。

[![デモ動画を YouTube で見る（3 分、日本語と英語の字幕つき）。サムネイル：世界地図に「世界の行政書類を、エージェントへ。185 か国・地域で測定、正確に書ける文書 約 238 万件」](https://img.youtube.com/vi/zB95Q_jUq5A/maxresdefault.jpg)](https://youtu.be/zB95Q_jUq5A)

**[デモ動画（3 分・日本語と英語の字幕つき）](https://youtu.be/zB95Q_jUq5A)** ・ **[ポータル：世界の調査と全結果](https://claude.ai/artifact/45P2GdSFYWkAqWrdLebmSH)**

```text
証明：元のファイル 110,831 バイトは先頭からバイト一致のまま、後ろに 18,121 バイトを足しました
```

```console
$ head -c 110831 hino_yogu_filled.pdf | cmp - hino_yogu.pdf && echo 一致
一致
```

エージェントの報告を信じる必要はない。`cmp` で検証できる。

[English](README.en.md)

## 30 秒で入れる

Claude Code で実行する（[uv](https://docs.astral.sh/uv/) が必要。対応する環境は macOS。Windows と Linux はβ）。

```text
/plugin marketplace add Ono-Katsuki/todoku
/plugin install todoku@todoku
```

新しい版にするときは、ターミナルで `claude plugin marketplace update todoku` と `claude plugin update todoku@todoku` を実行して、Claude Code を開き直す。

あとは自然言語で頼む。

```text
この様式に、私のことを書いて
この Excel の計算のきまりを説明して。区分 II にしたらいくら？
```

## 「書けた」と「正確に書けた」は違う

同じ申請書、同じ依頼、同じ Claude で比較した（公開の様式 20 通、値は架空。詳細は [ポータル](https://claude.ai/artifact/45P2GdSFYWkAqWrdLebmSH)）。

| | ただの Claude Code | Claude Code + とどく |
|---|---|---|
| 元の PDF をそのまま残す | 0/16 | **16/16** |
| 変わったのは書いた所だけと証明 | なし | **20/20** |

ある PDF は Poppler では正しく表示された。Apple PDFKit で開くと、氏名の欄は「16」。報告は「記入済み」だった。

とどくは PDF を増分保存し、元のバイトに触れない。書いた後は読み戻し、描画して、画素の変化まで確かめる。

## データと処理を、分けて渡す

値だけを渡せば、計算はエージェントの憶測になる。`extract_processing` は計算の規則をそのまま取り出す。

```text
新潟県の処遇改善計画書（公開の Excel）
数式 105,256 セル → 203 の規則
別のプロセスで動かし、保存された値と 105,256 / 105,256 セル一致

ROUNDDOWN(ROUND(L × P, 0), 0) × AA
```

順番を憶測すれば、1 円ずれる。

| 計算の順番（値は架空） | 見込額 |
|---|---|
| **様式のきまりどおり** | **84,692 円** |
| 全部掛けてから、最後に切り捨てる | 84,691 円 |
| 先に切り捨ててから、月数を掛ける | 84,690 円 |

`format: "python"` なら、標準ライブラリだけで動く `calculate(inputs) -> dict` として書き出す。

## 依存ゼロ、中身は Python

- **ファイル形式：** PDF（構文・暗号・フォント・XFA・増分の保存）、Excel（.xlsx・.xls・.xlsb）、Word（.docx・.doc）、一太郎、HWP／HWPX、ODF、RTF、PowerPoint、CSV、JSON、ZIP など。公開仕様のある形式は仕様書から実装し、一部は OSS の実装を読んで Python に書き直した（`THIRD_PARTY_NOTICES.md`）。
- **数式エンジン：** Excel と同じ規則（有効数字 15 桁、1900 年のうるう年、和暦の表示形式）で計算する。VBA も実行する。
- **決定的：** 読む・書く・証明・検査に AI を使わない。規則で決められない所は `？` で示す。事実は道具が出し、判断はエージェントが下す。
- **MCP の道具 14 個：** `open_document`・`read_window`・`list_fields`・`fill_document`・`audit_document`・`extract_processing`・`try_calculation`・`run_macro` など。大きな書類は範囲で順に読む。読めない所は、読めないと伝える。

約 15 万行・312 ファイル・テスト 5,703 件。

## 実物で測った

版を固定し（`eval-freeze-2`）、185 の国と地域の政府サイトの書類 935 件を実際に読み書きして測定した。うち 878 件は、固定するまで誰も開いていない新しい標本だ。

**Common Crawl が集めた政府サイトの書類のうち、正確に読める文書は推定約 245 万件、正確に書ける文書は推定約 238 万件。** 全カテゴリの表・区間・方法は [ポータル](https://claude.ai/artifact/45P2GdSFYWkAqWrdLebmSH)、評価コードは `scripts/eval/`。

- 正確に読める：**約 245 万件**（95% 区間 225〜264 万件）
- 正確に書ける：**約 238 万件**（95% 区間 217〜257 万件）
- 初めて見る日本の様式 46 件に日本語で記入：**45 件で正しく入り、証明と読み戻しは 46 件すべてで通った**

## 作った人

> エージェントがアクセスできないことは、アクセシビリティの問題です。

私は寝たきりの開発者で、経営者で、研究者だ。話せない。手も使えない。入力は頭のスイッチひとつ。千字書くのに一時間以上かかる。

このリポジトリは、そのスイッチで Claude と会話して、1 週間で作った。実物の書類でバグを見つけ、別のエージェントが確かめ、直す。その繰り返しだ。

私たちの事業所は申請がとても多い。これからは私も社員も、申請書をとどくで書く。

| Claude の使い分け | 何のために |
|---|---|
| Claude Code（CLI） | 開発の中心。設計・実装・検証を、頭のスイッチで会話して進めた |
| Remote Control | 寝たまま iPad から、Mac の Claude Code を動かす |
| サブエージェント（worktree） | 実物の書類でバグを探す・直す・別のエージェントが確かめる、を並行して回す |
| MCP サーバー（道具 14 個） | とどく本体。エージェントが書類を読み書きし、証明する入口 |
| スキル `shinsei` | 申請書を探して記入する進め方を、エージェントに渡す |
| プラグイン（マーケットプレイス） | MCP サーバーとスキルを一度に入れる |
| Claude.ai の Artifact | 世界の調査の結果を、ポータルとして公開する |

> Claude があれば、寝たきりも開発者になれる。

論文：Katsuki Ono, [*Reconfiguration of Ability: An Autoethnography of LLM-Mediated Development by a Developer with Severe Physical Disability*](https://doi.org/10.1145/3797867.3828990), ASSETS '26

Built with Claude: Japan 2026 応募作品（開発期間 2026 年 11 月 2 日〜8 日）。

---

## 詳しく

[詳細レポート](docs/REPORT.md)：MCP の使い方、道具ごとの説明、扱える形式、証明と段階の読み方、測定の条件。

## 書類と仕様書の出どころ

集めた書類と仕様書は**リポジトリに含めない**（`corpus/` は `.gitignore`）。取得元の URL・掲載ページ・SHA-256 を
[`sources/`](sources/) に記録してあり、同じものを取り直せる。取得は robots.txt に従い、1 件ずつ間隔をあけて行った。
ボット対策は破っていない。XFA 3.3 の仕様書は Internet Archive の写しから取得した（再配布なし）。

書類ごとの一覧は [詳細レポート](docs/REPORT.md) にある。

## ライセンス

Apache License 2.0（`LICENSE`・`NOTICE`）。個人でも会社でも、自由に使える。中に入れた他者のデータ（PDF に足す字の書体 Noto Sans JP（SIL OFL 1.1）・Adobe の文字の表・Unicode の表・RFC 6386 の表・地図のデータなど）は、それぞれのライセンスに従う（`THIRD_PARTY_NOTICES.md`・`LICENSES/`）。

본 제품은 한글과컴퓨터의 한/글 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.

Microsoft、Excel、Word、PowerPoint、Access、一太郎、한글、Hancom、RAR、Adobe、PDF などは各社の商標または登録商標だ。とどくはこれらの会社と提携しておらず、承認も受けていない。
