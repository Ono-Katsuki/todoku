# 他者のデータ

とどくのコードは Apache License 2.0 です（`LICENSE`・`NOTICE`）。次のデータは、それぞれのライセンスに従います。とどくは外部のライブラリに依存しません。

| データ | 置き場 | 元 | ライセンス |
|---|---|---|---|
| CID と Unicode の対応表 | `src/todoku/pdf/data/cid2unicode.json.gz`, `uni2cid.json.gz` | Adobe mapping-resources-pdf, cmap-resources | BSD 3-Clause（`src/todoku/pdf/data/LICENSE-adobe-mapping-resources.txt`） |
| 字の組み立ての表 | `src/todoku/pdf/data/shaping.json.gz` | Unicode Character Database | Unicode License（`src/todoku/pdf/data/LICENSE-unicode.txt`） |
| 文字の性質の表（Scripts, BidiBrackets, PropList） | `src/todoku/chars.py` | Unicode Character Database | Unicode License（同上） |

試験の見本（`tests/fixtures/public/`）は、とどくのために作ったもので、CC0 です。
