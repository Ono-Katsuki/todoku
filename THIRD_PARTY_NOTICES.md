# 他者のデータ

とどくのコードは Apache License 2.0 です（`LICENSE`・`NOTICE`）。次のデータは、それぞれのライセンスに従います。とどくは外部のライブラリに依存しません。

| データ | 置き場 | 元 | ライセンス |
|---|---|---|---|
| CID から Unicode への対応表 | `src/todoku/pdf/data/cid2unicode.json.gz` | Adobe mapping-resources-pdf | BSD 3-Clause（`src/todoku/pdf/data/LICENSE-adobe-mapping-resources.txt`） |
| Unicode から CID への対応表 | `src/todoku/pdf/data/uni2cid.json.gz` | adobe-type-tools/cmap-resources。Copyright 1990-2023 Adobe | BSD 3-Clause（`src/todoku/pdf/data/LICENSE-adobe-cmap-resources.txt`） |
| 字の組み立ての表 | `src/todoku/pdf/data/shaping.json.gz` | Unicode Character Database | Unicode License（`src/todoku/pdf/data/LICENSE-unicode.txt`） |
| 文字の性質の表（Scripts, BidiBrackets, PropList） | `src/todoku/chars.py` | Unicode Character Database | Unicode License（同上） |
| VP8 の復号の表 | `src/todoku/imaging/webp_vp8.py` | RFC 6386 §20.26。Copyright (c) 2010, 2011, Google Inc. | BSD 3-Clause（`LICENSES/LICENSE-webm.txt`） |
| Helvetica などの字幅 | `src/todoku/pdf/glyphs.py`, `src/todoku/pdf/fonts.py` | Adobe Core 14 AFM | Helvetica の字幅は Adobe Core 14 AFM による。Copyright (c) 1985, 1987, 1989, 1990, 1997 Adobe Systems Incorporated. All Rights Reserved. |
| 世界地図の国境 | `world/portal/countries-110m.json` | world-atlas 2.0.2。Copyright 2013-2019 Michael Bostock。元のデータは Natural Earth（public domain）。Made with Natural Earth. | ISC（`LICENSES/LICENSE-world-atlas.txt`）。Natural Earth は public domain |
| 政府のドメインの一覧 | `world/domains.json` | Public Suffix List の抜き出し。Wikidata | Public Suffix List は MPL-2.0（`LICENSES/LICENSE-MPL-2.0.txt`）。Wikidata は CC0 |

## 他者の実装に基づく部品

次の部品は、公開されている実装を読み、Python に書き直した部分を含みます。

| 部品 | 元 | ライセンス |
|---|---|---|
| `src/todoku/imaging/jpegarith.py`（算術符号の復号） | IJG libjpeg の `jdarith.c`・`jaricom.c` を Python に書き直したもの。This software is based in part on the work of the Independent JPEG Group. | IJG の条件（`LICENSES/LICENSE-ijg.txt`） |
| `src/todoku/imaging/webp_vp8.py`（`_large`・`_coeffs`・`_BANDS`・定数 101581） | libwebp | BSD 3-Clause（`LICENSES/LICENSE-libwebp.txt`） |
| `src/todoku/pdf/shape.py`・`src/todoku/pdf/otl.py`（字の組み立て） | HarfBuzz | MIT（`LICENSES/LICENSE-harfbuzz.txt`） |
| `src/todoku/imaging/jpx.py`・`src/todoku/imaging/jpx_t1.py`（一部） | OpenJPEG | BSD 2-Clause（`LICENSES/LICENSE-openjpeg.txt`） |
| `src/todoku/legacy/branch.py`・`src/todoku/legacy/ppmd7.py` | LZMA SDK・XZ | public domain・0BSD |

## 公開文書を参考にした形式

본 제품은 한글과컴퓨터의 한/글 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.

試験の見本（`tests/fixtures/public/`）は、とどくのために作ったもので、CC0 です。
