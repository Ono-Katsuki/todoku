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

## zlib の puff.c（Deflate64 を解く所：src/todoku/legacy/sevenzip.py の deflate64 と符号の表）

符号の表の作り方と 1 字ずつ解く形が、zlib の contrib/puff/puff.c と同じ形に見えます。由来を記録で確かめられないので、派生として扱い、変更した版（Python で書き直し、Deflate64 の長さと距離に広げたもの。元の puff.c そのものではありません）として次の表示を残します。

```
Copyright (C) 2002-2013 Mark Adler, all rights reserved

This software is provided 'as-is', without any express or implied
warranty.  In no event will the author be held liable for any damages
arising from the use of this software.

Permission is granted to anyone to use this software for any purpose,
including commercial applications, and to alter it and redistribute it
freely, subject to the following restrictions:

1. The origin of this software must not be misrepresented; you must not
   claim that you wrote the original software. If you use this software
   in a product, an acknowledgment in the product documentation would be
   appreciated but is not required.
2. Altered source versions must be plainly marked as such, and must not be
   misrepresented as being the original software.
3. This notice may not be removed or altered from any source distribution.
```
