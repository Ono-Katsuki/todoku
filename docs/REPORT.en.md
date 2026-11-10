# todoku: detailed report

How to use it, what each tool does, how each format is built, how to read the checks, and the conditions of each comparison. Full results are on the [portal](https://claude.ai/artifact/45P2GdSFYWkAqWrdLebmSH).

[日本語](REPORT.md)

**Summary**

- 935 government-site documents from 185 countries and regions were actually read and written after freezing the version (878 of them a new sample nobody opened before the freeze). Estimated documents read accurately: about 2.45 million; written accurately: about 2.38 million.
- The baseline count is an upper bound: a person picks the best tool per format and todoku points out the fields. Even so, baselines write 0 .doc and .xls files and read 0 .doc, .odt, .ods and .rtf files, and on .xlsx it is 29 to 11 for writing and 35 to 14 for reading. Requiring the original PDF bytes to be kept, documents written are about 2.38 million for todoku and about 1.92 million for the baselines.
- todoku proves, document by document, that nothing changed besides what it wrote and the resulting calculations. The baselines have no such proof.

## Results (`eval-freeze-2`)

Tool: `eval-freeze-2` (86bba80). Scoring: the pre-registered v2 (plan in Japanese: [再評価の計画](再評価の計画.md)). Documents: 935 from 185 countries and regions (list: [`sources/eval_freeze2/world.tsv`](../sources/eval_freeze2/world.tsv)).

- **878 documents (161 countries and regions):** a new sample nobody had opened before the freeze. 900 were chosen; 22 with truncated captures were excluded.
- **57 documents (24 countries):** strata with 50 or fewer candidates. Chosen from survey documents by seed and hash order, without looking at results. 11 overlap the frozen version's comparison. todoku passed 39 on reading and 38 on writing.

| World (stratum × weight 1/π estimate) | Read | Write |
|---|---|---|
| **Accurate with todoku** | **2,452,605** (2,247,931–2,637,377) | **2,377,412** (2,169,193–2,567,703) |
| of which Japan | 289,454 | 299,125 |
| From the 878 alone | 2,452,339 | 2,377,089 |

The frame is 3,358,926 government-site documents in Common Crawl (CC-MAIN-2026-43); measured strata cover 3,344,887. Intervals are 95%. Per-country samples are small, so country numbers are indicative.

Definition: "write" means a place was chosen and written, the value landed, and nothing was lost. Documents with no place to write count as not writable. Blanks inside text count as places, and places todoku refuses are re-chosen. A count close to the earlier definition (no blanks inside text, no re-choosing) gives 1,508,644 for todoku and 1,506,136 for the baselines. The 45 csv, odt, ods and rtf documents are not write-scored and count as not writable.

Raw counts on the 878: read 657, write 601. Read grades: S 20, A 461, B 139, C 72, E 180, none 6.

Rescoring the frozen version (`eval-freeze-1`) on its 1,079 documents with v2 gives read 2,273,554 and write 2,106,851 (2,196,131 and 1,822,350 before the scoring fixes). That difference comes from the scoring fixes and the change in the write definition.

### Japanese forms seen for the first time

48 documents from the listing pages of Fukushima, Akita, Fukuoka and Kashiwa (not random; [`sources/eval_freeze2/jp_forms.tsv`](../sources/eval_freeze2/jp_forms.tsv)). 2 with no place to write are excluded from writing.

| Format | Docs (write) | ASCII | Japanese | Proof, read-back | Baselines (Japanese) |
|---|---|---|---|---|---|
| PDF (no fields) | 26 | 24 | 25 | 26, 26 | pypdf 0; pypdf overlay 26 (original bytes not kept) |
| .doc | 6 | 6 | 6 | 6, 6 | cannot write |
| .xls | 7 | 7 | 7 | 7, 7 | cannot write |
| .xlsx | 6 | 6 | 6 | 6, 6 | openpyxl 3 (shapes, comments etc. lost) |
| .docx | 1 | 1 | 1 | 1, 1 | python-docx 1 |
| Total | 46 | 44 | 45 | 46, 46 | |

Reading: 40 of the 46 with text (.doc 6/6, .docx 1/1, .xls 7/7, .xlsx 6/6, PDF 20/28; PDF baselines pypdf 22, pdfplumber 20).

### Comparison under conditions that favour the baselines

The baseline column is what a person gets by choosing the best tool for each format, with todoku pointing out the fields, and checking against the known answer. It is not what an agent gets working alone. Tools: MarkItDown, openpyxl, python-docx, pdfplumber, pypdf, xlrd.

How to read it: this upper bound does not require keeping the original document or proving nothing else changed. The PDF overlay (pypdf overlay) kept the original bytes in none of its 502 writes. Requiring the original PDF bytes to be kept, documents written are 2,377,412 for todoku and 1,915,851 for the baselines (last row below). todoku kept the original bytes and passed its proof in all 511 PDFs it wrote.

| World | Read | Write |
|---|---|---|
| todoku | **2,452,605** | **2,377,412** |
| Any baseline | 2,388,684 (2,178,692–2,580,045) | 2,374,569 (2,164,878–2,566,164) |
| Pass with todoku, not with any baseline | 133,662 (94,794–182,564) | 139,515 (99,745–189,254) |
| Pass with baselines only | 69,741 (30,806–134,464) | 136,639 (83,321–210,092) |
| Written with the original PDF bytes kept: todoku / baselines | – | **2,377,412** / 1,915,851 (1,683,348–2,141,464) |

Counts by format on the 878 (todoku / baselines):

| Format | Docs | Read | Write |
|---|---|---|---|
| PDF | 730 | 519 / 518 | 511 / 529 |
| PDF (written, original bytes kept) | 730 | – | **511** / 398 |
| .xlsx | 36 | **35** / 14 | **29** / 11 |
| .docx | 35 | **34** / 28 | **34** / 32 |
| CSV | 21 | 20 / 21 | – |
| .doc | 18 | **17** / 0 | **17** / 0 |
| .xls | 14 | **13** / 6 | **10** / 0 |
| .odt | 13 | **12** / 0 | – |
| .ods | 9 | **5** / 0 | – |
| .rtf | 2 | **2** / 0 | – |

Features measured on `eval-freeze-1` (documents not used in development; all tables on the portal):

| | todoku | Baselines |
|---|---|---|
| World government documents opened | **552/557** | 494/557 |
| Furigana (Tokyo forms) | **698/713** | 77/713 |
| Text in shapes and text boxes (world, excluding .doc) | **253/333** | 220/333 |
| Character recall on legacy Word .doc (Tokyo forms) | **1.0** | 0.0 |

### Known weaknesses

- **39 PDFs only todoku failed to write (of the 878):** about 5 correct refusals to protect certified signatures; about 15 written cleanly but misread by the checking OCR; about 12 written over photos or logos that are not fields; about 5 overlapping printed text; 2 hard to read (black on a dark background, clipped at the page edge). The breakdown was sorted by eye.
- **Non-fields offered as fields:** on documents that are not forms, photos and logos can be offered as places to write.
- **Unreadable text can still be written:** the render check looks at pixel change and contrast, so it does not stop every case of text lost in the background.
- **Read grades that are too kind:** 19 documents graded S or A missed the accuracy bar: 5 Arabic PDFs (all characters present, bigrams .89–.97; baselines .47–.63), 4 .ods/.odt, 3 PDFs with text in images (one with character recall 0.2 graded A), 7 just under the bar (.989–.995).
- **Excel calculation:** full recalculation uses the same calculator as writing (checked against real Excel on 3 input variations, 306 cells). Functions with undetermined dependencies such as `INDIRECT`, `OFFSET`, `TODAY` are only counted. The big-workbook path was tried on 3 real files and on no .xlsm. A big-workbook failure after writing leaves the file name in place. Id collisions between same-named checkboxes are unchecked.
- **105,256 cells:** the Niigata form is blank, so 87% of the results are empty values. 13,815 cells hold values, and all cells matched, empty ones included.
- **Legacy Word form fields:** fields that take no input, calculated fields and current-date/time fields are read from the field settings, shown as "not a place to enter", and not written without `force` (.docx, .doc, RTF). Word 6.0/95 .doc files are read without these settings. Number and date formats of a field (`#,##0` etc.) are shown as facts but not applied to written text. Calculated fields are not recalculated (when a written field is set to recalculate on exit, the result says so).
- **Reading text-only files:** the sample shown for each candidate encoding is the 48 characters at the first place where the candidates differ, so single-byte encodings can show identical samples (they differ elsewhere in the file). There is no argument yet to choose a CSV delimiter and reread (a file where two delimiters both fit is read with comma or tab, and the column count under the other delimiter is told). The vertical bar "|" is not detected as a delimiter. A `sep=` declaration line appears as the first row of the window.
- **Limits of tracked changes, strikethrough and font colour:** text deleted under tracked changes is shown apart from the body in .docx, .odt and RTF (in .doc it is shown as "hidden text"). Text written to .odt and RTF is not recorded as a change (when tracking is on, the result says so). The RTF change count is the number of runs of marked characters. Strikethrough and font colour are shown for .docx, .xlsx, .odt, RTF, .pptx, .hwpx and .hwp; Ichitaro (.jtd) character decoration and PDF strikethrough (a line drawn over the glyphs) are not read. For .pptx only settings on the run itself are read (colour and strikethrough inherited from the shape, layout or master are not), and a colour with a brightness or similar transform is not computed but shown as "#RRGGBB, modified". HWP 5 strikethrough and colour were implemented from the specification and not checked against a file made by Hangul. The large-.docx path treats black as the "normal colour" (in a document whose normal colour is not black, every paragraph gets a colour mark).
- **Limits of PDF layers:** visibility is decided from the default configuration (`D` of `OCProperties`: `BaseState`, `ON`, `OFF`), the usage switches for screen and print (`AS` with View, Print, Export) and membership marks (OCMD `P` and `VE`). The Zoom, Language, User and PageElement usages, other configurations (`Configs`) and the state after a person toggles layers in a viewer are not evaluated. Text on a layer that is off on screen is shown line by line, not shaped into tables; lines and shapes are only reported as present. Text that is off on screen but printed is not part of the overlap guard for writing. The split into layers is not checked by a second reading method (the omission check verifies that the character counts match, the separated text included). Where a replacement-text mark (ActualText) encloses a region on an off layer, the omission check can report a mismatch. Of the 835 collected PDFs, 15 have layer settings and 6 have a layer that is off on screen or in print (5 of them are XFA forms from USCIS and Canada with buttons and guidance text on a screen-only layer).
- **Limits of underscore ("_") blanks:** in PDF, underscores narrower than one em become a place only when they stand in a row separated by spaces ("_ _ _ _") or chained by separators ("_/_/_"). A lone short underscore gives no place. A spaced row is one place, not one box per character (the written text may shrink to fit its width, and the result says so). An underscore run touching a name character ("Tel.____", "1____"; letters and digits of scripts that separate words with spaces, such as Latin, Cyrillic and Greek) is a "?" place; touching a script that can break lines inside words (UAX #14 line-break classes ID, CJ, H2, H3, JL, JV, JT, SA: Han, kana, Hangul, Thai and others; `chars.SPACELESS_SCRIPTS`) it is a rule field. In Excel, Word and the like, a single underscore between name characters ("file_name", "имя_файла") is not a place. Of the 835 collected PDFs, one has rows of short underscores (a Ukrainian form: two places, number and date).
- **Limits of blanks made with tabs:** underlined tabs are a rule blank when three or more stand in a row, and a "?" place when there are one or two. A tab in a paragraph that has a tab stop with a leader line is a "?" place, and the paragraph's leader kinds are given as a fact (there is no layout, so todoku does not decide which tab reaches that stop). Leader tab stops are read in .docx and RTF, not in .odt (`style:leader-style`), HWPX or .doc. A written value goes in before the tab, in the tab's formatting; the tab is kept, and the result says so. A value written before a leader tab touches the label (add a space before the value to separate them). The remaining tabs of a written line are not offered again. Ichitaro (.jtd) tabs carry no underline information. A Word 6.0/95 .doc cannot take an insertion where there is no character. In RTF, the tab after a list number is read as part of the number. Of the 494 collected Word-family documents, 2 have blanks that contain tabs (two versions of one form, 6 places of underlined tabs), and leader tabs appear only in table-of-contents lines. No .hwpx or .hwp is in the collection; these were checked on hand-built documents.
- **Limits of field results:** the result text of a field (table of contents, page number, date and so on) is rebuilt when Word updates the field, so no place is shown inside it and an anchored write into it is refused (.docx, .doc, RTF, and the bodies of tables of contents, indexes and bibliographies in .odt; also a table of contents whose result runs over many paragraphs). A field with no end mark is read as ordinary text. A legacy form field whose result spans paragraphs is read as a field in its first paragraph only. The "locked" flag of a .doc field (fLocked) is not read. In a paragraph that has field results, the blanks outside them are still shown as places. A zero-width place directly after a computed value ("3[?]") is not shown (an anchored write can still be made there). Only EQ (even distribution, enclosed characters, phonetic guides) is read as text the author wrote into the instruction and so as label text; the results of HYPERLINK, QUOTE and MACROBUTTON are treated like computed values. Table labels (the text of the cell to the left or above) include field results. In the 494 collected Word-family documents, compared with the previous version, 45 places are gone (directly after page numbers, in table-of-contents lines, and the "□" inside enclosed characters) and 43 are new (41 after labels laid out with EQ and after the colon in those paragraphs, 2 at the trailing space of page-number lines). One .doc has a field whose result spans paragraphs (the table of contents of a Bulgarian form, 41 paragraphs).
- **Limits of symbol fonts:** characters written in a symbol font (Symbol, Wingdings, Wingdings 2, Wingdings 3, Webdings, ZapfDingbats) are read as Unicode characters through a table per font (.docx, .doc, RTF, .odt, and Symbol characters in embedded MathType equations; the tables come from Adobe's Symbol and ZapfDingbats tables and from WG2 N4384, the proposal that added Wingdings and Webdings to Unicode). A code that is not in the table, and any character of a font that the document declares as a symbol font but that has no table (Marlett, MT Extra, JustUnitMark and so on), stays a private-use character (U+F0xx), is told as a character outside Unicode, and makes reading B (private-use characters in .txt, .pptx, .hwpx, .hwp and Ichitaro files are told the same way). The declaration is looked up per font entry (`w:charset` 02 in .docx, chs 2 in .doc, `\fcharset2` in RTF, `style:font-charset="x-symbol"` in .odt): some documents list the same font name once with the declaration and once without. A document converted by LibreOffice writes the sign "©" as a character of an Arial declared as a symbol font, so it is told as a character outside Unicode. Not looked at: fonts of table styles in .docx, the document default font in .doc, Word 6.0/95 .doc files, symbol fonts in .xlsx and .pptx, and the level font of an RTF list (the mark is read from the `\listtext` text). The checked character is written with the code of its own font only for `w:sym` and bullet marks in .docx, and a box used as a bullet mark becomes a field only in .docx and only when it is unchecked. Written text does not inherit a symbol font (if it did, the written text would turn into that font's pictures or Greek letters): .docx and .doc drop the symbol font from the character format, RTF returns to the font chosen before it, .odt returns to the font of the enclosing style, and the result says so. Where a paragraph style sets the symbol font (.docx, .doc) or there is no font to return to (RTF, .odt), the text is written as is, and the output is not written when the text read back differs from the text written. In the collected documents, the grade changed for 15 of 530 files (all of them specification .txt files, from A to B, because private-use characters are now told); in the 494 Word-family documents the window text changed only in the bullet-mark characters of 11 files ("●" to "⚫", "➢" to "⮚", "■" to "◼", a private-use character to "•": 160 marks in all) and the places are the same. One document has places directly after a symbol-font character (two places after the seal mark in "医師の氏名㊞" of an asbestos health-check form; the mark is a JustUnitMark character).
- **Shaping limits:** a character missing from the font (such as "." in a Thai font) is refused rather than taken from another font. The Arabic lam-alef ligature is not formed with some font versions (seen with Noto Sans Arabic on Ubuntu).
- **Countries and regions not measured (10):** Central African Republic, North Korea, Eritrea, Eswatini, Gabon, Guinea-Bissau, Micronesia, São Tomé and Príncipe, Tuvalu, Vanuatu. No government documents were found in the crawl. 247 strata with 50 or fewer candidates in the 161 countries are also unmeasured.

## Usage: as an MCP server

todoku is an MCP (Model Context Protocol) server. It works from any AI agent that speaks MCP.

**In Claude Code, install it as a plugin. It then works from every folder.**
The plugin brings the MCP server and the skill `shinsei`, the workflow for finding and filling in application forms.

1. Install the Python tool [uv](https://docs.astral.sh/uv/). It launches todoku; todoku itself runs on the Python standard library alone.

   ```
   curl -LsSf https://astral.sh/uv/install.sh | sh
   ```

   (With Homebrew: `brew install uv`. Supported on macOS; Windows and Linux are in beta.)

   On first run, uv downloads Python 3.11 (about 30 MB). To work offline, run `uv python install 3.11` first while connected.
2. Install the plugin inside Claude Code.

   ```
   /plugin marketplace add Ono-Katsuki/todoku
   /plugin install todoku@todoku
   ```

**From an iPad or phone:** todoku runs on your computer, and documents stay there. Use Claude Code's
[Remote Control](https://code.claude.com/docs/en/remote-control) to operate the computer's Claude Code from an iPad.
Run `claude remote-control` on the computer (or `/remote-control` in an open Claude Code conversation), then connect to the session from the Claude app on the iPad
or from [claude.ai/code](https://claude.ai/code) in a browser. The computer's MCP server (todoku) and files stay available. Keep the computer on. (Available on Pro, Max, Team, Enterprise and similar plans. Not with an API key.)

Then just talk: "Find grants I can use as a bedridden person and write the application." "Put our office's details into this form."
Opened inside this repository, todoku also loads from `.mcp.json`. For other MCP agents, configure:

```json
{ "mcpServers": { "todoku": { "command": "uv", "args": ["run", "--quiet", "--no-dev", "--directory", "<path to this repository>", "todoku-mcp"] } } }
```

| Tool | What it does |
|---|---|
| `open_document` | Opens a document; returns the format, the number of input fields and the start of the window |
| `read_window` | Returns the window (Markdown). Large documents are requested by range (`pages`, `sheet`+`rows`, `paragraphs`, `records`). Following `next` (`cursor`) reads everything |
| `list_fields` | Lists input fields (label, type, choices, current value, writability, context of the row) |
| `fill_document` | Writes to input fields and exports under another name. Returns the proof, the read-back result and what to tell the applicant (`say_to_user`) |
| `audit_document` | Checks whether this document can be read and written completely and accurately |
| `run_macro` | Runs an Excel macro (button) and exports the changed cells with a proof. Yes/no prompts go to the applicant first |
| `grade_document` | The document's grade (read S–F, write A–D) and accessibility alerts |
| `read_table` | Extracts Excel and Access tables row by row under their header columns (merged cells spread as on screen, each value with its source cell; `rows` requests a range; exports to CSV; rows not returned from the used range are reported in `omitted_rows` as a count and row-number ranges) |
| `extract_processing` | Extracts a document's calculation as processing. **Use this to learn the calculation rules.** Folds same-shaped formulas into one rule ("D3:D100000 = B{row}*C{row}") and returns inputs, lookup tables, outputs, order, branches, macros and PDF calculations. `format: "python"` exports a verified `calculate(inputs)` module (see "Extracting processing as processing" below) |
| `explain_calculation` | Explains one value: traces the result back to its inputs in a formula written with labels |
| `try_calculation` | Computes a result from a set of values without writing the document ("What if I choose category II?") |
| `view_page` | Shows a PDF page as a rendered image (glyph shapes, lines, images, field appearance; `dpi` sets the resolution), with mm rulers. For reading scans, checking frames before writing and checking the look afterwards |
| `view_image` | Shows an image inside a document (by the id in the window's image list or by part name; `region` views a part at original detail) |
| `to_pdf` | Wraps an image file in a PDF. Write to the resulting PDF with `fill_document`, giving a frame in mm |

Example dialogue:

> Applicant: "Put the same office details as last year into this pay improvement plan."
> Agent: (`open_document` → `list_fields` → `fill_document` for everything known)
> "I filled in what I could. Form 2-1, office name: XX Welfare Corporation / postal code [123]-[4567] / address: … — 18 places in all.
> I read the written file back and confirmed the values match. The one thing I need from you is the addition category (I–V). Which is it?"

The same operations are available from the command line (`uv run todoku --help`).

<details>
<summary>Defenses against broken documents</summary>

The MCP server runs tool calls in a separate process (a worker). Each call has time and memory limits (`TODOKU_TIME_LIMIT`, `TODOKU_MEMORY_LIMIT_MB`). Past a limit, the worker is stopped and the call says so.

- **Windows (beta):** this split (passing pipes) is unavailable, so calls run inside the server. The limits are the budgets for inflating, rendering and items (`TODOKU_MAX_ITEMS`) plus a watchdog thread for time and memory. It stops at Python-level boundaries, so it does not interrupt a single operation inside C. The command line runs on the same budgets.
- **macOS:** the memory limit RLIMIT_AS is not enforced. Long calls also have their worker's memory watched from outside with `ps`. Whether `ps` works is checked once at server startup. Where it does not, the server reports that this watch is off and only the worker's internal watchdog remains.
- **Linux:** RLIMIT_AS is enforced even without `ps`. Only the external resident-memory watch stops.

In both cases the server reports this on standard error and in its description (instructions).

</details>

## Extracting processing as processing (`extract_processing`)

Take only the data from a document and the processing becomes the agent's guess. `extract_processing` hands over the document's calculation rules in a complete, structured form (the Excel family: .xlsx, .xlsm, .xls, .xlsb, .ods and others; PDF forms). It runs on the standard library alone. No value is decided by AI.

- **Rules:** same-shaped formulas (equal in R1C1) fold into one rule per block. Each rule returns `text` (`D3:D100000 = B{row}*C{row}`), `formula` (as written in the file), `a1`, `r1c1` (normal form), ranges read, cell count and exceptions (different formulas and overriding constants inside the range). An empty cell splits a block; when the range is filled and rule cells are the majority, the odd cell is an exception. Workbooks with many rows are streamed with no per-cell records (100,000 rows × 2 columns of formulas: 4 rules, under 24 MB of added memory).
- **Structure:** inputs (labels as the fact "text to the left and above", the rules reading them, data validation), parameters (value cells read by absolute reference), lookup tables (ranges read by VLOOKUP, INDEX, MATCH, XLOOKUP and others; contents returned as data with a `read_table` request), names, outputs (cells nothing reads), rule order (what is read comes first) and cycles (distinguishing true cell loops from running totals that compute in order), branch trees for IF, IFS and SWITCH, array formulas (kept as formulas, not expanded), references to external workbooks, volatile functions and places not statically knowable (an `OFFSET` height of `COUNTA(A:A)`, for example), hidden sheets, rows and columns used in calculation, and flows between sheets.
- **Macros:** procedure source, cells read and written (only where the text decides them: `Cells(i, 5)` is "E?", and otherwise "unknown"), when they run (`Worksheet_Change` and its watched range, buttons, shapes), procedures called, and the relation to rules (whose inputs they write, whose results they overwrite).
- **`format: "python"`:** exports one `.py` with `calculate(inputs: dict) -> dict`. It depends on neither todoku nor anything outside the standard library. Calculation behavior (rounding, dates, errors) is cut from the engine's source, only as much as is used, so no hand-kept copy of those rules exists. An output that depends on a function it cannot compute raises an exception stating why (`NotReproducible`) (`outputs` selects other cells; `on_unknown="mark"` marks them).
- **Verified before returning:** the exported module runs in a separate process (`python -I -S`, with todoku unimportable), and every formula cell is compared against both the saved value and todoku's calculation. The result is `verification`: match rate, all mismatches (in a separate file when many), and `proven`/`mismatch`/`not_proven` per rule with reasons. Rules that depend on volatile functions, cells without saved values and rules that cannot be reproduced are reported per rule as not verified.

Measured on development forms: of 13 Excel files from Tokyo's 23 wards such as Katsushika and Sumida, 1 contains formulas (42 cells → 42 rules, 100% match; the rest have no formulas). The 9 files in corpus/forms: 149,630 cells → 1,383 rules (largest workbook 105,256 cells → 203 rules). Matches against both saved values and todoku's calculation are 100%. Against values computed by real Excel, with inputs varied 3 ways across 306 cells, there is no difference.

One value's explanation is `explain_calculation`. What-if calculation is `try_calculation`.

<details>
<summary>Extraction coverage (<code>limits</code> in the result) and PDF forms</summary>

Conditional formatting and data validation are listed as their own rule types (not part of the Python export). PivotTables, Power Query (M source) and data connections are shown as "processing that cell formulas cannot express", with definitions only (source range and fields); the calculation is not reproduced (.xlsx, .xls and .ods are read to the definition; .xlsb only reports that a pivot exists). INDIRECT and OFFSET dependencies are not statically knowable (the module computes with the actual references). When a dynamic array's result changes size, its spill range changes as in Excel.

For PDF forms, calculation order (/CO) and rules between fields (`AFSimple_Calculate`) are returned in the same form and exported to Python. Custom JavaScript and simplified field notation are shown as they are and marked as not evaluated (`evaluated: false`).

</details>

## Supported formats

Read, write and proof are defined per format. **In Excel, nothing changes except the written cells and their calculated results. In PDF, not one original byte changes.**

<details>
<summary>Format list</summary>

| Format | Read | Write | Proof |
|---|---|---|---|
| Excel (.xlsx, .xlsm, .xls, Excel 5.0/95) | Cells, formulas (checked in full by the in-house calculation engine), data validation and choices, conditional formatting, shapes, notes, form controls, ActiveX, macro source | Input fields, checkboxes, ActiveX. Dependent formulas are recalculated and their results written | Only the written cells and their calculated results changed. Cells that do not depend on them keep their values and formulas (`proof_scope`) |
| Word (.docx, .docm, .doc, Word 6.0/95) | Body, tables, text boxes, headers and footers, footnotes, form fields, content controls, ActiveX, charts, SmartArt, embeds | Fields, empty cells, underlines, writing anchored on text, embedded Excel tables (redrawing the display image too) | Only the written paragraphs changed |
| PDF | Text (including Japanese font mapping tables), form fields, fields inferred from rules and gaps, image descriptions, XFA forms, layers (content not shown on screen by default is kept out of the reading order and fields and shown apart; text and fields that are not printed are flagged) | Form fields and places without fields (incremental save keeps the original bytes; text in places without fields is drawing appended after the original drawing, plus a field holding the value, not an annotation; tagged PDFs also get additions to the structure tree; ○ and check marks carry their meaning as transparent text; PDFs with digital signatures and similar get added fields), XFA data | Not one byte of the original file changes |
| Ichitaro (.jtd) | Body, tables, frames, headers and footers, footnotes, ruby, links, documents saved compressed | Overwriting blanks, masked characters and □; writing into empty cells (writes that change character counts) | Only the written characters and the structure decided for repair changed |
| OpenDocument (.odt, .ods, .odp, .odg), RTF, CSV, TSV, JSON | Per each specification | Per each specification | Not one byte outside the written places changes |
| PowerPoint (.pptx, .pptm) | In slide order: shapes, tables (merged cells), grouped shapes, notes, text on layouts and masters, charts, SmartArt, picture descriptions | Empty table cells, blanks, □, writing anchored on text (same field rules as Word) | Only the slide part holding the written paragraph changed (other parts match down to the compressed bytes; the written paragraph matches outside its text elements) |
| Hangul (.hwpx; most Korean government forms) | Body, tables (merged cells), 누름틀 (click-to-write fields, with guide text), form controls (check, radio, combo, text), text in shapes, headers and footers, footnotes, notes, paragraph numbers, equations, picture descriptions, Hangul macro (JScript) source | 누름틀, form controls, empty table cells, blanks, □, writing anchored on text (same field rules as Word) | Only the body section part holding the written paragraph changed (other parts match down to the compressed bytes; for paragraphs with changed text, Hangul's layout cache hp:linesegarray is removed and Hangul re-lays it out on opening) |
| Hangul (.hwp, HWP 5) | Body, tables, 누름틀, 글상자, headers and footers, footnotes, equations, picture descriptions (distribution and password-protected documents cannot be read: F) | Not writable (writable after saving again as HWPX) | — |
| Excel binary (.xlsb) | Cells, formulas (tokens to text), formats, merges, sheets, names, data validation, notes (shapes and conditional formatting are not read, and the window says so: C) | Only overwrites of cells with values that keep the record size (empty cells cannot be written: write D) | Other parts match byte for byte; in the sheet part, only the value bytes changed |
| Old PowerPoint (.ppt), XPS (.xps, .oxps), EPUB | Text of slides, notes and tables / page text (position order; glyph-index-only text is not read and its count is reported) / body text | Not writable | — |
| tar, .tar.gz, .tgz, .gz | Members as "a.tar!/name" (same as ZIP) | Writes to a member and packs it again | tar: other entries match byte for byte. gz: extracted again and matched byte for byte |
| .tar.bz2, .tar.xz, .bz2, .xz | Members as "a.tar.xz!/name" | Not written (no recompression; extract and export under another name) | Read only |
| LZH (.lzh, .lha; self-extracting .exe without running it), mail (.eml, .msg; attachments are members) | Members as "a.lzh!/name" (-lh0-, -lh4- to -lh7-; -lh1- to -lh3- are not unpacked, and `needs_tool` returns example commands for another tool). Mail also shows sender, recipients, subject and body | Not writable | Checked with CRC-16 (LZH) |
| Excel 2–4, Lotus (.wk1), dBase (.dbf), Access (.mdb, .accdb) | Tables as sheets (window and table extraction). Lotus formulas: saved results only. Access rows show NULL columns as 〔NULL〕, wrap text with leading or trailing spaces in "…", and return all-NULL records as rows | Not writable | — |
| Word 1.x, 2.0, OpenOffice.org 1.x (.sxw, .sxc), iWork (.pages, .numbers, .key) | Text (iWork: the undisclosed format as far as it is known: C. Word fast-save is not read: F) | Not writable | — |
| Text (.txt) | Encoding (decided as for CSV), every line (empty lines and leading spaces kept, "line number｜text") | Blanks, □, writing anchored on text (same field rules as Word; no line breaks written) | Apart from the bytes of the written text, not one byte changes, including encoding, BOM and line endings |
| Image files (JPEG (arithmetic coding, 12-bit, lossless too), PNG, GIF, BMP, TIFF, WebP, JPEG 2000, JBIG2, EMF, WMF, PICT, SVG) | Decoded in-house and shown through `view_page` (page = which image). mm rulers with a resolution record, pixel rulers without. `region` decodes only the overlapping part at original resolution | Not directly writable. `to_pdf` wraps it in a PDF; write to a frame (mm) in that PDF | Proved on the PDF |
| RAR, HEIC, HEIF, AVIF | Not unpacked (format identified only). `needs_tool` in `open_document` returns the reason and example commands for another tool (`sips`, `magick`, `ffmpeg`, `bsdtar`, `unrar` and others, per OS). Open the resulting file again to read it. Read grade E (convert with another tool, then read) | Not writable. Write to the converted or extracted file (it does not go back to the original format) | — |
| Images inside documents (Word, Excel, PowerPoint, ODF, HWPX, .doc, .xls, RTF, PDF annotations) | Listed in the window (id, position, size, description, part name) and handed over through `view_image` (PNG and JPEG as they are, others as PNG; EMF and WMF drawn) | — | — |
| Scanned paper (image-only PDF pages) | Images (CCITT G3 and G4, JPEG, Flate, JBIG2, JPEG 2000 and others) decoded in-house; the rendered page is shown to the agent (`view_page`; JPEG 2000 is decoded at the resolution the screen needs, and the result says so) | Written to a frame (mm) the agent chooses after viewing the image. The frame is drawn and checked before writing | Not one byte of the original file changes |
| ZIP | Members as "a.zip!/inner name", read in their own format | Writes to a member; only that part of the ZIP is replaced | Everything but the written part matches byte for byte |

</details>

Korean **HWPX** (Hangul) is read and written. **HWP 5 (.hwp)** is read only (write grade D: writable after saving again as HWPX).
Distribution (배포용) and password-protected HWP files have encrypted bodies and cannot be read (read grade F).

**Password-protected documents** (Office encryption, old .xls/.doc encryption, PDF encryption) open with the password the applicant gives.
After writing, the file is protected again with the same password. **The password stays out of results and logs.**

**A seal or signature image can stand in for a stamp or signature (PDF).** I cannot stamp or sign by hand.

<details>
<summary>Placing seals and signatures, and the proof</summary>

**Only with the applicant's explicit consent** (nothing is placed without `"consent": true`), the applicant's seal or signature image (PNG, JPEG) goes on
a text anchor (`{"seal": image, "consent": true, "cover": "印"}`) or a page frame (`"page"`, `"box"` in mm; scanned paper too).
The image keeps its aspect ratio inside the frame. An opaque image that would cover page text is not placed.

The proof confirms that what was added is only the stamp annotation with its appearance and image, and that the image matches the applicant's image (JPEG by file bytes, PNG by decoded pixels and transparency).
`view_page` draws the placed image on the page.

Where a registered seal or handwritten signature is required, an image may not be accepted. The agent tells the applicant so (acceptance depends on the form's rules and is not decided by matching words). Word and Excel do not take seals yet (they refuse until a way of placing them with a proof exists).

</details>

**PDF text is written only with glyphs that can be drawn.** Glyphs from the bundled font (Noto Sans JP Regular, SIL OFL 1.1) are subset-embedded, so the text looks the same on machines without the font. A character that cannot be drawn is refused with a reason, never written as a blank.

<details>
<summary>Glyph routes and digital signatures</summary>

For characters in the bundled font (Japanese, Latin, Greek, Cyrillic, Vietnamese and others, 16,732 characters), only the written glyphs are embedded, regardless of word or language (usually a few KB to a few tens of KB added).

Characters outside the bundled font take these routes, per character, and are recorded in `fonts_not_embedded`:

- Latin extensions such as ş ğ ł ř ő use Helvetica (WinAnsi plus the Latin extensions in the standard 14 fonts' glyph tables).
- Hangul, Simplified Chinese and others use non-embedded CID fonts, only where the CMap (Adobe's public mapping tables, `scripts/make_write_cmaps.py`) gives a glyph within the declared character collection's supplement.
- Characters in neither embed only the written glyphs from a TrueType font on this computer (Arial and others; set with `TODOKU_PDF_FONT`; an empty value embeds nothing, the bundled font included), after checking the font's embedding permission.
- In documents declaring PDF/A, the computer's font is tried before non-embedded fonts.
- Arabic, Hebrew, Indic scripts and Thai are shaped by todoku (form selection, reordering, mark positioning) and drawn with glyphs from the computer's font. The check here compares the drawn glyph sequence with todoku's own shaping; it does not check the look itself. Those fields therefore get write grade B, and the agent checks the image with view_page (`check_visually` in the result).

Proof and read-back reread not only the value (/V) but the appearance (/AP) text, confirming that every character has a drawable glyph and is the character meant to show.

**Digitally signed PDFs** are written only within what the signature allows (no writing to locked fields or under certifying signatures that forbid changes; no new fields or ○ under signatures that allow only filling). When a signature exists, the applicant is told the write will show as a "change after signing".

</details>

## Facts from the tool, judgment from the agent

todoku decides only what specifications decide (the rules of ECMA-376, MS-DOC, PDF 32000, Unicode and others) and the guards against damage (proof, read-back, refusing clear overlaps).
What specifications leave open (which text labels a field, whether a place is for writing, how to align) is never decided by word lists or by numbers tuned to a corpus. The window and field list give the facts; the agent reads them and decides.

<details>
<summary>What is reported as fact</summary>

- **Labels** are the printed text, uncut. Table cells report the text to the left and above (`row_label`, `column_label`) and the number of cells between;
  Excel reports nearby text and its distance (`near`: cell count and mm). Text not directly adjacent is not shown as a label, but as "E12 ('氏名', 3 columns to the left)".
- **"?" places** come with the place's text, character count, approximate width, text before and after, and the character count of spaced-out runs (`facts`).
- **Place ids are determined by position.** Inside paragraphs of Word-like window formats (.docx, .doc, .rtf, .odt, .jtd, .hwpx, .txt, .pptx):
  `P<paragraph>.c<characters before>` (5 printed characters before gives `P12.c5`). Inside Excel cells: `Sheet!B3@<position>`. These are not sequence numbers,
  so they stay the same as other places come and go, and after writing to earlier places. PDF uses `p{page}.{structure letter}{left}-{top}` (mm; `t` empty cell,
  `m` row of empty cells, `g` gap within a line, `b` box characters, `u` underline rule), with facts in mm too (`rect_mm`, `width_mm`). Empty cells in a grid
  are grouped by their x and y sequences; space below or above text in a cell shows as "[↓余白 12.0 mm]".
  Old sequence ids in Word-like formats (`P12.1`, `P12.a1`) are accepted only when they certainly point to the same place
  (when the paragraph has exactly one place of that kind).
- **Check values** accept only `true`/`false` and the state values the field declares (PDF state names, XFA values, box characters).
  Words like "yes" or "有" are not reinterpreted (turning the applicant's answer into true/false is the agent's job).
- **The meaning of a written value** is stated by the agent: `{"value": value, "as": "date of birth"}` (not written to the document). In the dialogue list it reads
  "date of birth: 昭和60年4月1日 (p1.t45-112 / nearby text "住所")"; without `as`, "p1.t45-112 (nearby text "住所"): …", so nearby text is never presented
  as the field's name. Dialogue is how the applicant checks, so `as` is always passed for documents the applicant checks (omissions go to `written_without_as`).
- **Inserting before or after text:** Word, Excel and others work as PDF does, with `{"@1": {"before": "年", "occurrence": 2, "value": …}}` and `{"@1": {"after": "氏名", "value": …}}`.
  todoku does not add places by word rules where a paragraph or cell starts with text.
- **Word's ○** defaults to a character border (`w:bdr`; a □ shape, but as a character format it does not shift). Word does not save text positions on the page, so an ellipse shape
  shifts between renderers (only with an explicit `"style": "ellipse"`; the result says which was used).
- **Placement and range** are the agent's too: PDF `align` (left, center, right), `pad_mm`, `box`; for text-anchored writing in Word,
  the `near` range (`near_paras`). The defaults (left-aligned, same line or same passage) are a starting point when nothing is decided.
- **Look after writing:** for PDF, `view_page` draws glyph shapes, lines, images and field appearance as on the page, and the agent looks at what it wrote
  (non-embedded fonts use a substitute, missing glyphs show as gray squares, images that cannot be decoded as gray frames, and the result says so).
  Excel and Word are not rendered, so wrap and overflow estimates (marked as tentative when based on values the document does not state) and the list of written content
  (including unselected choice groups and the count of remaining "?" places) are checked in dialogue.

</details>

## What is verified

### Can everything be read and written accurately (`audit_document`)

1. **Nothing missed in reading:** a reader separate from the window builder picks up every character in the file and confirms each one appears in the window.
2. **Every field is writable:** test values that satisfy data validation are written to all input fields the window shows, at once.
3. **Proof:** without using the writer's internals, both files are reread, confirming only the written places (and their calculated results and appearance) changed.
4. **Read-back:** the written file is reread, confirming the written values appear unchanged.

### Excel: nothing changed except what was written and its calculated results

Writes to .xlsx, .xlsm and .xls pass two checks. The coverage comes back as counts in `proof_scope`.

- **Dependency check:** from the formula text (other sheets, ranges, defined names, tables, shared formulas, array formulas), the cells that depend on the written cells are found without calculating. Any non-dependent cell whose value or formula differs from the original fails the proof.
- **Whole-workbook calculation:** the written file is recalculated in full and compared with its saved values. A difference in a dependent cell fails. A cell found non-dependent that changes with the written value fails as a missed dependency.

For a large workbook, if the full calculation does not fit the time limit (`TODOKU_CALC_SECONDS`), the result comes back as `deferred` (unverified). Pass the exported file to `audit_document`: it calculates the whole workbook in a separate process and returns the result in `written_proof`.

<details>
<summary>Reading <code>proof_scope</code></summary>

- `not_dependent`: cells that do not depend on the written values, and how many matched the original.
- `dependent`: dependent cells, and how many changed.
- `undetermined`: cells whose dependency cannot be decided. Formulas whose references are computed (INDIRECT, OFFSET, range operators and others) and functions that change on every open (TODAY, RAND and others) are counted, not guessed.
- `recalculated`: the full calculation's result. `checked` and `same` (cells compared and matched), `cannot_calculate` (cells not compared because they cannot be computed), `stale` (cells not yet reflecting the written value, such as those with a recalculate-on-open mark).
- The full calculation is separate from the writer's partial calculation and reads every row. It uses the same `formula/` engine, so function implementations are checked separately, in full, against Excel's saved values.
- For large workbooks, the dependency check compares the original and written sheet XML (without the writer's index).
- If `written_proof` has `problems`, the file does not go to the applicant.

</details>

### Grades: read (S–F) and write (A–D)

| Read | Meaning | | Write | Meaning |
|---|---|---|---|---|
| S | The data itself (CSV, JSON) | | A | Writable as it is (fields marked by the creator) |
| A | Readable accurately by rules alone (built-in date formats whose look varies by locale are shown in locale-independent ISO 8601, and the result says so) | | B | Writable after analysis (inferred from layout, "?" in the window judged by the agent, written after choosing an encoding) |
| B | Accurate when an agent reads it (everything is in the window; what is undecided is decided by the agent from the window's evidence: encoding candidates, calculated results saved by software other than Excel, paragraph-number formats that cannot be reproduced) | | C | Entered through a screen (UI) (button macros, e-application screens) |
| C | An agent can convert it with effort (macros, formulas that cannot be verified, window omissions, unreadable as any encoding) | | D | Not writable (no access; signatures, seals, thumbprints, notarization and other manual steps) |
| D | Needs the agent's interpretation | | | |
| E | A format agents are not good at (scans, image-only sheets, documents whose content is images) | | | |
| F | No access | | | |

The read grade is decided by reading content alone (finding input fields belongs to the write grade). If content picked up by the separate reader is missing from the window
(a miss in `audit_document`), the read grade is C, not A or S. When a text-only file reads under several candidate encodings,
the window lists candidates and samples; passing the agent's choice as `encoding` rereads the file in that encoding (the choice stays in the reasons).

Grading uses either "rules certain from the file's structure" or "the agent's judgment after reading the window", never word matching.
The word "押印" (stamp) in a document may be part of "押印不要" (no stamp needed). Things whose meaning matters, such as stamps, handwritten signatures and e-application screens,
are posed by todoku as questions, and the agent answers after reading the window. Grades are not rankings. They show how to handle a document so it gets through.

### Accessibility alerts (separate from grades)

Images without a description (alternative text) do not reach people using screen readers (WCAG 1.1.1).
For each image in PDF, Word, Excel, PowerPoint, OpenDocument and RTF, todoku checks for a description or decorative mark, and reports images without one as alerts with their location (page, paragraph, nearby text).
Agents do not guess at image content. They tell the applicant the location.

## Large documents

Large documents may be read in smaller amounts, but they are never treated as unreadable. When the estimated window exceeds `TODOKU_CONTEXT_BUDGET_CHARS` (400,000 characters, the size that no longer fits a 200,000-token conversation), `open_document` returns an overview (size, pages, sheets, rows, what is where), and the agent requests what it needs by range. Small documents work as before.

<details>
<summary>Reading and writing per format</summary>

| Format | Read | Write |
|---|---|---|
| CSV, TSV | `rows`. The file is not loaded into memory (line-position index). Empty cells by range | Only the written rows' bytes are replaced and copied. Proof: comparison of written rows and hashes of unchanged ranges |
| JSON Lines | `records` (unreadable lines shown as raw text) | Same |
| .txt | `paragraphs`. Very long lines are split by character range and marked (id `P<line>.x<position>`) | Same |
| .xlsx | `sheet`+`rows`. Sheet XML is streamed to build a row index; a small workbook of just the requested rows goes through the window reader. Data validation expands for the requested rows | Only the changed rows' bytes are replaced. Only dependent formulas are recalculated (dependency closure). The full-calculation check is `written_proof` from `audit_document` |
| PDF | `pages`. Only the page tree is read; only requested pages are laid out | Only pages being written are laid out (incremental update) |
| Word (.docx) | `paragraphs`. document.xml is streamed, keeping only an index of paragraph and table positions; a small document of just the requested part goes through the window reader (numbers and ids match the whole) | As usual (reads the whole) |
| .xls | `sheet`+`rows`. Only record heads are followed from BoundSheet. Shared strings are read when used | As usual |
| .doc | `paragraphs`. Only paragraphs in the requested character range are built from the piece table. Streams are memory-mapped | As usual |
| HWP 5 | `paragraphs`. Sections are inflated as a stream and read in small pieces at record boundaries | Not writable (save again as HWPX) |
| Ichitaro, Word 6.0/95, RTF | `paragraphs`. The whole is read, but the window is built only for the requested part (RTF reading uses much memory; use `raw: true` if it stops) | As usual |
| .pptx | `paragraphs`. Only that slide's parts, per slide | As usual |
| .odt, .hwpx | `paragraphs`. A small package is built from the index of top-level body element positions | As usual |
| .ods, .xlsb | `sheet`+`rows`. Row index and small package (repeated rows keep their numbers) | As usual |
| JSON | `path_in` (JSON Pointer) or `cursor`. Tokens are streamed and only requested values read (UTF-8; other encodings read the whole) | As usual |
| Others (HTML, single-XML OpenDocument, and so on) | `part`, `cursor` (character position). `raw: true` (raw text) when building the window hits a limit | As usual |

`raw: true` is the last route: it reads raw text little by little without building a window (no fields found). It also works for compound-file documents (.xls, .doc, HWP, Ichitaro) and .xlsb.

</details>

<details>
<summary>Continuation, limits and records</summary>

- **Continuation:** every result carries the range read (`covered`), the remainder (`remaining`) and how to continue (`next`). Following it to the end reads everything. `list_fields` works the same way.
- **Size of `list_fields` results:** when an MCP result does not fit `TODOKU_RESULT_CHARS` (25,000 characters, sized to Claude Code's limit of 25,000 tokens), the facts needed for writing stay (id, label, type, choices, value, row context, writing rules), evidence (`near`, `evidence`, `why` and others) is dropped, and `omitted` states what was dropped and how to request it. `ids` returns all facts for those fields; `detail: true` returns all fields unabridged, in parts. What still does not fit continues through `next`. Python API and command results do not change.
- **Grades:** low-cost checks run on the whole; remaining checks are listed in `deferred` ("unverified: reading omissions across all pages" and similar). Grades are not decided from samples.
- **Long lines and long cells:** `list_fields` on long .txt lines examines 2,000,000 characters per call and continues with `next`. For a very long CSV cell, the list cuts the value at the first 2000 characters, with `value_chars` and `value_cut` stating the cut (`read_window` has it all).
- **Audit:** `audit_document` is a long job. It is not bound by one call's time, returns progress on each call, and resumes from its record after stopping (records in `~/.cache/todoku/jobs`, changeable with `TODOKU_JOBS_DIR`).
- **Records and personal information:** long-job records (saved `.window` files, audit results, what-if values) contain document content. They live in a private location (0700) and are deleted `TODOKU_JOBS_TTL_HOURS` (24 hours) after the last write. Deleting the original file deletes its records. With `0`, windows are deleted once read and job records once results are returned (unread records after 1 hour; large documents are rebuilt on each open). `todoku forget` deletes everything except running jobs at once. Deletion happens at todoku's next call or server startup, so use `todoku forget` when you want nothing left after you stop using todoku. Passwords are never written to records.
- **Limits:** these stop broken or malicious files. Inflation is measured as a ratio to compressed bytes (`TODOKU_MAX_INFLATE_RATIO`, 200×); the memory limit is 1/4 of physical memory (`TODOKU_MEMORY_LIMIT_MB`). No file is refused for size alone (files over `TODOKU_MAX_FILE_MB` are memory-mapped; PDF incremental updates stream the original bytes too). A stopped call reports a smaller range to request instead (`smaller_request`).
- **Large embedded tables:** Excel workbooks embedded in Word that do not fit the window (`TODOKU_CONTEXT_BUDGET_CHARS`) are not parsed on opening. Passing `sheet: "E1!Sheet"` and `rows` to `read_window` or `list_fields` parses just those rows (following `next` reads everything). Field names for writing do not change. The audit (`audit_document`) reads it all. Documents whose window build exceeds the time limit become long jobs (`jobs.py`) that return progress.
- **Large arrays:** arrays that fit a sheet (`SEQUENCE(1048576,16384)`, `ROW(1:1048576)*COLUMN(A:XFD)`, `EXPAND`, `TEXTSPLIT`, `MAKEARRAY` and others) are held as a "row, column → value" calculation with no cells created, and only requested cells are computed (beyond the sheet, `#NUM!` as in Excel). Aggregates such as sums, counts and maxima are answered by formula for arrays of known shape (polynomials of row and column numbers), and otherwise read in order, bit by bit. When no cell value or partial sum is exactly representable as a double, values are added in Excel's row order (matching real Excel to 15 digits). Sums over arrays beyond 260 million cells (a size Excel cannot hold) return the exact value and say so. Aggregates that do not finish in time report progress and how to continue (`cells` in `try_calculation`: continues from the record in a separate process; `TODOKU_CALC_SECONDS`).
- **Range places:** when data validation, formula references, tables, array-formula ranges or OpenDocument repeats are wide (one declaration over 16,384 cells), they are not listed cell by cell. They keep the same facts as "`S!C5:C40000` (39,996 empty cells, read by formulas, data validation …)" (the window's range places, `range_places` in JSON, `kind: "range"` in `list_fields`). Every cell inside can be read and written by address. `list_fields` with `sheet` and `rows` returns those rows cell by cell (for tables embedded in documents, `sheet` is `E1!Sheet`).
- **Size for the split-read route:** documents with under `TODOKU_LAZY_MIN_KB` (4 MB) of content bytes build the window first and then split it. On the split route too, following to the end gives the same lines and field ids as the whole window; this is checked per format (tests/test_lazy_more.py). Only the "?" mark on cells referenced by formulas on other sheets depends on reading those sheets, so large workbooks may lack it (the cells are writable by address).
- **Rest of .xlsx:** text anchors without `near` search shared strings and rows and then write. ZIP64 and sheets with out-of-order row numbers are readable (writing takes the read-the-whole route; ZIP64 with a 64-bit central directory is read only, and writes are refused with a reason). Macros triggered by value changes run when their cells are decided by the source text alone. If decided at run time, they do not run, and the result says so.
- **PDF text anchors:** without `page`, pages are laid out in order and searched (stopping at the time limit resumes from the record).
- **Table extraction and CLI audit:** `read_table` reads rows little by little and streams to `out_path` (`cursor` continues; `rows` limits the range). `todoku audit` also takes the split-read route for large documents.
- **Access:** opening builds only each table's page table (tables, row counts, column overview). `read_window` and `read_table` with `sheet` and `rows` decode only the pages holding the requested rows. The file is not loaded into memory. Apart from the amount returned per call there is no limit, and every row of every table can be reached.

</details>

## Full audit during development (as of November 3, 2026)

The documents in this section were used in development, so this is not an evaluation (failures were fixed, so results approach 100%). The evaluation is [Results](#results-eval-freeze-2) above.

Forms collected from public pages of national, prefectural and municipal governments in Japan and from abroad (US, UK, Canada, Germany, New Zealand, Australia) went through the same checks as `audit_document` (`scripts/audit_corpus.py`). Failures were not blamed on the form. The general cause on todoku's side was found and fixed.

| Collected forms | Count | Readable and writable, completely and accurately |
|---|---|---|
| Excel (.xlsx, .xlsm): care and disability welfare designation applications, construction licenses, grants, subsidies and others | 196 | 196 |
| Old Excel (.xls) | 46 | 46 |
| PDF | 83 | 83 |
| PDF (XFA forms) | 38 | 38 |
| Word (.docx) | 57 | 57 |
| Old Word (.doc) | 30 | 30 |
| Old-format Word and Excel (MEXT, MHLW) | 64 | 64 |
| Forms with ActiveX controls | 4 | 4 |
| Documents with charts and embeds | 27 | 27 |
| OpenDocument | 36 | 36 |
| RTF | 17 | 17 |
| CSV, JSON | 33 | 33 |
| Ichitaro | 88 | 88 |

- **Matching filled examples:** governments often publish a blank form together with a filled example. The example is saved by Excel, so Excel's own calculated answers remain in it.
  Only the example's inputs are written into the blank form, and todoku's recalculated values are compared with the example's values, all of them.
  On disability welfare pay improvement plans (Wakayama and Niigata prefectures), **9,569/9,569 formulas match**, and checkbox controls match the ones Excel saved **byte for byte**.
- **Other spreadsheet software is not ground truth.** Comparison began against LibreOffice. On a real plan, LibreOffice
  broke formulas containing the sheet name `【参考】数式用` on load and disagreed with Excel's saved values in 3,855 cells. Ground truth is always what Excel saved.
- **Speed:** half of the 1,636 collected documents open within 0.06 seconds. A plan with 50,000 fields takes about 15 seconds.

```sh
uv sync
uv run pytest          # 4,756 tests
uv run python scripts/audit_corpus.py corpus/wide   # batch audit of collected forms (fetch them first)
```

**Python version:** runs on Python 3.11 and later (`requires-python = ">=3.11"`). Of the character classes (`chars.py`),
general category, East Asian width and character names (colons, spaces, underscores, dashes and others) are read at run time from the running Python's `unicodedata`
(version in `chars.UCD_VERSION`). Only data absent from `unicodedata` (bracket pairs in BidiBrackets.txt, sentence terminators in
PropList.txt, scripts in Scripts.txt) is held in small versioned tables (`chars.TABLES_UCD_VERSION`), rebuildable with
`scripts/make_chars.py`.

<details>
<summary>Running the tests</summary>

- **Tests that run on any computer:** documents are built inside the tests or come from committed fictional documents (`tests/fixtures/public/`).
  `tests/fixtures/public/` is produced by `scripts/fixtures/make_public.py` (CC0; no personal information; no font data embedded). It covers
  PDF (AcroForm, flat paper, scan, rotation, encryption, vertical text, right-to-left, Devanagari), xlsx, xls and ods (merges, data validation, formulas,
  rich text, 1904 dates), docx, doc, odt and rtf (content controls, legacy form fields, text boxes, tracked changes),
  pptx, HWPX, Ichitaro, txt, csv (UTF-8, Shift_JIS), json and ZIP, along the path from opening a file (window, fields, writing, audit, grades)
  (`tests/test_public_fixtures.py`; ground truth is the text the generator put in, stored in `MANIFEST.json`).
- **Mark `corpus`:** tests using collected public forms and specifications (`corpus/`, not redistributed, so not in git). Without `corpus/`,
  they are skipped together and the end reports "skipped corpus N". On a computer with `corpus/`, try CI's conditions with
  `TODOKU_NO_CORPUS=1 uv run pytest`; run only the corpus tests with `uv run pytest -m corpus`.
- **Mark `fonts`:** tests using the computer's fonts (shaping, embedding) or `hb-shape` (HarfBuzz). Skipped with a reason when absent.
  On Linux, `fonts-noto-core`, `fonts-dejavu-core` and `libharfbuzz-bin` enable them (fonts can also be set with `TODOKU_PDF_FONT`,
  `TODOKU_PDF_FONT_ARABIC` and others).
- **CI** (`.github/workflows/tests.yml`): macOS, Linux and Windows × Python 3.11, 3.12 and 3.13 run
  `uv run pytest -q`, and check that the MCP server starts as `.mcp.json` and the plugin start it and answers `initialize` and `tools/list`
  (`.github/scripts/mcp_smoke.py`). The count breakdown appears in the job Summary.

</details>

## Sources of documents and specifications

Collected documents and specifications are **not in this repository** (`corpus/` is in `.gitignore`). Source URLs, listing pages and SHA-256 are recorded in
[`sources/`](../sources/), so the same files can be fetched again. Fetching followed robots.txt, one file at a time with intervals.
No bot protection was circumvented. The XFA 3.3 specification came from an Internet Archive copy (not redistributed).

| Documents | Source |
|---|---|
| Pay improvement plans, performance reports and filled examples | [Wakayama Prefecture](https://www.pref.wakayama.lg.jp/prefg/040400/shitei_jigyosyo/syogukaizen.html), [Niigata Prefecture](https://www.pref.niigata.lg.jp/sec/shougaifukushi/r8syougaisyogukaizen.html) (`tests/forms.json`) |
| Forms for the generality audit (Excel, PDF, Word, old formats, Ichitaro, OpenDocument, RTF, CSV) | MHLW, MEXT, Legal Affairs Bureaus, MIC, prefectures and municipalities, GOV.UK, Immigration Canada, USCIS, IRS and others (`sources/wide*.tsv`, `sources/wide_*.txt`, `tests/sources.json`) |
| Random sample of public Excel files | [data.go.jp](https://data.e-gov.go.jp/) CKAN API (`sources/data.e-gov.go.jp.tsv`) |
| Pairs of blank forms and filled examples | Each municipality's listing page (`sources/pairs.tsv`) |
| World survey documents | Common Crawl archives and government sites of each country (`world/sample.jsonl`) |
| Forms from Tokyo's 23 wards and Tama (the 20-form comparison, furigana, .doc measurement) | Each ward's or city's listing page (`sources/tokyo23.tsv`, `sources/tama.tsv`) |
| World forms (development and blank-form filling) | Government sites of each country (`sources/forms_world.tsv`) |
| Re-evaluation documents | Common Crawl archive (`sources/eval_freeze2/world.tsv`); Fukushima, Akita, Fukuoka, Kashiwa (`sources/eval_freeze2/jp_forms.tsv`) |
| Forms shown in the video (values are fictional) | [Katsushika Ward assistive device application form](https://www.city.katsushika.lg.jp/kenkou/1000053/1002246/1002304.html), [Hino City daily living equipment application form](https://www.city.hino.lg.jp/fukushi/nanbyo/service/josei/1003716.html) |
| Listing pages for unseen Japanese forms | Fukushima Prefecture, Akita City, Fukuoka City, Kashiwa City (`sources/jp_forms_pages.tsv`) |
| Encrypted Office documents (with ground truth) | Test files from [msoffcrypto-tool](https://github.com/nolze/msoffcrypto-tool) (MIT) |
| Encrypted PDFs (with ground truth) | Test files from [pypdf](https://github.com/py-pdf/pypdf) (BSD 3-Clause) and [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0) (`sources/crypto__pdf.txt`) |
| XFA forms | pdf.js test files (Apache-2.0) and public forms of the Canadian and US governments (`sources/wide_xfa.txt`) |
| Embed and chart examples | Test files from [Apache POI](https://github.com/apache/poi) (Apache-2.0) and public files from MIC, the Ministry of the Environment and BODIK (`sources/wide_embed.txt`) |
| Real Access databases | Test files from [Jackcess](https://github.com/jahlborn/jackcess) (Apache-2.0) and Apache Tika, mdbtestdata and others (not redistributed) |
| File format specifications | Microsoft Open Specifications (MS-XLS, MS-XLSX, MS-DOC, MS-OFFCRYPTO, MS-OVBA, MS-OFORMS, MS-ODRAW, MS-OLEDS, MS-CFB, MS-EMF, MS-WMF, MS-OI29500 and others), ISO 32000 (PDF), XFA 3.3, OpenDocument 1.3, OWPML (KS X 6101, HWPX) and "한글 문서 파일 형식 5.0" (Hancom's public document), OpenOffice.org's guide to the Excel format, a guide to the Word 6 format (`sources/specs.txt`, `scripts/fetch_specs.py`) |

The code of these projects is not used. Only their files are used, as test ground truth.
Documents in `tests/fixtures/` are fictional, made by this repository's scripts and for its tests (`THIRD_PARTY_NOTICES.md`). **They contain no personal information.**

## Layout

```
src/todoku/
  api.py          Entry points (window, fields, writing with proof and read-back, audit, grades)
  mcp_server.py   MCP server (JSON-RPC over stdio; no external libraries)
  grade.py        Grades (read S–F, write A–D) and accessibility alerts
  formula/        Formula parser and calculation engine (Excel's rules: 15 significant digits, the 1900 leap year, Japanese era formats and others)
  xlsx/ xls/      Excel (window, writing, proof, dependency checks, audit, matching filled examples, structural lint)
  xlsx/reach.py   How far a written value reaches (dependencies), and the proof_scope checks
  docx/ doc/      Word
  pdf/            PDF (syntax, encryption, fonts, fields inferred from layout, incremental save, XFA, image descriptions)
  jtd/            Ichitaro (Ssmg container, LHA compression)
  odf/ rtf/ datafile/   OpenDocument, RTF, CSV, JSON
  pptx/ txt/      PowerPoint and text-only files (window and field rules shared with Word)
  hwpx/           Hangul (HWPX read and write, HWP 5 read only; window and field rules shared with Word)
  xlsb/           Excel binary workbooks (BIFF12 records; workbook shape shared with xlsx)
  legacy/         Read-only old formats (.ppt, XPS, EPUB, LZH, mail, Excel 2–4, Word 2, Lotus, dBase, Access, .sxw, iWork), and formats that return steps for another tool (RAR, HEIC, AVIF)
  imaging/        Image decoders (JPEG, JPEG 2000, JBIG2, WebP and others)
  lazy/           The split-read route for large documents
  vba.py vbarun/  Macro source extraction and a VBA interpreter (runs button procedures without Excel)
  processing/     Processing extraction (extract_processing): folding formulas into rules, structure, static macro reading, PDF calculation, Python module export and verification
  activex.py ax*.py     ActiveX controls (MS-OFORMS)
  embed.py embeds/      Writing to embedded tables; reading charts, SmartArt and embeds
  cfb.py zipraw.py      Byte-level reading and writing of compound files and ZIP
  officecrypto.py bincrypto.py unlock.py   Password-protected documents
scripts/          Form fetching, batch audits, source export (fixtures/ builds fictional test documents)
sources/          Sources of collected documents and specifications
tests/fixtures/public/   Committed fictional documents (each under 100 KB) and MANIFEST.json
world/            World survey data (built and read by scripts/world_*.py; committed because the portal and reproduction need it.
                  sample.jsonl about 11 MB, results.jsonl about 10 MB, domains.json about 2 MB, the rest about 5 MB together.
                  No diffs, via .gitattributes)
```
