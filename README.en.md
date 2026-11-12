# todoku

**Claude Code can fill in a government form. But you cannot trust it with one. todoku proves nothing but what it wrote has changed.**

Existing tools could not be trusted with government forms. I struggled with them for months, so I built todoku. todoku reads back what it wrote and proves, document by document, that only the filled places and their calculated results changed.

An MCP server that lets AI agents read and write administrative documents accurately. `dependencies = []`. No Office. No LibreOffice.

[![Watch the demo video on YouTube (3 min, Japanese and English captions). Thumbnail: a world map with the text "The world's government documents, to agents. Measured in 185 countries and regions; about 2.38 million documents written accurately"](https://img.youtube.com/vi/zB95Q_jUq5A/maxresdefault.jpg)](https://youtu.be/zB95Q_jUq5A)

**[Demo video (3 min, Japanese and English captions)](https://youtu.be/zB95Q_jUq5A)** · **[Portal: world survey and full results](https://claude.ai/artifact/45P2GdSFYWkAqWrdLebmSH)** 

```text
Proof: the original 110,831 bytes are unchanged from the start of the file; 18,121 bytes were appended.
```

```console
$ head -c 110831 hino_yogu_filled.pdf | cmp - hino_yogu.pdf && echo identical
identical
```

You don't have to trust the agent's report. Check it with `cmp`.

[日本語](README.md)

## Install in 30 seconds

In Claude Code (requires [uv](https://docs.astral.sh/uv/); supported on macOS; Windows and Linux in beta):

```text
/plugin marketplace add Ono-Katsuki/todoku
/plugin install todoku@todoku
```

To move to the newest version, run `claude plugin marketplace update todoku` and `claude plugin update todoku@todoku` in a terminal, then restart Claude Code.

Then just ask:

```text
Fill in this form with my details.
Explain how this Excel sheet calculates the amount. What if I choose category II?
```

## "Filled in" is not "filled in correctly"

Same form, same request, same Claude. 20 public forms, fictional values. Details in the [portal](https://claude.ai/artifact/45P2GdSFYWkAqWrdLebmSH).

| | Plain Claude Code | Claude Code + todoku |
|---|---|---|
| Original PDF bytes kept | 0/16 | **16/16** |
| Proof that only the written places changed | none | **20/20** |

One PDF looked right in Poppler. Opened in Apple PDFKit, the name field read "16". The agent had reported it as complete.

todoku saves PDFs as incremental updates and never touches the original bytes. After writing, it reads the values back, renders the page, and checks that the pixels changed.

## Data and processing, handed over separately

Give an agent only the values and the calculation becomes a guess. `extract_processing` hands over the calculation rules themselves.

```text
Niigata Prefecture care-worker pay improvement plan (public Excel form)
105,256 formula cells → 203 rules
Run in a separate process: 105,256 / 105,256 cells match the saved values

ROUNDDOWN(ROUND(L × P, 0), 0) × AA
```

Guess the order and you are off by one yen.

| Order of calculation (fictional values) | Estimate |
|---|---|
| **As the form defines it** | **¥84,692** |
| Multiply everything, then round down | ¥84,691 |
| Round down first, then multiply by months | ¥84,690 |

With `format: "python"`, the rules are exported as `calculate(inputs) -> dict` that runs on the standard library alone.

## Zero dependencies, all in Python

- **File formats:** PDF (syntax, encryption, fonts, XFA, incremental save), Excel (.xlsx, .xls, .xlsb), Word (.docx, .doc), Ichitaro, HWP/HWPX, ODF, RTF, PowerPoint, CSV, JSON, ZIP and more. Formats with public specifications are implemented from them; some parts were rewritten in Python after reading open-source implementations (`THIRD_PARTY_NOTICES.md`).
- **Formula engine:** computes by Excel's rules (15 significant digits, the 1900 leap-year bug, Japanese era formats). Runs VBA.
- **Deterministic:** no AI in reading, writing, proving or auditing. Places that rules cannot decide are marked `?`. The tool provides facts; the agent makes the judgment.
- **14 MCP tools:** `open_document`, `read_window`, `list_fields`, `fill_document`, `audit_document`, `extract_processing`, `try_calculation`, `run_macro` and others. Large documents are read by range, in turn; what cannot be read is reported.

About 150,000 lines, 309 files, 5,140 tests.

## Measured on real documents

Version frozen (`eval-freeze-2`), then measured by actually reading and writing 935 documents from government sites in 185 countries and regions. 878 of them are a new sample nobody had opened before the freeze.

**Among the government-site documents collected by Common Crawl, an estimated 2.45 million can be read accurately and an estimated 2.38 million can be written accurately.** All categories, intervals and methods are in the [portal](https://claude.ai/artifact/45P2GdSFYWkAqWrdLebmSH); the evaluation code is in `scripts/eval/`.

- Read accurately: **about 2.45 million** (95% interval 2.25–2.64 million)
- Written accurately: **about 2.38 million** (95% interval 2.17–2.57 million)
- 46 Japanese forms seen for the first time, filled in Japanese: **correct in 45, with proof and read-back passing on all 46**

## Who built this

> When an agent cannot access something, that is an accessibility problem.

I am a bedridden developer, business owner and researcher. I cannot speak. I cannot use my hands. My only input is a single head switch. Writing 1,000 characters takes me over an hour.

I built this repository in one week, by talking to Claude through that switch. Real documents expose a bug, another agent verifies it, it gets fixed. Over and over.

Our office handles a great many applications. From now on, my staff and I will fill them in with todoku.

| How Claude is used | For what |
|---|---|
| Claude Code (CLI) | The center of development: design, implementation and verification, by conversation through a head switch |
| Remote Control | Driving Claude Code on the Mac from an iPad, lying in bed |
| Subagents (worktrees) | Finding bugs on real documents, fixing them and having another agent verify, in parallel |
| MCP server (14 tools) | todoku itself: the entry point where agents read, write and prove documents |
| Skill `shinsei` | Hands agents the way to find and fill in application forms |
| Plugin (marketplace) | Installs the MCP server and the skill at once |
| Claude.ai Artifact | Publishes the world survey results as the portal |

> With Claude, even someone bedridden can become a developer.

Paper: Katsuki Ono, [*Reconfiguration of Ability: An Autoethnography of LLM-Mediated Development by a Developer with Severe Physical Disability*](https://doi.org/10.1145/3797867.3828990), ASSETS '26

Built with Claude: Japan 2026 entry (development period: November 2–8, 2026).

---

## More

[Detailed report](docs/REPORT.en.md): MCP usage, each tool, supported formats, how to read the proof and the grades, and the measurement conditions.

## Sources of documents and specifications

Collected documents and specifications are **not in this repository** (`corpus/` is in `.gitignore`). Source URLs, listing pages and SHA-256 are recorded in
[`sources/`](sources/), so the same files can be fetched again. Fetching followed robots.txt, one file at a time with intervals.
No bot protection was circumvented. The XFA 3.3 specification came from an Internet Archive copy (not redistributed).

The per-document list is in the [detailed report](docs/REPORT.en.md).

## License

Apache License 2.0 (`LICENSE`, `NOTICE`). Free to use, for individuals and companies. Third-party data inside (the bundled font Noto Sans JP (SIL OFL 1.1), Adobe's character tables, Unicode tables, RFC 6386 tables, map data and others) follows its own license (`THIRD_PARTY_NOTICES.md`, `LICENSES/`).

본 제품은 한글과컴퓨터의 한/글 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.

Microsoft, Excel, Word, PowerPoint, Access, Ichitaro, Hangul, Hancom, RAR, Adobe, PDF and others are trademarks or registered trademarks of their respective owners. todoku is not affiliated with or endorsed by these companies.
