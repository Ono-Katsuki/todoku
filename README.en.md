# todoku

**Claude Code can already fill in a government form. todoku proves not a single byte moved.**

An MCP server that lets AI agents read and write administrative documents accurately. `dependencies = []`. No Office. No LibreOffice.

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

In Claude Code (requires [uv](https://docs.astral.sh/uv/)):

```text
/plugin marketplace add Ono-Katsuki/todoku
/plugin install todoku@todoku
```

Then just ask:

```text
Fill in this form with my details.
Explain how this Excel sheet calculates the amount. What if I choose category II?
```

## "Filled in" is not "filled in correctly"

Same form, same prompt, same Claude. 20 public forms, fictional values. Details in `scripts/eval/` and `world/portal/`.

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

## Everything is implemented from scratch

- **File formats:** PDF (syntax, encryption, fonts, XFA, incremental save), Excel (.xlsx, .xls, .xlsb), Word (.docx, .doc), Ichitaro, HWP/HWPX, ODF, RTF, PowerPoint, CSV, JSON, ZIP and more. Formats with public specifications are implemented from the specifications.
- **Formula engine:** computes by Excel's rules (15 significant digits, the 1900 leap-year bug, Japanese era formats). Runs VBA.
- **Deterministic:** no AI in reading, writing, proving or auditing. Places that rules cannot decide are marked `?`. The tool provides facts; the agent makes the judgment.
- **14 MCP tools:** `open_document`, `read_window`, `list_fields`, `fill_document`, `audit_document`, `extract_processing`, `try_calculation`, `run_macro` and others. Large documents are read by range, and nothing is left unreachable.

About 140,000 lines, 298 files, over 4,000 tests.

## Measured on real documents

Version frozen (`eval-freeze-1`), then measured on documents not opened before the freeze. Baselines include MarkItDown, openpyxl, python-docx and pdfplumber, taking the best per format.

| | todoku | Baselines |
|---|---|---|
| World government documents opened | **552/557** | 494/557 |
| Furigana (Tokyo forms) | **698/713** | 77/713 |
| Text in shapes and text boxes (world) | **389/487** | 220/487 |
| Character recall on legacy Word .doc (Tokyo forms) | **1.0** | 0.0 |

We surveyed 8,500 documents randomly sampled from government sites in 185 countries and estimated from 1,079 of them, actually read and written. **An estimated 1.82 million documents can be written accurately.** All categories, intervals and methods are in `world/portal/`; the evaluation code is in `scripts/eval/`.

## Who built this

> When an agent cannot access something, that is an accessibility problem.

I am a bedridden developer, business owner and researcher. I cannot speak. I cannot use my hands. My only input is a single head switch. Writing 1,000 characters takes me over an hour.

I built this repository in one week, by talking to Claude through that switch. Real documents expose a bug, another agent verifies it, it gets fixed. Over and over.

> With Claude, even someone bedridden can become a developer.

Paper: Katsuki Ono, [*Reconfiguration of Ability: An Autoethnography of LLM-Mediated Development by a Developer with Severe Physical Disability*](https://doi.org/10.1145/3797867.3828990), ASSETS '26

Built with Claude: Japan 2026 entry (development period: November 2–8, 2026). Further documentation is in the [Japanese README](README.md).
