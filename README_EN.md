# format-agent-skill

**[中文](README.md) | English**

An Agent Skill for reformatting Chinese and English Word documents from plain-language rules, built-in style packs, or a reference DOCX template. Document structure is interpreted first; deterministic OOXML code then produces an editable, auditable DOCX.

> This repository is the skill. `SKILL.md` lives at the root; clone or download the complete repository to install it.

The skill is intended for documents dominated by paragraphs and conventional tables. It is not a general desktop-publishing, OCR, or compliance-certification system. An official template or written rule from the relevant institution, court, client, or publisher always takes precedence over a built-in baseline.

## Install

Ask a compatible Agent:

```text
Install this skill: https://github.com/KaguraNanaga/format-agent-skill
```

Or clone it manually:

```bash
gh repo clone KaguraNanaga/format-agent-skill
pip install -r format-agent-skill/requirements.txt
```

The host must be able to read local files and run Python. Dependency installation and Word/WPS automation may require additional permission. Skill installation alone does not make every advanced feature available on every Agent host.

## Quick start

Run preflight on an unfamiliar document first:

```bash
python main.py --preflight-only --target source.docx --out output/result.docx
```

Then choose exactly one formatting source:

```bash
# Plain-language rules
python main.py --spec rules.txt --target source.docx --out output/result.docx

# Reference Word template
python main.py --template template.docx --target source.docx --out output/result.docx

# Built-in baseline
python main.py --style-pack apa7-student --target paper.docx --out output/paper.docx

# No model
python main.py --spec-json examples/spec_std.json \
  --rolemap-json examples/rolemap_std.json \
  --target examples/messy.docx --out output/result.docx
```

Use `--verify` only when visual QA is requested; it also needs a working Word/WPS/LibreOffice renderer and an image-capable model. On Windows with Microsoft Word, `--refresh-fields` can refresh local allowlisted fields such as TOC, PAGE, REF, SEQ, TA, and TOA.

See [CLI and outputs](references/cli-and-output.md) for commands, exit codes, and troubleshooting.

## Design and outputs

The formatting path accepts two model-generated contracts—`FormatSpec` and `RoleMap`. The model never edits the DOCX directly. Optional visual QA returns an issue list; fixes are still applied by code.

Outputs include:

- `result.docx`: the primary editable document.
- `result_tracked.docx`: a property-change copy covering `pPrChange/rPrChange` only.
- `result_report.docx` / `.md`: structural and paragraph-level change reports.
- `result_formatspec.json`, `_rolemap.json`, and `_stylemap.json`: intermediate evidence.
- `result_preflight.json`: stories, fields, sections, objects, warnings, and blockers.

This architecture reduces the risk of model-driven document corruption; it does not guarantee that every role is classified correctly. Semantic content is preserved, while audited structural transformations—such as converting manual numbering or creating an explicitly requested TOA heading—may change visible text. Any other text difference fails integrity validation and prevents replacement of the main output.

The tracked copy does not make page setup, sections, headers/footers, tables, cover pages, directories, or structural moves individually acceptable/rejectable. Review those changes in the report.

## Document-type coverage

| Document type | Status | Boundary |
|---|---|---|
| Chinese speeches, study materials, briefings, and ordinary reports | Supported | Paragraph, numbering, page, header/footer, and ordinary-table formatting; complex floating layouts are preserved only |
| Chinese notices, requests, reports, letters, and minutes | Conditional | `official-cn-gbt9704` is a baseline, not full GB/T 9704 compliance; it does not generate a complete red heading, seal, urgency/security block, or imprint |
| Chinese theses and dissertations | Conditional | Abstract/keywords, headings, cover fields, TOC, sections, and dynamic headers/page numbers; complex covers, declarations, English title pages, spines, and advanced numbering need manual review |
| English memos, essays, reports, and academic papers | Supported/conditional | English roles, Letter/A4, running heads, block quotes, and hanging references; no citation, DOI, bibliography-data, or language-quality validation |
| APA 7, MLA 9, Chicago 18, Turabian 9 | Conditional | Auditable layout baselines only; not full compliance checkers and do not create missing metadata or reconcile citations |
| IEEE journal manuscripts | Conditional/high risk | Conservative two-column baseline; complex spanning objects, author/affiliation blocks, and reference renumbering require the target publication template |
| Basic technical manuals | Conditional | Code/commands, steps, and WARNING/CAUTION/NOTE/TIP boxes; no DTP reconstruction, floating-object movement, screenshot interpretation, or automatic caption creation |
| US legal briefs and TA/TOA | Conditional | Existing TA/TOA are preserved; new fields require exact user-provided marks; no citation discovery, Bluebook/local-rule validation, or line-number generation |
| Contracts and legal opinions | Partial | Ordinary Article/Section hierarchy only; signing pages, complex exhibits, content controls, fill logic, and jurisdictional compliance remain outside scope |
| Bid documents, annual/financial/audit reports, books, long manuals | Partial | Conservative paragraph/table/section migration; wide-table systems, complex spanning art, chapter openers, indexes, and multi-part appendices need manual work |
| Arabic/Hebrew and other complex-script or RTL documents | Partial, not specifically validated | Template transfer can carry `font_cs`, language, and RTL/Bidi properties, but there is no dedicated profile or style pack |
| Résumés, brochures, newsletters, posters | Unsupported | Text-box, icon, column, floating, and absolute-position layouts are outside the paragraph engine |
| Complex forms and fillable contract templates | Unsupported | The skill does not create or alter content controls, checkboxes, field logic, protection, or fill workflows |

See [style packs](references/style-packs.md), [thesis mode](references/thesis-mode.md), and [advanced modes](references/advanced-modes.md) for detailed behavior.

## Word feature boundaries

- Paragraph styles, pagination rules, manual numbering, and native Word numbering are supported.
- Page size, orientation, margins, columns, page-number formats, and first/odd/even headers and footers can be transferred. Rebuilding an existing complex section structure or creating landscape table sections requires explicit `--allow-risky-structure` authorization.
- Conventional table geometry is supported. Complex nested/floating tables, calculation logic, and large financial tables require manual inspection.
- Footnotes and endnotes are scanned and preserved; they are formatted only when explicit `notes` rules exist. Their content is never rewritten.
- Caption conversion is limited to recognized single-paragraph captions with simple integer numbers. Chapter numbering, numbers split across runs, advanced equation numbering, and complex cross-references are not guessed.
- Citation, bibliography, index, TA, and TOA fields are preserved and only safe local fields are refreshed. Citation content and source metadata are not corrected.
- Text boxes, shapes, comments, content controls, revision stories, formulas, embedded objects, floating images, and SVG are inventoried and preserved, not semantically reformatted or repositioned.
- `strict` cleanup can remove meaningful character styles, emphasis, or foreign-language variables; use `preserve_emphasis` for mixed-language or emphasis-heavy documents.
- Unaccepted revisions, editing protection, digital signatures, macros, and `altChunk` block output by default.
- Visual QA is optional and limited to one repair pass. If rendering or the vision model fails, the skill cannot claim visual verification.

## Input formats and environment

All final output is `.docx`. Legacy inputs are converted to a temporary DOCX and then pass through full preflight and integrity validation; pagination, fonts, fields, and floating objects may change during conversion.

| Input | Status | Requirement / validation boundary |
|---|---|---|
| `.docx` | Native, validated | Standard OOXML; safety blockers still apply |
| `.doc` | Conditional, validated | Windows + pywin32 + Word or WPS; tested with a real OLE/CFBF binary DOC |
| `.rtf` | Conditional, validated | LibreOffice preferred; compatible Word/WPS may be used; conversion is lossy |
| `.wps` | Conditional, partially validated | OOXML-backed `.wps` routing is tested; legacy proprietary binary WPS is not yet covered by a real fixture |
| `.odt` | Conditional, failed in the current acceptance host | LibreOffice is recommended; conversion fails explicitly when no usable import filter exists |
| `.pdf` / scanned input | Unsupported | No OCR or PDF layout reconstruction |
| `.docm` | Unsupported | Macros are not executed or migrated |
| Encrypted, password-protected, or damaged files | Unsupported | No password bypass or package repair |

Runtime notes:

- Python 3.10+; use an isolated environment for `requirements.txt`.
- `.doc/.wps` import and Windows rendering require pywin32 plus Word/WPS; LibreOffice is recommended for `.odt/.rtf`.
- Every Word/WPS candidate runs in an isolated process. The default timeout is 45 seconds, after which the next candidate is tried; only Office processes started by that attempt are cleaned up. Set `FORMAT_AGENT_COM_TIMEOUT_SECONDS` between 5 and 300 when needed.
- `--refresh-fields` still requires Windows + Microsoft Word. If Word is unavailable, the document is marked to update fields when opened.
- Natural-language rule extraction and semantic labeling need a model. `--spec-json` plus `--rolemap-json` bypasses the model completely.

## Real-environment acceptance

Acceptance completed on 2026-09-02 on Windows with Microsoft Word and WPS:

- Native DOCX, a real binary DOC, a real RTF, and the OOXML-backed `.wps` route passed end to end.
- Ten rendered pages across five documents were visually inspected with no missing glyphs, clipping, overlap, table movement, or pagination drift.
- Regression suite: `42 passed`.
- ODT did not pass on that host because LibreOffice was absent and the available Office import filters were unusable.
- A real legacy proprietary binary `.wps` sample is still required.

The current result is therefore **conditionally accepted**, not an unconditional compatibility claim for every Word genre or Office environment.

## Example

| Before | After |
|---|---|
| ![before](docs/images/en-board-before.png) | ![after](docs/images/en-board-after.png) |

## Related

- Archived hackathon edition with GUI and case gallery: [format-agent](https://github.com/KaguraNanaga/format-agent)
