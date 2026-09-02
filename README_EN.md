# format-agent-skill

**[中文](README.md) | English**

Give an Agent a formatting standard and a Word draft. The skill identifies document structure, transfers styles and page rules, checks the result, and produces an editable, auditable DOCX.

The formatting source can be:

- plain-language rules;
- a well-formatted Word template;
- a built-in APA, MLA, IEEE, Chicago, Turabian, Chinese official-document, technical-manual, or legal-brief style pack;
- hand-prepared `FormatSpec` and `RoleMap` JSON.

> This repository is the skill. `SKILL.md` lives at the root; clone or download the full repository to install it.

## What it does

- Formats Chinese and English headings, body text, numbering, fonts, spacing, indentation, and pagination.
- Transfers page size, orientation, margins, columns, page numbering, and first/odd/even headers and footers.
- Recognizes both manual numbering and native Word numbering.
- Formats conventional tables, including widths, cell margins, repeating headers, row splitting, and vertical alignment.
- Scans body stories, nested tables, headers/footers, notes, comments, text boxes, fields, bookmarks, content controls, revisions, and embedded objects.
- Produces a primary document, a property-change copy, change reports, preflight results, and structured evidence.
- Optionally renders the result for a visual-model QA pass.
- Accepts `.doc/.wps/.odt/.rtf` through conversion and always produces `.docx`.

## Documents it handles

| Scenario | Available capability |
|---|---|
| Speeches, study materials, briefings, and ordinary reports | Heading hierarchy, body styles, numbering, pages, headers/footers, and conventional tables |
| Chinese notices, requests, reports, letters, and minutes | Reads agency templates and includes an `official-cn-gbt9704` formatting baseline |
| Chinese theses and dissertations | Abstract/keywords, headings, references, cover fields, TOC, sections, dynamic headers, and page numbers |
| English memos, essays, reports, and academic papers | Abstract, Keywords, Chapter/Section, block quotes, running heads, captions, References, Bibliography, and Works Cited |
| APA 7, MLA 9, IEEE, Chicago 18, Turabian 9 | Auditable built-in style packs, or direct transfer from a school/course/journal template |
| Technical manuals | Code and commands, steps, WARNING/CAUTION/NOTE/TIP blocks, and pagination binding between figures and captions |
| US legal briefs | Court captions, case numbers, brief titles, TOC/TOA, attorney details, certificates, and explicitly configured TA/TOA fields |
| Contracts, legal opinions, bids, and long reports | Article/Section hierarchy, conventional tables, section-aware pages, and attachment headings |

An official template from the relevant institution, court, client, or publisher takes precedence over a general style pack.

## Example

| Before | After |
|---|---|
| ![before](docs/images/en-board-before.png) | ![after](docs/images/en-board-after.png) |

## Install

Ask a Skill-capable Agent:

```text
Install this skill: https://github.com/KaguraNanaga/format-agent-skill
```

Or install manually:

```bash
gh repo clone KaguraNanaga/format-agent-skill
pip install -r format-agent-skill/requirements.txt
```

The host needs local-file access and Python execution. Word, WPS, or LibreOffice is used only when a task needs legacy conversion, field refresh, or page rendering.

## Quick start

Preflight an unfamiliar document first:

```bash
python main.py --preflight-only --target source.docx --out output/result.docx
```

Then choose one formatting source:

```bash
# Plain-language rules
python main.py --spec rules.txt --target source.docx --out output/result.docx

# Reference Word template
python main.py --template template.docx --target source.docx --out output/result.docx

# Built-in style pack
python main.py --style-pack apa7-student --target paper.docx --out output/paper.docx

# No model
python main.py --spec-json examples/spec_std.json \
  --rolemap-json examples/rolemap_std.json \
  --target examples/messy.docx --out output/result.docx
```

Add `--verify` for visual QA. On Windows with Microsoft Word, add `--refresh-fields` to refresh safe local fields such as TOC, PAGE, REF, SEQ, TA, and TOA. See [CLI and outputs](references/cli-and-output.md) for all commands and exit codes.

## How it works

The model understands; deterministic code executes. The formatting path uses two JSON contracts:

```text
Formatting source (rules / style pack / Word template)
        │
        ▼
   FormatSpec (formatting rules) ◄── schema validation
        │
Target document ── story and structure extraction ──► paragraph list
        │                                               │
        │       deterministic rules + semantic fallback ▼
        │                                          RoleMap
        ▼                                               │
  Deterministic executor (named styles / OOXML) ◄───────┘
        │
        ▼
Primary DOCX + property-change copy + reports + JSON evidence
```

- The model does not edit the DOCX directly.
- Full-story and complex-structure preflight runs before formatting.
- Primary paragraph formatting uses named Word styles for continued editing.
- Candidate files are committed atomically only after text-integrity validation.
- Necessary structural text changes, such as manual-number conversion, are allowlisted and reported.

## Outputs

| Output | Purpose |
|---|---|
| `result.docx` | Editable primary document |
| `result_tracked.docx` | Paragraph and character property changes |
| `result_report.docx` / `.md` | Page, structure, and paragraph change report |
| `result_formatspec.json` / `_rolemap.json` / `_stylemap.json` | The Agent's interpretation of the rules and document |
| `result_preflight.json` | Stories, sections, objects, warnings, and blockers |
| `result_issues.json` | Created when visual QA finds issues |

## Inputs and runtime

- `.docx` is handled natively.
- `.doc/.wps` can be converted through local Word/WPS.
- `.odt/.rtf` prefers LibreOffice and can try compatible Word/WPS import filters.
- Final output is always `.docx`.
- Python 3.10+; an isolated environment is recommended for `requirements.txt`.
- Word/WPS conversion and rendering run in isolated, bounded processes so one unresponsive Office service does not stall the task.
- `--spec-json` plus `--rolemap-json` bypasses models; only `--verify` requires an image-capable model.

## Scope notes

The current engine is optimized for Word documents built mainly from paragraphs and conventional tables. Layout-heavy résumés, brochures, posters, and complex fillable forms are preserved and preflighted rather than rebuilt as desktop-publishing layouts. PDF and scanned-page reconstruction are separate workflows.

Citations, bibliographic data, legal authorities, and index entries are preserved rather than rewritten. Complex floating objects, content controls, and embedded objects are also handled conservatively. Large section rebuilds, two-column layouts, and landscape tables are preflighted before explicit structural changes are allowed.

Legacy conversion depends on locally installed Office components and may introduce font or pagination differences; LibreOffice is recommended for ODT. Detailed guidance:

- [Thesis mode](references/thesis-mode.md)
- [Style packs](references/style-packs.md)
- [Official, technical, legal, and legacy modes](references/advanced-modes.md)

## Real-environment validation

End-to-end validation completed on 2026-09-02 on Windows with Microsoft Word and WPS:

- Native DOCX, a real binary DOC, a real RTF, and the WPS route completed conversion, formatting, and text-integrity checks.
- Ten pages across five documents were visually inspected with no missing glyphs, clipping, overlap, table movement, or pagination drift.
- Regression suite: `42 passed`.

## Related

- Archived hackathon edition with GUI and case gallery: [format-agent](https://github.com/KaguraNanaga/format-agent)
