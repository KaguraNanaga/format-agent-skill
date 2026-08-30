# format-agent-skill

**[中文](README.md) | English**

A universal document-formatting Agent Skill. **AI understands, code executes, JSON in between.**

Give it a formatting standard (plain-language rules in text, or an already-formatted
Word document as the template) plus a messy .docx — the Agent reads the rules,
understands the document structure, rewrites the formatting paragraph by paragraph,
and delivers a production-ready Word document.

> This repository IS the skill: `SKILL.md` lives at the root. Clone it and your Agent can install it.

## Install (one sentence to your Agent)

```
Install this skill for me: https://github.com/KaguraNanaga/format-agent-skill
```

Your Agent reads `SKILL.md`, then performs its own security audit, dependency
setup, and smoke test. Verified working with Kimi Code and Tencent WorkBuddy;
any Agent environment with a skill mechanism (Trae, Qoder, Comate, etc.) works.

## Use

After installation, just describe the task:

> Reformat "draft.docx" according to the rules in "formatting-rules.txt".

> Use this board resolution template and apply its formatting to my draft.

**Verified in practice**: inside Kimi Code, hand the skill "a well-formatted document +
a document to fix" and it clones the formatting — fonts, sizes, spacing, indentation,
numbering — onto the target, producing a tracked-changes copy and a change report.

### Format cloning from an English document (example)

A messy English board resolution, reformatted against plain-English rules:

| Before | After |
|---|---|
| ![before](docs/images/en-board-before.png) | ![after](docs/images/en-board-after.png) |

Reproduce it yourself:

```bash
pip install -r requirements.txt
python main.py --spec examples/en/board_resolution_spec.txt \
    --target examples/en/board_resolution_messy.docx \
    --out out/board_resolution_formatted.docx
```

## Zero-config inside an Agent (no API key needed)

When running inside an Agent, **the host Agent itself does the understanding** —
no `.env`, no API key. The skill extracts the document structure, your Agent writes
the two JSONs (FormatSpec / RoleMap), and deterministic code does the rewriting.
Standalone/server use can optionally configure an OpenAI-compatible endpoint
(see `.env.example`; a multimodal model such as GPT-4o, Kimi K3, Qwen-VL or GLM-4V
is recommended if you want the visual self-check).

## Outputs

- `formatted.docx` — clean copy with named Word styles (fully editable afterwards)
- `formatted_tracked.docx` — **track-changes copy**: accept/reject every formatting edit in Word's Review pane
- `*_report.docx` / `.md` — paragraph-level change report
- `_formatspec.json` / `_rolemap.json` — the Agent's understanding, as evidence

## How it works

```
Formatting source (rules text / well-formatted template)
        │  ① Rule understanding (host Agent / LLM / deterministic reading)
        ▼
   FormatSpec (rules as JSON) ◄── schema-validated; failures fed back for self-correction
        │
Target docx ── ② Structure extraction (python-docx, numbering metadata) ──► paragraph list
        │                                                                  │
        │        ③ Role labeling (deterministic numbering conventions       │
        │           first, model semantics as fallback)                     ▼
        │                                                            RoleMap (roles as JSON)
        ▼                                                                  │
  ④ Executor (deterministic code, writes named Word styles) ◄──────────────┘
        │
        ▼
Clean copy + tracked-changes copy + change report + (optional) visual self-check
```

- The model only ever produces two JSONs — it never touches the document. Zero hallucinated formatting.
- **Rules come from your input, never from hardcoded conventions.** Every school, company,
  or country has its own formatting norms; the skill reads yours at runtime instead of
  pretending to know them all.
- Structure recognition is layered: high-frequency conventions (Chinese 一、（一）, legal
  第一章/第一条, native Word auto-numbering) are deterministic; anything unusual falls
  back to the model's general semantics — which is exactly what LLMs are good at.
- Trust by design: named styles, tracked changes, paragraph-level report, and an optional
  visual QA pass (renders pages and lets a multimodal model check them against the rules).

## Scope & boundaries

**Works well for** paragraph-based documents: board/shareholder/party-committee
materials, contracts, legal opinions, consulting reports, theses, press releases,
notices, work plans.

**Not covered** (left untouched, never corrupted): cover-page design.
Table content is preserved as-is on table-centric documents (rosters, price sheets).

## Related

- Hackathon edition (GUI demo, full case gallery, roadshow page):
  [format-agent-01](https://github.com/KaguraNanaga/format-agent-01) (archived)
