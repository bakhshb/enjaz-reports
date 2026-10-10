---
name: monday-meeting-report
description: Build the approved Arabic Monday Meeting PowerPoint from the current weekly topics, strategic-task, and Suhail Excel workbooks. The bundled PowerPoint is a design-pattern master; the final number of slides is determined by the current data.
---

# Monday Meeting Report

## Purpose and authority

Create the report from current final Excel inputs and this skill's approved master. Excel controls content; the master controls design. Read [business rules](references/business-rules.md) before generation and [acceptance checks](references/acceptance.md) before validation. Preserve this report's own mappings, ordering, omissions, pagination, and status rules. Sample records and slide counts are not runtime requirements.

## Inputs and output

1. Weekly topics workbook: `مواضيع الأسبوع.xlsx` or equivalent, with `Sheet1` and the columns `الترتيب`, `الموضوع`, `المتحدث`, `الوقت المطلوب من المالك`.
2. `ملخص وتفاصيل المهام الاستراتيجية.xlsx`, with `ملخص المهام` and `تفاصيل المهام`.
3. `ملخص_مشاريع_سهيل.xlsx`, with `ملخص مشاريع سهيل` and `تفاصيل مشاريع سهيل`.

Do not use old Library files or template text as weekly data. If an input workbook is missing or its required sheets/headers are not present, ask only for that missing/incorrect input.

## Master asset

- Runtime master: `assets/monday-meeting-master.pptx`
- Approved SHA-256: `1b7c72de3a5034ee0352a8b6a67f3f519a2e4bad5d95ad93424acdb5353ef255`
- `scripts/restore_template.py` verifies the checksum and restores a working copy when needed.
- Visual reference: `assets/monday-meeting-reference.pdf`. Use it to inspect the approved design, never as weekly data or a slide-count limit.

Treat the master as a library of approved patterns:

- cover;
- agenda table;
- task summary;
- task-detail page/table;
- Suhail summary;
- Suhail-detail page/table;
- section dividers and closing slides.

Never redesign these patterns. Clone and reuse them when more space is required.


Primary output: a new editable `MM_YYYY-MM-DD.pptx`. Generate a delivery PDF only if requested; a QA render remains mandatory.

## Table typography

All table text is **11 pt**, including table titles, headers, data, status cells and notes (and the Monday agenda). Apply this explicit size even when the master stores 10 or 14 pt. Preserve Abar, bold emphasis, RTL, fills and column widths. Adjust row heights and pagination when needed; never reduce the font to make a table fit.

Apply the per-column horizontal alignment in the business rules to **summary and detail tables**, including headers: sector/dates/status centered; task/project names, notes, updates, support text and challenges right-aligned.

## Business rules

Apply every rule in [the report-specific reference](references/business-rules.md). Do not infer content, change source workbooks, or modify the original master. Preserve the report's section order, dates, blank cells, and current status meaning. Template samples are never weekly data.

## Generate and validate

1. Check required workbooks, sheets, headers, and the approved master. Report only the missing or conflicting input that needs resolution.
2. Generate into a new candidate path:

```bash
uv run scripts/build_report.py --topics "<topics.xlsx>" --tasks "<tasks.xlsx>" --suhail "<suhail.xlsx>" --output "<candidate.pptx>"
```

On Windows, the builder measures a disposable populated copy in native PowerPoint before pagination. It does not alter the master or inputs. A native measurement failure stops publication. On other hosts, the conservative content estimate requires visual verification; `--layout-engine estimate` explicitly selects that fallback and `--layout-engine native` requires PowerPoint. `MONDAY_LAYOUT_ENGINE` can select the same engine in automated tests. No measurement cache is reused.

3. The builder includes the programmatic gate, including typography normalization. Do not repeat the gate for unchanged bytes. After a layout correction, rerun the gate once and render its resulting bytes.

```bash
uv run scripts/report_gate_v2.py --report "<candidate.pptx>" --topics "<topics.xlsx>" --tasks "<tasks.xlsx>" --suhail "<suhail.xlsx>"
```

The candidate is saved only after programmatic checks pass. Failed builds preserve existing outputs and inputs; output paths cannot alias an input or master. A programmatic pass is not visual acceptance.

## Visual review

Render the validated candidate and inspect every slide against this skill's template/reference, including Arabic, actual Abar usage, status fills, notes, chart labels, row continuity, table gaps and footer clearance. Verify native PowerPoint opening without repair and editable embedded charts. Follow [acceptance checks](references/acceptance.md); correct layout using the approved patterns and repeat affected checks. Final visual acceptance follows the last edit and programmatic check.

## Failure handling

Stop the affected report when inputs conflict, a required pattern/style is unapproved, a complete record cannot fit with the approved design, or faithful rendering is unavailable. Report the exact blocker and the condition needed to pass. Preserve the previous accepted report. Never mark a blocked or unperformed check as passed.

## Runtime and installation

Prefer `uv run` using the declared dependencies. Reuse an existing installation. If resolution is unavailable and an existing compatible Python satisfies the dependencies, use it and record the fallback; do not modify host Python. If installing uv is necessary, use Astral's official user-scoped installer, not pip. Both report folders require the sibling `skills/report_common/` directory, installed together from this repository. Runtime scripts must not import test fixtures or the other report's generator.

## Delivery

Deliver the editable PPTX bytes that passed data checks and final visual review, using an accessible file link. Deliver PDF only when requested. Keep QA evidence outside the slides: exact output/input/template hashes, performed checks, renderer, and any limitation. Any edit invalidates affected checks. Do not split sectors, schedule runs, or send reports to others unless requested.
