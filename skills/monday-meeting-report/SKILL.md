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
- `scripts/build_report.py` verifies this checksum before every build and reads the master directly from `assets/`.
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
Use the intended meeting date for the filename and cover. Pass `--meeting-date YYYY-MM-DD` for a specified meeting; otherwise the builder uses this/next Monday in Riyadh (UTC+3), including when prepared on Sunday.

## Table typography

Every PowerPoint table must use native right-to-left (RTL) table direction, including the agenda, summaries, details and continuation tables. Right-aligned text alone is insufficient. Keep numbering at the right, preserve each field's content and width, and retain valid merged titles. Reject output tables without RTL enabled.

The Monday topics/agenda table uses **14 pt** throughout, including headers, body cells and blank-cell defaults. Every other table uses **11 pt**, including titles, headers, statuses and notes. Apply these sizes explicitly, preserving Abar, bold emphasis, RTL, fills and column widths. The same pass writes a fixed line spacing of 1.2 times the font size into every table paragraph (13.2 pt at 11 pt, 16.8 pt at 14 pt, and the matching value at a reduced agenda size). Keep the agenda on one slide using content-sized rows and compact padding. When it does not fit at 14 pt, the font shrinks in 0.5 pt steps down to a floor of 8 pt, re-measured each time. If it still does not fit at 8 pt, the remaining rows continue on a second agenda slide with the header repeated. A notice records either step and the alternative of shorter topic text. Summaries and details can continue onto additional slides. Monday differs from the consolidated deck on purpose: its summary and detail slides are record lists, and the approved pattern for a list is pagination with the header repeated. Only the agenda is a single-slide item, and it shrinks to 8 pt and then continues on a second slide with a notice. Verify 14 pt for topics (or the one reduced size named in the notice) and 11 pt for other tables, with readable Arabic in the PPTX and QA PDF. Regenerate row measurements at the requested sizes.

Apply the per-column horizontal alignment in the business rules to **summary and detail tables**, including headers: sector/dates/status centered; task/project names, notes, updates, support text and challenges right-aligned.

## Business rules

Apply every rule in [the report-specific reference](references/business-rules.md). Do not infer content, change source workbooks, or modify the original master. Preserve the report's section order, dates, blank cells, and current status meaning. Template samples are never weekly data.

## Generate and validate

1. Check required workbooks, sheets, headers, and the approved master. Report only the missing or conflicting input that needs resolution.
2. Generate into a new candidate path:

```bash
uv run scripts/build_report.py --topics "<topics.xlsx>" --tasks "<tasks.xlsx>" --suhail "<suhail.xlsx>" --output "<candidate.pptx>"
```

On Windows, the builder measures a disposable populated copy in native PowerPoint. In `auto` mode, recoverable PowerPoint/PowerShell failures, timeouts, and missing or invalid measurement results produce a visible warning and use fresh conservative estimates for every table. The bundled measurement script runs with a process-only execution-policy override; no machine policy changes. On other hosts, `auto` measures every body row in LibreOffice when LibreOffice and the Abar Mid font are both installed; without Abar it uses the estimate, because substitute fonts have different widths. Estimates wrap each paragraph word by word at Abar Mid widths calibrated on the approved reference PDF (5.2 pt per character regular, 5.6 pt bold, 14 pt line pitch at 11 pt, scaled for 14 pt topics), including bullet indents. `--layout-engine native` requires valid PowerPoint measurements and stops with the failure reason; `--layout-engine libreoffice` requires LibreOffice measurements; `--layout-engine estimate` explicitly selects estimates. Every engine still requires programmatic checks and final visual verification before delivery. `MONDAY_LAYOUT_ENGINE` selects the same engine in automated tests. No measurement cache is reused. The master and inputs are unchanged.

3. The builder includes the programmatic gate, including typography normalization. Do not repeat the gate for unchanged bytes. After a layout correction, rerun the gate once and render its resulting bytes.

```bash
uv run scripts/report_gate_v2.py --report "<candidate.pptx>" --topics "<topics.xlsx>" --tasks "<tasks.xlsx>" --suhail "<suhail.xlsx>"
```

The candidate is saved only after programmatic checks pass. Failed builds preserve existing outputs and inputs; output paths cannot alias an input or master. A programmatic pass is not visual acceptance.

## Visual review

Judge row heights only in a render that really uses Abar Mid; the renderer warns when Abar is not installed, because substitute fonts make fitting rows look clipped or overlapping. For LibreOffice PDF previews, use `uv run ../report_common/render.py --report "<candidate.pptx>" --output "<candidate.pdf>"`. The shared exporter materializes the RTL visual column order in a disposable copy because LibreOffice ignores the native DrawingML RTL table setting. It leaves the editable PPTX unchanged. Inspect every rendered table: numbering must be on the right, followed by the task/project and remaining fields to its left. An RTL flag alone is not visual acceptance. `tests/check_rendered_pdf.py` verifies numbering from actual PDF coordinates as well as fonts and page counts.

Render the validated candidate and inspect every slide against this skill's template/reference, including Arabic, actual Abar usage, status fills, notes, chart labels, row continuity, table gaps and footer clearance. Verify native PowerPoint opening without repair and editable embedded charts. Follow [acceptance checks](references/acceptance.md); correct layout using the approved patterns and repeat affected checks. Final visual acceptance follows the last edit and programmatic check.

## Failure handling

Stop only when the report cannot be produced at all: a missing or unreadable input workbook, sheet or required header, a damaged master, or a built file that fails the gate. Lack of space does not stop it. An agenda that does not fit one slide at 14 pt shrinks its font in 0.5 pt steps down to 8 pt and, if it still does not fit, continues on a second agenda slide with the header repeated; a single record taller than a slide takes a slide of its own; an unknown task or project status keeps its own text with the neutral gray fill (`معلق` for tasks, `لم تبدأ` for projects); a KPI that disagrees with the sector counts is transferred as written. All of these are written to `<output>.notices.json` and printed as «تنبيه: …». The gate runs in its own process and passes its notices back to the builder, so they are in the same file. Read that file after every build. The skill does not wait: it delivers the build with the default, and in the same reply repeats every notice (what did not fit, what default was applied) and asks whether the user wants shorter source text, for example shorter topics for the agenda, and a rebuild. If the user had already stated a preference, it is used. Monday has no `--summary-*` flags. Preserve the previous accepted report. Never mark a blocked or unperformed check as passed.

After a manual edit of a delivered deck in PowerPoint (a new row, longer text), run `uv run ../report_common/reflow.py --report "<edited.pptx>" --output "<fixed.pptx>"` to restore row heights, numbering and table stacking in a new file.

## Runtime and installation

Prefer `uv run` using the declared dependencies. Reuse an existing installation. If resolution is unavailable and an existing compatible Python satisfies the dependencies, use it and record the fallback; do not modify host Python. If installing uv is necessary, use Astral's official user-scoped installer, not pip. Shared code lives in `report_common/`: the sibling `skills/report_common/` in a repository or plugin install, or `scripts/report_common/` inside a standalone bundle from `tools/package_skills.py`. In a bundle, run the PDF exporter as `uv run scripts/report_common/render.py`. Runtime scripts must not import test fixtures or the other report's generator.

## Delivery

Deliver the editable PPTX bytes that passed data checks and final visual review, using an accessible file link. Deliver PDF only when requested. Keep QA evidence outside the slides: exact output/input/template hashes, performed checks, renderer, and any limitation. Any edit invalidates affected checks. Do not split sectors, schedule runs, or send reports to others unless requested.
