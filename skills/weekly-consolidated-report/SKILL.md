---
name: weekly-consolidated-report
description: Populate the approved Ministry of Hajj and Umrah weekly consolidated PowerPoint report from the final tasks, Suhail projects, and transactions Excel summaries. Use for التقرير الأسبوعي الموحد or تعبئة تقرير الإنجاز الأسبوعي from these three outputs, preserving the supplied template exactly. Does not build the upstream Excel summaries or split the report by sector.
---

# Weekly Consolidated Report

## Purpose and authority

Create the report from current final Excel inputs and this skill's approved master. Excel controls content; the master controls design. Read [business rules](references/business-rules.md) before generation and [acceptance checks](references/acceptance.md) before validation. Preserve this report's own mappings, ordering, omissions, pagination, and status rules. Sample records and slide counts are not runtime requirements.

## Inputs and output

- ملخص وتفاصيل المهام الاستراتيجية.xlsx: ورقتا ملخص المهام وتفاصيل المهام.
- ملخص_مشاريع_سهيل.xlsx: ورقتا ملخص مشاريع سهيل وتفاصيل مشاريع سهيل.
- تقرير_المعاملات_الموحد.xlsx: ورقة المعاملات.
- قالب PowerPoint المعتمد الحالي: `assets/weekly-report-master.pptx`، مع PDF الموافق مرجعًا بصريًا. هذا القالب هو مصدر التصميم المعتمد الحالي.
- At runtime, use the new Excel files attached to the request. Restore the approved PowerPoint template from the assets attached to this skill; the user should not need to attach the template again. Ask for the template only if the skill assets are unavailable or fail integrity verification. Never substitute sample values.
- Primary output: `تقرير الإنجاز الأسبوعي بتاريخ الإنشاء.pptx`, saved as a new file. Create a matching PDF only when requested. Never modify the inputs or overwrite the original template.


Verify the bundled master size `9545257` bytes and SHA-256 `3e9b018c223c76cbb0e435cd543732801ec8685fc4ec827d4a7f99b29e0423e5` before use. Never fall back to an earlier master.

## Table typography

All table text is **11 pt**, including table titles, headers, data, status cells and notes (and the Monday agenda). Apply this explicit size even when the master stores 10 or 14 pt. Preserve Abar, bold emphasis, RTL, fills and column widths. Adjust row heights and pagination when needed; never reduce the font to make a table fit.

Apply the per-column horizontal alignment in the business rules to **summary and detail tables**, including headers: sector/dates/status centered; task/project names, notes, updates, support text and challenges right-aligned.

## Business rules

Apply every rule in [the report-specific reference](references/business-rules.md). Do not infer content, change source workbooks, or modify the original master. Preserve the report's section order, dates, blank cells, and current status meaning. Template samples are never weekly data.

## Generate and validate

1. Check required workbooks, sheets, headers, and the approved master. Report only the missing or conflicting input that needs resolution.
2. Generate into a new candidate path:

```bash
uv run scripts/build_report.py --tasks "<tasks.xlsx>" --suhail "<suhail.xlsx>" --transactions "<transactions.xlsx>" --output "<candidate.pptx>"
```

3. The builder includes read-only programmatic validation. Its initial placement is provisional: render before accepting page breaks, table gaps, or footer clearance. Follow the existing continuous-flow rules in the business reference; move complete records and render affected slides again. After any correction, rerun validation. Do not repeat it for unchanged bytes.

```bash
uv run scripts/validate_report.py --report "<candidate.pptx>" --tasks "<tasks.xlsx>" --suhail "<suhail.xlsx>" --transactions "<transactions.xlsx>"
```

The candidate is saved only after programmatic checks pass. Failed builds preserve existing outputs and inputs; output paths cannot alias an input or master. A programmatic pass is not visual acceptance.

## Visual review

Render the validated candidate and inspect every slide against this skill's template/reference, including Arabic, actual Abar usage, status fills, notes, chart labels, row continuity, table gaps and footer clearance. Verify native PowerPoint opening without repair and editable embedded charts. Follow [acceptance checks](references/acceptance.md); correct layout using the approved patterns and repeat affected checks. Final visual acceptance follows the last edit and programmatic check.

When a Windows PowerPoint render reveals overlapping tables or premature page breaks, `scripts/measure_layout.ps1 -Report <candidate.pptx> -Tasks <tasks.xlsx> -Suhail <suhail.xlsx> -Transactions <transactions.xlsx>` reads actual row heights without modifying the deck. Rerun the same builder with `--measurements <candidate.layout.json>` to repaginate complete records using those heights, then render again. Measurements are bound to the exact input/template hashes. They support the visual decision; they do not establish visual acceptance. Other environments may correct the candidate with available template-preserving tools and rerun validation.

## Failure handling

Stop the affected report when inputs conflict, a required pattern/style is unapproved, a complete record cannot fit with the approved design, or faithful rendering is unavailable. Report the exact blocker and the condition needed to pass. Preserve the previous accepted report. Never mark a blocked or unperformed check as passed.

## Runtime and installation

Prefer `uv run` using the declared dependencies. Reuse an existing installation. If resolution is unavailable and an existing compatible Python satisfies the dependencies, use it and record the fallback; do not modify host Python. If installing uv is necessary, use Astral's official user-scoped installer, not pip. Both report folders require the sibling `skills/report_common/` directory, installed together from this repository. Runtime scripts must not import test fixtures or the other report's generator.

## Delivery

Deliver the editable PPTX bytes that passed data checks and final visual review, using an accessible file link. Deliver PDF only when requested. Keep QA evidence outside the slides: exact output/input/template hashes, performed checks, renderer, and any limitation. Any edit invalidates affected checks. Do not split sectors, schedule runs, or send reports to others unless requested.
