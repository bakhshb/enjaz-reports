---
name: monday-meeting-report
description: Build the approved Arabic Monday Meeting PowerPoint from the current weekly topics, strategic-task, and Suhail Excel workbooks. The bundled PowerPoint is a design-pattern master; the final number of slides is determined by the current data.
---

# Monday Meeting Report

Build the Monday Meeting deck from the three current Excel workbooks and the bundled master `assets/monday-meeting-master.pptx`.

Excel is the source of truth for weekly content. The PowerPoint master is the source of truth for visual design and reusable slide/table patterns. **The master is not a fixed slide-count contract.** It currently contains the approved base patterns, but a generated report may have 18, 23, 35, or any other number of slides required by the data.

## Required inputs

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

## Dynamic generation rules

### Agenda

Read agenda rows in source order and stop after including `ملخص الاجتماع`. Ignore rows below it.

If the approved agenda table fills, clone the agenda slide and continue automatically. Repeat the same header and visual design. Do not fail merely because the agenda needs more than one page.

### Task summary

Update the task KPIs and native chart from the current summary sheet. Blank chart cells remain truly blank and `dispBlanksAs="gap"` is preserved.

Read the four independent Excel sections `المهام المكتملة`, `المهام المتأخرة`, `المهام المعلقة`, and `طلبات الدعم`. Populate the matching four-column PowerPoint table for every nonempty section, preserving its source rows and notes or support requests. Remove the entire corresponding table, including its title and headers, when the Excel section has no records. Never display template example rows. Flow overflowing records onto continuation summary pages, repeating each continued table's title and headers. Keep KPIs and the chart on the first summary page.

### Task details

Read every `مهام مصدر (...)` section from `تفاصيل المهام` in source order.

There is **no whitelist of task-source names**. Existing names such as `اجتماع القيادات`, `اجتماع التجربة الرقمية`, or `لجنة المتابعة` are examples only.

Use the approved six-column task-detail page as the reusable pattern. For each source:

- populate records in source order;
- when the page fills, clone the same detail pattern and continue;
- when a new task source appears, clone the same pattern and use the new source name;
- keep the approved title/header/table geometry and typography.

The Excel task detail sheet has a seventh field, `الجهات ذات العلاقة`. The user intentionally removed this column from the PowerPoint template; do not transfer it to the deck or append it to `ملاحظات`.

Longer text consumes more page capacity than short text. Prefer an extra continuation page over compressed or clipped content.
If a single record is taller than an otherwise empty approved page, stop with a clear error; never silently cap its estimated height.

### Suhail summary

Update the Suhail KPIs and native chart from the current summary sheet. Read its three independent sections: `أبرز التحديثات`, `المشاريع المتأخرة`, and `التحديات`. Populate a matching four-column template table only when that Excel section contains records; otherwise remove the table with its title and headers. Preserve each section's exact Excel content, including notes and challenge text. Flow overflow onto continuation summary pages, repeating the applicable title and headers. Keep KPIs and the chart on the first summary page.

### Suhail project details

Read every sector from `تفاصيل مشاريع سهيل` in source order.

There is **no whitelist of sector names** and no fixed project count. The sector names present in the master are examples, not limits.

Use the approved six-column Suhail-detail page as the reusable pattern. For every sector:

- create as many pages as required;
- preserve project order;
- repeat the sector title and full header on every continuation page;
- if a completely new sector appears, create pages from the same approved pattern and use the source sector name;
- never fail merely because a sector contains more projects than the master example.

For example, 20 projects may naturally produce three or more Suhail-detail pages. That is expected behavior.

## Text and formatting rules

- Preserve slide size, backgrounds, logos, masters, layout, colors, borders, alignment, RTL behavior, column widths, and approved fonts.
- Agenda tables use `Abar Mid` 14 pt.
- Other dynamic tables use `Abar Mid` 11 pt; preserve `Abar Mid SemiBold` where the approved pattern uses it.
- Do not summarize, rewrite, proofread, infer, or recalculate source content.
- Blank source cells clear old template content; never carry previous-week text forward.
- Remove inherited bullets from dynamic table cells unless the bullet is part of the source text itself.
- In `ما تم حتى تاريخه`, preserve the source line breaks and approved emphasis for the existing `تاريخ التحديث` line.
- The cover date is the runtime creation date in the runtime timezone, formatted like `27 سبتمبر 2026`.

## Runtime

Prefer `uv run` for the bundled scripts. If package/network resolution is unavailable but the active Python already satisfies the declared versions and can import the required packages, run the same script with that Python interpreter rather than failing solely because package download is unavailable.

Do not modify the host Python environment just to generate the report.

## Run sequence

```bash
uv run scripts/build_report.py \
  --topics "<current topics workbook>" \
  --tasks "<current task workbook>" \
  --suhail "<current Suhail workbook>" \
  --output "MM_YYYY-MM-DD.pptx"

uv run scripts/validate_report.py --report "MM_YYYY-MM-DD.pptx"

uv run scripts/report_gate_v2.py \
  --report "MM_YYYY-MM-DD.pptx" \
  --topics "<same current topics workbook>" \
  --tasks "<same current task workbook>" \
  --suhail "<same current Suhail workbook>"
```

Use the system Python fallback described above when `uv` cannot resolve packages but compatible packages are already installed.

## QA philosophy

QA must protect correctness **without limiting valid data growth**.

The final gate checks that:

- the PPTX package and XML are valid and every internal relationship resolves (opening in PowerPoint is a separate visual acceptance step);
- there are no external workbook relationships;
- agenda records match the current topics workbook and end at `ملخص الاجتماع`;
- all task-detail records match the current task workbook;
- all Suhail-detail records match the current Suhail workbook;
- record order matches the source, including when totals are unchanged;
- task and project KPIs reconcile with detail records and sector counts;
- chart categories, series names, and values match both the source summaries and the chart's embedded workbook;
- every populated task and Suhail summary table matches its own Excel section, and every empty section has no table in the deck;
- task and Suhail tables retain their title/header structure;
- dynamic tables use approved fonts and sizes.
- every detail status cell copies its approved master example: completed (blue), on-plan (green), delayed (light red), suspended tasks (light gray), and not-started Suhail projects (light gray). `لم يبدأ` and `لم تبدأ` denote the same project status; it is not a task status. Reject unknown statuses or missing/conflicting master examples. Validate status fill, Abar font, size, weight, and alignment without silently repairing a mismatched status style.
- generation validates a temporary candidate before replacing the requested output. A failed build or gate preserves any existing output and all source files; temporary candidates are cleaned up. Output paths may not alias the master or an input workbook.

Do **not** reject a deck because it has more than 18 slides, a new task source, a new Suhail sector, or more projects/rows than the master examples.

## Mandatory visual QA

After the programmatic gate passes, render the same gated PPTX and inspect the generated dynamic sections. Temporarily unhide appendix slides only in a disposable QA copy if necessary.

Do not deliver if the render shows clipped text, overlaps, broken Arabic glyphs, missing/duplicated rows, inherited bullets, broken table titles/headers, or content outside the approved page geometry. If visual density is too high, the correct repair is normally to create another continuation page, not to shrink the design.

## Delivery

Deliver the actual `.pptx` bytes that passed the final gate and visual QA. Never return only a local filesystem path.

## Updated template acceptance

- Use the user-approved current master and matching PDF reference. Preserve Abar typography when populating empty cells and duplicating rows or slides. Verify the generated PowerPoint and exported PDF, especially transaction bodies and ملاحظات columns.
- Omit a sector from charts only when every plotted status is blank or zero. Retain any sector with a positive count in any plotted status. Apply this to chart caches and embedded workbook categories, without changing source workbooks, totals, or detail records. For entirely empty data, show no named sector or bar and retain zero KPI indicators.
- Acceptance requires normal, empty, new-sector and overflow cases to reconcile with source records, open without PowerPoint repair, and pass final visual inspection for Abar, readable text, no clipping, no overlap and correct pagination.

- Verify actual font usage in the exported PDF, not just font declarations in PowerPoint. On the validated Windows host, PowerPoint SaveAs PDF substituted Calibri even for the unchanged approved master; Adobe PDFMaker preserved Abar and matches the supplied reference export route. If an exporter substitutes fonts, use a verified available exporter and repeat the visual check. Never mark font acceptance passed from successful export alone.
