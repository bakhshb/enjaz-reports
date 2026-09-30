# Report acceptance

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

## Updated template acceptance

- Use the user-approved current master and matching PDF reference. Preserve Abar typography when populating empty cells and duplicating rows or slides. Verify the generated PowerPoint and exported PDF, especially detail bodies and ملاحظات columns.
- Omit a sector from charts only when every plotted status is blank or zero. Retain any sector with a positive count in any plotted status. Apply this to chart caches and embedded workbook categories, without changing source workbooks, totals, or detail records. For entirely empty data, show no named sector or bar and retain zero KPI indicators.
- Acceptance requires normal, empty, new-sector and overflow cases to reconcile with source records, open without PowerPoint repair, and pass final visual inspection for Abar, readable text, no clipping, no overlap and correct pagination.

- Verify actual font usage in the exported PDF, not just font declarations in PowerPoint. On the validated Windows host, PowerPoint SaveAs PDF substituted Calibri even for the unchanged approved master; Adobe PDFMaker preserved Abar and matches the supplied reference export route. If an exporter substitutes fonts, use a verified available exporter and repeat the visual check. Never mark font acceptance passed from successful export alone.
