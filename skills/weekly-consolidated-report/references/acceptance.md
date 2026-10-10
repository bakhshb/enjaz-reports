# Report acceptance

## Acceptance checks before delivery

- `إجمالي المهام = مكتملة + على المخطط + متأخر + معلق`; each status total across sectors must equal the summary, and the sum of all task-detail sections must equal the total.
- `إجمالي مشاريع سهيل = مكتملة + على المخطط + متأخر + لم تبدأ`; sector totals and the number of project-detail records must match the summary.
- `إجمالي المعاملات = مكتملة + على المخطط + متأخر`; sector totals must match the summary, and the number of records in the delayed-transactions table must match the `متأخرة` indicator. `لم تبدأ` must not appear in transaction indicators, chart data, or acceptance arithmetic.
- Verify that completed and delayed summary tables contain the corresponding detail records, that `المعلقة` comes from task details, and that `طلبات الدعم` comes from its own source table. Never auto-fix source conflicts; report the conflicting cells.
- Every required source record must appear once in its destination table, except for intentional appearances in both summary and details. No empty table may remain, and no stale update text may remain where the source field is blank.
- Verify every detail-table status cell against the status-style map after all row copies, splits, and slide moves. The same normalized status must never appear with different fills, and `مكتملة` must not retain the green `على المخطط` fill.
- On every slide containing multiple tables, inspect all visible inter-table gaps together at full size. They must be consistently small; no pair may touch, overlap, or leave a noticeably larger blank band than the other pairs.
- All displayed numbers, including chart data labels, must match their source. Embedded workbook data and chart caches must be consistent.
- Open the generated file and recheck PowerPoint integrity and relationships. It must open directly without `Repair`, `Found a problem with content`, or any repair prompt, and it must contain no external or missing Excel links. Verify that `Edit Data` for every chart uses the workbook embedded inside PowerPoint. Render the slides and compare them visually against the PDF/template for fonts, sizes, positions, RTL behavior, clipping, overlaps, and table boundaries. A summary continuation slide is valid only when it contains overflow table records and does not repeat the summary indicators or chart; no blank continuation slide may remain.
- Verify that the template fonts are actually available, including `Abar Mid` and `Abar Mid SemiBold`. If the runtime substitutes fonts or cannot edit the template while preserving formatting, explain the specific limitation before producing a differently formatted report.
- Do not claim that the report was tested in PowerPoint or visually validated unless those checks were actually performed. Creating or updating this skill alone is not a full population test.

## Updated template acceptance

- Use the user-approved current master and matching PDF reference. Preserve Abar typography when populating empty cells and duplicating rows or slides. Verify the generated PowerPoint and exported PDF, especially transaction bodies and ملاحظات columns.
- Omit a sector from charts only when every plotted status is blank or zero. Retain any sector with a positive count in any plotted status. Apply this to chart caches and embedded workbook categories, without changing source workbooks, totals, or detail records. For entirely empty data, show no named sector or bar and retain zero KPI indicators.
- Acceptance requires normal, empty, new-sector and overflow cases to reconcile with source records, open without PowerPoint repair, and pass final visual inspection for Abar, readable text, no clipping, no overlap and correct pagination.

- Verify actual font usage in the exported PDF, not just font declarations in PowerPoint. On the validated Windows host, PowerPoint SaveAs PDF substituted Calibri even for the unchanged approved master; Adobe PDFMaker preserved Abar and matches the supplied reference export route. If an exporter substitutes fonts, use a verified available exporter and repeat the visual check. Never mark font acceptance passed from successful export alone.

## Suhail update rich-text acceptance

- Leading `تاريخ التحديث` and its date are bold; all following prose remains regular, including same-line text and other dates.
- Two or more source paragraphs after the date become native bullets in source order, including unmarked items like the supplied 08 October report. Explicit lists also qualify, without duplicate typed markers. A single paragraph/item receives no bullets.
- Reconcile source wording allowing only conversion of source list markers to native bullets; preserve all substantive text; leading/trailing empty paragraphs may be trimmed.
- Final visual review must confirm Abar, Arabic direction, readable bullet alignment and no clipping after formatting. These checks remain pending until run on the changed implementation.

## Table font size acceptance

Every visible table run must be 11 pt, including title/header rows, blank-cell defaults, statuses, notes and the Monday agenda. Verify generated PPTX content, not only the template or instructions. The weekly read-only gate rejects a 10 pt run; the Monday gate normalizes ordinary cells and validates status cells against 11 pt approved styles. Preserve other formatting and recheck rendered row heights, clipping and pagination at 11 pt. Previous 10 pt layout measurements must be regenerated.

## Table column alignment

Apply to both summaries and details, to column headers and every body paragraph. Center القطاع, تاريخ الإنجاز المخطط (including الانجاز spelling), تاريخ البداية, تاريخ النهاية, الحالة and حالة المشروع. Right-align المهمة, المشروع/اسم المشروع, ملاحظات, التحديث, طلب الدعم, التحدي and ما تم حتى تاريخه. Preserve merged table titles, numbering, other columns, vertical alignment, RTL, colors and emphasis. Font size is 11 pt throughout.

## Suhail summaries and details regression

Generate both report types with a dated multi-item update and a dated single paragraph in Suhail summary and detail cells. Confirm the supplied date label/date alone are bold, body text is regular, multiple source list items have native RTL bullets without duplicate markers, and prose has none. Remove bold or a bullet from a summary output deliberately: the delivery gate must reject it. Reconcile each summary against its own input. Render final bytes in PowerPoint and inspect summary and detail cells at 11 pt for clipping, overlaps and readable Arabic. Repeat the build using the installed skills and verify installed code hashes match the tested repository files.
