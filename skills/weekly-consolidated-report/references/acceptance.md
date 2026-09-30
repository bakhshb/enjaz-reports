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
