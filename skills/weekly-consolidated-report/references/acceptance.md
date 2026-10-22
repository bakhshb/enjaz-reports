# Report acceptance

## Acceptance checks before delivery

- `إجمالي المهام = مكتملة + على المخطط + متأخر + معلق`; each status total across sectors must equal the summary, and the sum of all task-detail sections must equal the total.
- `إجمالي مشاريع سهيل = مكتملة + على المخطط + متأخر + لم تبدأ`; sector totals and the number of project-detail records must match the summary.
- `إجمالي المعاملات = مكتملة + على المخطط + متأخر`; sector totals must match the summary, and the number of records in the delayed-transactions table must match the `متأخرة` indicator. `لم تبدأ` must not appear in transaction indicators, chart data, or acceptance arithmetic.
- Verify that completed and delayed summary tables contain the corresponding detail records, that `المعلقة` comes from task details, and that `طلبات الدعم` comes from its own source table. Never auto-fix source conflicts; report the conflicting cells.
- Every required source record must appear once in its destination table, except for intentional appearances in both summary and details. No empty table may remain, and no stale update text may remain where the source field is blank.
- Verify every detail-table status cell against the status-style map after all row copies, splits, and slide moves. The same normalized status must never appear with different fills, and `مكتملة` must not retain the green `على المخطط` fill.
- On every slide containing multiple tables, inspect all visible inter-table gaps together at full size. They must be consistently small; no pair may touch, overlap, or leave a noticeably larger blank band than the other pairs.
- Summary table title and header rows must have explicit positive heights in generated slides. The validator rejects unresolved auto-height rows and overlapping table frames; PowerPoint rendering confirms the visible rows and gaps, especially on «ملخص المهام» in both models.
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


## Detail artwork acceptance

- Normal, empty, new-sector and overflow builds must pass data checks. Every generated detail/continuation slide must retain the approved layout/master artwork and visible inherited graphics.
- All four header buttons (text and background) must target their correct current sections, including transactions.
- Reject deliberately hidden slide/layout graphics, deleted logo/footer/button, wrong layout, missing/wrong navigation links, moved artwork and table IDs colliding with grouped shapes. Validation must not repair the damaged input.
- Open the output in native PowerPoint without repair. Compare header/footer regions on every detail/continuation page against an unchanged-master export, including two consecutive exports of identical PPTX bytes. Inspect content/layout separately; artwork checks do not replace font, table and source reconciliation checks.
- Preserve input/template hashes; verify that the installed generator and validator match the tested repository files.

## Two-model acceptance

- When only the three Excel files arrive and no model is specified, the agent asks which approved model to use before generating. An explicit choice is honored without asking again. Choosing the current model retains its original master and three-input workflow.
- Choosing the digital-experience model requires the current PDF. Verify the bundled second-master hash, render PDF page one, inspect the extractor's JSON against that page, correct text-layer errors, and reject unreviewed JSON or a PDF hash mismatch.
- The digital slide contains the PDF's four independent percentages, sector rows (including blanks and `قيد الدراسة` notes), and challenge text in source order. No numerical or challenge sample from the master remains. Card labels, doughnuts, progress bars, sector data labels, chart caches, embedded workbooks, and challenge cells agree with the reviewed source.
- The digital summary retains the approved fixed «آلية الاحتساب» note between the indicators and sector chart: 10% for readiness (requirements, meetings and coordination), 5% for target dates, and the remainder for actual completion of guide requirements. Its Arabic wording and bold label remain legible, with no overlap or clipping.
- The new slide is in the approved position; five header destinations work for the digital model, including grouped buttons. Existing summary/detail reconciliation, typography, status, layout-artwork and visual checks pass for both models. A source with more than five sector rows or four challenge rows stops with a clear capacity error rather than losing data.
- Open each model's generated PPTX in PowerPoint without repair. Compare the digital slide to page one of the supplied PDF and the approved second master at full size; verify Arabic, 11 pt table text, no clipping or overlap, correct sector-note placement, and intact footer/logo. Final visual acceptance follows the last edit and validation.
- The updated digital master has six detailed pages after the detail divider. Each output card, sector title, table row (including section rows, codes and values), and challenge matches the reviewed corresponding PDF page two through seven. No duplicate sample row remains; the source PDF's page-three 10% values override the master examples. A missing source page or overflowing table fails clearly.
- The divider's digital-details button opens the first digital detail page. All six pages retain the approved header/footer artwork and 11 pt tables; no page-one sample values or detail sample data survive. Both PDF page-six challenges appear in the new approved table on the corresponding detail slide; no challenge is omitted.
