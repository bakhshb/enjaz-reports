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

Do not deliver if the render shows clipped text, overlaps, broken Arabic glyphs, missing/duplicated rows, unintended bullets, broken table titles/headers, or content outside the approved page geometry. If visual density is too high, the correct repair is normally to create another continuation page, not to shrink the design.

## Updated template acceptance

- Use the user-approved current master and matching PDF reference. Preserve Abar typography when populating empty cells and duplicating rows or slides. Verify the generated PowerPoint and exported PDF, especially detail bodies and ملاحظات columns.
- Omit a sector from charts only when every plotted status is blank or zero. Retain any sector with a positive count in any plotted status. Apply this to chart caches and embedded workbook categories, without changing source workbooks, totals, or detail records. For entirely empty data, show no named sector or bar and retain zero KPI indicators.
- Acceptance requires normal, empty, new-sector and overflow cases to reconcile with source records, open without PowerPoint repair, and pass final visual inspection for Abar, readable text, no clipping, no overlap and correct pagination.

- Verify actual font usage in the exported PDF, not just font declarations in PowerPoint. On the validated Windows host, PowerPoint SaveAs PDF substituted Calibri even for the unchanged approved master; Adobe PDFMaker preserved Abar and matches the supplied reference export route. If an exporter substitutes fonts, use a verified available exporter and repeat the visual check. Never mark font acceptance passed from successful export alone.

## Compact pagination and Suhail update acceptance

- All tables are explicitly RTL, keep the approved visual column order/widths and valid merged cells. Native PowerPoint must show the text; XML text presence alone is insufficient.
- Summary rows use content height, never sample-height multipliers. Multiple nonempty summary tables share available space. No row, title or header overlaps another table or enters the footer area.
- Task-source and Suhail-sector detail tables share a page when space permits, retaining separate titles and headers and source record order. Repeated headers count once per continued table. Do not require a fixed slide count for other weeks.
- For the supplied 4 October regression content: preserve all 9 task records and 20 project records from the generated report; task details fit two pages and the four Suhail highlight rows fit one summary page. The manually edited deck is a layout reference, not permission to drop its missing project or alter totals.
- The first Suhail update content line has bold `تاريخ التحديث` plus its date; body text is regular. Multiple source update paragraphs use native RTL bullets; a single paragraph has none. Leading/trailing blank paragraphs do not create excess height. Data comparison allows only this specified whitespace/list-marker normalization.
- Regressions cover consecutive tables, overflow with repeated headers, an oversized record that stops safely, RTL merged-cell/schema preservation, a single paragraph, multiple points, and deliberate removal of bold/bullets. Run the existing data, status, output-protection and weekly report checks after shared formatter changes.
- Final acceptance includes native rendering of the exact gated bytes, readable Abar at approved sizes, no blank tables, clipping or overlap, and actual font inspection in the QA PDF. Record input provenance: when only PPTX inputs were supplied, reconstructed regression workbooks are not independently verified original Excel exports.

## Table font size acceptance

Every visible table run must be 11 pt, including title/header rows, blank-cell defaults, statuses, notes and the Monday agenda. Verify generated PPTX content, not only the template or instructions. The weekly read-only gate rejects a 10 pt run; the Monday gate normalizes ordinary cells and validates status cells against 11 pt approved styles. Preserve other formatting and recheck rendered row heights, clipping and pagination at 11 pt. Previous 10 pt layout measurements must be regenerated.

## Table column alignment

Apply to both summaries and details, to column headers and every body paragraph. Center القطاع, تاريخ الإنجاز المخطط (including الانجاز spelling), تاريخ البداية, تاريخ النهاية, الحالة and حالة المشروع. Right-align المهمة, المشروع/اسم المشروع, ملاحظات, التحديث, طلب الدعم, التحدي and ما تم حتى تاريخه. Preserve merged table titles, numbering, other columns, vertical alignment, RTL, colors and emphasis. Font size is 11 pt throughout.
