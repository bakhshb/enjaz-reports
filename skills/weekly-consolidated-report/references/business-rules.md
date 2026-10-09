# Report business rules

## Authoritative rules

1. Transfer data, statuses, and dates from the final Excel outputs. Do not recalculate lateness from dates, and do not change a project status because its progress description appears to imply another status.
2. Transfer text from the final Excel output exactly as provided. Do not proofread, summarize, rewrite, remove details, or alter numbers, dates, or identifiers. Do not convert the word `مرفق` into a guessed document name, and do not automatically add `الوثيقة الداعمة`.
3. Normalize `مكتمل` and `مكتملة` to `مكتملة` in the report. Keep the other source categories unchanged in meaning: `على المخطط`, `متأخر/متأخرة`, and `معلق/معلقة`. `لم تبدأ` remains valid for Suhail projects only. Do not display or chart `لم تبدأ` for transactions. Never convert one category into another. If a new or ambiguous category appears, ask only about that category.
4. Sector names must come from Excel exactly as supplied. Do not rely on outdated labels from the template and do not merge two different agencies by assumption. Charts and tables driven by the same source must use consistent source naming.
5. The cover date is the report creation date in the runtime timezone, not the template date and not a project's latest update date. Format it exactly like the template and do not add a time. Never use the report creation date to fill missing project update dates.
6. Blank task notes or updates in Excel must remain blank in the report; remove any previous-week text from both summary and detail locations. Do not fill any other missing field from the template or by inference. Preserve incomplete source text such as `XXX` exactly as-is and mention the missing source content in the delivery note.
7. A table with no records must not appear in the report. Column headers, section titles, and the phrase `لا يوجد` are not records. Any table with actual records must be included in full, even if the current template version has no matching table.
8. Table text must use 11 pt throughout, overriding the master's 10 pt examples. This includes titles, headers, body, status and notes cells. Preserve all other styling. Never change the design: slide size, masters, layouts, font families or weights, alignment and RTL behavior, margins, z-order, template colors and borders, charts, logos, or backgrounds. Move whole tables vertically and duplicate the appropriate slide pattern only as required by pagination. Retain each delayed-table title strip's color from the newly approved template, including continuation slides; do not impose an older red override.
9. Paginate tables before they touch the footer. Split records across copies of the appropriate slide while preserving widths, fonts, column widths, row styling, and record order. Do not shrink fonts, compress text, stretch rows, overlap tables, or place a table over the footer. Summary indicators and charts stay on the first summary slide; table overflow may use a continuation slide that preserves the same section navigation and footer without repeating indicators or charts. Page numbering and internal links must be updated after insertion or removal of slides, and the thank-you slide must remain last.
10. Preserve the section order of the template and the record order from the final Excel output within each section. Do not reapply sorting or deduplication policies from upstream skills. `#` is display numbering within a section; do not treat it as a global identifier and do not automatically remove records that have similar names.
11. Summary indicators and charts remain present even when a true status count is zero. Text-based numeric indicators must display `0` when the true value is zero, but chart series data must not write `0` into the embedded worksheet cell: leave that cell blank and transfer only non-zero numeric values into chart data. Do not remove an entire category or sector if it has values in another series. The empty-table rule does not remove `ملخص المعاملات` or a legitimate zero indicator.

## Data and object discovery

The slide numbers and cell ranges below document the inspected template and sample; they are not weekly constants. Detect each table's end from the next section heading and column headers, ignoring blank rows. Parse sections from their headings even when row counts or positions change. Do not assume named Excel Tables exist; the samples use formatted ranges without formulas or attachment links.

Identify each slide's role from its section heading, objects, and column headers, then use the slide number plus object name/ID as verification. Some navigation labels come from the layout or master. The chart together with the transaction indicator names is sufficient to identify `ملخص المعاملات`; do not require a transaction-number table in the summary.

When matching summary records to details, use source/section, sector, and task or project name. Internal matching keys may normalize whitespace, directional characters, and diacritics without changing the displayed text. Preserve every repeated source project. If matching is ambiguous or summary and detail data conflict, name the conflict and ask instead of choosing the first match or modifying Excel.

## Task mapping

`ملخص المهام`:

- B2 `إجمالي المهام` → slide 2, TextBox 10, ID 11.
- B3 `مكتملة` → TextBox 13, ID 14.
- B4 `على المخطط` → TextBox 15, ID 16.
- B5 `متأخرة` → TextBox 6, ID 7.
- B6 `معلقة` → TextBox 5, ID 6.
- Sample A9:E17 → Chart 23, ID 24, `chart1.xml`. A=`القطاع`, B=`مكتملة`, C=`على المخطط`, D=`متأخر`, E=`معلق`.
- The updated slide 2 contains four approved four-column table patterns: `المهام المكتملة` (ID 10), `المهام المتأخرة` (ID 23), `المهام المعلقة` (ID 8), and `طلبات الدعم` (ID 25). Identify each by its title and headers rather than relying on its ID or sample position.
- Read each matching section from `ملخص المهام` independently, in Excel record order. The first three use `# | المهمة | القطاع | ملاحظات`; `طلبات الدعم` uses `# | المهمة | القطاع | طلب الدعم` and comes only from its own section. Copy all four values exactly. Do not derive support requests from task status or derive notes from detail rows when the summary cell is blank.
- If a section has no numbered records, remove its entire PowerPoint table, including title and headers. Ignore `لا يوجد` and template example records. If populated tables do not fit, continue them on summary-table pages with the same approved title/header style; do not duplicate KPI cards or the chart.

Task details:

- The current task-detail patterns start on slide 6 and have six columns: `# | المهمة | القطاع | تاريخ الإنجاز المخطط | الحالة | ملاحظات`. Read every `مهام مصدر (...)` section from Excel and create as many pages as its records need; the template source names and slide count are examples.
- The Excel detail sheet also contains `الجهات ذات العلاقة`. The user intentionally removed it from the PowerPoint template. Omit that field from the deck; do not append it to `ملاحظات` or add a seventh column.
- New Excel sections are added after the existing sections using a copy of the appropriate detail slide. Empty sections must not be shown as tables.

## Suhail project mapping

`ملخص مشاريع سهيل`:

- A3 `الإجمالي` → slide 3, TextBox 9, ID 10.
- B3 `مكتملة` → TextBox 11, ID 12.
- C3 `على المخطط` → TextBox 13, ID 14. The template value `16` is split across text runs; replace the complete displayed value while preserving the object's formatting.
- D3 `متأخرة` → TextBox 7, ID 8.
- E3 `لم تبدأ` → TextBox 6, ID 7.
- A6:E15 → Chart 20, ID 21, `chart2.xml`. A=`القطاع`, B=`مكتملة`, C=`على المخطط`, D=`متأخر`, E=`لم تبدأ`.
- The updated slide 3 contains three approved four-column table patterns: `أبرز التحديثات` (ID 4), `المشاريع المتأخرة` (ID 23), and `التحديات` (ID 2). Identify them by title and headers.
- Read each matching section from `ملخص مشاريع سهيل` independently. Updates use `# | المشروع | القطاع | التحديث`; delayed projects use `# | المشروع | القطاع | ملاحظات`; challenges use `# | المشروع | القطاع | التحدي`. Preserve the Excel text and record order, including blank notes. Do not infer new entries from project details.
- Remove an entire table with its title and headers when its Excel section has no numbered records. `لا يوجد` is not a record. Continue populated tables on summary-table pages when necessary, without repeating indicators or the chart. The template does not contain a completed-project summary table; completed projects remain in the KPI, chart, and details.

Suhail project details:

- The current six-column Suhail detail patterns begin on slide 8 and continue through slide 10. These are design examples, not a sector whitelist or a fixed number of pages. Read every titled sector section in `تفاصيل مشاريع سهيل` and build enough pages for all its records, retaining source order and the approved title/header pattern.
- Column order for every Suhail detail table: A=`#`, B=`اسم المشروع`, C=`تاريخ البداية`, D=`تاريخ النهاية`, E=`الحالة`, F=`ما تم حتى تاريخه`. In column F, when the cell begins with an update-date label and date such as `تاريخ التحديث 15 سبتمبر 2026` or the same phrase using another valid day-month-year value, format only the leading label and date in bold (see Suhail detail update formatting below). Keep the remaining text in the cell in the existing regular style. Do not change, rewrite, infer, or add the date text; apply bold only to the supplied label and date.
- The sector is taken from the section heading above the table, not from a non-existent detail column.
- PowerPoint table placement differs from the sector order in Excel; match by sector name. Add new sectors using an existing detail pattern. A sector that only has a zero-valued statistical row must not generate a detail table.

## Transaction mapping

Transaction worksheet:

- B2 `الإجمالي` → slide 4, TextBox 23, ID 24.
- B3 `مكتملة` → TextBox 28, ID 29.
- B4 `على المخطط` → TextBox 30, ID 31.
- B5 `متأخرة` → TextBox 25, ID 26.
- القالب المعتمد الحالي لا يحتوي مؤشرًا مرئيًا لحالة معاملات `لم تبدأ`. لا تُنشئ هذا المؤشر ولا تُعده إلى الشريحة، ولا تقرأ B6 حتى لو كان موجودًا في ملف Excel قديم.
- Read the sector table as A:D only → Chart 4, ID 5, `chart3.xml`. A=`القطاع`, B=`مكتملة`, C=`على المخطط`, D=`متأخر`. إذا بقيت سلسلة legacy باسم `لم تبدأ` داخل chart XML أو الـembedded workbook/cache رغم عدم ظهورها في الشريحة، احذفها أثناء التحديث والتحقق. يجب أن تنتهي بيانات المعاملات بثلاث سلاسل فقط: `مكتملة`, `على المخطط`, `متأخر`.
- The updated slide 4 already includes a five-column `المعاملات المتأخرة` table pattern (ID 4), below its indicators and chart. Read numbered transaction records from the matching Excel section. Populate the template's displayed column order: `تاريخ الإنجاز المخطط له | الجهة الوارد منها المعاملة | تاريخ إنشاء المعاملة | موضوع المعاملة | رقم المعاملة`. The Excel `القطاع` field has no destination column in this table, so omit it from the consolidated PowerPoint table while retaining it in Excel. Remove the table with its title and headers when the section is empty. Keep the template's approved table color and styling on continuations.
- A subsequent split-by-sector run cannot recover the omitted transaction sector from the PowerPoint alone. When delayed transactions exist, that workflow must also receive the transactions Excel or an explicit sector mapping; never infer a sector from the transaction text.
- Preserve `رقم المعاملة` as text; do not append `.00` and do not remove leading zeros. Preserve dates and their calendar exactly as supplied by Excel. Do not apply any new deduplication to the consolidated transaction output.

## Shared elements and chart updates

- Cover, slide 1: TextBox 18, ID 19 inside Group 1 contains the report creation date. The title `تقرير الإنجاز الأسبوعي` is fixed and must not be renamed.
- Slide 5 is a fixed divider. Its two links must point to the first task-detail slide and the first project-detail slide. In the current approved template these are slide 6 for task details and slide 8 for Suhail project details. Update the links when continuation slides are inserted.
- Slide 11 is the fixed thank-you slide in the current approved template and must remain the final slide after all insertions.
- Update page numbers as text only while preserving their position and formatting; do not modify the master to redesign numbering.

The three charts are native editable PowerPoint charts. Each chart has an Excel workbook embedded inside the PowerPoint file, and the chart must depend only on that embedded workbook with no link to an external Excel file. When updating stacked bar chart data, keep the required sector/category rows and transfer only non-zero numeric values; when a status value is zero, leave the corresponding cell blank in the embedded workbook instead of writing `0`. Embedded workbook mapping:

- chart1.xml: ppt/embeddings/Microsoft_Excel_Worksheet.xlsx، Sheet1!A1:E9 في القالب.
- chart2.xml: ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx، Sheet1!A1:E10.
- chart3.xml: ppt/embeddings/Microsoft_Excel_Worksheet2.xlsx، use A1:D9 after removing the transaction `لم تبدأ` series.

Verify each chart's actual relationship to its embedded workbook instead of assuming a filename when the template version changes. The chart relationship must be internal and point only under `ppt/embeddings`; chart or embedded-workbook relationships must not contain any external Target to the source Excel file. Update the embedded workbook, series references and ranges, category-name caches, numeric caches, and point indexes in chart XML consistently, leaving zero-value cells blank in both the workbook and chart cache. Never replace a chart with an image or place an image over it. Sector order in the template may differ from Excel; preserve the existing display order for matching sectors, refresh their names from the source, and add new sectors without losing them. A blank status-count cell in the sector table means zero; a blank required summary indicator means missing source data. إذا زاد عدد القطاعات، يبقى الرسم داخل شريحة الملخص نفسها ولا تنشئ نسخة أو شريحة استمرار للملخص؛ حدّث بيانات الرسم داخل نفس العنصر مع الحفاظ على أبعاده وتصميم الشريحة.

Build a status-style map from the approved template before populating any detail table, keyed by the normalized status text rather than by row position or the copied table's existing fill. Apply the mapped status-cell formatting again after inserting every record so copied row formatting cannot override the true status. Use these approved template fills wherever those statuses appear: `مكتملة` = light blue `#DCE6F2`, `على المخطط` = light green `#EBF1DE`, and `معلق/معلقة` = light gray `#F2F2F2`. Preserve the corresponding template font, border, alignment, and margins. Never infer a status color from the destination row, alternating-row formatting, table position, or the previous week's value. For `متأخر/متأخرة` or any new status without a verified template exemplar, identify the exact approved status style before changing the design; do not invent a color. During validation, equal normalized statuses must have identical cell formatting throughout the report.

## Table flow and overflow handling

Duplicate the closest complete table style based on column count and table purpose; never copy Excel formatting into PowerPoint. Titles, headers, and row text may be updated. Remove copied tables or stale data that are not needed.

Lay out tables from top to bottom as one continuous visual flow:

- Render the slide before deciding where a table ends or where the next table begins. PowerPoint text wrapping and displayed row height are authoritative; calculated frame dimensions are not.
- Place each next table immediately below the visible bottom edge of the previous table, leaving only a very small clear gap. The gap should read as separation between two tables, not as an empty section of the page.
- On any slide containing two or more consecutive tables, compare every inter-table gap in the full-size rendered slide. All gaps on that slide must appear visually equal and very small. A slide fails validation if one pair has a larger gap while another pair touches or overlaps.
- Never overlap table frames to manufacture a small gap. Text wrapping can make the visible table extend beyond its nominal frame, so position the next table from the rendered bottom edge of the preceding table.
- Use the available page height fully. The last table should end just above the visible footer elements, with only the small clearance needed to avoid touching the footer line, logo, or page number.
- Do not split a table while there is visible room for the next complete row. Move a row to a continuation slide only after rendering confirms that the entire row cannot fit above the footer.
- Preserve the exact source order across page boundaries. Do not move a record to another slide merely to balance whitespace, equalize page fill, or improve visual symmetry. Pagination may move only the next complete record when it cannot fit above the footer after rendering.
- Never split one record across slides. A continuation table repeats its title strip and column headers, then starts with the next complete record.
- After a continuation table, place the following table directly beneath it with the same small visual gap. If its title, header, and at least one complete record cannot fit, move that whole table to the next slide.
- For summary slides, keep indicators and charts only on the first slide. If summary tables genuinely overflow, use a continuation slide containing only the continued and subsequent tables.
- For task and Suhail detail slides, pack consecutive tables into the remaining visible space. Remove a detail slide that becomes empty.
- Preserve fonts, column widths, cell margins, colors, borders, and row styling. A row may become taller when its exact source text requires it; never shrink the font or compress the row to force it onto the slide.
- After every table move or split, render the affected slide and the following slide at full size. Correct excessive gaps, premature splits, footer clearance, clipping, and overlap before delivery.
- Do not use a fixed gap measurement, fixed footer coordinate, fixed character-count estimate, or fixed number of records per slide as a substitute for visual inspection.
- Do not repeat records or count continuation rows twice. If one complete record cannot fit on an otherwise empty continuation slide with the approved style, stop and ask about that specific record.


## Reference validation example

The current template is a design/reference snapshot only; never use its displayed values as weekly defaults. For this template snapshot, the visible summary values are:

- Tasks: 12 total, 2 `مكتملة`, 10 `على المخطط`, 0 `متأخر`, 0 `معلق`.
- Suhail: 20 total, 4 `مكتملة`, 16 `على المخطط`, 0 `متأخر`, 0 `لم تبدأ`.
- Transactions: 589 total, 334 `مكتملة`, 255 `على المخطط`, 0 `متأخر`; `لم تبدأ` is not a valid displayed transaction status.

These values are for validating template mapping only. Runtime Excel remains the sole source of weekly data, and the PDF must never be used as a data source.


## Approved shared status colors (user confirmation, 29 September 2026)

The Monday status colors also apply to this weekly report: completed `DCE6F2`, on-plan `EBF1DE`, delayed `F2DCDB`, suspended `F2F2F2`, and not-started `F2F2F2`. The current weekly master has completed/on-plan detail examples. For its missing examples, copy the corresponding weekly task/project base status cell and apply the approved fill. Preserve that weekly cell's font, size, borders, margins, and alignment. This explicit approval resolves the requirement above to identify an approved style before using missing status examples. Do not import Monday typography or change either master. Reject new, unsupported statuses and conflicting template examples.

### Suhail detail update formatting

In `ما تم حتى تاريخه`, bold only the leading `تاريخ التحديث` label and its supplied date, including when regular body text follows on the same line. Keep the rest regular. Never invent a date. Use native PowerPoint bullets only for two or more explicitly marked source items (bullet, dash, or numbered item at the beginning of a line); remove their textual list markers so bullets are not duplicated. Preserve item wording and order. A single item, unmarked prose, line wrapping, and paragraphs remain plain text; do not infer a list from sentence punctuation. Preserve Abar, RTL, font size and cell margins. This applies to project details only.

## Table column alignment

Apply to both summaries and details, to column headers and every body paragraph. Center القطاع, تاريخ الإنجاز المخطط (including الانجاز spelling), تاريخ البداية, تاريخ النهاية, الحالة and حالة المشروع. Right-align المهمة, المشروع/اسم المشروع, ملاحظات, التحديث, طلب الدعم, التحدي and ما تم حتى تاريخه. Preserve merged table titles, numbering, other columns, vertical alignment, RTL, colors and emphasis. Font size is 11 pt throughout.
