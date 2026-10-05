# Report business rules

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
- when a new task source appears, add its separately titled table below the previous table if a header and at least one complete row fit; otherwise continue on a new page;
- keep the approved title/header/table geometry and typography.

The Excel task detail sheet has a seventh field, `الجهات ذات العلاقة`. The user intentionally removed this column from the PowerPoint template; do not transfer it to the deck or append it to `ملاحظات`.

Size rows from their actual text at the approved font and column widths, not multiples of the master example height or its sample row count. Pack successive tables on the same page with a 9 pt gap when they fit. Count each table title/header once per chunk. Preserve source order, repeat headers on continuation pages, and reserve 55 pt above the slide bottom for the footer. Never shrink fonts or clip content to reduce page count.
If a single record is taller than an otherwise empty approved page, stop with a clear error; never silently cap its estimated height.

### Suhail summary

Update the Suhail KPIs and native chart from the current summary sheet. Read its three independent sections: `أبرز التحديثات`, `المشاريع المتأخرة`, and `التحديات`. Populate a matching four-column template table only when that Excel section contains records; otherwise remove the table with its title and headers. Preserve each section's exact Excel content, including notes and challenge text. Flow overflow onto continuation summary pages, repeating the applicable title and headers. Keep KPIs and the chart on the first summary page.

### Suhail project details

Read every sector from `تفاصيل مشاريع سهيل` in source order.

There is **no whitelist of sector names** and no fixed project count. The sector names present in the master are examples, not limits.

Use the approved six-column Suhail-detail page as the reusable pattern. For every sector:

- share available page space with the preceding sector table, keeping separate sector titles/headers and a 9 pt gap; create a continuation only when the next complete row and its headers will not fit;
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
- Remove inherited bullets from ordinary dynamic cells. Apply the Suhail detail exception below.
- In `ما تم حتى تاريخه`, put the existing `تاريخ التحديث` and its day/month/year date on the first content line in bold. Body text is not bold. Ignore leading/trailing empty paragraphs. Two or more nonempty source paragraphs after this line are separate update points: render native bullets using the approved master bullet paragraph. A single sentence/paragraph stays unbulleted. Do not split prose by punctuation or visual wrapping, invent a date, or rewrite source wording.
- All tables must explicitly use RTL, with `#` at the right. When converting a summary pattern stored left-to-right, reverse its underlying columns and merged groups as well as enabling RTL, so visible column order and approved widths remain unchanged.
- The cover date is the runtime creation date in the runtime timezone, formatted like `27 سبتمبر 2026`.


## Summary layout

Apply content-based row sizing and continuous packing to both summary sections too. A nonempty table does not automatically require a new slide. Preserve the first-page KPI/chart area and approved table width; only the summary tables continue. Short rows must not inherit the height of a tall master example.
