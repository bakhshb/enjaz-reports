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

