# Design rules

The master deck is the template. When an optional table is absent, use the same
approved shape from `weekly-consolidated-report/assets/weekly-report-master.pptx`.
Reuse its shapes; change text, status fills,
accent colour and geometry, nothing else. No new layouts, no new charts, no new
KPIs, no analysis or recommendations added anywhere.

## Page geometry (A4 portrait, 8.27" × 14.7")

| item | value (EMU) |
|---|---|
| content column x / width | `474245` / `6623050` |
| first table top (where the chart was) | `1780255` |
| gap between stacked tables | `200000` |
| everything below this y is body | `1700000` |
| footer starts around | `12687759` |

Tables stack from the top with a fixed gap. Leaving whitespace at the bottom is
fine — the master does the same. Don't shrink fonts to force content onto one
slide; continue onto a second slide with the identical design if it truly won't
fit.

## Navigation

The master has four tabs; sector decks keep three (المهام / مشاريع سهيل /
المعاملات) and drop التفاصيل. The tabs live in the slide **layouts**, not the
slides, so edit `1_Title Slide`, `2_Title Slide`, `5_Title Slide`.

* the التفاصيل tab sits at x `428628` — delete both its shapes
* the remaining three sit at x `5741129`, `3970297`, `2199462`
* shift each left by `1012345` so the group is centred on the page

Keep the active-tab colours as they are: black on المهام, green on سهيل, navy on
المعاملات. Hyperlinks pointing at deleted slides must be stripped, not retargeted.

## Status colours

Cell fill on the الحالة column — the master's own legend:

| status | fill |
|---|---|
| مكتمل | `DCE6F2` |
| على المخطط | `EBF1DE` |
| متأخر | `F2DCDB` |
| لم يبدأ / معلق | `F2F2F2` |

## Title rows

A table's title row carries **no border and no fill at all**. The only visible
element is the small accent bar (`lnL`, width `57150`, which PowerPoint mirrors to
the right edge in an RTL table). Its colour encodes the table:

| table | accent |
|---|---|
| المتأخرة (tasks / projects / transactions) | `87452B` |
| المكتملة | `174766` |
| على المخطط | `0A8062` |
| أبرز التحديثات · الدعم المطلوب | `D9A86A` |
| المعلقة · لم تبدأ | `A6A6A6` |

These are the same colours the master uses on its KPI numbers, which is why they
match without being sampled by eye.

## Table sources

| table | cloned from | columns |
|---|---|---|
| tasks by status | matching titled status table on ملخص المهام, or the approved asset when absent | as in that table, including ملاحظات or التحديث |
| طلبات الدعم | its own table on ملخص المهام | as in the master |
| أبرز التحديثات | its own table on ملخص مشاريع سهيل | `#`, المشروع, القطاع, التحديث |
| التحديات | its own table on ملخص مشاريع سهيل | `#`, المشروع, القطاع, التحدي |
| projects by status | a تفاصيل مشاريع سهيل table, rescaled to the content width | `#`, اسم المشروع, تاريخ البداية, تاريخ النهاية, الحالة, ما تم حتى تاريخه |
| المعاملات المتأخرة | its own table, rows replaced | as in the master; sector is joined from final Excel when absent |

Project status tables take the gold-bar title row from the أبرز التحديثات table
(expanded to six columns) so every table on that slide shares one title style.

On-plan task tables clone the six-column task-detail pattern: `# | المهمة | القطاع | تاريخ الإنجاز المخطط | الحالة | ملاحظات`. Keep the sector and add its existing planned date in the detail-table order. Other task status tables retain their summary patterns. Preserve detail column proportions when scaling to the content width.

## Cover

Group order top to bottom: the master's own title, the sector name **plain, with no
parentheses**, date, gold rule. Room for the sector line is made by shifting everything below the
title down `700000` and growing the group, then recentring it. The sector textbox
is cloned from the date box at 16pt white, `spAutoFit` removed so a long name
can't push into the date.

## Project rows are copied, not retyped

A project row is cloned as XML straight out of the master's تفاصيل مشاريع سهيل
table; only the `#` cell is rewritten and the الحالة fill re-applied. That is what
keeps the ما تم حتى تاريخه column looking right: in the master, the leading
`تاريخ التحديث 3 سبتمبر 2026` is bold and everything after it is not. Rebuilding
the cell from a plain string collapses it to one run and the whole paragraph goes
bold — a visible regression the user will ask about. Never retype these cells.

## Table geometry must be self-consistent

Every `<a:tr>` carries its measured height and the frame's `cy` equals their sum.
This isn't cosmetic: PowerPoint recomputes row heights with the real Abar Mid
metrics, and a frame that claims to be shorter than its content ends up under the
next table, whose frame then swallows the clicks — the user can no longer put the
cursor in the lower rows. `refine_positions` handles it; anything you add that
touches heights must preserve `sum(row h) == frame cy`.

## Text handling

Keep project, task and transaction wording exactly as in the master, `<a:br/>`
line breaks included. Shorten only if space genuinely demands it, and keep the
full meaning when you do. Renumber `#` per table (`1, 2, 3…`; transactions use
`01, 02, 03…`).
