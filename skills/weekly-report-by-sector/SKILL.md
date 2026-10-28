---
name: weekly-report-by-sector
description: Split the ministry-wide weekly report deck (تقرير المهام / التقرير الأسبوعي, Ministry of Hajj and Umrah) into one standalone PowerPoint per وكالة / قطاع / مركز, reusing the master deck's exact design. Use this whenever the user uploads that weekly report and asks for per-sector reports, تقرير لكل وكالة, تقرير مستقل لكل قطاع, تقسيم التقرير الأسبوعي, splitting a consolidated report by جهة, or a sector-level version of the weekly deck — including when they re-describe the requirements from scratch instead of naming this skill, and including follow-up rounds of design tweaks on decks this skill already produced.
---

# Weekly report, split by sector

Turn one consolidated weekly deck into a standalone report per جهة. The master

deck supplies the report content and design. If an optional status table is absent

from the populated deck, clone its approved shape from the bundled weekly master

asset. When its delayed-transactions table

has no sector column, the final transactions Excel supplies the sector for those

exact rows. Never invent a

number, never redesign anything, never introduce a layout the master doesn't

already contain.

## Run it

Prefer `uv` for isolated dependencies. If resolution is unavailable but a compatible Python already has the declared packages, use that interpreter and record the fallback. Do not modify the host Python environment.

```bash
uv run --with python-pptx --with lxml --with openpyxl --with pdfplumber --with pillow scripts/split_report.py MASTER.pptx --transactions FINAL_TRANSACTIONS.xlsx --outdir out
uv run --with python-pptx --with lxml --with openpyxl --with pdfplumber --with pillow scripts/qa_report.py out --master MASTER.pptx --transactions FINAL_TRANSACTIONS.xlsx --render qa
```

The QA command requires the same source master used for the split. Pass the same
`--transactions`, `--aliases`, and `--only` arguments to QA when they were used for
the build. QA compares every output record and its status grouping/order against
the source rather than trusting manifest totals. It also rejects broken internal
relationships, external links, and unused slide parts left in the package.
Programmatic success does not replace the mandatory render and visual review.

`--only <name>` builds a single sector, which is the fast way to test a change

before rebuilding all ten.

When LibreOffice is unavailable, run the splitter with `--no-refine` and run

`qa_report.py` without `--render`. The initial pass still sets each table's row

heights to match its frame and creates continuation pages. Open the resulting

deck in PowerPoint for a visual check before delivery.

`--transactions` is required when delayed transactions use the approved five-column

table (which omits sector). Use the **final** `المعاملات` Excel summary from the

same reporting week. The splitter matches all five visible transaction fields to

the Excel row; it stops if a row is missing or matches multiple sectors. If there

are no delayed rows, or an older deck includes a sector column, this argument is

optional.

`split_report.py` reads the master, groups everything by sector, writes one deck

per sector into `out/`, plus `out/manifest.json`. It also re-encodes the master's

full-bleed background JPEGs at the same pixel size, which takes a deck from ~9.4 MB

to ~2.7 MB with no visible change. `qa_report.py` validates the

files and prints the row-count / leakage / empty-table checks. `--render` also

writes one composite JPEG per deck (all slides side by side) — review those

rather than previewing slides one at a time.

Then present the generated `.pptx` files as downloadable attachments.

## What the master looks like

Slide roles are detected by content, not by slide number, so a reordered master

still works. Detection keys on table header text:

| role | how it's found |
| --- | --- |
| cover | first visible slide |
| ملخص المهام | chart slide with `إجمالي المهام` KPI; titled chartless continuation tables are also read |
| ملخص مشاريع سهيل | chart slide with `إجمالي مشاريع سهيل` KPI; titled chartless continuation tables are also read |
| المعاملات | slide(s) with a table whose header has `رقم المعاملة` |
| تفاصيل المهام | six-column table with `المهمة` and `الحالة` in the header |
| تفاصيل مشاريع سهيل | six-column table with `اسم المشروع` in the header |
| شكرا | last visible slide containing `شكرا` |

Hidden slides (`show="0"`) are template leftovers — the master usually carries one

after شكرا — and are ignored throughout.

The master changes shape from week to week, so treat these as optional:

- **المعاملات may be absent.** When no table has a `رقم المعاملة` header there are
    
    no delayed transactions that week and every deck simply has no transactions
    
    slide. The script says so and carries on; do not fabricate the section.
    
- **ملخص المهام status tables gain and lose columns.** Each status table is

    selected by its title, and every output cell is filled from the matching

    detail row. The `طلبات الدعم` table is read independently and appears only

    for sectors with support requests. المهام المعلقة carried a
    
    `التحديث` column in September 2026. Any status table on that slide whose master
    
    version has the extra column is cloned from that table, so the sector deck keeps
    
    its content but labels the output column `ملاحظات`. `التحديات` and `أبرز التحديثات` are also read independently from Suhail

    summary pages and appear only when that sector has rows.
    

If detection fails the script stops with a clear message — fix the keys rather

than falling back to hard-coded indices.

## Deck structure per sector

Omit the `القطاع` column from every output table in sector reports, including tasks, support, Suhail updates/challenges and any legacy transaction table that contains it. Continue using the source sector values internally to assign records to the correct report. Keep the sector name on the cover. Remove the column before measuring/paginating, redistribute its width across the remaining fields without shrinking fonts, and preserve all records, dates, notes, numbering, status formatting and RTL. This exception applies only to sector reports; consolidated and Monday reports retain their own columns.

Keep the `المهمة` column exactly the same width in every task-status and support table, using the on-plan task-detail pattern after removing sector/status as the reference. Keep every table at its full content width. Give the space of a missing date column to `ملاحظات` (or `طلب الدعم` in the support table), so notes are wider in tables without dates. Normalize widths before measuring and paginating, including continuation tables. QA rejects inconsistent task-column widths. Label the suspended-task note column `ملاحظات`, preserving the source `التحديث` text.

Cover → ملخص المهام → ملخص مشاريع سهيل → المعاملات → شكرا.

A section slide is **deleted entirely** when the sector has no data for it, and a

status table is **deleted entirely** when no row has that status. Never emit an

empty table, a "لا توجد بيانات" line, or a zero inside a table. A sector with

nothing in any of the three sections gets no deck at all.

Cover keeps the master's own title (`تقرير الإنجاز الأسبوعي` in the current

master), then the sector name on its own line **without parentheses**, then the

master's date, unchanged.

Table order — this sequence matters, it's how the report is read:

- Tasks: المتأخرة → المعلقة → المكتملة → على المخطط → طلبات الدعم
- Projects: أبرز التحديثات → التحديات → المتأخرة → على المخطط → لم تبدأ → المكتملة

### Formatting أبرز التحديثات

When an update contains a standalone first line in the form `تاريخ التحديث [يوم] [شهر] [سنة]`, keep that line on its own and make **only that entire date line bold** in the PowerPoint cell. Keep every line of the update body after it regular (not bold), including any other dates mentioned in the body. Implement this as separate rich-text runs; never apply bold to the whole cell. If the date line is absent, keep the full update text regular.

KPI cards show the sector's own counts only. Drop the cross-sector comparison

chart — a single-sector report has nothing to compare against.

**Transactions carry only one KPI:** `إجمالي المعاملات المتأخرة`. The master has

no per-sector totals for completed / on-plan / not-started transactions, so

computing them would be fabrication. Recentre the surviving gold number in the

KPI card and delete the rest of the card's contents.

Start the transactions table at the same vertical content origin as task and project tables (`TOP_Y = 1780255` EMU), directly below the KPI area. Do not reuse its lower vertical position from the consolidated master. Keep its horizontal position and width; continuation tables also start at this common top. QA rejects a misplaced transaction table.

Make the title `المعاملات المتأخرة` explicitly bold on every page, including continuations. QA rejects a nonbold title. Preserve the body and column-header formatting.

## Sector names

`references/sector-aliases.json` maps every spelling to one canonical name.

Hamza and definite-article variants (`تطوير الاعمال` / `تطوير الأعمال`,

`امتثال أعمال` / `امتثال الأعمال`) always merge.

The المعاملات table often uses **old ministry agency names** that don't match the

ones in المهام / مشاريع سهيل. Merge those into the current names — otherwise one

agency gets two half-empty reports. When a legacy name has no clear counterpart

(e.g. `وكالة الوزارة للخدمات المشتركة`), keep it as its own sector, and when the

mapping is genuinely ambiguous, ask rather than guess. Add whatever you settle on

to the aliases file so next week runs clean.

## Design rules

Read `references/design-rules.md` before changing anything visual. In short:

every table, row and KPI element is cloned from a shape that already exists in

the master; only text, status fills, accent colour and geometry change. The three

nav tabs (التفاصيل removed) sit centred on the page as one group. Title rows carry

no border and no fill — only a small accent bar coloured by table type.

## Pitfalls that produce corrupt or ugly files

These all cost a rebuild once. They are handled in the scripts; keep them handled.

- **Adding a `<a:tc>` after `<a:extLst>`** in a row breaks schema order, PowerPoint
    
    ignores `hMerge`, and the unmerged cells show black default borders. Insert
    
    before `extLst`.
    
- **Out-of-order `<a:tcPr>` children** make PowerPoint discard the whole block and
    
    fall back to black borders. Order is `lnL, lnR, lnT, lnB`, then the fill.
    
- **`.//a:ext`** matches the `extLst` inside `cNvPr` before the real
    
    `spPr/a:xfrm/a:ext`. Always go through `spPr` → `xfrm`.
    
- **python-pptx shape proxies are new objects each iteration** — compare by
    
    `shape_id`, never with `is`.
    
- **Deleting slides via python-pptx** collides part names on save. Remove the
    
    `<p:sldId>` entries from `presentation.xml` and retain only package parts
    reachable from the root relationships, as the bundled splitter does,
    
    then strip layout→slide back-reference rels and the nav `hlinkClick` entries
    
    that now dangle, or the file opens as repaired.
    
- Row heights auto-grow, and a table whose declared geometry lies about its size
    
    becomes uneditable in PowerPoint. If the row `h` values are stale minimums while
    
    the frame `cy` is something else, PowerPoint recomputes with the real Abar Mid
    
    metrics, the table grows past the frame below it, and clicking a lower row hits the
    
    next table's frame instead — the user lands in the wrong cell. The refine pass
    
    therefore measures **every row**, writes the real height onto each `<a:tr>`, sets
    
    frame `cy` to the sum, and restacks. Keep `sum(row h) == frame cy` in anything you
    
    add.
    
- **Stacked tables overlap in the render**, which would corrupt that measurement, so
    
    each refine pass sends every table except the one being measured off the page and
    
    renders again. That's one render per table position — slower, but the measurement
    
    is unambiguous.
    
- **The preview's font isn't Abar Mid**, so measured heights run slightly short of
    
    PowerPoint's. `SAFETY` (1.12) covers the difference. Lowering it risks overlap;
    
    raising it wastes page.
    

## Before delivering

The refinement and QA render paths use the shared `report_common/render.py` exporter. It materializes RTL column order only in a disposable LibreOffice export copy, keeping the editable source PPTX unchanged. Verify numbering on the right in every rendered table; an RTL XML property alone is insufficient. Run `tests/check_rendered_pdf.py` on exported PDF/PPTX pairs to check actual numbering coordinates, fonts and page counts.

First smoke-test one sector through `uv run` with `--only <name>`. Then run the full build and `qa_report.py` through `uv run` using the commands above. Confirm: sector name right on every cover; totals match the

master's detail slides; each table holds only its own status; no empty tables; no

other sector's data anywhere; every sector with data got a deck and no sector

without data did. It also flags any table whose bottom passes the footer line —

when that fires, move the overflowing rows onto a second slide of the same design

rather than shrinking the font. Then eyeball the composites.

Two checks report expected noise: a sector name legitimately appears inside another

sector's project update when the master's own text names the agencies it coordinates

with, and the footer check reads the safety margin as well as the content.


## Updated template acceptance

- Use the user-approved current master and matching PDF reference. Preserve Abar typography when populating empty cells and duplicating rows or slides. Verify the generated PowerPoint and exported PDF, especially transaction bodies and ملاحظات columns.
- Omit a sector from charts only when every plotted status is blank or zero. Retain any sector with a positive count in any plotted status. Apply this to chart caches and embedded workbook categories, without changing source workbooks, totals, or detail records. For entirely empty data, show no named sector or bar and retain zero KPI indicators.
- Acceptance requires normal, empty, new-sector and overflow cases to reconcile with source records, open without PowerPoint repair, and pass final visual inspection for Abar, readable text, no clipping, no overlap and correct pagination.

- Verify actual font usage in the exported PDF, not just font declarations in PowerPoint. On the validated Windows host, PowerPoint SaveAs PDF substituted Calibri even for the unchanged approved master; Adobe PDFMaker preserved Abar and matches the supplied reference export route. If an exporter substitutes fonts, use a verified available exporter and repeat the visual check. Never mark font acceptance passed from successful export alone.

## Planned dates for delayed and on-plan tasks

For `المهام المتأخرة` and `المهام على المخطط`, clone the master's task-detail table and omit `القطاع` and `الحالة`, producing four columns in this order: `# | المهمة | تاريخ الإنجاز المخطط | ملاحظات`. Keep both titles bold on every page, including continuations. Copy the planned date exactly, including blank dates. Retain detail typography and remaining column proportions; fit the existing content width and paginate rather than shrink text. Completed and suspended task tables retain their summary fields with the sector column omitted and without this added date. Project tables retain their status column.

Acceptance: no output table contains a sector column; delayed and on-plan task tables show the planned date between task and notes, with no status column. Every date matches its source detail row. Preserve task order, numbering, counts and notes. QA rejects a restored sector header, a status column in either dated task table, or a nonbold on-plan title. The output must pass source reconciliation and final visual review for clipping, wrapping and footer clearance.

## Approved table typography

Every generated PowerPoint table must have native RTL table direction, including support, project status, transactions and continuation tables. Right-aligned Arabic text alone is insufficient. Keep numbering at the right and preserve field content, column widths and merged titles when converting an LTR source pattern. QA must reject any table without RTL enabled.

Apply to every summary/detail/status table, including patterns cloned from the bundled master or a supplied 10 pt report. All table text must be **11 pt**: titles, column headers, body, status, notes and empty-cell defaults. This is the explicit exception to preserving the source's font size/alignment.

Apply horizontal alignment to the column headers and body paragraphs:

- Center تاريخ الإنجاز المخطط / تاريخ الانجاز المخطط, تاريخ البداية, تاريخ النهاية, الحالة and حالة المشروع. The sector column is omitted from output tables.
- Right-align المهمة, المشروع / اسم المشروع, ملاحظات, التحديث, طلب الدعم, التحدي and ما تم حتى تاريخه.
- Preserve merged title alignment, numbering and unspecified columns, vertical alignment, RTL, Abar, bold emphasis, bullets and colors.

The splitter applies the shared `report_common/table_typography.py` rules before refinement, and QA rejects a different font size or specified column alignment. Install this skill with the sibling `skills/report_common/` directory, or use the standalone bundle from `tools/package_skills.py`, which carries `scripts/report_common/` and the weekly master in `assets/`. Measure/refine and repaginate at 11 pt; never shrink text to fit.

Acceptance: test a 10 pt source with incorrect alignment and a current 11 pt source; both must produce 11 pt tables with the above alignment, reconcile all source records/counts/statuses/planned dates, and pass native visual review for clipping/overlap/footer clearance. Deliberately changing an output to 10 pt or right-aligning a centered column must fail QA. Templates and source workbooks remain unchanged.
