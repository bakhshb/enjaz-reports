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

Use `uv` as the supported runner so the skill uses an isolated, reproducible dependency set. Do not call these scripts with the system Python.

```bash
uv run --with python-pptx --with lxml --with openpyxl --with pdfplumber --with pillow scripts/split_report.py MASTER.pptx --transactions FINAL_TRANSACTIONS.xlsx --outdir out
uv run --with python-pptx --with lxml --with openpyxl --with pdfplumber --with pillow scripts/qa_report.py out --master MASTER.pptx --render qa
```

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
    
    it. `التحديات` and `أبرز التحديثات` are also read independently from Suhail

    summary pages and appear only when that sector has rows.
    

If detection fails the script stops with a clear message — fix the keys rather

than falling back to hard-coded indices.

## Deck structure per sector

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
    
    `<p:sldId>` entries from `presentation.xml` and run the pptx skill's `clean.py`,
    
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

First smoke-test one sector through `uv run` with `--only <name>`. Then run the full build and `qa_report.py` through `uv run` using the commands above. Confirm: sector name right on every cover; totals match the

master's detail slides; each table holds only its own status; no empty tables; no

other sector's data anywhere; every sector with data got a deck and no sector

without data did. It also flags any table whose bottom passes the footer line —

when that fires, move the overflowing rows onto a second slide of the same design

rather than shrinking the font. Then eyeball the composites.

Two checks report expected noise: a sector name legitimately appears inside another

sector's project update when the master's own text names the agencies it coordinates

with, and the footer check reads the safety margin as well as the content.

