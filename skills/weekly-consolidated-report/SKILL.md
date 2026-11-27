---
name: weekly-consolidated-report
description: Populate either approved Ministry of Hajj and Umrah weekly consolidated PowerPoint model from the final tasks, Suhail projects, and transactions Excel summaries. The digital-experience model also uses its current PDF. After delivery, replies with a short and a long copyable Arabic message to the Minister built from the delivered report. Ask which model to use when the user has not chosen one. Does not build upstream Excel summaries or split by sector.
---

# Weekly Consolidated Report

## Purpose and authority

Create the report from current final Excel inputs and one of this skill's approved masters. Excel controls the existing sections; the optional digital-experience PDF controls its added section; the chosen master controls design. Read [business rules](references/business-rules.md) before generation and [acceptance checks](references/acceptance.md) before validation. Preserve this report's own mappings, ordering, omissions, pagination, and status rules. Sample records and slide counts are not runtime requirements.

## Choose the model

When the user sends the Excel files without choosing a model and is there to answer, ask once: **«أي نموذج تريد: التقرير الموحد الحالي، أم التقرير الموحد مع نضج التجربة الرقمية؟»** If the user already selected one, use it without repeating the question. When there is no reply, or the run has no one to ask, do not wait: build the digital model when the digital-maturity PDF is among the inputs and the standard model otherwise, and say in the final reply which model was built and why.

- **التقرير الموحد الحالي** (`standard`): the existing master and three Excel inputs, with unchanged business rules.
- **التقرير الموحد مع نضج التجربة الرقمية** (`digital`): the second bundled master and the same three Excel inputs, plus the current digital-maturity PDF (the landscape «horizontal» export). The PDF is a data source only: page one fills the summary; every later page belongs to one sector, and a sector may span several source pages titled «(k/n)». Merge each sector's pages into one record set and re-flow it into the approved portrait detail slides; never reproduce the source pagination, one slide per PDF page, or a half-empty slide. Read [digital-maturity mapping](references/digital-maturity.md) only for this model. A missing PDF is a missing required input: ask for it (see «When to stop»).

## Inputs and output

- ملخص وتفاصيل المهام الاستراتيجية.xlsx: ورقتا ملخص المهام وتفاصيل المهام.
- ملخص_مشاريع_سهيل.xlsx: ورقتا ملخص مشاريع سهيل وتفاصيل مشاريع سهيل.
- تقرير_المعاملات_الموحد.xlsx: ورقة المعاملات.
- قالب PowerPoint المعتمد الحالي: `assets/weekly-report-master.pptx`، مع PDF الموافق `assets/weekly-report-reference.pdf` مرجعًا بصريًا (تقرير 8 أكتوبر 2026، مُصدَّر من PowerPoint). هذا القالب هو مصدر التصميم المعتمد الحالي.
- قالب النموذج الرقمي الإضافي: `assets/weekly-report-digital-master.pptx`. يبقى النموذج الحالي متاحًا دون استبدال. مرجعه البصري المعتمد للمخرج هو `assets/weekly-report-digital-reference.pdf` (تقرير 5 نوفمبر 2026، مُصدَّر من PowerPoint): يبيّن كيف يُعاد صبّ قطاعات ملف التجربة الرقمية الأفقي في شرائح التفاصيل العمودية، وجدول «أبرز التحديثات» في الملخص، وجدولي «أبرز التحديثات» و«التحديات» في نهاية كل قطاع، وهو مرجع تصميم لا مصدر بيانات.
- عينة المصدر الأفقي المعتمد لملف التجربة الرقمية محفوظة في المستودع: `tests/fixtures/digital-maturity-horizontal.pdf` (تحديث 1448/5/17هـ). يختبرها `tests/test_weekly_digital_maturity.py` استخراجًا ودمجًا وبناءً؛ أي تغيير في المستخرج يجب أن يبقي هذه العينة ناجحة.
- At runtime, use the new Excel files attached to the request. Restore the approved PowerPoint template from the assets attached to this skill; the user should not need to attach the template again. Ask for the template only if the skill assets are unavailable or fail integrity verification. Never substitute sample values.
- Primary output: `تقرير الإنجاز الأسبوعي بتاريخ الإنشاء.pptx`, saved as a new file. Create a matching PDF only when requested. Never modify the inputs or overwrite the original template.


Verify the bundled master size `9540224` bytes and SHA-256 `173d48ace27661a4900cc1d6428826e41515d6d4a764a26385171701ead5d96a` before use. Never fall back to an earlier master.

For `digital`, verify SHA-256 `4f44e50242617b4448ff4f9dcde2b29e609a3a37e89132a05ef41ee47aafef98` for its bundled master. Do not use the current model's master for this choice.

## Table typography

Every PowerPoint table must use native right-to-left (RTL) table direction in both models, including support, transactions, digital summary/detail/challenge tables and all continuations. Right-aligned text alone is insufficient. Keep numbering at the right; convert LTR-stored patterns together with their column widths and merged cells so every value stays with its field. Reject output tables without RTL enabled. Latin identifiers may retain LTR text within an RTL table.

All table text is **11 pt**, including table titles, headers, data, status cells and notes (and the Monday agenda). Apply this explicit size even when the master stores 10 or 14 pt. Preserve Abar, bold emphasis, RTL, fills and column widths. Adjust row heights and pagination when needed; never reduce the font to make a detail table fit.

Every table run is set in the master's Abar Mid for Latin, East Asian and Arabic text, including note cells whose master example names no font (they otherwise fall back to the theme font); the validator rejects any other table font.

Summaries are the one exception: each stays on its one slide, and when they do not all fit at 11 pt the summary slides together take compact cell margins and a shared smaller size, down to 10 pt by default (digital experience down to 8 pt). Margins and size are decided once for all summary slides so they look the same (see «When space is tight» and business rule 9, user request of 26 November 2026).

The same typography pass keeps a table sound when the user edits it by hand in PowerPoint: every paragraph ends with complete end-of-paragraph properties (11 pt Abar), so text typed into an emptied cell or an inserted row keeps the table's look, and every row has its own row ID. See «After a manual edit».

Apply the per-column horizontal alignment in the business rules to **summary and detail tables**, including headers: sector/dates/status centered; task/project names, notes, updates, support text and challenges right-aligned.

## Business rules

Apply every rule in [the report-specific reference](references/business-rules.md). Do not infer content, change source workbooks, or modify the original master. Preserve the report's section order, dates, blank cells, and current status meaning. Template samples are never weekly data.

## Generate and validate

1. Check required workbooks, sheets, headers, and the chosen approved master. For `digital`, render page one of the current PDF and extract a review file:

```bash
uv run scripts/digital_maturity.py --pdf "<digital.pdf>" --output "<review.json>"
```

Compare every extracted card percentage, sector/bar color and value, sector note, and page-one note with rendered page one. Review every merged sector in `detail_sectors` (title without suffix, four cards, every table row in source order, every source note from all of its pages) against the rendered source pages listed in its `source_pages`. Correct PDF text-layer reversal, numbers, codes, `(cid:…)` glyph gaps, or ligature errors in the review file; never invent missing values. Then classify each sector's notes into `updates` (أبرز التحديثات) and `challenges` (التحديات) by content, splitting a mixed note, and rewrite every update and challenge, including the summary `updates`, in formal government Arabic (فصحى رسمية، جمل خبرية قصيرة، بلا عامية): an update is the statement itself with no title, a challenge reads «عنوان قصير: النص»; no task codes, every number and date kept. (see [digital-maturity mapping](references/digital-maturity.md)). Set `reviewed` to `true` only after the full comparison and classification. The builder rejects unreviewed data, a PDF hash mismatch, a title that still carries «(k/n)», a row that still carries «(تتمة)» and unclassified source notes. These are corrections to your own review file: fix it and run again, without asking the user. A note that still carries a task code or a colloquial word, an update with a title, a machine-sounding phrase and a card status the template does not list do not stop the build: each is a notice (kind `wording` or `status`). When the notices file has wording notices about text you wrote yourself, correct the review file and rebuild once without asking the user, then report in the final reply whatever notice remains. Keep the review file as QA evidence outside the delivered slides. If the extractor rejects the PDF (layout, page count, table shape), the required source cannot be read: stop and report the exact error; do not hand-build the digital slides page by page.
2. Build the report with one command. `--measure` is the delivery path: it builds a first candidate, measures every row through LibreOffice, rebuilds with the measured heights and validates the result.

```bash
uv run scripts/build_report.py --measure --tasks "<tasks.xlsx>" --suhail "<suhail.xlsx>" --transactions "<transactions.xlsx>" --output "<report.pptx>"
```

For the second model, add `--model digital --digital-pdf "<digital.pdf>" --digital-data "<review.json>"`. Pass the same model, PDF and review-file options to `validate_report.py`.

The command writes next to the report:

- `<report>.layout.json`: the measured row heights (see «Measure and rebuild»);
- `<report>.pptx.notices.json`: every best assumption the build made instead of stopping (see «When to stop»). No file means no notices.

Its last line says `measured build` or `PROVISIONAL build, not for delivery`. A provisional build (no `--measure`, or LibreOffice could not run) uses estimated row heights; in PowerPoint its tables can overlap or run past the footer, so it is never delivered as final. When `--measure` cannot measure through LibreOffice, it falls back to a provisional build with a `layout` notice instead of stopping. A provisional deck goes to the user only as a draft, with that notice repeated to him.

3. The builder includes programmatic validation. If it finds two tables overlapping or a table frame reaching the footer on a slide, the builder repairs those slides once (frame heights from the rows, tables stacked again with the approved gap) and validates again; whatever remains is a `layout` notice naming the slide, and the final reply tells the user which slide to open and check. It never stops the build. Content findings (a damaged package, or a number, record or chart value that differs from the inputs) stop the build, because the report would state something the inputs do not. Every other finding (layout, style, navigation, a workbook that disagrees with itself) is published with the report and listed in the notices as «فحص التقرير: …»; correct what you can and rebuild. To check a file again:

```bash
uv run scripts/validate_report.py --report "<report.pptx>" --tasks "<tasks.xlsx>" --suhail "<suhail.xlsx>" --transactions "<transactions.xlsx>"
```

The standalone gate is strict (any finding fails); add `--lenient` to fail on content findings only. Failed builds preserve existing outputs and inputs; output paths cannot alias an input or master. A programmatic pass is not visual acceptance.

The gate also verifies inherited detail-layout artwork against the chosen master, visible master graphics, unique shape IDs including grouped children, and the chosen model's header buttons and destinations. Keep `scripts/weekly_artwork.py` with the installed generator and validator. A missing logo, footer, button, or layout link appears among the notices; rebuild from the approved master rather than deliver with one.

## Measure and rebuild

This step is required before delivery, and `build_report.py --measure` runs it for you. The builder's first-pass row heights are a cautious estimate, and a deck built without measurements must not be delivered. The measurement can also be run on its own, on a candidate built without `--measure`; it writes `<candidate>.layout.json` from a disposable copy and does not modify the deck:

```bash
uv run scripts/measure_layout.py --report "<candidate.pptx>" --tasks "<tasks.xlsx>" --suhail "<suhail.xlsx>" --transactions "<transactions.xlsx>"
```

For `digital`, add `--digital-pdf "<digital.pdf>" --digital-data "<review.json>"`. It measures every record row through `report_common/measure.py`: LibreOffice draws rows at their declared height and never grows them, so each row is rendered alone in a tall one-row table anchored to the top, and the distance from the row's top edge to its lowest glyph, plus the bottom inset, is its height. Install the template font (Abar Mid) first; with a substitute font the heights are wrong. For `digital` the digital-maturity rows are measured too (summary updates, every sector record, update and challenge row), and the measured rebuild requires them. The summary rows are measured at 11 pt with the template margins, and on a copy with the compact margins at every summary size from 11 down to 8 pt. The layout file stores these under `tables[<title>].compact` (keyed by size) and `digital.summary_compact`; the old template-margin smaller sizes (`sizes`, `summary_sizes`) are no longer written. The builder decides from these measured heights.

With the separate measurement, rerun the builder into a new path with `--measurements "<candidate.layout.json>"` (instead of `--measure`) and do the visual review on that rebuilt file only. A table whose measurements do not describe its current records falls back to its estimate with a notice; measure again. A `--measurements` file that is stale (an input or the template changed), malformed or incomplete (for example without the digital-maturity tables) does not stop the build: the measurements are dropped, a `layout` notice says why, and the build continues as provisional with estimated heights, exit code 0. Measure again with `--measure` before delivery. Measurements are bound to the exact input/template hashes, so any input or template change needs a new measurement. A measured rebuild is normally much shorter (the 2026-11-05 inputs went from 28 to 15 slides on the standard model and from 38 to 25 on the digital one). Measurements support the visual decision; they do not establish visual acceptance.

On Windows with PowerPoint, `scripts/measure_layout.ps1 -Report <candidate.pptx> -Tasks <tasks.xlsx> -Suhail <suhail.xlsx> -Transactions <transactions.xlsx>` writes the same `<candidate>.layout.json` (same hashes, same format) and can be used for this step instead. It measures a temporary copy whose record rows are first collapsed so PowerPoint fits them to their text. It covers the standard model only: it does not write the digital-maturity rows or the compact-margin data, so for `digital` use `measure_layout.py` (the measured digital rebuild refuses a layout file without digital rows). A layout file without compact data (an older one, or one from this script) does not stop the build: for a summary that overflows, the builder uses a safe upper bound computed from the 11 pt heights.

## Visual review

For LibreOffice PDF previews, use `uv run ../report_common/render.py --report "<candidate.pptx>" --output "<candidate.pdf>"`. The shared exporter uses a disposable copy with the RTL column order materialized for LibreOffice; the editable PPTX remains unchanged. Inspect every summary, detail, transaction and digital table: numbering must visibly appear on the right. Do not accept the RTL property alone. `tests/check_rendered_pdf.py` checks numbering using actual PDF coordinates, alongside fonts and page counts.

Render the measured, rebuilt and validated candidate and inspect every slide against this skill's template/reference, including Arabic, actual Abar usage, status fills, notes, chart labels, row continuity, table gaps and footer clearance. Verify native PowerPoint opening without repair and editable embedded charts. Follow [acceptance checks](references/acceptance.md); correct layout using the approved patterns and repeat affected checks. Final visual acceptance follows the last edit and programmatic check.

## When space is tight

Space never stops the report. The rule, in order:

1. **Each summary is one slide** (tasks, Suhail, transactions, digital experience), and the four must look the same, so margins and size are decided once for all of them. When every summary fits at 11 pt with the template margins, nothing changes. When any one of the four does not fit, every table on every summary slide gets the compact cell margins: the margin the template uses in the body cells of its detail tables, read from the template's detail pattern at build time (0.72 pt in the current masters, against 3.6 pt in the summary cells). It is set on all four sides of the body cells; title and header rows keep the template margins.
2. **Shared size.** With the compact margins, each of the three non-digital summaries (tasks, Suhail, transactions) finds the largest size from 11 down to the floor, in half-point steps, at which it fits. The shared size is the smallest of the three, and every table on those three slides uses it. The default floor is 10 pt.
3. **Past the floor, whole tables move to an appendix slide.** A non-digital summary that does not fit even at the floor with the compact margins keeps the tables that fit complete at the shared size, and every table that does not fit moves complete to an appendix slide right after it, at 11 pt with template margins. The appendix slide has no label of its own: it keeps the summary's header and each table keeps its own title. Tables are walked in order, and one stays on the summary slide whenever it fits in the room left under the tables already kept there, even if an earlier table moved. A table is never cut. Only a table that is taller than a whole slide by itself is split, on its own slide, with its title and header rows repeated on each part.
4. **Digital summary.** It takes the shared margins and the shared size when its «أبرز التحديثات» table fits with them. Only when it does not, that table alone goes smaller, in half-point steps down to 8 pt with the compact margins, and past 8 pt its rows are squeezed into the space. It has no appendix (its table belongs under the indicators), and the notice names the slide to check. The other three stay at the shared size. If the digital summary alone needs the compact margins, all summary slides still get them.
5. **Detail tables keep 11 pt** and continue on further slides. A single record taller than a whole slide takes a slide of its own. Digital update/challenge tables that cannot follow their sector's last record even on a fresh slide are set smaller on that slide.

The alternative for the digital summary is shorter updates in the review file, then a rebuild.

Every such decision is a `space` notice. The notice is short formal Arabic: which slide, what was done (margins, font size, appendix), and one suggestion (shorten the text in the file). It carries no command-line flags. The skill does not wait for an answer. It builds with this default, and in the same reply that delivers the file it tells the user in plain words what did not fit and what default was applied, and offers the alternatives below so he can ask for a rebuild. If the user had already stated a preference (a floor size, an appendix), that preference is used instead of the default. The alternatives, with their flags for the agent:

- a larger floor size, with more on the appendix: `--summary-min-pt 10.5`; a smaller one, down to 8: `--summary-min-pt 8`. The option is optional. Not given: the shared floor is 10 pt and the digital floor is 8 pt. Given: that value is the floor for all four summaries;
- 11 pt on every summary (template margins when all fit, otherwise compact margins on all summary slides) and the overflow of a non-digital summary on its appendix slide as whole tables: `--summary-overflow appendix`;
- shorter wording of the long cells in the Excel file (or, for the digital notes you wrote, in the review file), then a rebuild. Never shorten source text on your own: business rule 2.

## When to stop

Stop only when the report cannot be produced at all:

- a required input is missing or unreadable: one of the three workbooks, a required sheet or column, the digital PDF for the digital model, or a PDF the extractor rejects;
- the bundled master is missing or fails its checksum;
- the built file would be damaged or would state a number, record or chart value that differs from the inputs (the content findings of the gate).

Anything else is a best assumption plus a notice, and the work continues. Read `<report>.pptx.notices.json` after every build and repeat every notice in the final reply. The notices are written in Arabic; repeat them as written, with the assumption that was made. Examples of what no longer stops the build: a status the template has no style for (kept as written on a neutral gray fill), a workbook whose summary and details disagree (both transferred as written), a summary that does not fit its slide, a record taller than a slide, a table without measurements (estimated), a stale or malformed `--measurements` file, an empty digital detail card, a digital section heading with no records (the heading is dropped), a layout or navigation finding of the gate. A notice of kind `layout` that says the build is estimated, or of kind `check`, means the file needs another pass before it is final; say so plainly rather than present it as accepted. Preserve the previous accepted report, and never mark an unperformed check as passed.

## After a manual edit

PowerPoint does not move the table under one that grew, and a row the user inserts copies the height of the row beside it. When the user edits a delivered report by hand (a new row, longer text, a deleted record) and the tables no longer sit right, run the shared tool on the edited file:

```bash
uv run ../report_common/reflow.py --report "<edited.pptx>" --output "<fixed.pptx>"
```

It keeps each table's own text size and the approved typography, counts the «#» column again, gives every row the height its text needs (measured through LibreOffice) and stacks the tables of each slide again with the approved gap. It does not move records between slides: a slide whose tables now pass the footer is named in the notices, and the fix is to move those rows in PowerPoint and run the tool again, or to rebuild from the Excel files.

## Runtime and installation

Prefer `uv run` using the declared dependencies. Reuse an existing installation. If resolution is unavailable and an existing compatible Python satisfies the dependencies, use it and record the fallback; do not modify host Python. If installing uv is necessary, use Astral's official user-scoped installer, not pip. Shared code lives in `report_common/`: the sibling `skills/report_common/` in a repository or plugin install, or `scripts/report_common/` inside a standalone bundle from `tools/package_skills.py`. In a bundle, run the PDF exporter as `uv run scripts/report_common/render.py` and the manual-edit tool as `uv run scripts/report_common/reflow.py`. Runtime scripts must not import test fixtures or the other report's generator.

## Delivery

Deliver the editable PPTX bytes that were rebuilt with measurements and passed data checks and final visual review, using an accessible file link. Deliver PDF only when requested. List every notice of the build in the reply.

After delivering the report, always reply with the two Arabic messages to the Minister, short and long, each in its own copyable code block, and save both next to the report. Generate them with `scripts/minister_messages.py` from the delivered report, follow [the minister messages reference](references/minister-summary.md) for both templates and the wording rules, and run its `--check` on both finished messages before sending them. Keep QA evidence outside the slides: exact output/input/template hashes, performed checks, renderer, and any limitation. Any edit invalidates affected checks. Do not split sectors, schedule runs, or send reports to others unless requested.
