---
name: weekly-consolidated-report
description: Populate either approved Ministry of Hajj and Umrah weekly consolidated PowerPoint model from the final tasks, Suhail projects, and transactions Excel summaries. The digital-experience model also uses its current PDF. Ask which model to use when the user has not chosen one. Does not build upstream Excel summaries or split by sector.
---

# Weekly Consolidated Report

## Purpose and authority

Create the report from current final Excel inputs and one of this skill's approved masters. Excel controls the existing sections; the optional digital-experience PDF controls its added section; the chosen master controls design. Read [business rules](references/business-rules.md) before generation and [acceptance checks](references/acceptance.md) before validation. Preserve this report's own mappings, ordering, omissions, pagination, and status rules. Sample records and slide counts are not runtime requirements.

## Choose the model

When the user sends the Excel files without choosing a model, ask: **«أي نموذج تريد: التقرير الموحد الحالي، أم التقرير الموحد مع نضج التجربة الرقمية؟»** Wait for the choice before generating. If the user already selected one, use it without repeating the question. Do not select a model from the mere presence or absence of a PDF.

- **التقرير الموحد الحالي** (`standard`): the existing master and three Excel inputs, with unchanged business rules.
- **التقرير الموحد مع نضج التجربة الرقمية** (`digital`): the second bundled master and the same three Excel inputs, plus the current digital-maturity PDF (the landscape «horizontal» export). The PDF is a data source only: page one fills the summary; every later page belongs to one sector, and a sector may span several source pages titled «(k/n)». Merge each sector's pages into one record set and re-flow it into the approved portrait detail slides; never reproduce the source pagination, one slide per PDF page, or a half-empty slide. Read [digital-maturity mapping](references/digital-maturity.md) only for this model. If the PDF is missing, ask for it before building.

## Inputs and output

- ملخص وتفاصيل المهام الاستراتيجية.xlsx: ورقتا ملخص المهام وتفاصيل المهام.
- ملخص_مشاريع_سهيل.xlsx: ورقتا ملخص مشاريع سهيل وتفاصيل مشاريع سهيل.
- تقرير_المعاملات_الموحد.xlsx: ورقة المعاملات.
- قالب PowerPoint المعتمد الحالي: `assets/weekly-report-master.pptx`، مع PDF الموافق مرجعًا بصريًا. هذا القالب هو مصدر التصميم المعتمد الحالي.
- قالب النموذج الرقمي الإضافي: `assets/weekly-report-digital-master.pptx`. يبقى النموذج الحالي متاحًا دون استبدال. مرجعه البصري المعتمد للمخرج هو `assets/weekly-report-digital-reference.pdf` (تقرير 29 أكتوبر 2026): يبيّن كيف يُعاد صبّ قطاعات ملف التجربة الرقمية الأفقي في شرائح التفاصيل العمودية، وهو مرجع تصميم لا مصدر بيانات.
- عينة المصدر الأفقي المعتمد لملف التجربة الرقمية محفوظة في المستودع: `tests/fixtures/digital-maturity-horizontal.pdf` (تحديث 1448/5/17هـ). يختبرها `tests/test_weekly_digital_maturity.py` استخراجًا ودمجًا وبناءً؛ أي تغيير في المستخرج يجب أن يبقي هذه العينة ناجحة.
- At runtime, use the new Excel files attached to the request. Restore the approved PowerPoint template from the assets attached to this skill; the user should not need to attach the template again. Ask for the template only if the skill assets are unavailable or fail integrity verification. Never substitute sample values.
- Primary output: `تقرير الإنجاز الأسبوعي بتاريخ الإنشاء.pptx`, saved as a new file. Create a matching PDF only when requested. Never modify the inputs or overwrite the original template.


Verify the bundled master size `9546137` bytes and SHA-256 `4944d523075aafb37d38f90d1ee12fd0b2b80e46033f1af81880c646ad5fc689` before use. Never fall back to an earlier master.

For `digital`, verify SHA-256 `6af842ee98af6f1e4bf44e9e6889e8d9cc206685e22acf2adc88fd3078510819` for its bundled master. Do not use the current model's master for this choice.

## Table typography

Every PowerPoint table must use native right-to-left (RTL) table direction in both models, including support, transactions, digital summary/detail/challenge tables and all continuations. Right-aligned text alone is insufficient. Keep numbering at the right; convert LTR-stored patterns together with their column widths and merged cells so every value stays with its field. Reject output tables without RTL enabled. Latin identifiers may retain LTR text within an RTL table.

All table text is **11 pt**, including table titles, headers, data, status cells and notes (and the Monday agenda). Apply this explicit size even when the master stores 10 or 14 pt. Preserve Abar, bold emphasis, RTL, fills and column widths. Adjust row heights and pagination when needed; never reduce the font to make a table fit.

Apply the per-column horizontal alignment in the business rules to **summary and detail tables**, including headers: sector/dates/status centered; task/project names, notes, updates, support text and challenges right-aligned.

## Business rules

Apply every rule in [the report-specific reference](references/business-rules.md). Do not infer content, change source workbooks, or modify the original master. Preserve the report's section order, dates, blank cells, and current status meaning. Template samples are never weekly data.

## Generate and validate

1. Check required workbooks, sheets, headers, and the chosen approved master. For `digital`, render page one of the current PDF and extract a review file:

```bash
uv run scripts/digital_maturity.py --pdf "<digital.pdf>" --output "<review.json>"
```

Compare every extracted card percentage, sector/bar color and value, sector note, and challenge line with rendered page one. Review every merged sector in `detail_sectors` (title without suffix, four cards, every table row in source order, every challenge from all of its pages) against the rendered source pages listed in its `source_pages`. Correct PDF text-layer reversal, numbers, codes, `(cid:…)` glyph gaps, or ligature errors in the review file; never invent missing values and never merge, reword or drop challenges. Set `reviewed` to `true` only after the full comparison. The builder rejects unreviewed data, a PDF hash mismatch, a title that still carries «(k/n)», or a row that still carries «(تتمة)». Keep the review file as QA evidence outside the delivered slides. If the extractor rejects the PDF (layout, page count, table shape), stop and report the exact error; do not hand-build the digital slides page by page.
2. Generate into a new candidate path:

```bash
uv run scripts/build_report.py --tasks "<tasks.xlsx>" --suhail "<suhail.xlsx>" --transactions "<transactions.xlsx>" --output "<candidate.pptx>"
```

For the second model, add `--model digital --digital-pdf "<digital.pdf>" --digital-data "<review.json>"`. The standard command remains valid. Pass the same model, PDF and review-file options to `validate_report.py`.

3. The builder includes read-only programmatic validation. Its initial placement is provisional: render before accepting page breaks, table gaps, or footer clearance. Follow the existing continuous-flow rules in the business reference; move complete records and render affected slides again. After any correction, rerun validation. Do not repeat it for unchanged bytes.

```bash
uv run scripts/validate_report.py --report "<candidate.pptx>" --tasks "<tasks.xlsx>" --suhail "<suhail.xlsx>" --transactions "<transactions.xlsx>"
```

The candidate is saved only after programmatic checks pass. Failed builds preserve existing outputs and inputs; output paths cannot alias an input or master. A programmatic pass is not visual acceptance.

The gate also verifies inherited detail-layout artwork against the chosen master, visible master graphics, unique shape IDs including grouped children, and the chosen model's header buttons and destinations. Keep `scripts/weekly_artwork.py` with the installed generator and validator. A missing logo, footer, button, or layout link is a blocking failure; do not bypass the gate to deliver.

## Visual review

For LibreOffice PDF previews, use `uv run ../report_common/render.py --report "<candidate.pptx>" --output "<candidate.pdf>"`. The shared exporter uses a disposable copy with the RTL column order materialized for LibreOffice; the editable PPTX remains unchanged. Inspect every summary, detail, transaction and digital table: numbering must visibly appear on the right. Do not accept the RTL property alone. `tests/check_rendered_pdf.py` checks numbering using actual PDF coordinates, alongside fonts and page counts.

Render the validated candidate and inspect every slide against this skill's template/reference, including Arabic, actual Abar usage, status fills, notes, chart labels, row continuity, table gaps and footer clearance. Verify native PowerPoint opening without repair and editable embedded charts. Follow [acceptance checks](references/acceptance.md); correct layout using the approved patterns and repeat affected checks. Final visual acceptance follows the last edit and programmatic check.

When a Windows PowerPoint render reveals overlapping tables or premature page breaks, `scripts/measure_layout.ps1 -Report <candidate.pptx> -Tasks <tasks.xlsx> -Suhail <suhail.xlsx> -Transactions <transactions.xlsx>` reads actual row heights without modifying the deck: it measures a temporary copy whose record rows are first collapsed so PowerPoint fits them to their text, and skips the digital-maturity slides, which paginate with their own model. Rerun the same builder with `--measurements <candidate.layout.json>` to repaginate complete records using those heights, then render again. Measurements are bound to the exact input/template hashes. They support the visual decision; they do not establish visual acceptance. Without Windows PowerPoint, use the LibreOffice counterpart, which writes the same `<candidate>.layout.json` (same hashes, same format) from a disposable copy:

```bash
uv run scripts/measure_layout.py --report "<candidate.pptx>" --tasks "<tasks.xlsx>" --suhail "<suhail.xlsx>" --transactions "<transactions.xlsx>"
```

For `digital`, add `--digital-pdf "<digital.pdf>" --digital-data "<review.json>"`. It measures every record row through `report_common/measure.py`: LibreOffice draws rows at their declared height and never grows them, so each row is rendered alone in a collapsed one-row table and its glyph extent is measured. Install the template font (Abar Mid) first; with a substitute font the heights are wrong. The builder's first-pass heights are a deliberately cautious estimate, so a measured rebuild is normally much shorter (the 2026-11-05 inputs went from 28 to 15 slides on the standard model and from 38 to 25 on the digital one).
For the digital model, also pass `-Template <digital-master.pptx> -DigitalPdf <digital.pdf> -DigitalData <review.json>` so the measurement is tied to its PDF and reviewed values.

## Failure handling

Stop the affected report when inputs conflict, a required pattern/style is unapproved, a complete record cannot fit with the approved design, or faithful rendering is unavailable. Report the exact blocker and the condition needed to pass. Preserve the previous accepted report. Never mark a blocked or unperformed check as passed.

## Runtime and installation

Prefer `uv run` using the declared dependencies. Reuse an existing installation. If resolution is unavailable and an existing compatible Python satisfies the dependencies, use it and record the fallback; do not modify host Python. If installing uv is necessary, use Astral's official user-scoped installer, not pip. Shared code lives in `report_common/`: the sibling `skills/report_common/` in a repository or plugin install, or `scripts/report_common/` inside a standalone bundle from `tools/package_skills.py`. In a bundle, run the PDF exporter as `uv run scripts/report_common/render.py`. Runtime scripts must not import test fixtures or the other report's generator.

## Delivery

Deliver the editable PPTX bytes that passed data checks and final visual review, using an accessible file link. Deliver PDF only when requested. Keep QA evidence outside the slides: exact output/input/template hashes, performed checks, renderer, and any limitation. Any edit invalidates affected checks. Do not split sectors, schedule runs, or send reports to others unless requested.
