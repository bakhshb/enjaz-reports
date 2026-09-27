---
name: monday-meeting-report
description: Populate the approved Arabic Monday Meeting PowerPoint deck from the weekly topics workbook, the final task summary/details workbook, and the final Suhail summary/details workbook while preserving the supplied PowerPoint template exactly. Use this skill for Monday Meeting, اجتماع الاثنين, عرض الاثنين, جدول أعمال اجتماع القيادات, ملخص المهام, or ملخص مشاريع سهيل when the requested output is the approved Monday Meeting presentation. This skill is standalone and does not depend on the weekly-consolidated-report skill.
---

# Monday Meeting Report

Populate the approved bundled PowerPoint design master with the current weekly data. The legacy filename `MM_W38.pptx` is only the internal asset name; **W38 is not a report version, data source, or week selector**. Excel from the current request is the sole source of weekly content, and the PowerPoint master is the sole source of design.

This skill is intentionally independent from `weekly-consolidated-report`. Do not import transaction rules, summary-table rules, slide order, or redesign behavior from the weekly consolidated report. If another reporting skill conflicts with this page, this skill controls the Monday Meeting deck.

## Inputs and output

- `مواضيع الأسبوع.xlsx` or the equivalent weekly topics workbook: use `Sheet1` as the agenda source. The sample structure is the same as `32 مواضيع الأسبوع.xlsx`; the weekly values may change.
- `ملخص وتفاصيل المهام الاستراتيجية.xlsx`: sheets `ملخص المهام` and `تفاصيل المهام`.
- `ملخص_مشاريع_سهيل.xlsx`: sheets `ملخص مشاريع سهيل` and `تفاصيل مشاريع سهيل`.
- Approved PowerPoint design master: bundled inside this skill as complete bundled master under the legacy internal name `MM_W38.pptx`. Treat that name as an implementation detail only. The latest approved visual QA reference is `MM_2026-09-20.pptx`, and no weekly values may ever be taken from either PowerPoint file. The user does **not** upload the master during normal use.
- At runtime, always use the new Excel files attached to the request. The files used to define this skill are structural examples only; never reuse their weekly values as production data.
- Save the result as a new `.pptx`. Never overwrite the restored template or any Excel input.
- Create a PDF only when the user asks for one.
- **Mandatory preflight — all 3 Excel workbooks are required before generation.** Before opening or building the PowerPoint, verify that the current request contains all three production workbooks: (1) `مواضيع الأسبوع.xlsx` or its weekly equivalent, (2) `ملخص وتفاصيل المهام الاستراتيجية.xlsx`, and (3) `ملخص_مشاريع_سهيل.xlsx`. Do not start a partial report with only one or two files. Do not substitute structural examples, older Library copies, or files bundled with the skill for the current-week inputs.
- Validate the three workbooks structurally, not only by filename: the topics workbook must contain `Sheet1` with `الترتيب`, `الموضوع`, `المتحدث`, and `الوقت المطلوب من المالك`; the tasks workbook must contain `ملخص المهام` and `تفاصيل المهام`; the Suhail workbook must contain `ملخص مشاريع سهيل` and `تفاصيل مشاريع سهيل`.
- If any required Excel workbook is missing or does not match the expected structure, stop before generation and ask only for the missing/incorrect workbook, naming it explicitly. Do not ask for the PowerPoint template during normal use: the approved template is bundled inside this skill and must be restored from the bundled assets.

Approved template SHA-256 for the bundled master `MM_W38.pptx`: `0a91475077b86cb2c0d8e2089eb8016ea665cf8c0eb86dc655bf45f40b7eb69e`.

Latest visual QA reference reviewed on 20 September 2026: `MM_2026-09-20.pptx`. It is a populated 18-slide reference used to verify the current slide/shape placement and visual result; it is **not** a runtime input and must never be requested from the user. The bundled `MM_W38.pptx` remains the restorable master asset used for generation.

## Bundled implementation

```text
monday-meeting-report/
├── SKILL.md
├── scripts/
│   ├── build_report.py
│   ├── validate_report.py
│   ├── report_gate_v2.py
│   └── restore_template.py
└── assets/
    └── MM_W38.pptx
```

The complete approved master is bundled as `assets/MM_W38.pptx` (legacy name only). Its SHA-256 is `0a91475077b86cb2c0d8e2089eb8016ea665cf8c0eb86dc655bf45f40b7eb69e`. `restore_template.py` verifies that asset and copies it into the working `template/` directory when needed. Run `build_report.py` with the three current Excel inputs; `report_gate_v2.py` is the current mandatory final repair and QA gate.

## Python runtime and automatic bootstrap

Use `uv` for every Python command in this skill. The required Python version and packages are declared as inline script metadata, so do not install them with `pip` and do not modify the host Python environment.

Before the first run, execute `uv --version`:

- If `uv` is available, keep the existing installation and continue.
- If it is missing, install it automatically with Astral's official user-scoped standalone installer, then refresh `PATH` or call the installed executable by its full path and verify `uv --version` again.

Windows PowerShell:

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

macOS or Linux:

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

Do not reinstall or upgrade a working `uv`. If the skill was copied without running its installation step and `uv` is missing at report time, perform the same bootstrap automatically. Stop only when the official installer cannot run because of a real environment restriction, and report the exact error.

## Authoritative rules

1. **Preserve the template, do not redesign it.** Keep slide size, masters, layouts, backgrounds, logos, photos, decorative elements, fonts, font sizes, weights, colours, borders, spacing, margins, alignment, RTL behaviour, table widths, element positions, dimensions, and z-order exactly as supplied. Replace content only.
2. The PowerPoint template is not a fallback data source. Remove all previous-week text or values from fields that are blank in the new source. Never carry an old value forward because the new Excel cell is empty.
3. Transfer source text, names, numbers, dates, statuses, and notes as supplied by the final Excel outputs. Do not summarize, rewrite, proofread, infer missing content, or recalculate status from dates.
4. Preserve source order. Do not sort agenda items, tasks, projects, sectors, or updates unless the source workbook already presents them in that order.
5. Summary slides remain summary slides. Populate only summary objects that exist in the approved Monday Meeting template. **Do not add new summary sections or tables merely because an Excel summary contains additional sections.** Detail data is handled in the detail section.
6. When detail content needs more space, duplicate the matching **detail** slide/table pattern and continue there. Never shrink fonts, compress text, stretch elements, or redesign a table to force more records onto a slide.
7. If a summary table has more records than can fit in its approved single-slide geometry at the existing font sizes, stop and report the overflow instead of creating a continuation summary slide or shrinking the design.
8. Keep every PowerPoint chart as a real editable native chart. Never replace a chart with an image and never link it to an external Excel file.
9. For chart data, mirror the Excel source cells exactly: **a blank source cell stays blank in the chart's embedded worksheet; never write `0` into a blank chart cell.** Numeric values remain numeric. Keep `dispBlanksAs` as `gap`.
10. Text KPI indicators may display `0` when the Excel KPI itself explicitly contains zero. The blank-cell rule above is specifically for the chart data matrix.
11. Blank notes, updates, or relationship fields stay blank. Do not insert `لا يوجد`, `مرفق`, `الوثيقة الداعمة`, or any other filler unless it exists in the source cell.
12. The cover date is the **presentation creation date in the runtime timezone**, not a date from any Excel file and not the latest project update date. Format it exactly like the template, for example `15 سبتمبر 2026`, with no time.

## Current approved slide order

Use the section title and object structure as the primary identifier; the slide numbers below document the approved reference template and serve as verification.

1. Cover — `Monday Meeting` + creation date.
2. Agenda table.
3. `افتتاحية`.
4. `ملخص الإنجاز الأسبوعي`.
5. `ملخص المهام`.
6. `إحاطات`.
7. `ملخص الاجتماع`.
8. `شكرا لكم`.
9. `الملحقات`.
10. `تفاصيل المهام`.
- Slides 11–12 — Task detail tables.
- Slide 13 — `تفاصيل مشاريع سهيل`.
- Slide 14 — `ملخص مشاريع سهيل`.
- Slides 15–18 — Suhail project detail tables.

Static section-divider slides remain unchanged except for changes explicitly defined in this skill.

## Cover mapping

Reference slide 1 contains a group with:

- `Monday Meeting` — keep unchanged.
- Date text box, current visual-reference text `20 سبتمبر 2026` — replace with the runtime presentation creation date.

Do not change the date position, font, colour, alignment, or separator line.

## Agenda source and stopping rule

Use `Sheet1` in the weekly topics workbook. The relevant source columns are:

- `الترتيب` → agenda item number.
- `الموضوع` → `جدول الأعمال`.
- `المتحدث` → `المسؤول`.
- `الوقت المطلوب من المالك` → `المدة الزمنية (بالدقيقة)`.

Rules:

- Use agenda rows with a valid `الترتيب` value, preserving the worksheet order.
- Format visible item numbers like the template: `01`, `02`, `03`, …
- **`ملخص الاجتماع` is the final agenda item. Include it, then stop. Ignore every row after it completely, even if later rows contain a speaker, topic, duration, or order value.**
- Normalize surrounding whitespace only for detecting the stop phrase; display the source text normally.
- Rows such as `لم تتم الإفادة` that occur after `ملخص الاجتماع` are not added.
- If `ملخص الاجتماع` is missing, do not guess the stopping point; ask the user before producing the deck.
- The last visible body row in the PowerPoint agenda must be `ملخص الاجتماع`. Remove old template rows below it; never leave old topics after the meeting summary.

Reference target: slide 2, `Table 5`, shape ID `7`, four columns in this exact visual order:

`م | جدول الأعمال | المسؤول | المدة الزمنية (بالدقيقة)`.

Preserve the current column widths, fills, font, alignment, row style, and table position. Add or remove body rows only by cloning the existing row style. If the agenda cannot fit on the approved slide at the existing geometry and typography, report the overflow rather than redesigning the slide.

## Task summary mapping

Source sheet: `ملخص المهام`.

### KPI indicators

- `B2` (`إجمالي عدد المهام`) → slide 5, `TextBox 10`, shape ID `21`.
- `B3` (`المهام المكتملة`) → slide 5, `TextBox 13`, shape ID `23`.
- `B4` (`المهام على المخطط`) → slide 5, `TextBox 15`, shape ID `25`.
- `B5` (`المهام المتأخرة`) → slide 5, `TextBox 6`, shape ID `19`.
- `B6` (`المهام المعلقة`) → slide 5, `TextBox 5`, shape ID `18`.

Replace the complete displayed value inside each object while preserving its existing formatting.

### Task status chart

Source range begins at the heading row `القطاع | مكتملة | على المخطط | متأخر | معلق` and continues through the sector rows until the next blank/section boundary. In the current source structure this is `A9:E17`.

Reference target: slide 5, `Chart 2`, shape ID `3`, `chart1.xml`, embedded workbook `Microsoft_Excel_Worksheet.xlsx`.

Embedded chart worksheet structure:

`القطاع | مكتملة | على المخطط | متأخر | معلق`.

Requirements:

- Replace sector labels and series values from the Excel summary.
- Keep source sector order.
- A blank Excel summary cell must remain blank in the embedded chart workbook; **do not convert blanks to zero**.
- Keep the chart native and editable through PowerPoint `Edit Data`.
- Update the embedded workbook and the chart cache if the editing method does not refresh the cache automatically.
- Keep `dispBlanksAs="gap"`.
- Do not create an external workbook relationship.
- Do not change chart type, size, location, colours, axes, legend behaviour, gap width, series formatting, or plot area.

### Completed-task table

Source section: `المهام المكتملة`, with source columns `# | المهمة | القطاع`; in the current structure the header is at row 20 and records follow until the next section.

Reference target: slide 5, `Table 47`, shape ID `48`.

The PowerPoint table visually displays `القطاع | المهمة | #`; map source fields to those existing columns without changing widths or styling. Replace all old body rows with the current source records.

The source sheet may also contain sections such as `المهام المتأخرة` and `طلبات الدعم`. The approved Monday Meeting summary slide does **not** contain matching summary tables, so do not introduce new ones. Their statuses still affect the KPI/chart, and their records remain available in the detail section when present there.

## Task details

Source sheet: `تفاصيل المهام`.

Detect a task section from a heading shaped like `مهام مصدر (اسم المصدر)` followed by the seven-column header. Use the source section order exactly.

Every task-detail table uses this exact source order:

`# | المهمة | القطاع | تاريخ الانجاز المخطط | الحالة | ملاحظات | الجهات ذات العلاقة`.

In the PowerPoint preserve the approved header spelling and visual order already present in the template:

`# | المهمة | القطاع | تاريخ الإنجاز المخطط | الحالة | ملاحظات | الجهة ذات العلاقة`.

Current reference mappings:

- `اجتماع القيادات` → slide 11, `Table 4`, shape ID `5`; continuation in the reference template → slide 12, `Table 2`, shape ID `3`.
- `اجتماع التجربة الرقمية` → slide 12, `Table 17`, shape ID `18`. The old template wording `لجنة التجربة الرقمية` refers to the same section; display the current source section name while preserving heading formatting.
- `لجنة المتابعة` → slide 12, `Table 3`, shape ID `4` in the current 18-slide visual reference.

Rules:

- Replace every old row with the current source row.
- Preserve source numbering inside each source section.
- Do not infer or rewrite notes.
- An empty source cell clears the corresponding old PowerPoint cell.
- If a section needs additional capacity, duplicate the appropriate task-detail slide/table pattern and continue the same section there.
- If a new task source appears, clone the approved task-detail pattern and insert it after the existing task-source sections and before `تفاصيل مشاريع سهيل`.
- If a source section has no records, do not show stale rows from the template. Remove or clear that section without moving unrelated content to fill the space.

## Suhail summary mapping

Source sheet: `ملخص مشاريع سهيل`.

### KPI indicators

Current structure uses row 3:

- `A3` (`إجمالي مشاريع سهيل`) → slide 14, `TextBox 10`, shape ID `5`.
- `B3` (`المشاريع المكتملة`) → slide 14, `TextBox 13`, shape ID `8`.
- `C3` (`المشاريع على المخطط`) → slide 14, `TextBox 15`, shape ID `11`.
- `D3` (`المشاريع المتأخرة`) → slide 14, `TextBox 6`, shape ID `17`.
- `E3` (`المشاريع لم تبدأ`) → slide 14, `TextBox 5`, shape ID `13`.

The approved reference slide contains an old KPI label `معلقة` beside shape ID `13`. When the Excel summary category is `لم تبدأ`, replace that stale label with **`لم تبدأ`** while preserving the exact text-box formatting, dimensions, and position. Do not relabel any other status by inference.

### Suhail status chart

Source range begins at `القطاع | مكتملة | على المخطط | متأخر | لم تبدأ` and continues through the sector rows until the next blank/section boundary. In the current structure this is `A6:E15`.

Reference target: slide 14, `Chart 23`, shape ID `24`, `chart2.xml`, embedded workbook `Microsoft_Excel_Worksheet1.xlsx`.

Requirements:

- Embedded worksheet columns remain `القطاع | مكتملة | على المخطط | متأخر | لم تبدأ`.
- Replace sector labels and series values from the current Excel summary.
- Keep source sector order.
- **Blank chart-source cells stay blank. Never insert zero into a blank source cell.**
- Keep the chart as a native editable PowerPoint chart using its embedded workbook.
- Update chart cache when necessary, keep `dispBlanksAs="gap"`, and preserve all chart design properties.

### أبرز التحديثات

Source section: `خامسا: أبرز التحديثات`, with columns `# | المشروع | القطاع | التحديث`; in the current structure the header is at row 29.

Reference target: slide 14, `Table 43`, shape ID `44`.

The PowerPoint visually displays `التحديث | القطاع | المشروع | #`. Map the source fields into those existing columns without changing the design. Replace all old body rows. If the source contains no update records, clear the old body content; do not keep previous-week updates.

The source summary may also contain `المشاريع المكتملة` and `المشاريع المتأخرة`. The approved Monday Meeting summary slide does not contain those tables, so **do not add them** to this presentation.

## Suhail project details

Source sheet: `تفاصيل مشاريع سهيل`.

Detect each sector from a sector heading followed by this six-column header:

`# | اسم المشروع | تاريخ البداية | تاريخ النهاية | الحالة | ما تم حتى تاريخه`.

Match PowerPoint tables by sector name, not by Excel row number. Current reference mappings:

- `وكالة شؤون الحج` → slide 15, `Table 3`, shape ID `4`; continuation pattern in the current reference → slide 16, `Table 3`, shape ID `4`.
- `وكالة تطوير الأعمال وخدمة المستفيدين` → slide 16, `Table 5`, shape ID `6`.
- `وكالة التعاون الدولي والشراكات` → slide 16, `Table 9`, shape ID `10`.
- `وكالة التوعية وتنمية المهارات` → slide 17, `Table 5`, shape ID `6`.
- `وكالة التحول الرقمي وتقنية المعلومات` → slide 17, `Table 2`, shape ID `3`.
- `مركز منصة نسك الرقمية` → slide 18, `Table 7`, shape ID `8`.
- `مركز امتثال أعمال` → slide 18, `Table 3`, shape ID `4`.

Rules:

- Replace all old project rows with the current source rows.
- Preserve source project order inside each sector.
- Use the sector heading from Excel; never invent a missing sector column.
- Keep the exact six-column table layout and widths.
- In `ما تم حتى تاريخه`, if the source cell already begins with a standalone line such as `تاريخ التحديث 15 سبتمبر 2026` or `تاريخ التحديث: 15 سبتمبر 2026`, make only that existing date line bold to match the approved table style. Do not invent, normalize, or add an update date when the source does not contain one.
- If a sector needs more room, duplicate the matching Suhail-detail pattern and continue there.
- If a new sector appears, add it using a clone of the approved Suhail-detail table pattern after the existing sector detail tables.
- If a sector is absent in the current source, remove its previous-week records. When a slide contains another valid sector table, leave that other table in its approved position; do not move it merely to fill space.

## PowerPoint chart integrity

The reference deck contains two native charts with internal Excel packages:

- Task chart: `chart1.xml` → `Microsoft_Excel_Worksheet.xlsx`.
- Suhail chart: `chart2.xml` → `Microsoft_Excel_Worksheet1.xlsx`.

When editing the package directly or through a library:

- Preserve the existing chart relationship and embedded package relationship.
- Preserve series order and category references unless the number of sector rows requires expanding/contracting the referenced range.
- If the number of sector rows changes, update the chart formulas and caches to the exact used range while keeping the same four status series.
- Do not write formulas or external links into the embedded workbook.
- Empty cells must remain truly empty, not strings such as `""` and not numeric zero, unless the source cell explicitly contains zero.

## Delivery verification

Before delivering the presentation, verify all of the following:

- The `.pptx` opens without PowerPoint repair, corruption, or a reference-to-another-file warning.
- No external Excel links were introduced.
- The cover date equals the presentation creation date and uses the template format.
- The agenda ends visibly with `ملخص الاجتماع`; nothing after it from the topics workbook appears in the agenda.
- Agenda order, speaker, topic, and duration match the topics workbook.
- Task KPI values match `ملخص المهام` exactly.
- Task chart sectors and values match the summary sheet, and every blank source chart cell is still blank in the embedded chart workbook.
- The completed-task summary table matches the Excel summary.
- Every task-detail row matches `تفاصيل المهام` and no previous-week note survives in a blank source field.
- Suhail KPI values match `ملخص مشاريع سهيل` exactly, including `لم تبدأ`.
- Suhail chart sectors and values match the Excel summary and blank cells remain blank.
- `أبرز التحديثات` matches the Excel summary exactly.
- Every Suhail-detail table matches `تفاصيل مشاريع سهيل` by sector.
- Both charts remain editable with PowerPoint `Edit Data` and visually use the original chart design.
- No font, layout, background, logo, colour, table width, or unrelated element changed.

If any check fails, fix the output before delivery rather than explaining the error away.

## Production gate v2.1 — 27 September 2026

This section is the **current production rule** and overrides every earlier hotfix, postprocess, QA, or delivery-gate section on this page when they conflict.

The failed week-40 deck showed three distinct failure classes that must be prevented together: stale/current-request data mismatch, agenda/title-row corruption, and table typography drift. Therefore a build is not complete when `build_report.py` exits successfully.

### Required execution order

1. Bind the three Excel files from the **current request attachments only** by workbook structure and keep their exact runtime paths. Never search Notion, Library, older uploads, previous runs, or bundled examples for weekly data.
2. Run the existing `build_report.py` with those exact paths.
3. Run the existing `validate_report.py`.
4. Use the bundled `scripts/report_gate_v2.py` as the current v2.1 gate. The older postprocessing and QA attachments have been superseded.
5. Run the mandatory final gate on the **same output PPTX and the same three current-request Excel paths**:

```bash
uv run scripts/report_gate_v2.py \
  --report "MM_YYYY-MM-DD.pptx" \
  --topics "<exact current-request topics path>" \
  --tasks "<exact current-request tasks path>" \
  --suhail "<exact current-request Suhail path>"
```

1. Deliver only if the final command prints `QA PASSED - report repaired and safe to deliver` and exits with code 0.


### What v2 repairs before validation

- Rebuilds the agenda from the current topics workbook only, stopping at `ملخص الاجتماع`.
- Restores the exact agenda header row: `م | جدول الأعمال | المسؤول | المدة الزمنية (بالدقيقة)`.
- Removes duplicate/stale agenda body rows.
- Forces the approved table typography explicitly instead of relying on inherited theme formatting:
    - agenda table: `Abar Mid`, 14 pt;
    - all other dynamic tables: `Abar Mid`, 11 pt;
    - `Abar Mid SemiBold` remains allowed where the approved template already uses it.
- Replaces known table-font substitutions such as Arial, Calibri, and `+mn-lt` with `Abar Mid` while preserving bold/italic, fills, borders, alignment, widths, and geometry.

### What v2 blocks

The gate must fail and the deck must **not** be delivered when any of the following is true:

- agenda rows do not exactly match the current topics workbook;
- task-detail records do not exactly match the current `تفاصيل المهام` sheet;
- Suhail-detail records do not exactly match the current `تفاصيل مشاريع سهيل` sheet;
- a task-detail table lacks its merged section-title row or seven-column header row;
- a Suhail-detail table lacks its sector/title row or six-column header row;
- a completed-task or `أبرز التحديثات` table loses its title/header structure;
- any visible table run uses an unapproved font or the wrong approved font size;
- the PPTX ZIP/XML package is malformed or contains an external relationship;
- the generated PowerPoint cannot be re-opened by the PowerPoint parser after the final save.

### Continuation-slide rule

When a detail table overflows, clone the **whole approved table pattern**, not body rows alone. The v2.1 gate also repairs a pure Suhail continuation slide that was placed after unrelated sectors by moving it directly after the sector it continues, and then blocks non-contiguous sector continuations. Every continuation table repeats:

- its section/sector title row;
- its full column-header row;
- then the next complete source record.

Do not shrink 11 pt text to 10 pt, omit a title row, or reuse a previous-week row to make content fit. Add a continuation slide instead.

### Template-version rule

The legacy asset name `MM_W38.pptx` is an internal filename only. It must never be interpreted as week 38 data. The current visual acceptance reference remains `MM_2026-09-20.pptx`. If the restored bundled master renders differently from that reference in layout, title rows, fonts, sizes, or table geometry, stop and report a template-asset mismatch rather than silently continuing.

### Delivery rule

Return the actual generated `.pptx` as a real downloadable attachment. Never return only a local path such as `/mnt/...`, `/tmp/...`, a `file://` URI, or a workspace-only path. The exact bytes attached to the final reply must be the bytes that passed the final v2 gate.

