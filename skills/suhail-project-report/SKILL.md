---
name: suhail-project-report
description: Build the two-sheet Arabic «ملخص مشاريع سهيل» Excel report (ملخص مشاريع سهيل + تفاصيل مشاريع سهيل) from the raw «التقرير التنفيذي لمشاريع سهيل» export, always reading status and progress from the newest weekly Wxx columns. Use this skill whenever the user uploads a Suhail projects Excel file or mentions مشاريع سهيل, محفظة سهيل, التقرير التنفيذي لمشاريع سهيل, ملخص مشاريع سهيل, تفاصيل مشاريع سهيل, or asks for the weekly Suhail status/summary — even if they re-describe the whole layout from scratch instead of naming this skill, even if they only want the sector breakdown or one of the tables, and even for follow-up tweaks to a file this skill already produced.
---

# Suhail Project Report

Turns the raw Suhail portfolio export into one Excel file with exactly two RTL Arabic sheets.

## The rule that matters most: newest week wins

The source has repeating weekly columns:

- `الحالة الفعلية W36`, `الحالة الفعلية W37`, …
- `ما تم حتى تاريخة W35`, `ما تم حتى تاريخة W36`, …

Detect every column containing `W` + a number, take the **highest** number for each family, and use only that column. W38 beats W37 beats W36 — never assume a fixed week, and never merge several weeks together.

**Never use the plain `الحالة` column** when `الحالة الفعلية Wxx` columns exist. The same applies to `ما تم حتى تاريخه` — newest week only.

The chosen week drives everything: the KPI counts, the sector table, the completed and delayed tables, the status column in the details sheet, and the sort order inside each sector.

## التدقيق الإملائي قبل إنشاء ملف Excel

قبل اعتماد المخرج، راجع جميع النصوص المنقولة إلى ملف Excel وصحح الأخطاء الإملائية والطباعية الواضحة فقط. لا تلخص، ولا تعيد الصياغة، ولا تغير المعنى، ولا الأرقام، ولا التواريخ، ولا المعرفات، ولا أسماء الحقول أو التصنيفات المعتمدة. إذا كان التصحيح المحتمل قد يغير المعنى أو كان النص غامضا، اتركه كما هو واذكر الملاحظة للمستخدم. ملف Excel الناتج من هذه المهارة هو المصدر النصي المعتمد للمراحل اللاحقة، لذلك يجب أن يتم التصحيح هنا قبل التسليم.

## Statuses

Only four, in this canonical spelling: `مكتملة` · `على المخطط` · `متأخر` · `لم تبدأ`.

The source often writes `مكتمل` and `لم يبدأ` — map them. **A project whose newest-week status cell is blank counts as `على المخطط`** (the source frequently fills the narrative column but leaves the status empty). Report how many rows this affected so the user can correct the source if needed.

## Python runtime and automatic bootstrap

When installing or updating this skill locally, first run `uv --version`.

- If `uv` is already available, keep the existing installation and continue.
- If `uv` is missing, install it automatically with Astral's official user-scoped standalone installer:
    
    **Windows (PowerShell):**
    
    ```powershell
    powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
    ```
    
    **macOS or Linux:**
    
    ```bash
    curl -LsSf https://astral.sh/uv/install.sh | sh
    ```
    
- After installation, make the installed executable available to the current process (refresh `PATH` or use the executable's installed path) and verify with `uv --version`.
- Do not reinstall or upgrade `uv` when a working installation already exists.
- Do not use `pip` or modify the host Python environment. Python and package requirements are declared in `scripts/build_report.py` using inline metadata.
- If the skill was copied without running its installation step and `uv` is missing at report time, perform the same bootstrap automatically before running the report.
- Stop only if the official installer cannot run because of a real environment restriction such as blocked network access, denied execution, or unsupported platform; report the exact failure.

## Run it

```bash
uv run scripts/build_report.py <source.xlsx> [output.xlsx]
```

Defaults to `/mnt/user-data/outputs/ملخص_مشاريع_سهيل.xlsx`. It prints the week it picked, the project count, and the status breakdown — check that line before delivering. Then present the file.

## Sheet 1 — «ملخص مشاريع سهيل»

Six sections in this order:

1. **أولا: الملخص العام** — five KPI tiles: إجمالي مشاريع سهيل, المشاريع المكتملة, المشاريع على المخطط, المشاريع المتأخرة, المشاريع لم تبدأ.
2. **ثانيا: تفصيل حالات المشاريع حسب القطاع** — columns `القطاع | مكتملة | على المخطط | متأخر | لم تبدأ`. Zeros render as blank cells. Every sector in the fixed list appears even with no projects.
3. **ثالثا: المشاريع المكتملة** — `# | المشروع | القطاع`.
4. **رابعا: المشاريع المتأخرة** — `# | المشروع | القطاع | ملاحظات`, with each delayed project's note copied from the main projects sheet. Blank notes stay blank; shows `لا يوجد` when empty.
5. **خامسا: أبرز التحديثات** — taken from the `أبرز التحديثات` sheet: `# | المشروع | القطاع | التحديث`.
6. **سادسا: التحديات** — `# | المشروع | القطاع | التحدي`. Take nonblank challenges from the `التحديات` sheet when present, one row per challenge. Otherwise use the `التحدي` or `التحديات` column in the main projects sheet. Preserve source order and show `لا يوجد` when no challenges are present. If the source has neither a challenges sheet nor a challenge column, report this rather than inventing entries.

Fixed sector order for section 2:

مركز امتثال أعمال · وكالة التوعية وتنمية المهارات · مركز منصة نسك الرقمية · وكالة التعاون الدولي والشراكات · وكالة تطوير الأعمال وخدمة المستفيدين · وكالة التحول الرقمي وتقنية المعلومات · وكالة شؤون الحج · وكالة شؤون العمرة والزيارة · مركز التخطيط الاستراتيجي

## Sheet 2 — «تفاصيل مشاريع سهيل»

One table per sector, columns `# | اسم المشروع | تاريخ البداية | تاريخ النهاية | الحالة | ما تم حتى تاريخه`.

Sector order: وكالة شؤون الحج → وكالة تطوير الأعمال وخدمة المستفيدين → وكالة التعاون الدولي والشراكات → وكالة التوعية وتنمية المهارات → وكالة التحول الرقمي وتقنية المعلومات → مركز منصة نسك الرقمية → مركز امتثال أعمال. Sectors with no projects are omitted here, so `وكالة شؤون العمرة والزيارة` and `مركز التخطيط الاستراتيجي` do not appear in this sheet — they still get their row in the sector table on sheet 1.

Any sector found in the data but missing from the list above is appended after the last named one rather than dropped. Mention it when it happens.

Inside each sector sort by status: **متأخر → على المخطط → لم تبدأ → مكتملة**, keeping the source order within a status group.

## Sector name normalisation

The export spells two sectors differently from the report. Map them:

- `مركز امتثال الأعمال` → `مركز امتثال أعمال`
- `وكالة تطوير الاعمال وخدمة المستفيدين` → `وكالة تطوير الأعمال وخدمة المستفيدين`

A sector in the data that matches nothing in the fixed lists is appended at the end rather than dropped — mention it to the user.

## Hard limits

- Exactly two sheets. No dashboard, no charts, no extra sheets, no extra columns, no analysis or recommendations sections.
- Both sheets right-to-left, Arabic, wrap text on.
- **Font: `Abar Mid`, size 10, in every single cell** — headers, section titles, KPI figures, table body, all of it. No exceptions and no larger display sizes. Bold and colour carry the hierarchy instead. The constants `F` and `SZ` at the top of the script are the only place this is set.
- Never invent missing data. Dates are copied as-is from the source (they are text, not date values).
- Arabic text carries no diacritics.

## Colours

Ministry palette, already in the script: green `038060` (headers, على المخطط), dark blue `164767` (section titles, مكتملة), gold `D7A562`, `CFA76C` for متأخر, grey `8C8C8C` for لم تبدأ.

