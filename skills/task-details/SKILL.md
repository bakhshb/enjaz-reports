---
name: task-details
description: "Consolidate a raw 'حالة المهام الاستراتيجية' (strategic tasks status) Excel export — a tasks sheet, with a blank manual-entry support-requests table in the output — into one clean two-sheet Arabic workbook: 'ملخص المهام' (overall counts, per-sector breakdown, completed-tasks table, delayed-tasks table, suspended-tasks table, support-requests table) and 'تفاصيل المهام' (every task grouped into its own section by المصدر / source). Use this skill whenever the user uploads a tasks/متابعة المهام Excel file and asks for a ملخص المهام, تفاصيل المهام, تصنيف المهام حسب القطاع or حسب المصدر, or a consolidated tasks status report — even if they re-describe the layout from scratch instead of naming this skill by name."
---

# Task Details — ملخص وتفاصيل المهام الاستراتيجية

Builds one workbook, exactly two sheets, from a raw tasks export. No charts, no

percentages, no dashboards, no analysis or recommendations, no Pivot Tables in

the delivered file, no extra sheets or columns — the request is for a plain,

literal restructuring of the source data, and anything beyond spec is a

defect, not a bonus.

## Source file shape

Only the main tasks sheet is required in the upload:

- **Main tasks sheet** (name varies — detect it by the presence of a الحالة
    
    column, don't assume a fixed sheet name). Relevant columns: المهمة, القطاع,
    
    المصدر, تاريخ الإنجاز المخطط (spelling varies — "الانجاز" / "المخطط له" are
    
    the same field), الحالة, الملاحظات, جهات ذات علاقة. Other columns in the
    
    source (مجمعة حسب, مدة التأخير باليوم, تاريخ الإنجاز الأساسي, عدد تكرار طلبات
    
    الدعم, …) are **not** part of any requested table — leave them out entirely.
    
`scripts/build_report.py` detects the main tasks sheet and its required columns by alias. Missing task columns remain errors. Do not search for, require, validate, or import a support-requests sheet: طلبات الدعم is a manual-entry output table, even if support data exists in the source. Do not report “لا توجد طلبات دعم”.

## التدقيق الإملائي قبل إنشاء ملف Excel

قبل اعتماد المخرج، راجع **جميع النصوص المنقولة إلى ملف Excel مراجعة إملائية وطباعية شاملة**، وصحح الأخطاء الواضحة في جميع الحقول النصية دون استثناء.

يجب أن يشمل ذلك نصوص المهام، الملاحظات، طلبات الدعم، الجهات ذات العلاقة، أسماء المصادر، وأي نصوص أخرى يتم نقلها إلى الملف.

لا تلخص، ولا تعيد الصياغة، ولا تغير المعنى، ولا الأرقام، ولا التواريخ، ولا المعرفات، ولا أسماء الحقول أو التصنيفات المعتمدة. إذا كان التصحيح المحتمل قد يغير المعنى أو كان النص غامضا، اتركه كما هو واذكر الملاحظة للمستخدم.

ملف Excel الناتج من هذه المهارة هو المصدر النصي المعتمد للمراحل اللاحقة، لذلك يجب أن يتم التدقيق والتصحيح هنا قبل التسليم.

## Status normalisation

Map every spelling variant onto exactly these four categories — never invent a

fifth:

| In the source | In the report |
| --- | --- |
| مكتمل، مكتملة، منجز، منجزة | مكتملة |
| على المخطط، علي المخطط | على المخطط |
| متأخر، متاخر، المتأخر، متأخرة | متأخر |
| معلق، معلقة | معلق |

If a status shows up that isn't covered here, stop and surface it to the

user instead of forcing it into one of the four buckets — this mirrors the

original spec's "لا تخمّن التصنيف" rule. `build_report.py` already does this:

it exits with the unmapped value(s) named rather than producing a file.

A status category with zero matching tasks still gets its line in the general

summary with the value `0` (e.g. "المهام المكتملة: 0" is a real, expected

answer, not a bug) — only the *per-sector* table hides zeros (see below).

## Output layout

Two sheets, both right-to-left, font **Abar Mid, size 10** throughout (titles,

headers, and body cells alike — this is a fixed rule, not a default to

override), Arabic text with **no diacritics** anywhere Claude writes it

(titles, headers, labels).

### Sheet 1 — "ملخص المهام"

In this exact order, each table preceded by a dark navy (`1F4E78`) title band

with white bold text spanning the full table width, and a light-gray

(`D9D9D9`) bold header row for each data table:

1. **الملخص العام للمهام** — five label/value lines, values computed directly
    
    from the data (never hardcoded from a prior run): إجمالي عدد المهام،
    
    المهام المكتملة، المهام على المخطط، المهام المتأخرة، المهام المعلقة.
    
2. **تفصيل حالات المهام حسب القطاع** — table `القطاع | مكتملة | على المخطط |
   متأخر | معلق`, one row per sector that has at least one task (order:
    
    first-appearance order in the source, which is deterministic and mirrors
    
    how the ministry lists sectors). A cell whose count is 0 is left blank,
    
    never written as "0" — this is the one place zeros are hidden.
    
3. **المهام المكتملة** — `# | المهمة | القطاع | ملاحظات`, only tasks whose normalised
    
    status is مكتملة. Carry each task's الملاحظات from the source; keep an empty
    note blank. Sequential numbering from 1. If there are none, the title
    
    and header row still render, with no data rows underneath — an empty
    
    section is a correct answer when the source has no completed tasks, not a
    
    sign something broke.
    
4. **المهام المتأخرة** — `# | المهمة | القطاع | ملاحظات`, filtered to متأخر.
    Carry each task's الملاحظات from the source; keep an empty note blank.
5. **المهام المعلقة** — `# | المهمة | القطاع | ملاحظات`, filtered to معلق.
    Carry each task's الملاحظات from the source; keep an empty note blank.
    Render the title and header even when there are no suspended tasks.
6. **طلبات الدعم** — `# | المهمة | القطاع | طلب الدعم`. Always render the title, headers, and one formatted empty entry row. Leave all four cells blank, including numbering; the user fills this table manually. Never generate or copy support requests.

### Sheet 2 — "تفاصيل المهام"

One section per **المصدر** value actually present in the data, each its own

title band reading `مهام مصدر (اسم المصدر)`, followed by a header row and then

only that source's tasks:

`# | المهمة | القطاع | تاريخ الانجاز المخطط | الحالة | ملاحظات | الجهات ذات
العلاقة`

**Section order is fixed, not data-driven:** اجتماع القيادات، ثم اجتماع

التجربة الرقمية، ثم لجنة المتابعة. A source in that list that has no tasks in

this export is simply skipped (no empty section). Any source *not* in the list

is appended after the three, in first-appearance order.

**Row order inside each section: status first, always.** Order the tasks by

status — **متأخر ← معلق ← مكتملة ← على المخطط** — so the section opens on what

needs attention and closes on what is done. Within one status block, order by

القطاع using the same sector order as the per-sector table on sheet 1 (first-

appearance order in the source). Rows tying on both keys keep their source-file

order (stable sort). Note this means a sector's tasks are deliberately split

across status blocks rather than kept together.

The الحالة column shows the **normalised** status (مكتملة، على المخطط، متأخر،

معلق), not the raw spelling from the export, so it matches the wording used in

the summary sheet and the sort order above.

Numbering (`#`) restarts at 1 inside every source section — it follows the

sorted order, so it is a display sequence, not the source row number. Never

merge two differently-spelled source names unless they are unmistakably the

same value with only incidental whitespace differences — when in doubt, keep

them separate and let the user confirm.

### Formatting rules that apply everywhere

- Wrap text on: المهمة, الملاحظات, الجهات ذات العلاقة, طلب الدعم (any column
    
    that can hold a full sentence).
    
- Thin gray borders on every table cell; numeric/status columns centered,
    
    text columns right-aligned.
    
- Column widths sized generously for the columns that carry the long text
    
    above (~45–55) even though this leaves surplus whitespace in the narrower
    
    numeric tables sharing the same column letters — that trade-off is expected
    
    and preferred over truncated task text.
    
- Dates are kept exactly as formatted in the source (already consistent
    
    `dd-mm-yyyy` strings in every export seen so far) — don't reparse or
    
    reformat them unless a future export mixes formats, in which case unify to
    
    `dd-mm-yyyy` and say so.
    
- **No formulas.** The numbers are computed once in Python directly from the
    
    source data and written as values — there's no live reference to a "raw
    
    data" sheet inside the delivered workbook to formula against, so a formula
    
    would have nothing to recalculate from. Verify the counts in Python instead
    
    (the script's printed checks below) and don't run a LibreOffice recalc pass
    
    on this file.
    

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

## Running it

```bash
uv run scripts/build_report.py --input /mnt/user-data/uploads/<source>.xlsx \
  --output "/mnt/user-data/outputs/ملخص وتفاصيل المهام الاستراتيجية.xlsx"
```

The script prints the total, the four status counts, the sector list, the

source list, and two cross-check confirmations (sector counts sum to the

totals; source-group counts sum to the total). Read that output before

replying — if it exits with a missing-column or unmapped-status message,

relay that to the user rather than trying to patch around it silently.

## Final verification before delivering

- إجمالي = مكتملة + على المخطط + متأخر + معلق (the script asserts this).
- مجموع كل حالة عبر القطاعات = نفس رقم الملخص العام لتلك الحالة (asserted).
- مجموع عدد المهام عبر كل مصادر تفاصيل المهام = الإجمالي (asserted).
- ترتيب أقسام تفاصيل المهام يطابق الترتيب الثابت أعلاه، وداخل كل قسم تتدرج
    
    المهام متأخر ثم معلق ثم مكتملة ثم على المخطط، والقطاعات مرتبة داخل كل حالة.
    
- جداول المهام المكتملة والمتأخرة والمعلقة لا تحتوي إلا على حالات "مكتملة"
    
    و"متأخر" و"معلق" على التوالي (true by construction — built via filtered dataframes).
    
- لا تكرار غير مقصود لنفس نص المهمة (script warns on stderr; review before
    
    sending if it fires — it's a warning, not a hard stop, since two genuinely
    
    distinct tasks can coincidentally share wording).
    

Then hand the user only the file plus, if relevant, one short note on data

caveats (statuses that had to be flagged, an empty مكتملة/متأخر section) — the

workbook itself carries no analysis, so the reply shouldn't smuggle any in.


## Manual-entry acceptance

A tasks-only input produces the normal two sheets and a formatted, empty support table. Existing source support rows must not populate it. No missing-support warning or fabricated request is permitted. Preserve all task counts, status rules, dates, and ordering.
