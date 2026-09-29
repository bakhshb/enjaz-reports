---
name: delayed-transactions-summary
description: "Consolidate per-sector Arabic transaction exports (تقرير انجاز المعاملات الواردة للوزارة) into one Excel file with a four-line summary (اجمالي / مكتملة / على المخطط / متأخرة), a status breakdown by sector (تفصيل حالات المعاملات حسب القطاع), and a table of the delayed transactions only, grouped by sector. Use this skill whenever the user uploads several .xls or .xlsx files that each represent a قطاع / وكالة / مركز and asks for a ملخص المعاملات, تقرير موحد, المعاملات المتأخرة, or a consolidated transactions file — even if they describe the layout from scratch instead of naming this skill, and even if they only ask for the delayed-transactions table, the per-sector status table, or only for the counts."
---

# ملخص المعاملات المتأخرة — Delayed transactions summary

Builds one consolidated Excel workbook from several per-sector exports of the

ministry's incoming-transactions report (`تقرير انجاز المعاملات الواردة للوزارة`).

The workbook has exactly three blocks and nothing else: a four-line summary at the

top, a status breakdown by sector beneath it, and then a table of the

transactions whose status is `متأخرة`. No charts, no percentages, no dashboards,

no total row, no extra columns, no recommendations — the user asks for this file to be plain, and extra content is

treated as a defect, not a bonus.

## The two questions to settle before building

Ask these up front (in one short message) unless the user already answered them:

1. **Duplicates.** The same transaction number often appears in two sectors, or
    
    twice inside one file, because it was routed to more than one party. Should
    
    the summary count every record (`اجمالي = عدد السجلات`) or count each
    
    transaction number once? Both are legitimate readings; the counts differ by
    
    the number of duplicate records, so never pick silently. Report the duplicate
    
    count either way so the user can sanity-check the total.
    
2. **Sector order** in the sector breakdown and delayed tables, if they care.
    
    Default is Arabic alphabetical, which is deterministic and easy to defend; both
    
    tables use the same order.
    

Everything else in this skill is fixed and needs no asking.

## Source file anatomy

The exports are legacy `.xls` (OLE2) files written by the ministry's reporting

system. `pandas.read_excel` needs `xlrd` for these; `xlrd` is declared in the script's inline `uv` dependencies and must not be installed with `pip`.

Every file has the same shape, so read it positionally rather than by column

name:

- Row 1, column 12: the sector name (`مركز امتثال الأعمال`, `وكالة شؤون الحج`, …).
    
    Use it as the القطاع value. Fall back to the file name only if that cell is empty.
    
- The header row is the first row containing `رقم المعاملة` (row 12 in every file
    
    seen so far — detect it, don't hardcode it). Data starts on the next row.
    
- Column indices in the header row: `19` الحالة · `23` تاريخ الانجاز المخطط له ·
    
    `26` الجهة الوارد منها المعاملة · `27` تاريخ انشاء المعاملة · `28` موضوع المعاملة ·
    
    `30` رقم المعاملة. The other columns (تفاصيل الانجاز, مدة البقاء, تاريخ الاحالة,
    
    مدة التاخير) are not needed.
    

Two traps in the raw values:

- **Bidi control characters.** Dates arrive wrapped in `U+202B … U+202C`. Strip
    
    `\u202a-\u202e`, `\u200e`, `\u200f`, `\u061c` from every field before use, or
    
    the values sort and compare wrongly and look broken in the output.
    
- **Transaction numbers must stay text.** They are identifiers, not amounts. If
    
    they become numeric, a long one like `381117010117` renders as
    
    `381117010117.00`. Write them as strings with number format `@`.
    

## التدقيق الإملائي قبل إنشاء ملف Excel

قبل اعتماد المخرج، راجع جميع النصوص المنقولة إلى ملف Excel وصحح الأخطاء الإملائية والطباعية الواضحة فقط. لا تلخص، ولا تعيد الصياغة، ولا تغير المعنى، ولا الأرقام، ولا التواريخ، ولا المعرفات، ولا أسماء الحقول أو التصنيفات المعتمدة. إذا كان التصحيح المحتمل قد يغير المعنى أو كان النص غامضا، اتركه كما هو واذكر الملاحظة للمستخدم. ملف Excel الناتج من هذه المهارة هو المصدر النصي المعتمد للمراحل اللاحقة، لذلك يجب أن يتم التصحيح هنا قبل التسليم.

## Status normalisation

The files use `مكتمل` / `على المخطط` / `متأخر`. Map every value onto exactly three

categories and invent no fourth one:

| In the files | In the report |
| --- | --- |
| مكتمل، مكتملة، منجز، منجزة | مكتملة |
| على المخطط، علي المخطط | على المخطط |
| متأخر، متاخر، المتأخر | متأخرة |

If a status appears that this table doesn't cover, stop and show it to the user

rather than forcing it into a bucket. A category with no rows still gets its

line in the summary with the value `0`; omitting a supported category would

look like an oversight.

## Output layout

One sheet, right-to-left, Arabic without diacritics.

Font: `Abar Mid` at 10 pt for every cell in the workbook — headings, summary

labels, numbers, table header and data alike. This is fixed and has no

exceptions: do not scale headings up for emphasis. Section bands and the table

header are already distinguished by bold weight and fill colour, which is enough

at a uniform size. Write the name exactly as `Abar Mid`; openpyxl only records

the name, so it renders correctly on machines where the font is installed and

falls back to the system default elsewhere.

```jsx
A1:B1   الملخص العام للمعاملات        (section band, not merged)
A2:B2   اجمالي عدد المعاملات      | n
A3:B3   المعاملات المكتملة        | n
A4:B4   المعاملات على المخطط      | n
A5:B5   المعاملات المتأخرة        | n
(row 6 blank)
A7:D7   تفصيل حالات المعاملات حسب القطاع   (section band, not merged)
row 8   header: A القطاع | B مكتملة | C على المخطط | D متأخر
row 10+ one row per sector
(one blank row)
next    جدول المعاملات المتأخرة       (section band A:F, not merged)
next    header row
then    one delayed transaction per row
```

### No merged cells anywhere

The user copies ranges out of this sheet, and merged cells break copy/paste —

they asked explicitly for no merges. So nothing in the workbook is merged: not

the section bands, not the headers, not the sector column. A section band is

built by giving every cell in its span the green fill and writing the text in

column A only; the text overflows across the empty filled cells and looks the

same as a merged band. The script's verification fails if any merged range

exists.

### Sector breakdown table

- Headers exactly as the user wrote them: `القطاع | مكتملة | على المخطط | متأخر`
    
    (note `متأخر`, not `متأخرة`, in this table only).
    
- القطاع is the sector name as it appears inside the source file itself (row 1,
    
    column 12) — never a shortened, renamed or "official" variant. It sits in
    
    column A alone (never merged); the three counts go in B, C, D. Column A is
    
    44 wide so the longest sector name fits on one line, with wrap as a fallback.
    
- One row for every sector file read, even if all its counts are zero, in the
    
    same order as the delayed table.
    
- A zero count is an empty cell, matching the user's sample. (The four summary
    
    lines above still show `0`.)
    
- Counts are computed from the same rows as the summary (after `--dedupe` if it
    
    was used), so each column sums to its summary line and the whole table sums to
    
    اجمالي. No total row.
    

### Delayed transactions table

Table columns, in this order and no others (column A is the rightmost on an RTL

sheet, so this order reads correctly as written):

`رقم المعاملة | موضوع المعاملة | القطاع | الجهة الوارد منها المعاملة | تاريخ إنشاء المعاملة | تاريخ الإنجاز المخطط`

Sorting: by sector, so a sector's rows are contiguous; within a sector by

`تاريخ الإنجاز المخطط` oldest first. The dates are Hijri text — sort on the

`(year, month, day)` triple parsed out of `dd/mm/yyyy`, never on the string.

Dates keep their original values; only the display form is unified to

`dd/mm/yyyy`. A required field that is missing stays an empty cell — never guess it.

Styling that matches the ministry decks: section bands in dark green `0F5C3F`

with white bold text, table header in gold `D7A562` with green bold text,

alternating row fill `F2F6F4`, thin grey borders, landscape and fit-to-width for

printing. No freeze panes: the delayed header now sits below the sector table,

and freezing there would lock most of the screen.

## Do not put formulas in this workbook

This is the one counter-intuitive rule. Formulas would force a LibreOffice

recalculation pass, and that pass rewrites text-formatted transaction numbers as

real numbers — which is exactly the `381117010117.00` defect above. The four

summary numbers are counts over external source files that the sheet does not

contain, so a formula could not recompute them anyway. Write them as values,

verify them in Python, and record where they came from in a cell comment on

`A1` (source files, dedupe decision, duplicate count, status mapping).

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

`scripts/build_report.py` does the whole job:

```bash
uv run scripts/build_report.py --input-dir /mnt/user-data/uploads \
  --output "/mnt/user-data/outputs/تقرير_المعاملات_الموحد.xlsx" [--dedupe]
```

`--dedupe` counts each transaction number once; without it every record counts.

The script prints the counts, the duplicate count, any unmapped status, and the

verification line — read that output before replying, and do not open a

LibreOffice recalculation on the result.

Every selected source file must contain the transaction header and required cells; a malformed file stops the run rather than being skipped. With `--dedupe`, duplicate transaction numbers that disagree on status or sector stop for a decision instead of keeping an arbitrary first record. The script publishes the output path only after its saved-workbook verification passes; a failed verification exits nonzero and leaves any earlier output untouched.

## Final verification before delivering

- `اجمالي = مكتملة + على المخطط + متأخرة`. If it doesn't balance,
    
    there is an unmapped status; surface it instead of forcing the numbers.
    
- Re-open the saved file and confirm the transaction-number cells are still
    
    strings (`cell.data_type == 's'`).
    
- Confirm the sheet has no merged cell ranges at all.
- Confirm the sector table has one row per sector file, and that each of its
    
    three columns sums to the matching summary line (blank = 0).
    
- Confirm the delayed table row count equals the `المعاملات المتأخرة` line.
- Confirm every populated cell is `Abar Mid` at size 10.

Then tell the user the four numbers, the sector breakdown of the delayed rows,

and any data caveats worth knowing (duplicates found, statuses absent from the

data, dates being Hijri). Keep it to that — the file itself carries no analysis,

so the message shouldn't smuggle any in either.

