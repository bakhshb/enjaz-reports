# Digital-experience model

All digital PowerPoint tables, including every detail table and challenge table, must use native RTL direction. Preserve the review JSON's PDF field order (`أساس التقييم | الإنجاز | المهمة | م`); map those values into the corresponding RTL table columns without changing the JSON schema. Numbering/codes appear at the right, and Latin task codes retain LTR text direction inside the RTL table. Validate direction independently of paragraph alignment.

Use this reference only after the user chooses **التقرير الموحد مع نضج التجربة الرقمية**. The added summary follows the transactions summary and precedes the detail divider. The divider's «تفاصيل مؤشر نضج التجربة الرقمية» button opens the populated indicator-detail slides after the task and Suhail details; the detail data must not appear only in the summary. The current model remains separate.

## Source contract: the landscape PDF

The current digital-maturity PDF is a required fourth input and arrives as the landscape («horizontal») export. It is a **data source only**; its layout, page breaks and page count are never reproduced.

- **Page one** holds the four perspective percentages, one sector bar row per agency/center (a gray bar with the note inside the label's parentheses marks a sector that is not assessed separately), and the challenges table (`# | المنظور | …`).
- **Every later page belongs to one sector.** A sector may span several consecutive pages titled «القطاع (1/3)», «القطاع (2/3)», …; a continued section heading on a later page reads «المنظور: 15% (تتمة)»; a sector's challenges table («التحديات» or «أبرز التحديثات») may appear on any of its pages. All of these are **source pagination artifacts**. The unit of data is the sector, not the page.
- Pages without a «(k/n)» suffix are single-page sectors. Sectors keep their source order.

`scripts/digital_maturity.py` reads the landscape pages, restores logical Arabic from the visual text layer (Latin codes, digits, percentages and brackets included), and folds each sector's pages into one record set. The JSON keys are `source_pdf_sha256`, `source_date_hijri`, `metrics` (`experience`, `satisfaction`, `complaints`, `tools`), `sectors` (name, four metrics, optional note), summary `challenges` (perspective and text), `detail_sectors` (one entry per sector: `title` without suffix, `source_pages`, `cards`, four-column `rows` including section headings but no «(تتمة)», `challenge_title`, `challenge_header`, `challenges` from all of the sector's pages in source order) and `reviewed`. The extractor begins with `reviewed: false` because the PDF text layer can still reverse glyphs, lose characters as `(cid:…)`, or wrap codes; compare every field with the rendered source pages, correct the JSON, then set it to `true`. If any row or challenge is illegible or ambiguous, stop and request a clearer source. Do not infer runtime data from master examples. Never merge, reword, renumber or drop challenges while reviewing; they transfer exactly as written.

The master's detail slides are **design patterns only**: the first one supplies the title strip, the header row, one section-heading row per perspective (its fill color), the body row, the four cards and the challenge table. Never carry sample percentages, rows, or challenge prose from the PowerPoint master into a generated report. The master summary still contains sample values that differ from the supplied PDF; the PDF controls runtime values. Preserve the master's fixed «آلية الاحتساب» note between the four indicators and sector chart, including its wording and emphasis; the source's own «أساس الاحتساب» line is not extracted.

## Output pagination (portrait slides)

The builder re-flows every sector with the same continuous-flow rules used for task and Suhail details:

- One sector starts on a fresh detail slide; its records fill the slide down to the footer clearance. The number of digital detail slides follows the data, not the number of source pages.
- Records are never split, and a section heading never ends a slide. When a section continues on the next slide, the continuation starts with the same heading plus «(تتمة)» (generated from the output pagination, never copied from the source).
- A sector that needs more than one slide titles them «القطاع (1/n)» … «القطاع (n/n)» where *n* is the number of **output** slides; a single-slide sector has no suffix. The four cards repeat with identical values on every slide of the sector.
- The sector's challenges table appears **once**, directly below the last records of the sector with the small approved gap, titled as in the source («التحديات / التحدي» or «أبرز التحديثات / التحديث»). If it cannot fit under the last record, the trailing record moves with it to the next slide; a slide never holds only a challenge table.
- The task/indicator column keeps the master's fixed header; the source's «المؤشر» header is not transferred.
- Surplus master detail slides are removed; the remaining block stays contiguous after the Suhail details and before the thank-you slide, and the divider button targets its first slide.

The summary slide has room for up to five sector rows and four challenge rows; if a future source exceeds that space, stop for a revised layout rather than silently drop records or shrink the 11 pt table font. The four perspective values and sector percentages are independent rates, not portions that must sum to 100. Preserve every number as shown. A missing perspective for one sector stays blank. A sector marked with a note has a labeled gray bar and no invented percentages. Keep the challenge wording and perspective in source order.

Reconcile the generated PPTX with the reviewed JSON and its source PDF hash. Inspect the native chart caches, embedded data, percentage labels, progress fills, note bars, every detail table and challenge table. Render all populated digital slides and compare each sector with its source pages and the unchanged master design. Correct any clipping, RTL reversal, stale sample value, overlap, premature page break or footer collision before delivery. Never silently truncate or shrink the 11 pt table font.
