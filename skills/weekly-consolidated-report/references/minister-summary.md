# Minister summary message (رسالة ملخص التقرير لمعالي الوزير)

After the final report is delivered, reply with one short Arabic message to the Minister inside a single fenced code block (```` ```text ````) so the user can copy it. Write nothing inside the block except the message.

The user approved this format on 26 November 2026. It replaces the earlier long message with numbered sections; do not go back to it.

## Shape

One opening word, one line per part of the report, one closing word. Nothing else: no date line, no counts sentence, no section headings, no task or project list.

```text
معاليك
- <سطر المهام>
- <سطر المعاملات>
- <سطر مبادرات سهيل>
- <سطر مؤشر نضج التجربة الرقمية>
تحياتي
```

- `standard` model: drop the digital-experience line. The other three stay in this order.
- `digital` model: all four lines.
- Plain text only: no Markdown bold, no emoji, no tables. Each part line starts with «- ».

## Where each line comes from

Build every line only from the delivered report (its summary slides and, for the digital line, its digital-experience detail slides). Never add a fact, number, name or judgement that is not in the report.

**المهام.** Name the task sources from the report's «تفاصيل المهام (…)» titles, in report order, joined with «و».
- No delayed and no suspended tasks: «كل المهام المسندة من <المصادر> تسير على المخطط».
  Example with one source: «كل المهام المسندة من اجتماع القيادات تسير على المخطط».
- Otherwise state what is late or suspended with its count and sector, from the «المهام المتأخرة» / «المهام المعلقة» tables, e.g. «المهام المسندة من اجتماع القيادات على المخطط عدا مهمة متأخرة لدى وكالة شؤون الحج».

**المعاملات.**
- No delayed transactions: «لا يوجد اي تأخير في المعاملات».
- Otherwise: «يوجد <العدد> معاملات متأخرة» with correct number–noun agreement (معاملة واحدة، معاملتان، 3–10 معاملات، 11+ معاملة).

**مبادرات سهيل.**
- No delayed and no not-started projects: «مبادرات سهيل على المخطط».
- Otherwise: «مبادرات سهيل على المخطط عدا <العدد> متأخرة» (or «لم تبدأ»), naming the projects from «المشاريع المتأخرة» when there are one or two.

**مؤشر نضج التجربة الرقمية** (`digital` only).
- Start with «مؤشر نضج التجربة الرقمية على المخطط».
- If a sector's challenges (التحديات) or the summary «أبرز التحديثات» report a delay (a title or text with «تأخر»), add «، فيه تأخير من <القطاع>» naming each such sector as the report names it.
- How serious the delay is, whether it needs escalation, and who is working on it are the user's own judgement and are not in the report. Never write them from the report. Put them on one line right after the digital line for the user to complete, starting with «ولكن» and leaving the unknown part as «…», for example:
  «ولكن لا يستدعي التصعيد الآن، شغالين مع … على تجاوز التأخير.»
  Tell the user, outside the code block, that this line is theirs to complete.

## Approved example (26 November 2026, digital model)

The user's own wording; values and names belong to that week only.

```text
معاليك
- كل المهام المسندة من اجتماع القيادات تسير على المخطط
- لا يوجد اي تأخير في المعاملات
- مبادرات سهيل على المخطط
- مؤشر نضج التجربة الرقمية على المخطط في تأخير طفيف من وكالة شؤون الحج، ولكن لا يستدعي التصعيد الان ، شغالين مع الدكتور عمار الصائغ من نسك على تجاوز التأخير.
تحياتي
```

In that example, «طفيف», «لا يستدعي التصعيد» and the named contact came from the user, not from the report.

If the report is rebuilt after an edit, regenerate the message from the new report.
