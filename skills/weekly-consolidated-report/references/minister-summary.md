# Minister messages (رسالتا معالي الوزير)

After the final report is delivered, reply with two Arabic messages to the Minister, each in its own fenced code block (```` ```text ````) so the user can copy it, and save both as text files next to the report:

1. the **short message** (`رسالة معالي الوزير - مختصرة.txt`), and
2. the **long message** (`رسالة معالي الوزير - مفصلة.txt`).

Write nothing inside a code block except the message. Both formats were set by the user on 26 November 2026.

## Source

Both messages are built only from the delivered report. Run:

```bash
uv run scripts/minister_messages.py --report "<delivered.pptx>"
```

It reads the delivered PPTX (never the Excel or PDF inputs) and writes next to it:

- `<report>.message-facts.json`: every count, table row, percentage and note the messages may use;
- `<report>.short-message.txt`: the short message, complete;
- `<report>.long-message.draft.txt`: the long message with every fixed and counted part already written (salutation, date, counts sentences with number–noun agreement, completed tasks, percentages, the sentences for empty sections, closing). Only the parts in `[…]` are left to write.

Never change a number, date, name or count the script wrote, and never add a fact, number, name or judgement that is not in the facts file. When the report is rebuilt, run the script again on the new report.

## Wording

Formal government Arabic, the same register as the report's digital notes: فصحى رسمية، جمل خبرية قصيرة، بلا عامية, no judgement or adjective the report does not state. Plain text only: no Markdown, emoji or tables; «-» for bullets and «1.» for numbered blocks. The short message keeps the user's own fixed phrases below as written (including «اي» and «شغالين»), because they are the user's style for that message.

## Message 1: short

```text
معاليك
- <المهام>
- <المعاملات>
- <مبادرات سهيل>
- <مؤشر نضج التجربة الرقمية>        (digital model only)
تحياتي
```

One line per part, nothing else. The script applies these rules:

- **المهام**: no delayed or suspended task → «كل المهام المسندة من <مصادر المهام> تسير على المخطط», the sources taken from the report's «تفاصيل المهام (…)» titles; otherwise «… على المخطط عدا <العدد> متأخرة/معلقة».
- **المعاملات**: «لا يوجد اي تأخير في المعاملات», or «يوجد <العدد> معاملات متأخرة».
- **مبادرات سهيل**: «مبادرات سهيل على المخطط», or «… عدا <العدد> متأخرة/لم تبدأ».
- **مؤشر نضج التجربة الرقمية**: «مؤشر نضج التجربة الرقمية على المخطط», plus «، فيه تأخير من <القطاع>» for each sector whose notes report «تأخر». The line then ends with the user's own judgement, which is not in the report: «، ولكن لا يستدعي التصعيد الآن، شغالين مع … على تجاوز التأخير.» Tell the user, outside the code block, that the «…» and the escalation judgement are theirs to confirm. The digital part is always one line.

## Message 2: long

```text
معالي الوزير حفظه الله

السلام عليكم ورحمة الله وبركاته،

تجدون أدناه ملخص حالة المهام ومشاريع سهيل والمعاملات، بالإضافة إلى مستجدات مؤشر نضج التجربة الرقمية، وفق تقرير الإنجاز الأسبوعي المحدث بتاريخ <تاريخ الغلاف>م.

أولا: المهام
<جملة العدد>
المهام المكتملة:
1. <الجهة>
المهمة: <نص المهمة>
المستجد: <الملاحظة>
(ثم «المهام المتأخرة:» و«المهام المعلقة:» بالشكل نفسه إن وجدت)

ثانيا: مشاريع سهيل
<جملة العدد>
أبرز المستجدات:
- <المشروع>: <التحديث>
(ثم «المشاريع المتأخرة:» إن وجدت)
التحديات:
- <المشروع>: <التحدي>

ثالثا: المعاملات
<جملة العدد>

رابعا: مؤشر نضج التجربة الرقمية
تتواصل أعمال رفع مؤشر نضج التجربة الرقمية عبر أربعة مناظير رئيسية، حيث بلغت نسبة الإنجاز في <المناظير الأربعة بنسبها>.
وتعكس النسب أدناه تقدم استكمال المتطلبات والأدلة، بما يشمل الجاهزية وتحديد تواريخ التنفيذ والأعمال المنجزة، ولا تمثل نتيجة التقييم النهائي.
1. منظور <المنظور> – نسبة الإنجاز <النسبة>
- بلغت نسبة الإنجاز لدى <الجهة> <النسبة>، ولدى …
- <المستجدات>
- أبرز التحديات: <التحديات>
(لكل منظور من الأربعة، بترتيب البطاقات)

خامسا: طلبات الدعم المرفوعة لمعاليكم
<الطلبات، أو جملة عدم وجودها>

وتفضلوا بقبول خالص التحية والتقدير.
```

Blank lines separate every paragraph, heading and block, as in the draft. For the `standard` model the intro drops «بالإضافة إلى مستجدات مؤشر نضج التجربة الرقمية»، the digital section is left out and support requests become «رابعا»; the script numbers the sections.

What the agent writes in the `[…]` parts:

- **Completed, delayed and suspended tasks**: the script copies each row of the report's table; correct only grammar (e.g. «تم مشاركة … لمركز» → «تمت مشاركة … مع مركز»), keep the meaning and every number and date.
- **Suhail «أبرز المستجدات» and «التحديات»**: one bullet per row of the report's tables, in order. Condense the update into one or two formal sentences; keep its dates, counts and targets, drop the «تاريخ التحديث» line.
- **Each digital perspective**: the script writes the sector percentages from the sector chart, highest first; put a sector shown with a note (gray bar) at the end of the first perspective's line as a formal clause. Then write «المستجدات» and «أبرز التحديات» from the facts file: the summary «أبرز التحديثات» rows of that perspective, the sector updates and challenges that concern it, and the most telling requirement rows of that perspective (`sectors_detail.<sector>.rows.<perspective>`: task, achievement, basis). Progress, completions and expected dates go under «المستجدات»; a pending decision, an obstacle, a returned or blocked file, a dependency or a requirement still under study goes under «أبرز التحديات». Write «لم يتضمن التقرير تحديات لهذا المنظور.» only when none of these sources has one. A note that spans several perspectives may be used under each it concerns; do not repeat the same sentence twice in one perspective.
- **Support requests**: one bullet per row of the report's «طلبات الدعم» table. When there is none, keep the sentence the script wrote; it may be followed by «، فيما تبرز الحاجة إلى …» listing only decisions or actions that the report's challenges explicitly call for (e.g. a pending decision, a missing response, a deadline).

### Empty sections

A section with no data is never dropped; the script writes a formal sentence in its place:

| Case | Sentence |
| --- | --- |
| No completed tasks | لا توجد مهام مكتملة ضمن التقرير. |
| No Suhail updates | لم يتضمن التقرير مستجدات لمشاريع سهيل. |
| No Suhail challenges | لم يتضمن التقرير تحديات لمشاريع سهيل. |
| No sector percentage in a perspective | لم يتضمن التقرير نسب إنجاز للجهات في هذا المنظور. |
| No challenge for a perspective | لم يتضمن التقرير تحديات لهذا المنظور. |
| No support request | لم يتضمن التقرير طلبات دعم مستقلة مرفوعة لمعاليكم. |
| A zero total | لم يتضمن التقرير مهام. / لم يتضمن التقرير مشاريع لسهيل. / لم يتضمن التقرير معاملات. |

Delayed and suspended task blocks and the delayed-project block appear only when the report has rows for them; their absence is already stated in the counts sentence («ولا توجد مهام متأخرة أو معلقة …»).
