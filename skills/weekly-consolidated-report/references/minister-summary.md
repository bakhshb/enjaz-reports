# Minister messages (رسالتا معالي الوزير)

After the final report is delivered, reply with two Arabic messages to the Minister, each in its own fenced code block (```` ```text ````) so the user can copy it, and save both as text files next to the report:

1. the **short message** (`رسالة معالي الوزير - مختصرة.txt`), and
2. the **long message** (`رسالة معالي الوزير - مفصلة.txt`).

Write nothing inside a code block except the message. The structure of both messages was set by the user on 26 November 2026; their wording was made formal on 27 November 2026 (see «Wording»).

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

When both messages are written, check them before sending:

```bash
uv run scripts/minister_messages.py --report "<delivered.pptx>" --check "<long message.txt>" --check "<short message.txt>"
```

The check lists, line by line: a bracket or «…» left from the draft, a number that is not in the facts file, a colloquial word, a machine-sounding phrase, a dash or other markup, and words in Latin letters. Fix every finding yourself and run it again; do not ask the user. Two findings are the user's to settle and stay in the reply as a question: the «…» in the short message's digital line, and a Latin name that has no Arabic form.

## Wording

The Minister reads these messages as he reads a memo from his office: formal government Arabic written by a person. فصحى رسمية، جمل خبرية قصيرة، بلا عامية. Plain text only: no Markdown, emoji or tables; «-» for bullets and «1.» for numbered blocks.

Write each sentence so that it states one fact from the report:

- Start with what happened or what the number is: «اكتمل…»، «اعتمد…»، «بلغت نسبة الإنجاز…»، «تأخر…»، «يجري العمل على…».
- Keep the report's own names. One name for one thing through the whole message: «مشاريع سهيل» stays «مشاريع سهيل», a sector keeps the name the report gives it.
- Keep every number, date and percentage as the report shows it. No «نحو» or «قرابة», no rounding.
- Join with «و» and «ثم». A sentence that needs more than two clauses is two sentences.
- Say a need as a need: «المطلوب: التوجيه باعتماد قرار توحيد مركز الاتصال.»

Never write the patterns below. They are what makes a text read as machine-written. `report_common/wording.py` catches the common ones («يعكس»، «بالإضافة إلى» at the start of a sentence، «ليس فقط … بل»، «يشار إلى أن»، «ختاما»), but the list in this file is what you follow, and a clean check does not replace rereading the text:

| Pattern | Examples | Write instead |
| --- | --- | --- |
| Inflated importance | «يعدّ من أبرز»، «يلعب دورا محوريا»، «حيوي»، «جوهري»، «رائد»، «نقلة نوعية» | the fact and its number |
| A tail the report did not state | «، مما يعكس…»، «، مما يسهم في…»، «، ما يؤكد…»، «بما يعزز…» | end the sentence at the fact |
| «يعكس» for «is» | «تعكس النسب تقدم…» | «تقيس هذه النسب…»، «هذه النسب هي…» |
| Lead-ins | «تجدر الإشارة»، «لا يخفى»، «من المهم الإشارة»، «في إطار الجهود»، «في ظل»، «حرصا على» | the information itself |
| Sentence openers | «علاوة على ذلك»، «بالإضافة إلى ذلك» | start with the subject |
| Continuity filler | «تتضافر الجهود» | what was done and its percentage |
| The challenge formula | «رغم … إلا أنه يواجه تحديات» then a hopeful line | the obstacle alone |
| Denying what no one said | «ليس مجرد … بل …» | what it is |
| Closing summaries | «في الختام»، «خلاصة القول»، «باختصار» | nothing; the closing line is fixed |
| Vague sources | «يرى المختصون»، «يشير المراقبون» | the entity by name, or leave it out |
| Dashes | the long and the short dash | a comma, a full stop or brackets |

Do not line up three adjectives or three short phrases for effect, and do not use a new name for the same thing to avoid repeating a word; repeat the word. Well-known names are written in Arabic letters («نسك»، «كونتنت سكوير»); a web address or a name with no Arabic form stays as it is. «رحلة الحاج» and «رحلة المستفيد» are the ministry's own terms and are fine.

## Message 1: short

```text
معالي الوزير حفظه الله
- <المهام>
- <المعاملات>
- <مشاريع سهيل>
- <مؤشر نضج التجربة الرقمية>        (digital model only)
مع التحية والتقدير
```

One line per part, nothing else. The script applies these rules:

- **المهام**: no delayed or suspended task → «جميع المهام المسندة من <مصادر المهام> على المخطط.», the sources taken from the report's «تفاصيل المهام (…)» titles; otherwise «المهام المسندة من <مصادر المهام> على المخطط، عدا <العدد> متأخرة/معلقة.»
- **المعاملات**: «لا توجد معاملات متأخرة.», or «توجد <العدد> معاملات متأخرة.»
- **مشاريع سهيل**: «مشاريع سهيل على المخطط.», or «مشاريع سهيل على المخطط، عدا <العدد> متأخرة/لم تبدأ.»
- **مؤشر نضج التجربة الرقمية**: «مؤشر نضج التجربة الرقمية على المخطط.», and when a sector's notes report «تأخر»: «مؤشر نضج التجربة الرقمية على المخطط، مع تأخر لدى <القطاع>.» The message states the delay as a fact and stops there. A judgement about escalation or about who is handling it is added only if the user supplies it; nothing is invented and no placeholder is left in the text. The digital part is always one line.

## Message 2: long

```text
معالي الوزير حفظه الله

السلام عليكم ورحمة الله وبركاته،

أرفع لمعاليكم ملخص تقرير الإنجاز الأسبوعي المحدث بتاريخ <تاريخ الغلاف>م، ويشمل المهام ومشاريع سهيل والمعاملات ومؤشر نضج التجربة الرقمية.

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
بلغت نسب الإنجاز في مناظير المؤشر الأربعة: <المناظير الأربعة بنسبها>.
تقيس هذه النسب التقدم في استكمال المتطلبات والأدلة، وتشمل الجاهزية وتحديد تواريخ التنفيذ والأعمال المنجزة، وليست نتيجة التقييم النهائي.
1. منظور <المنظور> (نسبة الإنجاز <النسبة>)
- بلغت نسبة الإنجاز لدى <الجهة> <النسبة>، ولدى …
- <المستجدات>
- أبرز التحديات: <التحديات>
(لكل منظور من الأربعة، بترتيب البطاقات)

خامسا: طلبات الدعم المرفوعة لمعاليكم
<الطلبات، أو جملة عدم وجودها>

وتفضلوا بقبول خالص التحية والتقدير.
```

Blank lines separate every paragraph, heading and block, as in the draft. For the `standard` model the intro lists «المهام ومشاريع سهيل والمعاملات» only, the digital section is left out and support requests become «رابعا»; the script numbers the sections.

What the agent writes in the `[…]` parts:

- **Completed, delayed and suspended tasks**: the script copies each row of the report's table; correct only grammar, keep the meaning and every number and date. Prefer the direct verb to «تم + مصدر» («شاركت الوكالة…»، «اكتمل الربط»); when the source sentence is quoted as written and uses «تم», keep it and fix only its agreement («تم مشاركة … لمركز» becomes «تمت مشاركة … مع مركز»).
- **Suhail «أبرز المستجدات» and «التحديات»**: one bullet per row of the report's tables, in order. Condense the update into one or two formal sentences; keep its dates, counts and targets, drop the «تاريخ التحديث» line.
- **Each digital perspective**: the script writes the sector percentages from the sector chart, highest first; a sector shown with a note (gray bar) follows on the first perspective's line as «، أما <القطاع> ف<ملاحظته بصيغة رسمية>». Then write «المستجدات» and «أبرز التحديات» from the facts file: the summary «أبرز التحديثات» rows of that perspective, the sector updates and challenges that concern it, and the most telling requirement rows of that perspective (`sectors_detail.<sector>.rows.<perspective>`: task, achievement, basis). Progress, completions and expected dates go under «المستجدات»; a pending decision, an obstacle, a returned or blocked file, a dependency or a requirement still under study goes under «أبرز التحديات». Write «لم يتضمن التقرير تحديات لهذا المنظور.» only when none of these sources has one. A note that spans several perspectives may be used under each it concerns; do not repeat the same sentence twice in one perspective.
- **Support requests**: one bullet per row of the report's «طلبات الدعم» table. When there is none, keep the sentence the script wrote; it may be followed by « والمطلوب لاستكمال العمل: …» listing only decisions or actions that the report's challenges explicitly call for (e.g. a pending decision, a missing response, a deadline). Leave it out when the report calls for none.

### Empty sections

A section with no data is never dropped; the script writes a formal sentence in its place:

| Case | Sentence |
| --- | --- |
| No completed tasks | لا توجد مهام مكتملة ضمن التقرير. |
| No Suhail updates | لم يتضمن التقرير مستجدات لمشاريع سهيل. |
| No Suhail challenges | لم يتضمن التقرير تحديات لمشاريع سهيل. |
| No sector percentage in a perspective | لم يتضمن التقرير نسب إنجاز للجهات في هذا المنظور. |
| No challenge for a perspective | لم يتضمن التقرير تحديات لهذا المنظور. |
| No support request | لم يتضمن التقرير طلبات دعم مرفوعة لمعاليكم. |
| A zero total | لم يتضمن التقرير مهام. / لم يتضمن التقرير مشاريع لسهيل. / لم يتضمن التقرير معاملات. |

Delayed and suspended task blocks and the delayed-project block appear only when the report has rows for them; their absence is already stated in the counts sentence («ولا توجد مهام متأخرة أو معلقة …»).
