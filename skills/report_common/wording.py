"""Checks for Arabic text the agent writes into a report or a message.

The reports and the Minister messages are formal government Arabic written
the way a ministry officer writes it. This module holds the one list of
colloquial words and the one list of machine-sounding patterns, so the digital
notes, the Minister messages and any later text are held to the same rule.
Standard library only.
"""
import re

TASHKEEL = re.compile('[ً-ْٰـ]')
ARABIC_WORD = re.compile(r'[ء-ي]+')
LATIN_WORD = re.compile(r'[A-Za-z][A-Za-z0-9._-]*')

# Saudi and Gulf colloquial words that have no place in the formal register.
COLLOQUIAL = {'الحين', 'شغالين', 'شغالة', 'شغال', 'وش', 'ليش', 'زين', 'يبي', 'نبي', 'تبي', 'يبون',
              'مافي', 'مافيه', 'اللي', 'هذي', 'عشان', 'علشان', 'بعدين', 'ابشر', 'أبشر'}

# Machine-sounding patterns and what to write instead. «رحلة» is left out on
# purpose: «رحلة الحاج» and «رحلة المستفيد» are the ministry's own terms.
_PREFIX = r'(?<![ء-ي])[وف]?'
PATTERNS = (
    (r'[يت]عكس|مما يعكس|ما يعكس', 'اذكر الحقيقة نفسها بدل «يعكس»'),
    (r'[يت]سلط الضوء|تسليط الضوء', 'اذكر المعلومة مباشرة'),
    (r'[يت]لعب دورا?', 'اذكر ما يفعله فعلا'),
    (r'(?:ال)?محوري[ةا]?|(?:ال)?جوهري[ةا]?|(?:ال)?حيوي[ةا]|دور(?:ا|اً)? حيوي', 'احذف الوصف أو اذكر الأثر بالرقم'),
    # «يعد/تعد» without shadda is also "prepares" («يعد الفريق التقرير»), which cannot be told apart
    # cheaply. So only «يعد/تعد» + a short noun phrase + من/أحد/إحدى/أبرز/أهم/شاهدا/ركيزة/نموذجا is
    # flagged, plus the explicit «يعدّ/تعدّ» with shadda (tashkeel is stripped before matching, so
    # the shadda form is matched on the raw text in issues()).
    (r'[يت]عد(?: \S+){0,3}? (?:من|أحد|إحدى|أبرز|أهم|شاهدا|شاهدًا|ركيزة|نموذجا)(?![ء-ي])|[يت]عتبر', 'قل «هو» أو اذكر الحقيقة'),
    (r'[يت]مثل نقلة|نقلة نوعية', 'اذكر ما تغير'),
    (r'في ظل (?:ال)?مشهد|في ظل (?:ال)?(?:تطورات|متغيرات) (?:ال)?متسارع', 'اذكر السبب مباشرة'),
    (r'مما [يت]سهم|ما [يت]سهم|مما [يت]ؤكد|ما [يت]ؤكد|مما [يت]عزز|بما [يت]عزز|بما [يت]سهم', 'احذف التعقيب الذي لم يرد في التقرير'),
    (r'علاوة على ذلك|بالإضافة إلى ذلك|إضافة إلى ذلك', 'ابدأ الجملة بمضمونها'),
    (r'(?:^|[.!؟:]\s*)بالإضافة إلى', 'استخدم «و» أو «مع»'),
    (r'تجدر الإشارة|الجدير بالذكر|لا يخفى|من المهم الإشارة|يجب التنويه', 'اذكر المعلومة دون تمهيد'),
    (r'في الختام|ختاما|خلاصة القول|باختصار|إجمالا', 'لا تلخص في النهاية'),
    (r'(?:ي|يُ)شار إلى أن|يشار إلى أن', 'اذكر المعلومة دون تمهيد'),
    (r'ليس مجرد|ليست مجرد|لا يقتصر على|ليس فقط .{0,60}بل', 'اذكر ما هو دون نفي ما لم يقله أحد'),
    (r'(?:ال)?رائد[ةا]|رائد (?:في|من)|فريد[ةا]? من نوع|باقة متنوعة|في قلب', 'احذف لغة الإعلان'),
    (r'آفاق|نسيج|[يت]تعمق', 'استخدم لفظا مباشرا'),
    (r'تبرز أهمية|تتجلى', 'قل «يحتاج … إلى» أو «المطلوب»'),
    (r'تتضافر|في إطار (?:ال)?(?:جهود|حرص)|حرصا على', 'اذكر العمل المنجز ونسبته'),
    (r'رغم .{0,60}إلا أن(?:ه|ها) [يت]واجه|رغم (?:ال)?جهود', 'اذكر العائق وحده'),
    (r'[يت]رى (?:ال)?(?:خبراء|مختصون|متخصصون)|[يت]شير (?:ال)?مراقبون', 'اذكر المصدر باسمه'),
    (r'بكل سرور|بالتأكيد!|سؤال رائع|أتمنى أن يكون هذا مفيد[اً]?|لا تتردد|إليك(?![ء-ي])', 'احذف المجاملة وابدأ بالمعلومة'),
    (r'هل تريد[^.!؟\n]*؟\s*$', 'احذف السؤال الختامي'),
    (r'نعتذر عن أي إزعاج', 'احذف الاعتذار العام'),
)
_COMPILED = tuple((re.compile(_PREFIX + '(?:' + pattern + ')'), advice) for pattern, advice in PATTERNS)
SHADDA_AD = re.compile(r'(?<![ء-ي])[يت]ُ?عَ?دّ[ًٌٍَُِْ]*(?![ء-ي])')
MARKUP = re.compile(r'\*\*|^#{1,6} |[—–]|[\U0001F300-\U0001FAFF☀-➿]', re.M)


def plain(text):
    """`text` without tashkeel or tatweel, for matching."""
    return TASHKEEL.sub('', text or '')


def colloquial(text):
    """Colloquial words found in `text`, sorted."""
    return sorted({word for word in ARABIC_WORD.findall(plain(text)) if word in COLLOQUIAL})


def issues(text):
    """Every machine-sounding pattern found in `text` as ``(found, advice)``, in order."""
    found = []
    # «يُعدّ» with shadda is always "is considered"; it is read before tashkeel is stripped.
    for match in SHADDA_AD.finditer(text or ''):
        found.append((plain(match.group()).strip(), 'قل «هو» أو اذكر الحقيقة'))
    for line in plain(text).splitlines():
        for pattern, advice in _COMPILED:
            for match in pattern.finditer(line):
                found.append((match.group().strip(), advice))
    return found


def markup(text):
    """Markdown, dashes and emoji that plain formal text must not carry."""
    return sorted({match.group().strip() or match.group() for match in MARKUP.finditer(text or '')})


def latin_words(text):
    """Latin-letter words in `text`; on a phone they scramble an Arabic line."""
    return LATIN_WORD.findall(text or '')


def review(text, where):
    """Human-readable findings for `text`, each naming `where` it was found."""
    out = [f'{where}: كلمة عامية «{word}»؛ استخدم الفصحى الرسمية.' for word in colloquial(text)]
    out += [f'{where}: صياغة آلية «{found}»؛ {advice}.' for found, advice in issues(text)]
    out += [f'{where}: رمز غير مناسب لنص رسمي «{mark}».' for mark in markup(text)]
    return out
