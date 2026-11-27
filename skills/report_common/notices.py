"""Warn and keep going: the one place every skill records a best assumption.

A skill stops only when it cannot produce the report at all (a missing file,
sheet or required column, a damaged template). Anything else is recorded here,
the skill continues with the stated assumption, and the agent repeats every
notice in its final reply. Standard library only, so the Excel skills can
import it without extra dependencies.
"""
import json
import re
import sys
from pathlib import Path

from .dates import DateError, date_parts, report_date

_NOTICES = []


def warn(message, kind='assumption'):
    """Record one notice and echo it on stderr. Repeated messages are kept once."""
    entry = {'kind': kind, 'message': str(message)}
    if entry not in _NOTICES:
        _NOTICES.append(entry)
        print('تنبيه: ' + entry['message'], file=sys.stderr)
    return entry


def collected():
    return list(_NOTICES)


def reset():
    _NOTICES.clear()


def notices_path(output):
    """`<output>.notices.json`, next to the file the skill delivers."""
    output = Path(output)
    return output.with_name(output.name + '.notices.json')


def save(output):
    """Write the notices next to `output`; with none, remove a file left by an earlier run."""
    path = notices_path(output)
    if _NOTICES:
        path.write_text(json.dumps(_NOTICES, ensure_ascii=False, indent=1), encoding='utf-8')
        print(f'{len(_NOTICES)} تنبيه محفوظ في: {path}', file=sys.stderr)
    elif path.exists():
        path.unlink()
    return path


def clear_stale(output):
    """Delete a `<output>.notices.json` left by an earlier run, so a failed build leaves none beside it."""
    path = notices_path(output)
    if path.exists():
        path.unlink()
    return path


def tolerant_date(value, where, *, hijri=False):
    """A date as dd-mm-yyyy; a value that cannot be read is kept as written, with a notice."""
    try:
        return report_date(value, hijri=hijri)
    except DateError:
        text = re.sub(r'\s+', ' ', str(value)).strip()
        warn(f'تاريخ غير مقروء في {where}: «{text}». أبقيته كما ورد في المصدر.', 'date')
        return text


def tolerant_date_parts(value, where, *, hijri=False):
    """(text, sort key) for a date column; an unreadable value sorts last and is kept as written."""
    text = tolerant_date(value, where, hijri=hijri)
    try:
        parts = date_parts(value, hijri=hijri)
    except DateError:
        parts = None
    return text, parts or (9999, 99, 99)


# «منتهي/منتهية/منته» whole word means finished; «منتهي المدة» and «منتهية الصلاحية» mean expired, not completed.
_ENDED = re.compile(r'(?<![ء-ي])منته(?:ي|ية)?(?![ء-ي])(?!\s+(?:ال)?(?:مدة|صلاحية))')

# Stems that place a status the maps do not list; the first stem found wins.
STATUS_STEMS = (
    ('لم يبدأ', 'لم تبدأ'), ('لم تبدأ', 'لم تبدأ'), ('لم يتم البدء', 'لم تبدأ'),
    ('مكتمل', 'مكتملة'), ('منجز', 'مكتملة'), (_ENDED, 'مكتملة'),
    ('متأخر', 'متأخر'), ('متاخر', 'متأخر'), ('متعثر', 'متأخر'),
    ('معلق', 'معلق'), ('متوقف', 'معلق'), ('موقوف', 'معلق'),
    ('مخطط', 'على المخطط'), ('جار', 'على المخطط'),
)
NEGATIONS = ('غير', 'لم', 'لا', 'ليس', 'بدون')
_NEGATION = re.compile(r'(?<![ء-ي])(?:' + '|'.join(NEGATIONS) + r')(?![ء-ي])')


def _stem_at_word_start(stem, text):
    """The index of `stem` (text, or a compiled pattern for a whole word) at the start of a word in `text`, or -1."""
    if isinstance(stem, re.Pattern):
        match = stem.search(text)
        return match.start() if match else -1
    match = re.search(r'(?<![ء-ي])' + re.escape(stem), text)
    return match.start() if match else -1


def guess_status(raw, known, default, where):
    """The category for a status text the skill's own map does not list.

    `known` maps a canonical category (مكتملة، على المخطط، متأخر، معلق، لم تبدأ)
    to the label this skill uses for it. A stem at a word start picks the
    category; a negation before the stem («غير مكتمل»، «لم ينجز») means no
    guess, and `default` is used. Either way the choice is recorded.
    """
    text = re.sub(r'\s+', ' ', str(raw)).strip()
    negated = False
    for stem, category in STATUS_STEMS:
        at = _stem_at_word_start(stem, text)
        if at < 0 or category not in known:
            continue
        if _NEGATION.search(text[:at]) and not (isinstance(stem, str) and stem.startswith('لم ')):
            negated = True
            continue
        warn(f'حالة غير معروفة في {where}: «{text}». اعتبرتها «{known[category]}» لتشابه اللفظ.', 'status')
        return known[category]
    if negated or _NEGATION.search(text):
        warn(f'حالة غير معروفة في {where}: «{text}». فيها نفي فلم أخمّن معناها، واعتبرتها «{default}» افتراضا، وتحتاج مراجعة.', 'status')
        return default
    warn(f'حالة غير معروفة في {where}: «{text}». اعتبرتها «{default}» افتراضا، وتحتاج مراجعة.', 'status')
    return default
