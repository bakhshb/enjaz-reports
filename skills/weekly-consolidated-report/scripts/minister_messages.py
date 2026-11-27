# /// script
# requires-python = ">=3.12"
# dependencies = ["openpyxl>=3.1", "python-pptx>=1.0.2", "lxml>=5.3", "pdfplumber>=0.11"]
# ///
"""Read the delivered weekly report and write the facts for the two Minister messages.

Every number and record comes from the delivered PPTX, never from the Excel or
PDF inputs, so the messages match what the Minister will see. The output is:

- ``<stem>.message-facts.json``: every fact both messages use.
- ``<stem>.short-message.txt``: the short message, complete except the
  user's own judgement at the end of the digital line.
- ``<stem>.long-message.draft.txt``: the long message with every fixed and
  counted part written; the parts marked ``[…]`` are rewritten by the agent in
  formal Arabic from the facts file (see references/minister-summary.md).

``--check <message.txt>`` reads a finished message back and lists what still
needs the agent's hand: a bracket left in, a number that is not in the facts
file, a colloquial word or a machine-sounding phrase.
"""
import argparse
import json
import re
import sys
from pathlib import Path

from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from report_common import pptx_helpers as helpers, wording
from validate_report import SPECS
from weekly_data import COVER_DATE_ID, task_source
from update_format import update_parts

# Number–noun agreement: (one, two, 3-10, 11+) for each noun and each status.
NOUNS = {
    'task': ('مهمة واحدة', 'مهمتان', 'مهام', 'مهمة'),
    'project': ('مشروع واحد', 'مشروعان', 'مشاريع', 'مشروعا'),
    'transaction': ('معاملة واحدة', 'معاملتان', 'معاملات', 'معاملة'),
}
STATUS = {
    ('task', 'مكتملة'): ('مكتملة', 'مكتملتان', 'مكتملة', 'مكتملة'),
    ('task', 'متأخرة'): ('متأخرة', 'متأخرتان', 'متأخرة', 'متأخرة'),
    ('task', 'معلقة'): ('معلقة', 'معلقتان', 'معلقة', 'معلقة'),
    ('project', 'مكتملة'): ('مكتمل', 'مكتملان', 'مكتملة', 'مكتملا'),
    ('project', 'متأخرة'): ('متأخر', 'متأخران', 'متأخرة', 'متأخرا'),
    ('project', 'لم تبدأ'): ('لم يبدأ', 'لم يبدآ', 'لم تبدأ', 'لم يبدأ'),
    ('transaction', 'مكتملة'): ('مكتملة', 'مكتملتان', 'مكتملة', 'مكتملة'),
    ('transaction', 'متأخرة'): ('متأخرة', 'متأخرتان', 'متأخرة', 'متأخرة'),
}
PLURALS = {'task': 'مهام', 'project': 'مشاريع', 'transaction': 'معاملات'}
LABELS = {'task': 'المهام', 'project': 'مشاريع سهيل', 'transaction': 'المعاملات'}
# Statuses that close the counts sentence with «ولا توجد …» when they are zero.
NEGATIVE = {'task': ('متأخرة', 'معلقة'), 'project': ('متأخرة', 'لم تبدأ'), 'transaction': ('متأخرة',)}
NEGATIVE_TAIL = {'task': ' ضمن جدول متابعة المهام', 'project': '', 'transaction': ''}
DISCLAIMER = ('تقيس هذه النسب التقدم في استكمال المتطلبات والأدلة، وتشمل الجاهزية وتحديد تواريخ التنفيذ '
              'والأعمال المنجزة، وليست نتيجة التقييم النهائي.')
EMPTY = {
    'completed_tasks': 'لا توجد مهام مكتملة ضمن التقرير.',
    'suhail_updates': 'لم يتضمن التقرير مستجدات لمشاريع سهيل.',
    'suhail_delayed': 'لا توجد مشاريع متأخرة.',
    'suhail_challenges': 'لم يتضمن التقرير تحديات لمشاريع سهيل.',
    'support': 'لم يتضمن التقرير طلبات دعم مرفوعة لمعاليكم.',
    'perspective_sectors': 'لم يتضمن التقرير نسب إنجاز للجهات في هذا المنظور.',
    'perspective_challenges': 'لم يتضمن التقرير تحديات لهذا المنظور.',
}
ORDINALS = ('أولا', 'ثانيا', 'ثالثا', 'رابعا', 'خامسا')


def _form(count):
    if count == 1:
        return 0
    if count == 2:
        return 1
    return 2 if 3 <= count % 100 <= 10 else 3


def counted(count, kind, status=None):
    """«مهمتان مكتملتان»، «15 مهمة على المخطط»، «7 مشاريع مكتملة» with the right agreement."""
    form = _form(count)
    noun = NOUNS[kind][form]
    text = noun if form < 2 else f'{count} {noun}'
    if status == 'على المخطط':
        return f'{text} على المخطط'
    if status:
        return f'{text} {STATUS[(kind, status)][form]}'
    return text


def counts_sentence(kind, total, statuses):
    """The opening sentence of a section: total, every non-zero status, then the zero problem statuses."""
    label = LABELS[kind]
    if not total:
        return f'لم يتضمن التقرير {PLURALS[kind]}.' if kind != 'project' else 'لم يتضمن التقرير مشاريع لسهيل.'
    parts = [counted(value, kind, status) for status, value in statuses if value]
    sentence = f'بلغ إجمالي {label} {counted(total, kind)}'
    if parts:
        sentence += '، منها ' + ' و'.join(parts)
    zero = [s for s, v in statuses if s in NEGATIVE[kind] and not v]
    if zero:
        sentence += f'، ولا توجد {PLURALS[kind]} ' + ' أو '.join(zero) + NEGATIVE_TAIL[kind]
    return sentence + '.'


def sentence(text):
    """End a sentence with a full stop unless it already ends with punctuation."""
    text = text.strip()
    return text if not text or text[-1] in '.!؟:' else text + '.'


def _kpis(prs, label, ids):
    slide = helpers.find_slide_with_text(prs, label)
    return [int(helpers.shape_by_id(slide, sid).text.strip() or 0) for sid in ids]


def _rows(shape):
    """Table records keyed by header, in display order, whatever the column storage order."""
    table = shape.table
    headers = [c.text.strip() for c in table.rows[1].cells]
    return [{h: c.text.strip() for h, c in zip(headers, row.cells)} for row in list(table.rows)[2:]]


def _summary_tables(prs, titles, header):
    out = {title: [] for title in titles}
    for slide in prs.slides:
        for shape in helpers.tables(slide):
            title = helpers.table_title(shape)
            if title in out and header in [c.text.strip() for c in shape.table.rows[1].cells]:
                out[title].extend(_rows(shape))
    return out


def _flat(text):
    """One paragraph: drop the leading «تاريخ التحديث …» label and date, and join the update items."""
    lines = [value[date_end:].strip() for value, date_end, _ in update_parts(text)]
    return '، '.join(line.rstrip('.،') for line in lines if line)


def report_date(prs):
    return helpers.shape_by_id(prs.slides[0], COVER_DATE_ID).text.strip()


def facts(path):
    prs = Presentation(path)
    labels = [spec[0] for spec in SPECS]
    tk = _kpis(prs, labels[0], SPECS[0][4])
    pk = _kpis(prs, labels[1], SPECS[1][4])
    xk = _kpis(prs, labels[2], SPECS[2][4])
    tasks = _summary_tables(prs, helpers.TASK_SUMMARY_TITLES, 'المهمة')
    suhail = _summary_tables(prs, helpers.SUHAIL_SUMMARY_TITLES, 'المشروع')
    sources = []
    for slide in prs.slides:
        for shape in helpers.detail_tables(slide, 'task'):
            source = task_source(helpers.table_title(shape))
            if source is not None and source not in sources:
                sources.append(source)
    result = {
        'date': report_date(prs),
        'task_sources': sources,
        'tasks': {'total': tk[0], 'completed': tk[1], 'on_plan': tk[2], 'delayed': tk[3], 'suspended': tk[4],
                  'tables': tasks},
        'suhail': {'total': pk[0], 'completed': pk[1], 'on_plan': pk[2], 'delayed': pk[3], 'not_started': pk[4],
                   'tables': suhail},
        'transactions': {'total': xk[0], 'completed': xk[1], 'on_plan': xk[2], 'delayed': xk[3]},
        'digital': None,
    }
    try:
        import digital_maturity as digital
        slide = digital.digital_slide(prs)
    except (ValueError, ImportError):
        return result
    metrics = {axis: helpers.shape_by_id(slide, ids[1]).text.strip() for axis, ids in digital.CARDS.items()}
    chart = helpers.shape_by_id(slide, digital.SECTOR_CHART_ID).chart
    names = [str(c) for c in chart.plots[0].categories]
    series = {str(s.name): list(s.values) for s in chart.plots[0].series}
    notes = digital.note_boxes(slide)
    sectors = []
    for index, name in enumerate(names):
        values = {axis: series[label][index] for axis, label in zip(digital.AXES, digital.LABELS)}
        sectors.append({'name': name, 'note': notes.get(digital.NOTE_BOX.format(index + 1)),
                        'metrics': {a: (round(v * 100) if v is not None else None) for a, v in values.items()}})
    summary = digital.summary_table_shape(slide)
    details, sector_rows = {}, {}
    for detail in map(digital._DetailSlide, digital.detail_slides(prs)):
        entry = details.setdefault(detail.base, {'headings': [], 'updates': [], 'challenges': [], 'rows': {}})
        rows = sector_rows.setdefault(detail.base, {'cards': {axis: detail.card(axis) for axis in digital.AXES},
                                                    'rows': []})['rows']
        for row in detail.rows:
            if digital.is_continuation(row):
                continue
            rows.append(row)
            if digital.is_heading(row):
                entry['headings'].append(row[0])
        for table in detail.notes:
            entry[digital.NOTE_KEYS[table.cell(0, 0).text.strip()]].extend(
                digital._plain(r.cells[1].text).strip() for r in list(table.rows)[2:])
    # Requirement rows by perspective, with the same perspective the slides give each row.
    for title, sector in sector_rows.items():
        for index, axis in digital.row_axes(sector).items():
            row = sector['rows'][index]
            if axis and not digital.is_heading(row):
                basis, achievement, task, _code = row
                details[title]['rows'].setdefault(axis, []).append(
                    {'task': task, 'achievement': achievement, 'basis': basis})
    result['digital'] = {
        'metrics': metrics, 'labels': dict(zip(digital.AXES, digital.LABELS)), 'sectors': sectors,
        'summary_updates': [{'perspective': r['المنظور'], 'text': r['التحديث']} for r in _rows(summary)],
        'sectors_detail': details,
    }
    return result


def delayed_sectors(data):
    """Sectors whose digital notes report a delay («تأخر»), in report order."""
    digital = data['digital']
    if not digital:
        return []
    found = [name for name, entry in digital['sectors_detail'].items()
             if any('تأخر' in t for t in entry['challenges'] + entry['updates'])]
    for update in digital['summary_updates']:
        for sector in digital['sectors']:
            if 'تأخر' in update['text'] and sector['name'] in update['text'] and sector['name'] not in found:
                found.append(sector['name'])
    return found


def short_message(data):
    t, p, x = data['tasks'], data['suhail'], data['transactions']
    sources = ' و'.join(data['task_sources']) or 'الجهات المسندة'
    if not (t['delayed'] or t['suspended']):
        tasks = f'جميع المهام المسندة من {sources} على المخطط.'
    else:
        late = [counted(v, 'task', s) for s, v in (('متأخرة', t['delayed']), ('معلقة', t['suspended'])) if v]
        tasks = f'المهام المسندة من {sources} على المخطط، عدا ' + ' و'.join(late) + '.'
    transactions = ('لا توجد معاملات متأخرة.' if not x['delayed']
                    else f'توجد {counted(x["delayed"], "transaction", "متأخرة")}.')
    if not (p['delayed'] or p['not_started']):
        suhail = 'مشاريع سهيل على المخطط.'
    else:
        late = [counted(v, 'project', s) for s, v in (('متأخرة', p['delayed']), ('لم تبدأ', p['not_started'])) if v]
        suhail = 'مشاريع سهيل على المخطط، عدا ' + ' و'.join(late) + '.'
    lines = ['معالي الوزير حفظه الله', f'- {tasks}', f'- {transactions}', f'- {suhail}']
    if data['digital']:
        digital = 'مؤشر نضج التجربة الرقمية على المخطط'
        late = delayed_sectors(data)
        if late:
            # Only the fact from the facts file; any judgement about it is the user's to add.
            digital += '، مع تأخر لدى ' + ' و'.join(late)
        lines.append(f'- {digital}.')
    lines.append('مع التحية والتقدير')
    return '\n'.join(lines) + '\n'


def _percent_sentence(entries):
    entries = sorted(entries, key=lambda item: -item[1])
    if not entries:
        return EMPTY['perspective_sectors']
    first, *rest = entries
    text = f'بلغت نسبة الإنجاز لدى {first[0]} {first[1]}%'
    for name, value in rest:
        text += f'، ولدى {name} {value}%'
    return text


def long_draft(data):
    t, p, x, digital = data['tasks'], data['suhail'], data['transactions'], data['digital']
    parts = 'المهام ومشاريع سهيل والمعاملات' + (' ومؤشر نضج التجربة الرقمية' if digital else '')
    out = ['معالي الوزير حفظه الله', '', 'السلام عليكم ورحمة الله وبركاته،', '',
           f'أرفع لمعاليكم ملخص تقرير الإنجاز الأسبوعي المحدث بتاريخ {data["date"]}م، ويشمل {parts}.', '']
    sections = iter(ORDINALS)
    out += [f'{next(sections)}: المهام', '',
            counts_sentence('task', t['total'], [('مكتملة', t['completed']), ('على المخطط', t['on_plan']),
                                                 ('متأخرة', t['delayed']), ('معلقة', t['suspended'])]), '']
    for title, label in (('المهام المكتملة', 'المهام المكتملة'), ('المهام المتأخرة', 'المهام المتأخرة'),
                         ('المهام المعلقة', 'المهام المعلقة')):
        rows = t['tables'][title]
        if not rows and title != 'المهام المكتملة':
            continue
        out += [f'{label}:', '']
        if not rows:
            out += [EMPTY['completed_tasks'], '']
        for number, row in enumerate(rows, 1):
            out += [f'{number}. {row["القطاع"]}', '', f'المهمة: {sentence(row["المهمة"])}', '',
                    f'المستجد: {sentence(row.get("ملاحظات", row.get("التحديث", ""))) or "لم يتضمن التقرير مستجدا لهذه المهمة."}', '']
    out += [f'{next(sections)}: مشاريع سهيل', '',
            counts_sentence('project', p['total'], [('مكتملة', p['completed']), ('على المخطط', p['on_plan']),
                                                    ('متأخرة', p['delayed']), ('لم تبدأ', p['not_started'])]), '']
    for title, label, header, empty in (('أبرز التحديثات', 'أبرز المستجدات', 'التحديث', 'suhail_updates'),
                                        ('المشاريع المتأخرة', 'المشاريع المتأخرة', 'ملاحظات', 'suhail_delayed'),
                                        ('التحديات', 'التحديات', 'التحدي', 'suhail_challenges')):
        rows = p['tables'][title]
        if not rows and title == 'المشاريع المتأخرة':
            continue
        out += [f'{label}:', '']
        out += [f'- {row["المشروع"]}: [{_flat(row[header])}]' for row in rows] or [EMPTY[empty]]
        out.append('')
    out += [f'{next(sections)}: المعاملات', '',
            counts_sentence('transaction', x['total'], [('مكتملة', x['completed']), ('على المخطط', x['on_plan']),
                                                        ('متأخرة', x['delayed'])]), '']
    if digital:
        labels, metrics = digital['labels'], digital['metrics']
        listing = '، '.join(f'{"و" if i else ""}{labels[a]} {metrics[a]}' for i, a in enumerate(labels))
        out += [f'{next(sections)}: مؤشر نضج التجربة الرقمية', '',
                'بلغت نسب الإنجاز في مناظير المؤشر الأربعة: ' + listing + '.', '', DISCLAIMER, '']
        noted = [s for s in digital['sectors'] if s['note']]
        for number, (axis, label) in enumerate(labels.items(), 1):
            entries = [(s['name'], s['metrics'][axis]) for s in digital['sectors'] if s['metrics'][axis] is not None]
            line = _percent_sentence(entries)
            if number == 1 and noted:
                line += '، أما [' + '؛ '.join(f'{s["name"]}: {s["note"]}' for s in noted) + ']'
            out += [f'{number}. منظور {label} (نسبة الإنجاز {metrics[axis]})', '', f'- {line}.',
                    '- [المستجدات]', '- أبرز التحديات: [التحديات]', '']
    out += [f'{next(sections)}: طلبات الدعم المرفوعة لمعاليكم', '']
    support = t['tables']['طلبات الدعم']
    if support:
        out += [f'- {r["المهمة"]} ({r["القطاع"]}): {r["طلب الدعم"]}' for r in support] + ['']
    else:
        out += [EMPTY['support'] + '[ والمطلوب لاستكمال العمل: …]', '']
    out.append('وتفضلوا بقبول خالص التحية والتقدير.')
    return '\n'.join(out) + '\n'


NUMBER = re.compile(r'\d+(?:[.,]\d+)?')
ARABIC_DIGITS = str.maketrans('٠١٢٣٤٥٦٧٨٩', '0123456789')
LIST_NUMBER = re.compile(r'^\s*\d+[.)]\s')


def check_message(text, data):
    """What a finished message still needs, as a list of Arabic findings (empty when it is ready).

    A bracket or «…» left from the draft, a number that is not in the facts
    file, a colloquial word, a machine-sounding phrase, or markup. Latin
    words are listed as a reminder to write well-known names in Arabic letters.
    """
    findings = []
    known = set(NUMBER.findall(json.dumps(data, ensure_ascii=False).translate(ARABIC_DIGITS)))
    for number, line in enumerate(text.splitlines(), 1):
        where = f'السطر {number}'
        if '[' in line or ']' in line or '…' in line:
            findings.append(f'{where}: بقي فيه جزء من المسودة بين قوسين أو «…»؛ اكتبه من ملف الحقائق أو احذفه.')
        body = LIST_NUMBER.sub('', line.translate(ARABIC_DIGITS))
        for value in NUMBER.findall(body):
            if value not in known and value.replace(',', '.') not in known:
                findings.append(f'{where}: الرقم {value} غير موجود في ملف الحقائق؛ لا تضف رقما ليس في التقرير.')
        findings += wording.review(line, where)
        latin = wording.latin_words(line)
        if latin:
            findings.append(f'{where}: كلمات بحروف لاتينية ({"، ".join(latin)})؛ اكتب الاسم المعروف بحروف عربية إن وجد.')
    return findings


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--report', required=True, help='The delivered weekly report (PPTX)')
    parser.add_argument('--output-dir', help='Defaults to the report folder')
    parser.add_argument('--check', action='append', default=[], metavar='MESSAGE.txt',
                        help='Check a finished message against the report instead of writing the drafts')
    args = parser.parse_args(argv)
    report = Path(args.report)
    data = facts(report)
    if args.check:
        total = 0
        for path in args.check:
            findings = check_message(Path(path).read_text(encoding='utf-8-sig'), data)
            total += len(findings)
            print(f'{path}: ' + ('جاهزة' if not findings else f'{len(findings)} ملاحظة'))
            for finding in findings:
                print('  ' + finding)
        return 1 if total else 0
    folder = Path(args.output_dir) if args.output_dir else report.parent
    folder.mkdir(parents=True, exist_ok=True)
    stem = report.stem
    (folder / f'{stem}.message-facts.json').write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding='utf-8')
    (folder / f'{stem}.short-message.txt').write_text(short_message(data), encoding='utf-8')
    (folder / f'{stem}.long-message.draft.txt').write_text(long_draft(data), encoding='utf-8')
    print(folder / f'{stem}.long-message.draft.txt')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
