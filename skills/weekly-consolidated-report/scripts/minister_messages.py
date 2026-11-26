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
"""
import argparse
import json
import re
import sys
from pathlib import Path

from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from report_common import pptx_helpers as helpers
from validate_report import SPECS

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
NEGATIVE_WORDS = {'متأخرة': 'متأخرة', 'معلقة': 'معلقة', 'لم تبدأ': 'لم تبدأ'}
MONTHS = helpers.MONTHS
DISCLAIMER = ('وتعكس النسب أدناه تقدم استكمال المتطلبات والأدلة، بما يشمل الجاهزية وتحديد تواريخ التنفيذ '
              'والأعمال المنجزة، ولا تمثل نتيجة التقييم النهائي.')
EMPTY = {
    'completed_tasks': 'لا توجد مهام مكتملة ضمن التقرير.',
    'suhail_updates': 'لم يتضمن التقرير مستجدات لمشاريع سهيل.',
    'suhail_delayed': 'لا توجد مشاريع متأخرة.',
    'suhail_challenges': 'لم يتضمن التقرير تحديات لمشاريع سهيل.',
    'support': 'لم يتضمن التقرير طلبات دعم مستقلة مرفوعة لمعاليكم.',
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
    zero = [NEGATIVE_WORDS[s] for s, v in statuses if s in NEGATIVE[kind] and not v]
    if zero:
        sentence += f'، ولا توجد {PLURALS[kind]} ' + ' أو '.join(zero) + NEGATIVE_TAIL[kind]
    return sentence + '.'


def sentence(text):
    """End a sentence with a full stop unless it already ends with punctuation."""
    text = text.strip()
    return text if not text or text[-1] in '.!؟:' else text + '.'


def _texts(slide):
    return helpers.slide_texts(slide)


def _kpis(prs, label, ids):
    slide = next(s for s in prs.slides if label in _texts(s))
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
    """One paragraph: drop a leading «تاريخ التحديث …» line and join bullet lines."""
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    if lines and lines[0].startswith('تاريخ التحديث'):
        lines = lines[1:]
    return '، '.join(l.rstrip('.،') for l in lines)


def report_date(prs):
    return helpers.shape_by_id(prs.slides[0], 19).text.strip()


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
            found = re.fullmatch(r'تفاصيل المهام \((.*)\)', helpers.table_title(shape))
            if found and found.group(1) not in sources:
                sources.append(found.group(1))
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
    chart = helpers.shape_by_id(slide, 93).chart
    names = [str(c) for c in chart.plots[0].categories]
    series = {str(s.name): list(s.values) for s in chart.plots[0].series}
    notes = {s.name: s.text_frame.text.strip() for s in slide.shapes if s.name.startswith('Digital sector note ')}
    sectors = []
    for index, name in enumerate(names):
        values = {axis: series[label][index] for axis, label in zip(digital.AXES, digital.LABELS)}
        sectors.append({'name': name, 'note': notes.get(f'Digital sector note {index + 1}'),
                        'metrics': {a: (round(v * 100) if v is not None else None) for a, v in values.items()}})
    summary = next(sh for sh in slide.shapes if sh.has_table and len(sh.table.columns) == 3)
    details, rows = {}, {}
    for detail in digital.detail_slides(prs):
        title = digital.split_title(digital._detail_table(detail).cell(0, 0).text)[0]
        entry = details.setdefault(title, {'headings': [], 'updates': [], 'challenges': [], 'rows': {}})
        cards = {axis: helpers.shape_by_id(detail, sid).text_frame.paragraphs[1].text
                 for axis, sid in digital.DETAIL_CARD_IDS.items()}
        rows.setdefault(title, {'cards': cards, 'rows': []})
        for row in digital._slide_rows(digital._detail_table(detail)):
            if not digital.is_continuation(row):
                rows[title]['rows'].append(row)
            if digital.is_heading(row) and not digital.is_continuation(row):
                entry['headings'].append(row[0])
        for table in digital._note_tables(detail):
            key = 'updates' if table.cell(0, 0).text.strip() == digital.UPDATES_TABLE[0] else 'challenges'
            entry[key].extend(digital._plain(r.cells[1].text).strip() for r in list(table.rows)[2:])
    # Requirement rows by perspective, with the same heading styles the slides use.
    for title, sector in rows.items():
        axes = digital.heading_axes(sector)
        current = None
        for index, row in enumerate(sector['rows']):
            if index in axes:
                current = axes[index]
            elif current:
                basis, achievement, task, _code = row
                details[title]['rows'].setdefault(current, []).append(
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
        tasks = f'كل المهام المسندة من {sources} تسير على المخطط'
    else:
        late = [counted(v, 'task', s) for s, v in (('متأخرة', t['delayed']), ('معلقة', t['suspended'])) if v]
        tasks = f'المهام المسندة من {sources} على المخطط عدا ' + ' و'.join(late)
    transactions = ('لا يوجد اي تأخير في المعاملات' if not x['delayed']
                    else f'يوجد {counted(x["delayed"], "transaction", "متأخرة")}')
    if not (p['delayed'] or p['not_started']):
        suhail = 'مبادرات سهيل على المخطط'
    else:
        late = [counted(v, 'project', s) for s, v in (('متأخرة', p['delayed']), ('لم تبدأ', p['not_started'])) if v]
        suhail = 'مبادرات سهيل على المخطط عدا ' + ' و'.join(late)
    lines = ['معاليك', f'- {tasks}', f'- {transactions}', f'- {suhail}']
    if data['digital']:
        digital = 'مؤشر نضج التجربة الرقمية على المخطط'
        late = delayed_sectors(data)
        if late:
            digital += '، فيه تأخير من ' + ' و'.join(late)
            digital += '، ولكن لا يستدعي التصعيد الآن، شغالين مع … على تجاوز التأخير.'
        lines.append(f'- {digital}')
    lines.append('تحياتي')
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
    parts = 'المهام ومشاريع سهيل والمعاملات'
    intro = (f'تجدون أدناه ملخص حالة {parts}، بالإضافة إلى مستجدات مؤشر نضج التجربة الرقمية، '
             if digital else f'تجدون أدناه ملخص حالة {parts}، ')
    out = ['معالي الوزير حفظه الله', '', 'السلام عليكم ورحمة الله وبركاته،', '',
           intro + f'وفق تقرير الإنجاز الأسبوعي المحدث بتاريخ {data["date"]}م.', '']
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
                'تتواصل أعمال رفع مؤشر نضج التجربة الرقمية عبر أربعة مناظير رئيسية، حيث بلغت نسبة الإنجاز في '
                + listing + '.', '', DISCLAIMER, '']
        noted = [s for s in digital['sectors'] if s['note']]
        for number, (axis, label) in enumerate(labels.items(), 1):
            entries = [(s['name'], s['metrics'][axis]) for s in digital['sectors'] if s['metrics'][axis] is not None]
            line = _percent_sentence(entries)
            if number == 1 and noted:
                line += '، فيما [' + '؛ '.join(f'{s["name"]}: {s["note"]}' for s in noted) + ']'
            out += [f'{number}. منظور {label} – نسبة الإنجاز {metrics[axis]}', '', f'- {line}.',
                    '- [المستجدات]', '- أبرز التحديات: [التحديات]', '']
    out += [f'{next(sections)}: طلبات الدعم المرفوعة لمعاليكم', '']
    support = t['tables']['طلبات الدعم']
    if support:
        out += [f'- {r["المهمة"]} ({r["القطاع"]}): {r["طلب الدعم"]}' for r in support] + ['']
    else:
        out += [EMPTY['support'][:-1] + '[، فيما تبرز الحاجة إلى …]' + '.', '']
    out.append('وتفضلوا بقبول خالص التحية والتقدير.')
    return '\n'.join(out) + '\n'


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--report', required=True, help='The delivered weekly report (PPTX)')
    parser.add_argument('--output-dir', help='Defaults to the report folder')
    args = parser.parse_args(argv)
    report = Path(args.report)
    folder = Path(args.output_dir) if args.output_dir else report.parent
    folder.mkdir(parents=True, exist_ok=True)
    data = facts(report)
    stem = report.stem
    (folder / f'{stem}.message-facts.json').write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding='utf-8')
    (folder / f'{stem}.short-message.txt').write_text(short_message(data), encoding='utf-8')
    (folder / f'{stem}.long-message.draft.txt').write_text(long_draft(data), encoding='utf-8')
    print(folder / f'{stem}.long-message.draft.txt')


if __name__ == '__main__':
    main()
