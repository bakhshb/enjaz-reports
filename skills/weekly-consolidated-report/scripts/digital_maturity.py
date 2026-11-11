# /// script
# requires-python = ">=3.12"
# dependencies = ["pdfplumber>=0.11", "python-pptx>=1.0.2", "openpyxl>=3.1", "lxml>=5.3"]
# ///
"""Read the landscape digital-maturity PDF and fill the approved weekly slides.

The PDF is the data source only: page one supplies the summary, the following
pages supply one sector each, where a sector may span several source pages
with a «(k/n)» title suffix.  The source pagination is never reproduced; the
approved portrait master controls the design, and every sector is re-flowed
into as many detail slides as its records need.
"""
import argparse
import hashlib
import json
import re
import sys
import unicodedata
from collections import defaultdict
from copy import deepcopy
from pathlib import Path

import pdfplumber
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from report_common import pptx_helpers as helpers
from weekly_data import content_bottom, set_cell_exact, TABLE_GAP
from weekly_artwork import place_table_copy, is_digital_detail, is_digital_detail_table
from report_common.table_typography import set_table_directions


MASTER = Path(__file__).resolve().parents[1] / 'assets/weekly-report-digital-master.pptx'
MASTER_HASH = '4f44e50242617b4448ff4f9dcde2b29e609a3a37e89132a05ef41ee47aafef98'
CALCULATION_NOTE = ('آلية الاحتساب: 10% للجاهزية وتشمل المتطلبات والاجتماعات والتنسيق، '
                    'و5% لتحديد التواريخ المستهدفة، والنسبة المتبقية وفق الإنجاز الفعلي لمتطلبات الدليل.')
CALCULATION_NOTE_ID = 7
AXES = ('experience', 'satisfaction', 'complaints', 'tools')
LABELS = ('تجربة المستخدم', 'رضا المستفيد', 'التعامل مع الشكاوى', 'التقنيات والأدوات')
DETAIL_CARD_IDS = dict(zip(AXES, (6, 8, 10, 12)))
# Landscape source geometry (points): the four cards sit left-to-right as
# tools, complaints, satisfaction, experience on every page.
PAGE_WIDTH = (700, 740)
PAGE_HEIGHT = (390, 420)
CARD_COLUMNS = {'tools': (15, 185), 'complaints': (190, 355), 'satisfaction': (360, 525), 'experience': (530, 705)}
SUMMARY_CARD_BAND = (30, 130)
DETAIL_CARD_BAND = (45, 85)
SECTOR_BARS = {'x': (360, 570), 'y': (150, 330), 'height': (8, 22)}
SECTOR_LABELS = (565, 705)
COLORS = {
    'experience': (0.898, 0.737, 0.553),
    'satisfaction': (0.220, 0.435, 0.404),
    'complaints': (0.706, 0.482, 0.373),
    'tools': (0.557, 0.729, 0.859),
}
CARDS = {
    'experience': (27, 31, 45, 47),
    'satisfaction': (5, 9, 21, 23),
    'complaints': (73, 77, 89, 91),
    'tools': (51, 55, 67, 69),
}
# Titles the landscape source may give a sector's notes table.
CHALLENGE_TITLES = ('التحديات', 'أبرز التحديثات')
CHALLENGE_HEADERS = ('التحدي', 'التحديث')
# Output tables: the reviewer classifies every source note as an update or a
# challenge; each non-empty list becomes its own table, updates first.
UPDATES_TABLE = ('أبرز التحديثات', 'التحديث')
CHALLENGES_TABLE = ('التحديات', 'التحدي')
NOTE_TABLES = (('updates',) + UPDATES_TABLE, ('challenges',) + CHALLENGES_TABLE)
SUMMARY_TABLE = UPDATES_TABLE
TASK_HEADERS = ('المهمة', 'المؤشر')
CARD_STATUSES = ('معلق', 'لا ينطبق')
CONTINUATION = '(تتمة)'
SUFFIX = re.compile(r'^(.*?)\s*\((\d+)/(\d+)\)\s*$')
CODE = re.compile(r'[A-Z]{1,5}-[0-9A-Z./-]+')
CODE_IN_TEXT = re.compile(r'\b[A-Z]{1,5}-\d{2,3}(?:\.\.\d{2,3})?\b')
MIRRORED = str.maketrans('()[]', ')(][')
PERCENT_FIRST = re.compile(r'%(\d+(?:[.,]\d+)?)')
GAP = TABLE_GAP
# Calibrated against a native PowerPoint render of the approved 11 pt detail cells;
# the shared default (12) roughly doubles the deck length for this model.
CHARS_PER_INCH = 18


BRACKETS = {'(': ')', '[': ']'}
SEPARATORS = set('./:,-')


def _kind(ch):
    if ch.isascii() and ch.isalnum():
        return 'L' if ch.isalpha() else 'EN'
    if unicodedata.bidirectional(ch) in ('R', 'AL', 'AN'):
        return 'R'
    return 'ET' if ch == '%' else 'N'


def _logical(line):
    """Return logical-order text for one visual (left-to-right) glyph line.

    The source stores each line in visual order: Arabic glyphs reversed, Latin
    and digit runs readable left-to-right, brackets as mirrored shapes.  Latin
    and digit runs (with the neutrals enclosed between them) are kept intact,
    everything else is reversed, and brackets outside those runs un-mirrored.
    """
    chars = list(line)
    count = len(chars)
    kinds = [_kind(ch) for ch in chars]
    in_run = [kind in ('L', 'EN') for kind in kinds]
    for index, kind in enumerate(kinds):
        if kind == 'ET' and ((index and kinds[index-1] == 'EN') or (index+1 < count and kinds[index+1] == 'EN')):
            in_run[index] = True
    index, balance = 0, 0
    while index < count:
        if in_run[index]:
            balance += chars[index] in BRACKETS
            balance -= chars[index] in BRACKETS.values() and balance > 0
            index += 1
            continue
        if kinds[index] == 'R':
            balance, index = 0, index + 1
            continue
        end = index
        while end < count and not in_run[end] and kinds[end] != 'R':
            end += 1
        sequence = chars[index:end]
        left = index > 0 and in_run[index-1]
        right = end < count and in_run[end]
        internal = False
        if left and right:
            start = index - 1
            while start > 0 and in_run[start-1]:
                start -= 1
            left_has_letter = any(kinds[i] == 'L' for i in range(start, index))
            single_separator = len(sequence) == 1 and sequence[0] in SEPARATORS and kinds[index-1] == 'EN' and kinds[end] == 'EN'
            internal = single_separator or left_has_letter
        elif left and end == count and all(ch in SEPARATORS for ch in sequence):
            internal = True
        elif left and balance and sequence[0] in BRACKETS.values():
            # A bracket opened inside the run closes inside it, whatever follows.
            closing = index
            while closing < end and chars[closing] in BRACKETS.values() and balance > 0:
                in_run[closing] = True
                balance -= 1
                closing += 1
            index = closing
            continue
        if internal:
            for inner in range(index, end):
                in_run[inner] = True
                balance += chars[inner] in BRACKETS
                balance -= chars[inner] in BRACKETS.values() and balance > 0
        else:
            balance = 0
        index = end
    # An opening bracket glued to a run that closes it belongs to the run.
    index = 0
    while index < count:
        if in_run[index] and (index == 0 or not in_run[index-1]):
            end = index
            while end < count and in_run[end]:
                end += 1
            for opening, closing in BRACKETS.items():
                unmatched = ''.join(chars[index:end]).count(closing) - ''.join(chars[index:end]).count(opening)
                if unmatched > 0 and index > 0 and chars[index-1] == opening:
                    in_run[index-1] = True
            index = end
        else:
            index += 1
    out, index = [], count
    while index > 0:
        if in_run[index-1]:
            start = index
            while start > 0 and in_run[start-1]:
                start -= 1
            out.append(''.join(chars[start:index]))
            index = start
        else:
            out.append(chars[index-1].translate(MIRRORED))
            index -= 1
    text = unicodedata.normalize('NFKC', ''.join(out))
    return PERCENT_FIRST.sub(r'\1%', text).strip()


def _lines(page, box):
    """Visual text lines inside a box, each assembled left-to-right like the page."""
    words = [w for w in page.extract_words() if w['x0'] >= box[0] and w['x1'] <= box[2]
             and w['top'] >= box[1] and w['bottom'] <= box[3]]
    lines = []
    for word in sorted(words, key=lambda w: (w['top'], w['x0'])):
        if lines and abs(lines[-1][0] - word['top']) <= 3:
            lines[-1][1].append(word)
        else:
            lines.append([word['top'], [word]])
    return [((top + max(w['bottom'] for w in group)) / 2,
             ' '.join(w['text'] for w in sorted(group, key=lambda w: w['x0']))) for top, group in lines]


def _rtl(value):
    """Join the visual lines of one PDF field into logical Arabic text."""
    result = ''
    for line in (value or '').splitlines():
        part = _logical(line)
        if not part:
            continue
        if result and result[-1] in '/-' and part[0].isascii() and part[0].isalnum():
            result += part
        else:
            result = (result + ' ' + part).strip()
    return result


def _percent(value):
    match = re.fullmatch(r'\s*(?:%\s*([\d.,]+)|([\d.,]+)\s*%)\s*', value or '')
    if not match:
        raise ValueError('Cannot read a digital-maturity percentage: ' + repr(value))
    number = float((match.group(1) or match.group(2)).replace(',', '.'))
    if not 0 <= number <= 100:
        raise ValueError('Digital-maturity percentage outside 0–100')
    return int(number) if number.is_integer() else number


def _chart_number(value):
    return int(value) if float(value).is_integer() else value


def _check_page(page, number):
    if not PAGE_WIDTH[0] <= page.width <= PAGE_WIDTH[1] or not PAGE_HEIGHT[0] <= page.height <= PAGE_HEIGHT[1]:
        raise ValueError(f'Digital-maturity PDF page {number} is not the approved landscape source layout')


PERCENT_TOKEN = re.compile(r'%\s*[\d.,]+|[\d.,]+\s*%')


def _one_percent(page, box):
    """Read a summary card's headline percentage.

    A card may also carry a smaller weekly-change line (e.g. «+2% عن الأسبوع
    الماضي»), so only percentages set in the card's largest type count.
    """
    region = page.within_bbox(box)
    sizes = defaultdict(list)
    for char in region.chars:
        if char['text'].strip():
            sizes[round(char['size'], 1)].append(char)
    for size in sorted(sizes, reverse=True):
        text = region.filter(lambda obj, size=size: obj.get('object_type') != 'char'
                             or round(obj['size'], 1) == size).extract_text() or ''
        matches = PERCENT_TOKEN.findall(text)
        if matches:
            if len(matches) != 1:
                raise ValueError('Expected one headline percentage in a digital-maturity card')
            return _percent(matches[0])
    raise ValueError('Expected one percentage in a digital-maturity card')


def _axis_from_fill(color):
    if isinstance(color, (int, float)):
        color = (color, color, color)
    if not isinstance(color, (tuple, list)) or len(color) != 3:
        return None
    for axis, expected in COLORS.items():
        if max(abs(float(a) - b) for a, b in zip(color, expected)) < .025:
            return axis
    values = [float(a) for a in color]
    if max(values) - min(values) < .02 and .72 <= values[0] <= .95:
        return 'note'
    return None


def split_sector_note(name):
    """Split «sector (note)» from a landscape label; some exports keep the
    parentheses in visual (mirrored) order, i.e. «sector )note(»."""
    match = re.fullmatch(r'(.*?)\s*\((.+)\)', name) or re.fullmatch(r'(.*?)\s*\)(.+)\(', name)
    if not match:
        raise ValueError('Digital-maturity sector note is missing: ' + name)
    return match.group(1).strip(), match.group(2).strip()


def split_title(title):
    """Separate a source page title from its «(k/n)» pagination suffix."""
    match = SUFFIX.match(title or '')
    if not match:
        return (title or '').strip(), None, None
    return match.group(1).strip(), int(match.group(2)), int(match.group(3))


def heading_base(text):
    return (text or '').strip().removesuffix(CONTINUATION).strip()


def is_heading(row):
    return bool(row[0].strip()) and not any(cell.strip() for cell in row[1:])


def is_continuation(row):
    return is_heading(row) and row[0].endswith(CONTINUATION)


def last_heading(rows):
    return next((row[0] for row in reversed(rows) if is_heading(row)), None)


def _card_box(axis, band):
    x0, x1 = CARD_COLUMNS[axis]
    return (x0, band[0], x1, band[1])


def _card_value(page, axis, band):
    text = _rtl(page.within_bbox(_card_box(axis, band)).extract_text())
    label = LABELS[AXES.index(axis)]
    if not text.startswith(label):
        raise ValueError('Digital-maturity card label is missing: ' + axis)
    value = text[len(label):].strip()
    if not value:
        raise ValueError('Digital-maturity card value is missing: ' + axis)
    return f'{_percent(value):g}%' if '%' in value else value


def _extract_summary(page):
    _check_page(page, 1)
    metrics = {axis: _one_percent(page, _card_box(axis, SUMMARY_CARD_BAND)) for axis in AXES}
    bars = defaultdict(list)
    for rect in page.rects:
        if (SECTOR_BARS['y'][0] <= rect['top'] <= SECTOR_BARS['y'][1]
                and SECTOR_BARS['x'][0] <= rect['x0'] and rect['x1'] <= SECTOR_BARS['x'][1] + 1
                and rect['width'] > 15 and SECTOR_BARS['height'][0] <= rect['height'] <= SECTOR_BARS['height'][1]):
            axis = _axis_from_fill(rect['non_stroking_color'])
            if axis:
                key = next((k for k in bars if abs(k - rect['top']) <= 3), round(rect['top']))
                bars[key].append((axis, rect))
    if not 1 <= len(bars) <= 5:
        raise ValueError('Digital-maturity sector bars are missing or exceed template capacity')
    ordered = sorted(bars.items())
    keys = [key for key, _ in ordered]
    pitch = min([b - a for a, b in zip(keys, keys[1:])] or [22])
    labels = _lines(page, (SECTOR_LABELS[0], SECTOR_BARS['y'][0] - pitch, SECTOR_LABELS[1], SECTOR_BARS['y'][1] + pitch))
    sectors = []
    centres = [(segments[0][1]['top'] + segments[0][1]['bottom']) / 2 for _, segments in ordered]
    for centre, (_, segments) in zip(centres, ordered):
        name = ' '.join(_logical(text) for line, text in labels
                        if abs(line - centre) < pitch and min(centres, key=lambda c: abs(c - line)) == centre)
        if not name:
            raise ValueError('Digital-maturity sector label is missing')
        values = {axis: None for axis in AXES}
        note = None
        for axis, segment in sorted(segments, key=lambda item: item[1]['x0']):
            box = (segment['x0'] - 1, segment['top'] - 1, segment['x1'] + 1, segment['bottom'] + 1)
            label = page.within_bbox(box).extract_text()
            if axis == 'note':
                note = _rtl(label)
            else:
                if values[axis] is not None:
                    raise ValueError('Duplicate digital-maturity axis in one sector')
                values[axis] = _percent(label)
        if any(axis == 'note' for axis, _ in segments) and not note:
            # The landscape source keeps the note inside the label's parentheses.
            name, note = split_sector_note(name)
        if note and any(value is not None for value in values.values()):
            raise ValueError('Digital-maturity sector has both a note and percentages')
        sectors.append({'name': name, 'metrics': values, 'note': note})
    tables = []
    for table in page.extract_tables():
        rows = [[_rtl(cell) for cell in row] for row in table]
        header = next((i for i, row in enumerate(rows) if len(row) == 3 and row[2] == '#' and row[1] == 'المنظور'), None)
        if header is not None:
            tables.append(rows[header+1:])
    if len(tables) != 1:
        raise ValueError('Cannot identify the digital-maturity challenges table')
    challenges = []
    for index, row in enumerate(tables[0], 1):
        if len(row) != 3 or row[2] != str(index):
            raise ValueError('Digital-maturity challenge numbering is inconsistent')
        perspective, text = row[1], row[0]
        if not perspective or not text:
            raise ValueError('Digital-maturity challenge is incomplete')
        challenges.append({'perspective': perspective, 'text': text})
    if not 1 <= len(challenges) <= 4:
        raise ValueError('Digital-maturity challenges exceed the approved slide capacity')
    date = re.search(r'14\d\d/\d{1,2}/\d{1,2}', page.extract_text() or '')
    if not date:
        raise ValueError('Digital-maturity PDF update date is missing')
    return metrics, sectors, challenges, date.group()


def _extract_detail_page(page, number):
    _check_page(page, number)
    main, challenge = [], []
    for found in page.find_tables():
        rows = [[text for text in (_rtl(cell) for cell in row) if text] for row in found.extract()]
        if not any(rows):
            continue
        header = rows[1] if len(rows) > 1 else []
        if 'أساس التقييم' in header and 'م' in header:
            main.append(rows)
        elif rows[0] and rows[0][0] in CHALLENGE_TITLES and '#' in header:
            challenge.append(rows)
    if len(main) != 1:
        raise ValueError(f'Expected one detail table on PDF page {number}')
    if len(challenge) > 1:
        raise ValueError(f'Expected at most one challenge table on PDF page {number}')
    rows = main[0]
    title = rows[0][0] if rows[0] else ''
    if not title:
        raise ValueError(f'Digital-maturity detail title is missing on PDF page {number}')
    header = rows[1]
    if len(header) != 4 or header[1] != 'الإنجاز' or header[2] not in TASK_HEADERS:
        raise ValueError(f'Unexpected detail headers on PDF page {number}: {header}')
    draft_rows = []
    for raw_row in rows[2:]:
        if not raw_row:
            continue
        if len(raw_row) == 1:
            draft_rows.append([raw_row[0], '', '', ''])
        elif len(raw_row) == 4:
            draft_rows.append(raw_row)
        else:
            raise ValueError(f'Digital-maturity detail row is incomplete on PDF page {number}: {raw_row}')
    if not draft_rows:
        raise ValueError(f'Digital-maturity detail table is empty on PDF page {number}')
    cards = {axis: _card_value(page, axis, DETAIL_CARD_BAND) for axis in AXES}
    challenges, challenge_title, challenge_header = [], None, None
    if challenge:
        challenge_title = challenge[0][0][0]
        challenge_header = next((cell for cell in challenge[0][1] if cell != '#'), None)
        if challenge_header not in CHALLENGE_HEADERS:
            raise ValueError(f'Unexpected challenge headers on PDF page {number}')
        for index, row in enumerate(challenge[0][2:], 1):
            if len(row) != 2 or row[1] != str(index):
                raise ValueError(f'Digital-maturity challenge numbering is inconsistent on PDF page {number}')
            challenges.append(row[0])
    return {'page': number, 'title': title, 'cards': cards, 'rows': draft_rows,
            'note_title': challenge_title, 'notes': challenges}


def merge_detail_pages(pages):
    """Fold the source pagination back into one record set per sector."""
    sectors, declared = [], {}
    for page in pages:
        base, index, total = split_title(page['title'])
        if not base:
            raise ValueError(f'Digital-maturity PDF page {page["page"]} has no sector title')
        rows = [list(row) for row in page['rows']]
        current = sectors[-1] if sectors else None
        if current and current['title'] == base and index not in (None, 1):
            if index != len(current['source_pages']) + 1 or total != declared[base]:
                raise ValueError(f'Digital-maturity PDF page {page["page"]} breaks the {base} page sequence')
            if page['cards'] != current['cards']:
                raise ValueError(f'Digital-maturity cards differ between the pages of {base}')
            if rows and is_heading(rows[0]):
                rows[0][0] = heading_base(rows[0][0])
                if rows[0][0] == last_heading(current['rows']):
                    rows = rows[1:]
        else:
            if base in declared:
                raise ValueError(f'Digital-maturity sector {base} appears twice outside one page sequence')
            if index not in (None, 1):
                raise ValueError(f'Digital-maturity PDF page {page["page"]} starts {base} at page {index}')
            current = {'title': base, 'source_pages': [], 'cards': page['cards'], 'rows': [],
                       'source_notes': [], 'updates': [], 'challenges': []}
            sectors.append(current)
            declared[base] = total
        if any(is_continuation(row) for row in rows):
            raise ValueError(f'Digital-maturity PDF page {page["page"]} repeats a heading mid-page')
        current['rows'].extend(rows)
        current['source_pages'].append(page['page'])
        # Source notes are review evidence only: the reviewer rewrites each one and
        # files it under `updates` or `challenges` before the report is built.
        current['source_notes'].extend({'title': page['note_title'], 'text': text} for text in page['notes'])
    for sector in sectors:
        if declared[sector['title']] not in (None, len(sector['source_pages'])):
            raise ValueError(f'Digital-maturity sector {sector["title"]} is missing source pages')
    if not sectors:
        raise ValueError('Digital-maturity PDF has no detail pages')
    return sectors


def extract_pdf(path):
    """Extract the summary and merged sector drafts; every Arabic field needs visual review."""
    path = Path(path)
    with pdfplumber.open(path) as pdf:
        if len(pdf.pages) < 2:
            raise ValueError('Digital-maturity PDF needs a summary page and at least one detail page')
        metrics, sectors, challenges, date = _extract_summary(pdf.pages[0])
        pages = [_extract_detail_page(page, number) for number, page in enumerate(pdf.pages[1:], 2)]
    return {'source_pdf_sha256': hashlib.sha256(path.read_bytes()).hexdigest(),
            'source_date_hijri': date, 'metrics': metrics,
            'sectors': sectors, 'source_updates': challenges,
            'updates': [dict(item) for item in challenges],
            'detail_sectors': merge_detail_pages(pages), 'reviewed': False}


def load_reviewed(path, pdf):
    data = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    if data.get('source_pdf_sha256') != hashlib.sha256(Path(pdf).read_bytes()).hexdigest():
        raise ValueError('Reviewed digital-maturity data does not match the supplied PDF')
    if data.get('reviewed') is not True:
        raise ValueError('Review the extracted digital-maturity data against the rendered PDF before building')
    if set(data.get('metrics', {})) != set(AXES):
        raise ValueError('Four digital-maturity metrics are required')
    for value in data['metrics'].values():
        if not isinstance(value, (int, float)) or isinstance(value, bool) or not 0 <= value <= 100:
            raise ValueError('Invalid digital-maturity metric')
    sectors = data.get('sectors', [])
    if not 1 <= len(sectors) <= 5 or len({s['name'] for s in sectors}) != len(sectors):
        raise ValueError('Digital-maturity sectors must be unique and fit the approved chart')
    for sector in sectors:
        if not sector.get('name') or set(sector.get('metrics', {})) != set(AXES):
            raise ValueError('Incomplete digital-maturity sector')
        numbers = list(sector['metrics'].values())
        if any(v is not None and (not isinstance(v, (int, float)) or isinstance(v, bool) or not 0 <= v <= 100) for v in numbers):
            raise ValueError('Invalid digital-maturity sector percentage')
        if sector.get('note') and any(v is not None for v in numbers):
            raise ValueError('A digital-maturity sector note cannot have percentages')
    updates = data.get('updates')
    if not isinstance(updates, list) or not 1 <= len(updates) <= 4:
        raise ValueError('Digital-maturity updates must hold 1-4 rows for the summary table')
    if any(not isinstance(u, dict) or not u.get('perspective') or not _text_ok(u.get('text')) for u in updates):
        raise ValueError('Incomplete digital-maturity summary update')
    for update in updates:
        _no_codes(update['text'], 'summary')
    details = data.get('detail_sectors')
    if not isinstance(details, list) or not details:
        raise ValueError('Reviewed digital-maturity data must contain merged detail_sectors, one per sector')
    if len({d.get('title') for d in details}) != len(details):
        raise ValueError('Digital-maturity detail sectors must be unique')
    for detail in details:
        title = detail.get('title')
        if not isinstance(title, str) or not title.strip() or split_title(title)[1] is not None:
            raise ValueError('Digital-maturity detail title must be the sector name without a page suffix')
        if set(detail.get('cards', {})) != set(AXES):
            raise ValueError('Digital-maturity detail cards are incomplete: ' + title)
        for value in detail['cards'].values():
            if not isinstance(value, str) or not value.strip():
                raise ValueError('Digital-maturity detail card is empty: ' + title)
            if '%' in value:
                _percent(value)
            elif value not in CARD_STATUSES:
                raise ValueError('Unapproved digital-maturity detail status: ' + value)
        pages = detail.get('source_pages', [])
        if not pages or any(not isinstance(p, int) or isinstance(p, bool) or p < 2 for p in pages) or pages != sorted(pages):
            raise ValueError('Digital-maturity detail source pages are invalid: ' + title)
        rows = detail.get('rows', [])
        if not rows or any(len(row) != 4 or any(not isinstance(cell, str) for cell in row) for row in rows):
            raise ValueError('Digital-maturity detail rows are incomplete: ' + title)
        for row in rows:
            if not (all(cell.strip() for cell in row) or is_heading(row)):
                raise ValueError('Digital-maturity detail row must be complete or a section heading: ' + title)
            if is_continuation(row):
                raise ValueError('Reviewed rows must not carry source continuation headings: ' + title)
        if is_heading(rows[-1]):
            raise ValueError('Digital-maturity section heading has no records: ' + title)
        for key, _, _ in NOTE_TABLES:
            items = detail.get(key)
            if not isinstance(items, list) or any(not _text_ok(item) for item in items):
                raise ValueError(f'Digital-maturity detail {key} are incomplete: ' + title)
            for item in items:
                _no_codes(item, title)
        if detail.get('source_notes') and not (detail['updates'] or detail['challenges']):
            raise ValueError('Source notes are not classified into updates or challenges: ' + title)
    return data


def _text_ok(value):
    return isinstance(value, str) and bool(value.strip())


def _no_codes(text, where):
    found = CODE_IN_TEXT.search(text)
    if found:
        raise ValueError(f'Remove the task code {found.group()} from the rewritten note ({where}); keep dates')


def _chart(shape, categories, series, percent_labels=False):
    chart_data = CategoryChartData()
    chart_data.categories = categories
    for name, values in series:
        chart_data.add_series(name, values)
    shape.chart.replace_data(chart_data)
    if percent_labels:
        shape.chart.plots[0].data_labels.number_format = '0%'
        for chart_series in shape.chart.series:
            chart_series.data_labels.number_format = '0%'
    blank = shape.chart._chartSpace.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}dispBlanksAs')
    if blank is not None:
        blank.set('val', 'gap')


def digital_slide(prs):
    matches = [sl for sl in prs.slides if sum(sh.has_chart for sh in sl.shapes) == 5
               and any(sh.has_table and len(sh.table.columns) == 3 for sh in sl.shapes)]
    if len(matches) != 1:
        raise ValueError('Expected one approved digital-maturity slide')
    return matches[0]


def calculation_note(slide):
    note = helpers.shape_by_id(slide, CALCULATION_NOTE_ID)
    if ' '.join(note.text.split()) != ' '.join(CALCULATION_NOTE.split()):
        raise ValueError('Approved digital-maturity calculation note is missing or changed')
    return note


def detail_slides(prs):
    """Return the digital detail slides in deck order; the count follows the data."""
    matches = [slide for slide in prs.slides if is_digital_detail(slide)]
    if not matches:
        raise ValueError('Expected at least one approved digital-maturity detail slide')
    return matches


def _detail_shape(slide):
    return next(sh for sh in slide.shapes if is_digital_detail_table(sh))


def _detail_table(slide):
    return _detail_shape(slide).table


def _note_shapes(slide):
    return sorted((sh for sh in slide.shapes if sh.has_table and len(sh.table.columns) == 2
                   and sh.table.cell(0, 0).text in CHALLENGE_TITLES), key=lambda sh: sh.top)


def _challenge_shape(slide):
    shapes = _note_shapes(slide)
    return shapes[0] if shapes else None


def _challenge_table(slide):
    shape = _challenge_shape(slide)
    return None if shape is None else shape.table


def _note_tables(slide):
    return [shape.table for shape in _note_shapes(slide)]


def sector_notes(sector):
    """The sector's output note tables in order: (title, header, items), empty lists omitted."""
    return [(title, header, list(sector[key])) for key, title, header in NOTE_TABLES if sector.get(key)]


def _detail_columns(table, row):
    # A merged section heading stays in its merge origin, in either direction.
    if table.cell(1, 0).text == 'أساس التقييم' or row.cells[0].span_width == 4:
        return range(4)
    return range(3, -1, -1)


def _plain(value):
    # Older candidates wrapped codes in LRO/PDF marks; keep reading them.
    return value.replace('\u202d', '').replace('\u202c', '')


def _set_code_text(cell, value, pattern=CODE_IN_TEXT):
    """Write `value`, giving each Latin code its own run tagged English.

    PowerPoint lays out an Arabic-tagged run right-to-left, so a code such as
    «U-01..09» or «INF-02» inside one is shown reversed («U-09..01», «02INF-»).
    Bidi control characters are not honoured reliably, so the language tag of
    the run carries the direction instead.
    """
    set_cell_exact(cell, value)
    for paragraph, line in zip(cell.text_frame.paragraphs, value.replace('\r\n', '\n').split('\n')):
        parts = re.split(f'({pattern.pattern})', line)
        if len(parts) == 1 or not paragraph.runs:
            continue
        first = paragraph.runs[0]
        template = deepcopy(first._r.get_or_add_rPr())
        for run in paragraph.runs[1:]:
            run._r.getparent().remove(run._r)
        first.text = parts[0]
        for index, part in enumerate(parts[1:], 1):
            if not part:
                continue
            run = paragraph.add_run()
            run.text = part
            props = deepcopy(template)
            props.attrib.pop('rtl', None)
            if index % 2:
                props.set('lang', 'en-US')
                props.attrib.pop('altLang', None)
            run._r.insert(0, props)
        if not first.text:
            first._r.getparent().remove(first._r)


def _set_challenge_text(cell, value):
    _set_code_text(cell, value)
    _bullet_lines(cell)


BULLET_INDENT = 171450


def _bullet_lines(cell):
    """A note written on several lines is shown as native right-to-left bullets, one per line."""
    paragraphs = cell.text_frame.paragraphs
    if len(paragraphs) < 2:
        return
    a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    for paragraph in paragraphs:
        ppr = paragraph._p.get_or_add_pPr()
        for old in ppr.findall(a + 'buNone') + ppr.findall(a + 'buFont') + ppr.findall(a + 'buChar') + ppr.findall(a + 'buAutoNum'):
            ppr.remove(old)
        ppr.set('marL', str(BULLET_INDENT))
        ppr.set('indent', str(-BULLET_INDENT))
        font = ppr.makeelement(a + 'buFont', {'typeface': 'Arial'})
        char = ppr.makeelement(a + 'buChar', {'char': '\u2022'})
        anchor = next((child for child in ppr if child.tag in (a + 'tabLst', a + 'defRPr', a + 'extLst')), None)
        for element in (font, char):
            if anchor is None:
                ppr.append(element)
            else:
                anchor.addprevious(element)


def _axis_of(heading):
    base = heading_base(heading)
    return next((axis for axis, label in zip(AXES, LABELS) if base.startswith(label)), None)


class DetailPattern:
    """Row templates and height estimates taken from the approved master table."""

    def __init__(self, prs, seed):
        table_shape, challenge_shape = _detail_shape(seed), _challenge_shape(seed)
        if challenge_shape is None:
            raise ValueError('Approved digital detail pattern lacks its challenge table')
        self.table_element = deepcopy(table_shape._element)
        self.challenge_element = deepcopy(challenge_shape._element)
        self.top = table_shape.top
        self.bottom = content_bottom(prs, seed)
        rows = list(table_shape.table.rows)
        self.title_height, self.header_height = rows[0].height, rows[1].height
        self.heading_rows = {}
        for row in rows[2:]:
            if row.cells[0].span_width == 4:
                axis = _axis_of(row.cells[0].text)
                if axis and axis not in self.heading_rows:
                    self.heading_rows[axis] = deepcopy(row._tr)
                    self.heading_height = row.height
        if set(self.heading_rows) != set(AXES):
            raise ValueError('Approved digital detail pattern lacks a section heading for every perspective')
        bodies = [row for row in rows[2:] if row.cells[0].span_width != 4]
        self.body_row = deepcopy(bodies[0]._tr)
        heights = sorted({row.height for row in bodies})
        self.body_height = heights[0]
        self.line_height = (heights[1] - heights[0]) if len(heights) > 1 else round(heights[0] * .48)
        self.widths = [column.width / 914400 for column in table_shape.table.columns]
        challenge_rows = list(challenge_shape.table.rows)
        self.challenge_fixed = challenge_rows[0].height + challenge_rows[1].height
        self.challenge_body = deepcopy(challenge_rows[2]._tr)
        self.challenge_height = challenge_rows[2].height
        self.challenge_width = challenge_shape.table.columns[1].width / 914400
        if not all((self.title_height, self.header_height, self.heading_height, self.body_height,
                    self.challenge_fixed, self.challenge_height)):
            raise ValueError('Approved digital detail pattern has unresolved row heights')

    def row_height(self, row):
        if is_heading(row):
            return self.heading_height
        lines = max(helpers.estimate_lines(cell, width, 11, CHARS_PER_INCH) for cell, width in zip(row, self.widths))
        return self.body_height + (lines - 1) * self.line_height

    def challenge_row_height(self, text):
        lines = helpers.estimate_lines(text, self.challenge_width, 11, CHARS_PER_INCH)
        return self.challenge_height + max(0, lines - 2) * self.line_height

    def challenges_height(self, challenges):
        return self.challenge_fixed + sum(self.challenge_row_height(text) for text in challenges)

    def notes_height(self, notes):
        """Gap plus table height for every note table that follows the last records."""
        return sum(GAP + self.challenges_height(items) for _, _, items in notes)


def paginate_sector(pattern, sector):
    """Flow one sector into slide pages using the approved pattern's row heights.

    Returns a list of pages; each page is ``(rows, challenges)`` where rows may
    begin with a synthetic continuation heading.  Headings never end a page,
    records are never split, and the challenge table follows the last records.
    """
    rows, challenges = sector['rows'], sector_notes(sector)
    pages, index = [], 0
    while index < len(rows):
        used = pattern.title_height + pattern.header_height
        page_rows = []
        heading = last_heading(rows[:index])
        if heading is not None and not is_heading(rows[index]):
            page_rows.append([f'{heading} {CONTINUATION}', '', '', ''])
            used += pattern.heading_height
        while index < len(rows):
            row = rows[index]
            need = pattern.row_height(row)
            if is_heading(row):
                if index + 1 >= len(rows):
                    raise ValueError('Digital-maturity section heading has no records: ' + sector['title'])
                need += pattern.row_height(rows[index + 1])
            if pattern.top + used + need > pattern.bottom:
                break
            page_rows.append(row)
            used += pattern.row_height(row)
            index += 1
        page_challenges = []
        if index == len(rows) and challenges:
            if pattern.top + used + pattern.notes_height(challenges) <= pattern.bottom:
                page_challenges = list(challenges)
            else:
                # Only the trailing record (with its heading) moves to the next page with the challenges.
                page_rows.pop()
                index -= 1
                if page_rows and is_heading(page_rows[-1]) and not is_continuation(page_rows[-1]):
                    page_rows.pop()
                    index -= 1
        if not any(not is_heading(row) for row in page_rows):
            raise ValueError('A digital-maturity record cannot fit on an empty detail slide: ' + sector['title'])
        pages.append((page_rows, page_challenges))
    if challenges and not pages[-1][1]:
        raise ValueError('Digital-maturity challenges cannot fit with their last record: ' + sector['title'])
    return pages


def _append_row(table, template):
    table._tbl.append(deepcopy(template))
    return table.rows[len(table.rows) - 1]


def _place_table(slide, element, top):
    """Copy an approved table onto the slide at ``top`` with only its title and header rows."""
    shape = place_table_copy(slide, element)
    helpers.resize_table(shape.table, 2)
    shape.top = top
    return shape


def _fill_detail_table(slide, pattern, title, rows):
    shape = _place_table(slide, pattern.table_element, pattern.top)
    table = shape.table
    set_cell_exact(table.cell(0, 0), title)
    table.rows[0].height, table.rows[1].height = pattern.title_height, pattern.header_height
    for values in rows:
        if is_heading(values):
            axis = _axis_of(values[0])
            if axis is None:
                raise ValueError('Digital-maturity section heading has no approved perspective style: ' + values[0])
            row = _append_row(table, pattern.heading_rows[axis])
            set_cell_exact(row.cells[0], values[0])
        else:
            row = _append_row(table, pattern.body_row)
            columns = list(_detail_columns(table, row))
            for column, value in zip(columns, values):
                _set_code_text(row.cells[column], value)
            code_cell = row.cells[columns[3]]
            if CODE.fullmatch(values[3]):
                _set_code_text(code_cell, values[3], CODE)
                code_cell.text_frame.paragraphs[0]._p.get_or_add_pPr().set('rtl', '0')
        row.height = pattern.row_height(values)
    shape.height = sum(row.height for row in table.rows)
    return shape


def _fill_challenge_table(slide, pattern, top, note):
    title, header, items = note
    shape = _place_table(slide, pattern.challenge_element, top)
    table = shape.table
    set_cell_exact(table.cell(0, 0), title)
    set_cell_exact(table.cell(1, 1), header)
    for index, challenge in enumerate(items, 1):
        row = _append_row(table, pattern.challenge_body)
        set_cell_exact(row.cells[0], str(index))
        _set_challenge_text(row.cells[1], challenge)
        row.height = pattern.challenge_row_height(challenge)
    shape.height = sum(row.height for row in table.rows)
    return shape


def _set_cards(slide, cards):
    for axis, shape_id in DETAIL_CARD_IDS.items():
        shape = helpers.shape_by_id(slide, shape_id)
        paragraphs = shape.text_frame.paragraphs
        if len(paragraphs) != 2 or len(paragraphs[1].runs) != 1:
            raise ValueError('Approved digital-maturity card formatting changed')
        paragraphs[1].runs[0].text = cards[axis]


def populate_details(prs, data):
    slides = detail_slides(prs)
    seed, extras = slides[0], slides[1:]
    pattern = DetailPattern(prs, seed)
    for shape in [sh for sh in seed.shapes if sh.has_table]:
        shape._element.getparent().remove(shape._element)
    layout = [(sector, paginate_sector(pattern, sector)) for sector in data['detail_sectors']]
    # Clone the table-free seed once per page before filling, so no populated table is copied.
    slides, current = [seed], seed
    for _ in range(sum(len(pages) for _, pages in layout) - 1):
        current = helpers.clone_slide(prs, seed, after=current)
        slides.append(current)
    targets = iter(slides)
    for sector, pages in layout:
        for number, (rows, challenges) in enumerate(pages, 1):
            slide = next(targets)
            _set_cards(slide, sector['cards'])
            title = sector['title'] if len(pages) == 1 else f'{sector["title"]} ({number}/{len(pages)})'
            shape = _fill_detail_table(slide, pattern, title, rows)
            for note in challenges:
                shape = _fill_challenge_table(slide, pattern, shape.top + shape.height + GAP, note)
    for slide in extras:
        helpers.remove_slide(prs, slide)


def populate(prs, data):
    slide = digital_slide(prs)
    calculation_note(slide)
    for axis, (chart_id, label_id, track_id, fill_id) in CARDS.items():
        value = data['metrics'][axis]
        _chart(helpers.shape_by_id(slide, chart_id), ['المنجز', 'المتبقي'],
               [('نسبة الإنجاز', [_chart_number(value), _chart_number(100-value)])])
        label_shape = helpers.shape_by_id(slide, label_id)
        helpers.set_text(label_shape, f'{value:g}%')
        label_shape.text_frame.paragraphs[0]._p.get_or_add_pPr().set('rtl', '0')
        for run in label_shape.text_frame.paragraphs[0].runs:
            run._r.get_or_add_rPr().set('lang', 'en-US')
        track, fill = (helpers.shape_by_id(slide, sid) for sid in (track_id, fill_id))
        fill.left = track.left
        fill.width = round(track.width * value / 100)
    sector_chart = helpers.shape_by_id(slide, 93)
    _chart(sector_chart, [s['name'] for s in data['sectors']],
           [(label, [_chart_number(s['metrics'][axis]/100) if s['metrics'][axis] is not None else None
                     for s in data['sectors']])
            for axis, label in zip(AXES, LABELS)], percent_labels=True)
    if any(s.get('note') for s in data['sectors']):
        row_spacing = 1524000 / len(data['sectors'])
        for index, sector in enumerate(data['sectors']):
            if not sector.get('note'):
                continue
            left = sector_chart.left
            top = sector_chart.top + round(155000 + (5-len(data['sectors']))*70000 + index * row_spacing)
            box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, 4900000, 155000)
            box.name = f'Digital sector note {index+1}'
            box.fill.solid()
            box.fill.fore_color.rgb = RGBColor(229, 227, 224)
            box.line.fill.background()
            box.text = sector['note']
            for paragraph in box.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER
                paragraph.font.size = Pt(11)
                paragraph.font.name = 'Abar Mid'
                paragraph.font.color.rgb = RGBColor(55, 55, 55)
    table = next(sh.table for sh in slide.shapes if sh.has_table and len(sh.table.columns) == 3)
    set_cell_exact(table.cell(0, 0), SUMMARY_TABLE[0])
    set_cell_exact(table.cell(1, 2), SUMMARY_TABLE[1])
    for index, update in enumerate(data['updates'], 2):
        for cell, value in zip(table.rows[index].cells,
                               (str(index-1), update['perspective'], update['text'])):
            set_cell_exact(cell, value)
        _bullet_lines(table.rows[index].cells[2])
    for row in list(table.rows)[2+len(data['updates']):]:
        row._tr.getparent().remove(row._tr)
    populate_details(prs, data)
    set_table_directions(prs)
    return slide


def _slide_rows(table):
    return [[_plain(row.cells[column].text) for column in _detail_columns(table, row)]
            for row in list(table.rows)[2:]]


class _DetailSlide:
    """One generated digital detail slide, read once for every reconciliation check."""

    def __init__(self, slide):
        self.slide = slide
        self.table = _detail_table(slide)
        self.base, self.index, self.total = split_title(self.table.cell(0, 0).text)
        self.rows = _slide_rows(self.table)
        self.notes = _note_tables(slide)
        self.bottom = max(sh.top + sh.height for sh in slide.shapes if sh.has_table)

    def card(self, axis):
        return helpers.shape_by_id(self.slide, DETAIL_CARD_IDS[axis]).text_frame.paragraphs[1].text


def check_details(prs, data, errors, master=None):
    """Reconcile the generated detail slides with the reviewed sectors.

    The number of slides follows the data, so the checks cover: sector order,
    output-only «(k/n)» suffixes, identical cards on every slide of a sector,
    complete rows in source order with synthetic continuation headings only,
    a single challenge table on the sector's last slide, and no premature page
    break while a later slide of the same sector still holds records.
    """
    try:
        slides = detail_slides(prs)
    except ValueError as exc:
        errors.append(str(exc))
        return
    master = master or Presentation(MASTER)
    pattern = DetailPattern(master, detail_slides(master)[0])
    groups = []
    for detail in map(_DetailSlide, slides):
        if groups and groups[-1][0] == detail.base and detail.index not in (None, 1):
            groups[-1][1].append(detail)
        else:
            groups.append((detail.base, [detail]))
    expected = [sector['title'] for sector in data['detail_sectors']]
    if [base for base, _ in groups] != expected:
        errors.append('Digital-maturity detail sectors differ from the reviewed PDF data: ' + ', '.join(base for base, _ in groups))
        return
    for sector, (base, members) in zip(data['detail_sectors'], groups):
        count = len(members)
        rows = []
        for position, detail in enumerate(members, 1):
            if (count == 1 and detail.index is not None) or (count > 1 and (detail.index, detail.total) != (position, count)):
                errors.append(f'Digital-maturity detail title suffix does not match the output pagination: {base}')
            for axis in AXES:
                if detail.card(axis) != sector['cards'][axis]:
                    errors.append(f'Digital-maturity detail card differs: {base}: {axis}')
            if position < count and detail.notes:
                errors.append(f'Digital-maturity update/challenge tables must follow the last records of {base}')
            slide_rows = detail.rows
            if position > 1:
                if slide_rows and is_continuation(slide_rows[0]):
                    if heading_base(slide_rows[0][0]) != last_heading(rows):
                        errors.append(f'Digital-maturity continuation heading differs: {base}')
                    slide_rows = slide_rows[1:]
                elif last_heading(rows) is not None and slide_rows and not is_heading(slide_rows[0]):
                    errors.append(f'Digital-maturity continuation slide lacks its section heading: {base}')
            if any(is_continuation(row) for row in slide_rows):
                errors.append(f'Digital-maturity continuation heading appears mid-slide: {base}')
            rows.extend(slide_rows)
        if rows != sector['rows']:
            errors.append(f'Digital-maturity detail rows differ from the reviewed PDF data: {base}')
        notes = sector_notes(sector)
        actual = [(t.cell(0, 0).text, t.cell(1, 1).text, [_plain(row.cells[1].text) for row in list(t.rows)[2:]])
                  for t in members[-1].notes]
        if actual != notes:
            errors.append(f'Digital-maturity detail updates/challenges differ from the reviewed data: {base}')
        for position, (detail, following) in enumerate(zip(members, members[1:]), 2):
            following_rows = following.rows
            if following_rows and is_continuation(following_rows[0]):
                following_rows = following_rows[1:]  # a consequence of the break, not a cost of moving the record back
            next_record = next((row for row in following_rows if not is_heading(row)), None)
            if next_record is None:
                continue
            if position == count and following.notes:
                # The last records may move only together with the update/challenge tables.
                need = sum(pattern.row_height(row) for row in following_rows) + pattern.notes_height(notes)
            else:
                need = sum(pattern.row_height(row) for row in following_rows[:following_rows.index(next_record) + 1])
            if detail.bottom + need <= pattern.bottom:
                errors.append(f'Digital-maturity detail slide breaks the page early; the next record still fits: {base}')


def check(prs, report_path, data, errors, master=None):
    from report_common import validation
    slide = digital_slide(prs)
    try:
        calculation_note(slide)
    except ValueError as exc:
        errors.append(str(exc))
    for axis, (chart_id, label_id, track_id, fill_id) in CARDS.items():
        value = data['metrics'][axis]
        if helpers.shape_by_id(slide, label_id).text != f'{value:g}%':
            errors.append('Digital-maturity card label differs: ' + axis)
        track, fill = (helpers.shape_by_id(slide, sid) for sid in (track_id, fill_id))
        if abs(fill.width - round(track.width*value/100)) > 2 or fill.left != track.left:
            errors.append('Digital-maturity progress bar differs: ' + axis)
        chart = helpers.shape_by_id(slide, chart_id)
        validation.chart_metrics(report_path, str(chart.chart.part.partname).lstrip('/'),
                                 [['المنجز', _chart_number(value)], ['المتبقي', _chart_number(100-value)]],
                                 ['نسبة الإنجاز'], errors)
    sector_chart = helpers.shape_by_id(slide, 93)
    if any(series.data_labels.number_format != '0%' for series in sector_chart.chart.series):
        errors.append('Digital-maturity sector labels must display percentages')
    rows = [[s['name'], *[_chart_number(s['metrics'][axis]/100) if s['metrics'][axis] is not None else None
                         for axis in AXES]] for s in data['sectors']]
    validation.chart_metrics(report_path, str(sector_chart.chart.part.partname).lstrip('/'), rows, LABELS, errors)
    table = next(sh.table for sh in slide.shapes if sh.has_table and len(sh.table.columns) == 3)
    actual = [[c.text for c in row.cells] for row in list(table.rows)[2:]]
    expected = [[str(i), u['perspective'], u['text']] for i, u in enumerate(data['updates'], 1)]
    if actual != expected:
        errors.append('Digital-maturity summary updates differ from the reviewed data')
    if (table.cell(0, 0).text, table.cell(1, 2).text) != SUMMARY_TABLE:
        errors.append('Digital-maturity summary table must be titled «أبرز التحديثات»')
    notes = {s.name: s.text for s in slide.shapes if s.name.startswith('Digital sector note ')}
    expected_notes = {f'Digital sector note {i}': s['note'] for i, s in enumerate(data['sectors'], 1) if s.get('note')}
    if notes != expected_notes:
        errors.append('Digital-maturity sector notes differ from reviewed PDF data')
    check_details(prs, data, errors, master)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--pdf', required=True)
    parser.add_argument('--output', required=True)
    args = parser.parse_args(argv)
    output = Path(args.output)
    if output.resolve() == Path(args.pdf).resolve():
        raise ValueError('Extraction output must not overwrite the PDF')
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(extract_pdf(args.pdf), ensure_ascii=False, indent=2), encoding='utf-8')
    print(output)


if __name__ == '__main__':
    main()
