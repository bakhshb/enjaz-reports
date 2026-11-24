"""One date format for every report: dd-mm-yyyy.

Standard library only, so the Excel skills can import it without extra
dependencies. Callers pass values from date columns only; a value this module
cannot read raises DateError and the skill stops with the row and the value.
"""
import datetime as dt
import re

DIGITS = str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹', '01234567890123456789')
# Bidi marks that Arabic exports put around numbers.
MARKS = dict.fromkeys(map(ord, '‎‏؜‪‫‬‭‮'), None)

MONTHS = {
    'يناير': 1, 'فبراير': 2, 'مارس': 3, 'ابريل': 4, 'أبريل': 4, 'إبريل': 4,
    'مايو': 5, 'يونيو': 6, 'يونيه': 6, 'يوليو': 7, 'يوليه': 7,
    'اغسطس': 8, 'أغسطس': 8, 'سبتمبر': 9, 'اكتوبر': 10, 'أكتوبر': 10,
    'نوفمبر': 11, 'ديسمبر': 12,
}

TIME = r'(?:[ T]00:00(?::00)?)?'
DAY_FIRST = re.compile(r'(\d{1,2})([-/.])(\d{1,2})\2(\d{4})' + TIME)
YEAR_FIRST = re.compile(r'(\d{4})([-/.])(\d{1,2})\2(\d{1,2})' + TIME)
MONTH_NAME = re.compile(r'(\d{1,2})\s*(\S+?)\s*(\d{4})')
# Excel serial day numbers between 1955 and 2109; anything else is not a date.
SERIAL_RANGE = (20000, 76000)
EXCEL_EPOCH = dt.date(1899, 12, 30)
# Hijri years in the data are 13xx-15xx; Gregorian report years are 1900+.
HIJRI_YEARS = range(1300, 1600)


class DateError(ValueError):
    """A value in a date column that is not a readable date."""


def empty(value):
    return value is None or value != value or (isinstance(value, str) and not value.strip())


def _check(year, month, day, value, hijri):
    if year in HIJRI_YEARS:
        if not hijri:
            raise DateError(f'Hijri date not allowed here: {value}')
        if 1 <= month <= 12 and 1 <= day <= 30:
            return year, month, day
        raise DateError(f'bad date: {value}')
    try:
        dt.date(year, month, day)
    except ValueError:
        raise DateError(f'bad date: {value}') from None
    return year, month, day


def date_parts(value, *, hijri=False):
    """Return (year, month, day) for a date value, or None for an empty cell.

    hijri=True accepts Hijri text dates (the transactions exports use them);
    everywhere else a Hijri date stops the skill."""
    if empty(value):
        return None
    if isinstance(value, (dt.datetime, dt.date)):
        return value.year, value.month, value.day
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if SERIAL_RANGE[0] <= value <= SERIAL_RANGE[1]:
            day = EXCEL_EPOCH + dt.timedelta(days=int(value))
            return day.year, day.month, day.day
        raise DateError(f'not a date: {value}')
    text = re.sub(r'\s+', ' ', str(value).translate(MARKS).translate(DIGITS)).strip()
    text = re.sub(r'\s*م$', '', text)  # trailing Gregorian marker
    if re.search(r'هـ?$', text):
        if not hijri:
            raise DateError(f'Hijri date not allowed here: {value}')
        text = re.sub(r'\s*هـ?$', '', text)
    if m := DAY_FIRST.fullmatch(text):
        return _check(int(m[4]), int(m[3]), int(m[1]), value, hijri)
    if m := YEAR_FIRST.fullmatch(text):
        return _check(int(m[1]), int(m[3]), int(m[4]), value, hijri)
    if (m := MONTH_NAME.fullmatch(text)) and m[2] in MONTHS:
        return _check(int(m[3]), MONTHS[m[2]], int(m[1]), value, hijri)
    raise DateError(f'not a date: {value}')


def report_date(value, *, hijri=False):
    """Any readable date as dd-mm-yyyy text; an empty cell stays ''."""
    parts = date_parts(value, hijri=hijri)
    if parts is None:
        return ''
    year, month, day = parts
    return f'{day:02d}-{month:02d}-{year}'
