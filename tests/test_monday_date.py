"""The cover follows the meeting date, including the Riyadh midnight boundary."""
import datetime as dt
import sys
import unittest

from test_deck_integrity import MONDAY, load_module

builder = load_module('monday_date_builder', MONDAY / 'build_report.py')


class MondayDateTests(unittest.TestCase):
    def test_default_is_this_or_next_monday_in_riyadh(self):
        for instant, expected in [
            ('2026-10-25T12:00:00+00:00', '2026-10-26'),
            ('2026-10-26T12:00:00+00:00', '2026-10-26'),
            ('2026-10-26T21:01:00+00:00', '2026-11-02'),
            ('2026-12-31T12:00:00+00:00', '2027-01-04'),
        ]:
            with self.subTest(instant=instant):
                self.assertEqual(builder.meeting_date(now=dt.datetime.fromisoformat(instant)),
                                 dt.date.fromisoformat(expected))

    def test_explicit_meeting_date_is_used_and_cli_rejects_invalid_dates(self):
        args = ['--topics', 'a', '--tasks', 'b', '--suhail', 'c', '--output', 'd',
                '--meeting-date', '2026-11-09']
        self.assertEqual(builder.parse_args(args).meeting_date, dt.date(2026, 11, 9))
        with self.assertRaises(SystemExit):
            builder.parse_args(args[:-1] + ['2026-02-30'])


if __name__ == '__main__':
    unittest.main()
