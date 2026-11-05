"""Regressions from the weekly digital-model handoff: package names, row density, PDF reader."""
import collections
import io
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

import pdfplumber
from pptx import Presentation
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'skills'))
sys.path.insert(0, str(ROOT / 'skills/weekly-consolidated-report/scripts'))
from report_common import pptx_helpers as helpers
import digital_maturity as digital


class SlidePartNameTests(unittest.TestCase):
    def test_clone_after_remove_keeps_package_names_unique(self):
        prs = Presentation(digital.MASTER)
        slides = list(prs.slides)
        helpers.clone_slide(prs, slides[1])
        helpers.clone_slide(prs, slides[1])
        helpers.remove_slide(prs, slides[2])
        helpers.clone_slide(prs, slides[1])
        helpers.remove_slide(prs, slides[3])
        helpers.clone_slide(prs, slides[1], after=slides[0])
        buffer = io.BytesIO()
        prs.save(buffer)
        names = zipfile.ZipFile(buffer).namelist()
        self.assertEqual([n for n, k in collections.Counter(names).items() if k > 1], [])
        self.assertEqual(len(Presentation(io.BytesIO(buffer.getvalue())).slides), len(prs.slides))


class DensityTests(unittest.TestCase):
    def test_shared_default_is_unchanged_and_digital_density_is_explicit(self):
        text = 'أ' * 60
        self.assertEqual(helpers.estimate_lines(text, 2, 11), 3)
        self.assertEqual(helpers.estimate_lines(text, 2, 11, digital.CHARS_PER_INCH), 2)


class PdfReaderTests(unittest.TestCase):
    def pdf_page(self, draw):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        path = Path(temp.name) / 'card.pdf'
        pdf = canvas.Canvas(str(path), pagesize=(400, 200))
        draw(pdf)
        pdf.save()
        doc = pdfplumber.open(path)
        self.addCleanup(doc.close)
        return doc.pages[0]

    def test_headline_percentage_ignores_weekly_change_line(self):
        def draw(pdf):
            pdf.setFont('Helvetica', 28); pdf.drawString(50, 120, '%72.5')
            pdf.setFont('Helvetica', 9); pdf.drawString(50, 90, '+2.5% vs last week')
        self.assertEqual(digital._one_percent(self.pdf_page(draw), (0, 0, 400, 200)), 72.5)

    def test_single_percentage_still_reads(self):
        page = self.pdf_page(lambda pdf: (pdf.setFont('Helvetica', 20), pdf.drawString(50, 120, '15%')))
        self.assertEqual(digital._one_percent(page, (0, 0, 400, 200)), 15)

    def test_two_headline_percentages_rejected(self):
        def draw(pdf):
            pdf.setFont('Helvetica', 28); pdf.drawString(50, 120, '%72'); pdf.drawString(200, 120, '%30')
        with self.assertRaisesRegex(ValueError, 'headline'):
            digital._one_percent(self.pdf_page(draw), (0, 0, 400, 200))

    def test_sector_note_accepts_mirrored_parentheses(self):
        self.assertEqual(digital.split_sector_note('مركز نسك (قيد الدراسة)'), ('مركز نسك', 'قيد الدراسة'))
        self.assertEqual(digital.split_sector_note('مركز نسك )قيد الدراسة('), ('مركز نسك', 'قيد الدراسة'))
        with self.assertRaisesRegex(ValueError, 'note is missing'):
            digital.split_sector_note('مركز نسك')


if __name__ == '__main__':
    unittest.main()
