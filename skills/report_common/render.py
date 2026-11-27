# /// script
# requires-python = ">=3.12"
# dependencies = ["python-pptx>=1.0.2", "lxml>=5.3"]
# ///
"""Export RTL PowerPoint tables through LibreOffice without changing the source."""
import argparse
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from report_common.table_typography import all_tables, reverse_table_columns, DEFAULT_INSET
from report_common.status_styles import publish


def abar_installed():
    """True when LibreOffice can find the template font; without it widths and row heights are wrong."""
    try:
        fonts = subprocess.run(['fc-list', ':family'], capture_output=True, text=True, timeout=30).stdout
    except (OSError, subprocess.SubprocessError):
        return False
    return 'abar mid' in fonts.lower()


A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
# Abar Mid's descent as a share of the font size (700 of 1000 units; its ascent is 1100).
ABAR_DESCENT = 0.7
# PowerPoint puts the baseline of a line with a fixed spacing at this share of the spacing from the line's top.
FIXED_BASELINE = 0.8


def lower_fixed_spacing_text(table):
    """Move the text of `table` down in a disposable copy, to where PowerPoint draws it.

    The tables carry a fixed line spacing of 1.2 times the font size. For such
    a line LibreOffice keeps the font's whole descent and takes the rest of the
    spacing as ascent (5.5 pt of 13.2 at 11 pt Abar), so the text sits against
    the top of its cell and tall letters are cut at the cell edge. PowerPoint
    puts the baseline at 0.8 of the spacing. The difference is added to the top
    cell margin and taken from the bottom one, so the row heights and the text
    block's height stay as they are. Preview only: the measure step does not
    use it, because it reads the bottom of the last line, which is the same in
    both programs.
    """
    for tc in table._tbl.iter(A + 'tc'):
        shifts = []
        for paragraph in tc.iter(A + 'p'):
            exact = paragraph.find(A + 'pPr/' + A + 'lnSpc/' + A + 'spcPts')
            sizes = [int(node.get('sz')) for node in paragraph.iter(A + 'rPr', A + 'endParaRPr') if node.get('sz')]
            if exact is None or not sizes:
                continue
            spacing, size = int(exact.get('val')) / 100, max(sizes) / 100
            shifts.append(FIXED_BASELINE * spacing - (spacing - ABAR_DESCENT * size))
        shift = round(max(shifts, default=0) * 12700)
        if shift <= 0:
            continue
        props = tc.find(A + 'tcPr')
        if props is None:
            props = tc.makeelement(A + 'tcPr', {})
            tc.append(props)
        top, bottom = int(props.get('marT', DEFAULT_INSET)), int(props.get('marB', DEFAULT_INSET))
        anchor = props.get('anchor', 't')
        taken = min(bottom, shift)
        if anchor == 'ctr':
            # Centred text moves by half the difference of the two margins.
            props.set('marT', str(top + 2 * shift - taken))
            props.set('marB', str(bottom - taken))
        elif anchor == 'b':
            props.set('marB', str(bottom - taken))
        else:
            props.set('marT', str(top + shift))
            props.set('marB', str(bottom - taken))


def export_pdf(report, output, soffice=None, preview=True):
    """Render `report` to PDF through LibreOffice.

    With ``preview`` (default) table text is drawn where PowerPoint puts it
    (see `lower_fixed_spacing_text`); the measure step passes False and reads
    LibreOffice's own placement.
    """
    report, output = Path(report).resolve(), Path(output).resolve()
    if output == report or (output.exists() and output.samefile(report)):
        raise ValueError('PDF output must not overwrite the source PowerPoint')
    executable = soffice or shutil.which('soffice')
    if not executable:
        raise RuntimeError('LibreOffice is required for PDF rendering')
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='.report-pdf-', dir=output.parent) as temporary:
        root = Path(temporary)
        prs = Presentation(report)
        for table in all_tables(prs):
            if table._tbl.tblPr.get('rtl') in ('1', 'true'):
                # LibreOffice ignores DrawingML RTL table layout. Materialize its
                # visual order in a disposable LTR copy; keep the editable source RTL.
                reverse_table_columns(table)
                table._tbl.tblPr.set('rtl', '0')
            if preview:
                lower_fixed_spacing_text(table)
        candidate = root / 'report.pptx'
        prs.save(candidate)
        result = subprocess.run([executable, '-env:UserInstallation=' + (root / 'profile').as_uri(),
                                 '--headless', '--convert-to', 'pdf:impress_pdf_Export',
                                 '--outdir', str(root), str(candidate)],
                                capture_output=True, text=True, timeout=300)
        rendered = root / 'report.pdf'
        if result.returncode or not rendered.exists():
            raise RuntimeError('PDF export failed: ' + result.stdout + result.stderr)
        publish(rendered, output)
    return output


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--report', required=True)
    parser.add_argument('--output', required=True)
    args = parser.parse_args()
    if not abar_installed():
        print('WARNING: the Abar Mid font is not installed for LibreOffice. The PDF uses substitute '
              'fonts with different widths, so table rows can look clipped or overlapping although '
              'they fit in PowerPoint with Abar. Install Abar before judging row heights or fonts.',
              file=sys.stderr)
    print(export_pdf(args.report, args.output))
