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
from report_common.table_typography import all_tables, reverse_table_columns
from report_common.status_styles import publish


def export_pdf(report, output, soffice=None):
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
        candidate = root / 'report.pptx'
        prs.save(candidate)
        result = subprocess.run([executable, '-env:UserInstallation=' + (root / 'profile').as_uri(),
                                 '--headless', '--convert-to', 'pdf:impress_pdf_Export',
                                 '--outdir', str(root), str(candidate)],
                                capture_output=True, text=True, timeout=90)
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
    print(export_pdf(args.report, args.output))
