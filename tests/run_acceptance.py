# /// script
# requires-python = ">=3.12"
# dependencies = ["pandas>=2.2,<3", "openpyxl>=3.1,<4", "python-pptx>=1.0.2", "lxml>=5.3", "pillow", "xlrd>=2"]
# ///
"""Run all synthetic regression tests with one isolated dependency environment."""
from pathlib import Path
import unittest

if __name__=='__main__':
    suite=unittest.defaultTestLoader.discover(str(Path(__file__).parent))
    result=unittest.TextTestRunner(verbosity=2).run(suite)
    raise SystemExit(not result.wasSuccessful())
