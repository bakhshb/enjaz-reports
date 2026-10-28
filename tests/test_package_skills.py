import json, subprocess, sys, tempfile, unittest, zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'tools'))
import package_skills

CLI = {'monday-meeting-report': ['build_report.py', 'report_gate_v2.py'],
       'weekly-consolidated-report': ['build_report.py', 'validate_report.py', 'digital_maturity.py'],
       'weekly-report-by-sector': ['split_report.py', 'qa_report.py'],
       'delayed-transactions-summary': ['build_report.py'],
       'task-details': ['build_report.py']}


class PackageSkillsTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.out = Path(cls.tmp.name)
        cls.built = {p.name: p for p in package_skills.build(cls.out)}

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_every_skill_bundled_with_matching_name(self):
        expected = {p.parent.name for p in (ROOT / 'skills').glob('*/SKILL.md')}
        self.assertEqual(set(self.built), expected)
        for name, path in self.built.items():
            head = (path / 'SKILL.md').read_text(encoding='utf-8').split('---')[1]
            self.assertIn(f'name: {name}\n', head)
            with zipfile.ZipFile(f'{path}.zip') as z:
                self.assertIn(f'{name}/SKILL.md', z.namelist())
                self.assertFalse([n for n in z.namelist() if '__pycache__' in n])

    def test_bundled_scripts_import_without_repository(self):
        for name, scripts in CLI.items():
            for script in scripts:
                with self.subTest(skill=name, script=script):
                    r = subprocess.run([sys.executable, str(self.built[name] / 'scripts' / script), '--help'],
                                       cwd=self.out, capture_output=True, text=True)
                    self.assertEqual(r.returncode, 0, r.stderr)
        render = self.built['monday-meeting-report'] / 'scripts/report_common/render.py'
        r = subprocess.run([sys.executable, str(render), '--help'], cwd=self.out, capture_output=True, text=True)
        self.assertEqual(r.returncode, 0, r.stderr)

    def test_sector_bundle_carries_weekly_master(self):
        self.assertEqual((self.built['weekly-report-by-sector'] / 'assets/weekly-report-master.pptx').read_bytes(),
                         (ROOT / 'skills/weekly-consolidated-report/assets/weekly-report-master.pptx').read_bytes())

    def test_plugin_manifests(self):
        plugin = json.loads((ROOT / '.claude-plugin/plugin.json').read_text(encoding='utf-8'))
        market = json.loads((ROOT / '.claude-plugin/marketplace.json').read_text(encoding='utf-8'))
        self.assertEqual([p['name'] for p in market['plugins']], [plugin['name']])


if __name__ == '__main__':
    unittest.main()
