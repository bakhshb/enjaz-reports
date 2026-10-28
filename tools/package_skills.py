#!/usr/bin/env python3
"""Build standalone skill folders and zips for agents that install one skill at a time.

Each bundle carries the shared report_common package under scripts/ (the script
folder is already on Python's import path), so it runs without this repository.

Usage: python3 tools/package_skills.py [--outdir dist]
"""
import argparse, shutil
from pathlib import Path

SKILLS = Path(__file__).resolve().parents[1] / 'skills'
# Extra files a bundle needs from other skills: target path inside the bundle -> source.
EXTRAS = {'weekly-report-by-sector': {
    'assets/weekly-report-master.pptx': SKILLS / 'weekly-consolidated-report/assets/weekly-report-master.pptx'}}
IGNORE = shutil.ignore_patterns('__pycache__', '*.py[cod]')


def build(outdir):
    outdir = Path(outdir).resolve()
    if SKILLS in outdir.parents or outdir == SKILLS:
        raise SystemExit('Output folder must be outside skills/')
    built = []
    for skill in sorted(p.parent for p in SKILLS.glob('*/SKILL.md')):
        target = outdir / skill.name
        if target.exists():
            shutil.rmtree(target)
        shutil.copytree(skill, target, ignore=IGNORE)
        if any('report_common' in f.read_text(encoding='utf-8') for f in skill.glob('scripts/*.py')):
            shutil.copytree(SKILLS / 'report_common', target / 'scripts/report_common', ignore=IGNORE)
        for rel, source in EXTRAS.get(skill.name, {}).items():
            (target / rel).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, target / rel)
        # The zip root holds the skill folder, as claude.ai and ChatGPT skill uploads expect.
        shutil.make_archive(str(target), 'zip', outdir, skill.name)
        built.append(target)
    return built


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument('--outdir', default=str(SKILLS.parent / 'dist'))
    for path in build(parser.parse_args().outdir):
        print(f'{path}  {path}.zip')
