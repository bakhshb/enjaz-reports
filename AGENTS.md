# AGENTS.md

Guide for AI agents working in this repository.

## What this repo is

A Claude Code plugin (also used by Codex and ChatGPT) with skills that build the Ministry of Hajj and Umrah weekly achievement reports in Arabic. Each skill lives in `skills/<name>/` and its instructions are in `SKILL.md`.

## Skills and pipeline

1. `task-details`: raw tasks export → tasks summary Excel.
2. `suhail-project-report`: raw Suhail export → Suhail summary Excel (newest `Wxx` week wins).
3. `delayed-transactions-summary`: per-sector transaction files → one transactions Excel.
4. `weekly-consolidated-report`: the three Excel files → weekly PowerPoint, plus an Arabic summary message to the Minister.
5. `weekly-report-by-sector`: weekly deck → one deck per sector.

`monday-meeting-report` is separate: weekly topics + tasks + Suhail Excel → Monday meeting deck.

## Layout

- `skills/<name>/SKILL.md`: the skill's rules. Read it before changing that skill.
- `skills/<name>/scripts/`: Python build and validation scripts.
- `skills/<name>/assets/`: approved PowerPoint masters and reference PDFs. Do not edit them. Scripts check their SHA-256.
- `skills/<name>/references/`: business rules and acceptance checks.
- `skills/report_common/`: shared helpers imported by the PowerPoint skills. Keep it next to the skills.
- `tests/`: tests using synthetic data.
- `tools/package_skills.py`: builds standalone skill zips in `dist/`.
- `.claude-plugin/`: plugin and marketplace config.

## Running

Scripts declare their dependencies inline, so run them with `uv`:

```bash
uv run skills/task-details/scripts/build_report.py --help
uv run tests/run_acceptance.py          # all tests
uv run tests/test_upstream_integrity.py -v   # one test file
```

Needs Python 3.12+, `uv`, LibreOffice (`soffice`) for PDF previews, and the Abar font.

## Rules

- Output is Arabic and right-to-left. Keep it that way.
- Excel controls the content; the master controls the design. Do not add charts, columns, sheets, or text beyond what `SKILL.md` asks for.
- Never use template sample text or old files as weekly data.
- Visual review of the final file is required; passing tests alone is not acceptance.
- Use only synthetic data in tests. Do not commit real ministry files.
- Keep changes small. Update the skill's `SKILL.md` and tests when behavior changes.
