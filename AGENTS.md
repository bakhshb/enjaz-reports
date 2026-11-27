# AGENTS.md

Guide for AI agents working in this repository.

## What this repo is

A Claude Code plugin (also used by Codex and ChatGPT) with skills that build the Ministry of Hajj and Umrah weekly achievement reports in Arabic. Each skill lives in `skills/<name>/` and its instructions are in `SKILL.md`.

## Skills and pipeline

1. `task-details`: raw tasks export → tasks summary Excel.
2. `suhail-project-report`: raw Suhail export → Suhail summary Excel (newest `Wxx` week wins).
3. `delayed-transactions-summary`: per-sector transaction files → one transactions Excel.
4. `weekly-consolidated-report`: the three Excel files → weekly PowerPoint, plus a short and a long Arabic message to the Minister.
5. `weekly-report-by-sector`: weekly deck → one deck per sector.

`monday-meeting-report` is separate: weekly topics + tasks + Suhail Excel → Monday meeting deck.

## Layout

- `skills/<name>/SKILL.md`: the skill's rules. Read it before changing that skill.
- `skills/<name>/scripts/`: Python build and validation scripts.
- `skills/<name>/assets/`: approved PowerPoint masters and reference PDFs. Do not edit them. Scripts check their SHA-256.
- `skills/<name>/references/`: business rules and acceptance checks.
- `skills/report_common/`: shared helpers imported by every skill. Keep it next to the skills. Look here before writing a helper: dates (`dates.py`), warn-and-continue notices (`notices.py`), wording checks for Arabic text the agent writes (`wording.py`), table typography and edit-hardening (`table_typography.py`), row measurement (`measure.py`), the manual-edit repair tool (`reflow.py`), shared spacing and status fills.
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
- A skill stops only when it cannot produce its report at all (a missing file, sheet or required column, a damaged master, a built file that would state something the inputs do not). Anything else, such as one unreadable date, an unknown status or title, tight space, is a best assumption recorded with `report_common.notices.warn()`; the script continues and writes `<output>.notices.json`, and the agent repeats every notice in its final reply. Do not add a new `raise` or `sys.exit` for a condition the report can survive.
- Text the agent writes into a report or a message (digital updates and challenges, the Minister messages) is formal government Arabic written the way a person writes it. `report_common/wording.py` holds the banned machine-sounding patterns; extend that list rather than adding a second one.
