# Installing the enjaz-reports skills

Instructions for an AI agent (Claude Code or Codex) asked to install these skills. Run the steps for the agent you are, then verify. Ask the user before installing system packages.

## Claude Code

Install as a plugin. It updates from GitHub whenever `main` changes.

```bash
claude plugin marketplace add bakhshb/enjaz-reports
claude plugin install enjaz-reports@enjaz-reports
```

Verify:

```bash
claude plugin list
```

`enjaz-reports` must show as enabled. The skills load in the next session; tell the user to start a new one.

To update later:

```bash
claude plugin marketplace update enjaz-reports
```

## Codex

Clone the repository once and link each skill into the Codex skills folder. Links keep the shared `skills/report_common/` reachable, and `git pull` updates every skill.

```bash
git clone https://github.com/bakhshb/enjaz-reports ~/enjaz-reports
mkdir -p ~/.codex/skills
for s in ~/enjaz-reports/skills/*/SKILL.md; do ln -sfn "$(dirname "$s")" ~/.codex/skills/; done
```

If `~/enjaz-reports` already exists, run `git -C ~/enjaz-reports pull` instead of cloning.

Verify that the six skills are linked:

```bash
ls ~/.codex/skills
```

Expected: `delayed-transactions-summary`, `monday-meeting-report`, `suhail-project-report`, `task-details`, `weekly-consolidated-report`, `weekly-report-by-sector`. Tell the user to restart Codex.

To update later:

```bash
git -C ~/enjaz-reports pull
```

Do not copy single skill folders out of the repository: three skills import `skills/report_common/`. If links are not possible, build standalone copies with `python3 tools/package_skills.py` and copy the folders from `dist/`.

## ChatGPT

An agent cannot add plugins to ChatGPT. The user adds the repository `https://github.com/bakhshb/enjaz-reports` from the ChatGPT Plugins page (**Add**) and presses the refresh button there after changes to `main`.

## Prerequisites

Check each one and report what is missing:

```bash
python3 --version   # 3.12 or later
uv --version        # the skills install uv themselves if missing
soffice --version   # LibreOffice, for PDF previews and sector page measurement
fc-list | grep -i abar   # Abar font used by the reports
```

The Abar font is a licensed file the user must provide. Without LibreOffice, run the sector splitter with `--no-refine` and tell the user that footer overlap is not checked.

## Verify a skill runs

From the repository (or `~/enjaz-reports` for Codex):

```bash
uv run skills/weekly-consolidated-report/scripts/build_report.py --help
```

It must print the usage text without an import error.
