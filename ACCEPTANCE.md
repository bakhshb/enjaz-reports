# Acceptance criteria and evidence — 2026-09-29

## Scope

This change includes the six-skill audit remediation, the user's updated Monday and weekly masters and PDF references, and the follow-up chart and font acceptance checks. The masters themselves are retained as supplied. No generated production ministry report is included in the tests.

## Required checks

| Criterion | Evidence / command |
| --- | --- |
| Invalid input cannot silently drop records or overwrite a valid result | `uv run tests/run_acceptance.py` |
| Latest weekly source statuses, support requests and transaction identifiers are preserved | Upstream regression cases |
| KPIs, details, chart caches and embedded workbooks reconcile independently | Monday and weekly gates; deliberate-corruption tests |
| Entirely blank/zero sectors are absent from charts; any positive status keeps a sector | Template acceptance cases: one active, mixed, empty and new sectors, for both Monday and weekly |
| User-approved masters are the runtime assets | Updated SHA-256 constants and restore checks; original PPTX/PDF assets retained |
| Abar survives populated transaction and notes cells and added rows | Template typography regression; actual PDF font check |
| Empty, normal, new-sector and overflow scenarios open without repair | Four synthetic scenarios; 11 native PowerPoint open/export checks |
| Exported PDF page counts match, and Arabic text actually uses Abar | `uv run tests/check_rendered_pdf.py <report-directory>` |
| No clipped text, overlaps or footer collisions | Final rendered-page review, after source reconciliation |
| Installed skill files match the tested repository | 23 file hashes and installed-copy smoke checks |

## Verification results

- 48 regression tests passed (211.280 seconds). Three focused checks then passed, including extended Monday chart/workbook coverage and the final continuation-divider regression.
- Rendered table/chart pages were reviewed across all four scenarios. Weekly transaction continuation pages initially retained KPI dividers; those were removed, the tables moved up, and the exact corrected output was reconciled, reopened, re-exported and inspected.
- Seven installed-copy smoke tests passed (76.183 seconds).
- Normal, empty, new-sector and overflow weekly files reconcile after the final native layout adjustment; all sector outputs reconcile against those files and the transaction source.
- All 11 final Adobe PDFs passed actual Arabic-font and page-count checks. Exact presentation and PDF hashes are recorded in `tests/acceptance-evidence.json`.
- 11 presentations opened and exported natively without PowerPoint repair. Export did not change source presentation hashes.
- The direct PowerPoint PDF export failed Abar fidelity even on the unchanged master. This route is not accepted. Adobe PDFMaker, the exporter used for the supplied references, preserves Abar. The batch helper permits one retry for its intermittent error -8 and stops if that fails.

## Reproduction

```text
uv run tests/run_acceptance.py
uv run tests/run_weekly_workflow.py --outdir <new-directory>
```

Inspect rendered layouts, adjust only generated copies where required, and rerun the workflow's read-only `verify` function and sector QA after any adjustment. On Windows with Adobe PDFMaker installed:

```powershell
./tests/export_acceptance.ps1 -ReportRoot <report-directory>
uv run tests/check_rendered_pdf.py <report-directory>
```

A successful export does not establish font or visual acceptance. Record the final file hashes after the checks.

## Scope limits carried forward from the audit

The weekly visual scenarios use completed/on-plan task and project details, plus delayed transactions, because those task/project status patterns have verified template examples. Other task/project statuses are covered by data checks; complete visual approval of their detail-cell styles requires an approved matching exemplar. The original Monday limitation is superseded by the follow-up below. Do not claim that this synthetic run approves every possible production input or status design. Legacy `.xls` and live ministry exports were not exercised in this run.

Definition of full release completion: every required check above passes for the exact final files, and any additional status styles used by a production report have an approved reference and pass its visual review. Pushing the reviewed branch does not itself merge or release it.


## Monday follow-up: updated status examples and protected output

The user supplied master `1b7c72de3a5034ee0352a8b6a67f3f519a2e4bad5d95ad93424acdb5353ef255` supplies all five required styles. `لم يبدأ` / `لم تبدأ` is valid for Suhail projects only; `معلق` is valid for tasks. The matching PDF reference was regenerated with Adobe PDFMaker from an unchanged master copy.

Acceptance requires exact master status formatting, rejection of unknown/missing/conflicting styles, source/chart/KPI reconciliation, and preservation of prior output through oversized input or save/chart/gate/publication failures. Publication uses a validated temporary candidate and a sibling file with normal inherited Windows access. Input and master paths are protected against output collisions.

Results: 54 full-suite tests passed, then 24 final focused tests passed after the Windows publication correction. Four fresh synthetic scenarios (normal, empty, new-sector, overflow) reconcile and visually pass, covering all valid task and project statuses. Five native PowerPoint opens preserved source bytes. Five Adobe PDFs, totaling 83 pages including the 18-page master, passed actual Abar usage and page-count checks. Every generated dynamic page was visually inspected. See `tests/monday-acceptance-evidence.json` for exact output hashes. This resolves the Monday status-style limitation; the older weekly visual scope above remains unchanged. Production input acceptance remains a separate per-report check.

Installed refresh: all 24 files match the tested repository. Four installed-copy smoke checks passed, including normal generation and empty/new-sector/overflow scenarios. The previous installed snapshot remains backed up.
