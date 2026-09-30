# Shared report workflow acceptance

Scope: Monday Meeting and Weekly Consolidated remain separate skills with separate masters and business rules. Baseline is main `0a56fcb` (merged Monday fixes). The two PPTX masters and PDF references must remain byte-for-byte unchanged.

## Rule traceability

| Original rule group | Maintained authority | Execution/check |
| --- | --- | --- |
| Monday agenda, task summary/details, Suhail summary/details | Monday `references/business-rules.md` | Existing builder and report gate; neutral helpers moved without changing bodies |
| Monday exact text, six-column omissions, runtime date and typography | Monday business reference | Existing gate and original regression suite |
| Weekly authoritative rules 1–11 | Weekly `references/business-rules.md` | Weekly builder and validator; visual fit remains a required separate decision |
| Weekly source/object discovery and task mapping | Same reference, corresponding headings | Role discovery, source reconciliation, exact text/order tests |
| Weekly Suhail and transactions mapping | Same reference, corresponding headings | Independent sections, update-date emphasis, leading-zero transaction identity tests |
| Weekly chart/numbering/link behavior | Same reference, Shared elements and chart updates | Native charts, cache/workbook checks, PowerPoint inspection |
| Weekly continuous table flow | Same reference, Table flow and overflow handling | Initial candidate placement, render, correction, final visual review |
| Existing acceptance checks for both reports | Each `references/acceptance.md` | Automated suite plus exact final output inspection |
| Shared status color approval, 29 September 2026 | Weekly business reference, approval section | Weekly role-specific formatting plus approved color map; unsupported statuses fail |
| Skill invocation and scope | Each `SKILL.md` frontmatter and inputs | Frontmatter validation; unchanged distinct triggers and workbooks |

The weekly suspended-task pattern labels its notes column `التحديث`; its destination still receives the suspended-task summary notes. Preserve the template label. The user-approved color extension is the only added styling authority. No Monday typography is transferred into weekly cells.

## Required results

1. Both documented builders run from installed folders with sibling `report_common`; no production import depends on tests or the other report's generator.
2. Normal, empty, new-sector and overflow cases pass for both reports. Each valid task and project status is exercised.
3. Source records, exact text, blank notes, dates, grouping, order and intentional column omissions reconcile. Repeated names are retained and transaction leading zeros survive.
4. KPI, sector counts, native charts, embedded workbooks and caches agree. Zero-only sectors disappear without removing active sectors or changing source totals.
5. Invalid data, unknown status, font/style corruption and failed build/validation/publication are rejected. Previous outputs and sources remain intact.
6. PowerPoint opens every accepted file without repair. All final slides pass visual review for approved Abar usage, Arabic, table flow, continuation headers, status fills and footer clearance.
7. Same-machine normal/overflow generation plus validation timing is recorded over three runs. Investigate repeatable regression over 10%; report stricter-check costs separately from original generation behavior. Rendering/review time is recorded separately and is never called a generation speedup.
8. Evidence records final hashes, scenario scope, actual performed checks and limitations. Any file edit invalidates affected results. A generated candidate is not an accepted report until visual review passes.

## Completion record

See `tests/workflow-acceptance-evidence.json` for the current run's measured results. Synthetic acceptance is not a claim of testing ministry production data. Installation and final visual results must be recorded only after those steps actually occur.
