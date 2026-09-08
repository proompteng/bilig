# Bilig model workspace

## Product decision

Build a workspace for answering questions with executable models. The basic
loop is to name the assumptions, express the relationship as formulas, inspect
the results, and save a scenario before changing the assumptions again.

The existing WorkPaper engine owns calculation and portable workbook state.
The application owns model organization, input controls, scenario history,
and storage. Both people and agents can use the exported workbook.

## First release

- A library at `/models`, also the default entry at `/` without workbook query
  parameters. Existing document links and `/workbook` open the spreadsheet.
- Contribution, project budget, and capacity starters, plus a custom model.
- Named numeric inputs and editable result formulas with visible A1 references.
- Calculation in a dedicated worker. Formula errors remain visible, and a
  timeout cannot freeze the interface.
- Immutable scenario snapshots of both assumptions and formulas. Comparison
  uses saved definitions, even after the current model changes.
- IndexedDB records with revision checks in the same transaction as writes.
  A conflicting tab cannot overwrite newer work. Saving is acknowledged only
  after the transaction completes.
- Model backup import/export, standalone WorkPaper export, archive/restore,
  search, and stable links to saved models.

## Design

Use the existing warm neutral surfaces, green accent, shared control radius,
and system typography. A narrow navigation rail, a model list, and one working
area provide orientation. Inputs and results sit beside each other; comparison
is a table. Focus and hover transitions identify actions. Reduced-motion
preferences disable transitions. On small screens the rail becomes a header
and the working area becomes one column.

## Implementation sequence

1. Define and validate the model document. Add executable WorkPaper tests for
   input changes, formula failures, frozen scenarios, export, and restore.
2. Implement the calculation worker and transactional browser repository.
3. Implement library, model editor, comparison, backup, and recovery behavior.
4. Integrate routing, keep workbook links working, and update route-specific
   browser tests.
5. Run focused tests, the production build, repository checks, and actual
   browser flows including reload, import, conflict, keyboard, and mobile use.

## Acceptance

- Create a contribution model. Change units from 100 to 150 and read revenue
  changing from 10,000 to 15,000 through WorkPaper.
- Save a baseline, change assumptions and formulas, and compare against the
  original saved result. Restore the baseline without changing its snapshot.
- Reload and recover the same model and scenarios. Open a second model and
  prove its state is independent.
- Export a model backup, import it as a new model, and obtain equal calculated
  results. Export a WorkPaper, restore it in the runtime, and read equal values.
- Keep a conflicting edit available for backup without overwriting the newer
  browser record. Show storage failures explicitly.
- Build a custom model by adding an input and a result formula.

## Boundaries

This release stores models in the current browser. Export provides a portable
backup; clearing browser data removes local models. It does not claim cloud
sync, XLSX fidelity, or agent execution. The existing spreadsheet and remote
workbook paths retain their own persistence and access contracts.

The browser repository follows the transaction lifecycle described in
[MDN's IndexedDB guide](https://developer.mozilla.org/en-US/docs/Web/API/IndexedDB_API/Using_IndexedDB).

## Delivery evidence

- Fourteen focused tests pass for model validation, calculation, frozen
  scenarios, backup identity, routing, and package version compatibility.
- Seven Chromium flows pass against the production build, including storage
  failure recovery, scenario restore, backup import, independent model state,
  stale-tab conflict recovery, custom formulas, and a 390px mobile viewport.
- The browser export is restored through the installed Node package in a
  separate process. `Results!B2` reads `15000` with formula
  `=Inputs!B2*Inputs!B3` after editing `Inputs!B2` to `150`.
- The package-owned service evaluator returns `verified: true`, with computed
  readback `24000 -> 38400` and restored readback `38400`.
- The existing production dependency audit required Fastify `5.12.1` and
  fast-uri `3.1.6`. The patched lockfile reports no known production
  vulnerabilities. These are upstream security releases:
  [Fastify](https://github.com/fastify/fastify/releases/tag/v5.12.1) and
  [fast-uri](https://github.com/fastify/fast-uri/releases/tag/v3.1.6).

The browser checks live in `e2e/tests/web-shell-models.pw.ts` and are tagged
`@browser-ci`, so the regular CI smoke suite continues to exercise the workflow.

Local validation on September 7, 2026 passed the production build, generated
checks, semantic checks, lint, type checking, dependency analysis, and the core,
formula, server, and browser-runtime correctness suites. The complete
`pnpm run ci` release gate remains incomplete: Vitest encountered `ENOSPC`
while writing temporary files, and the same disk-full condition interrupted
the follow-up XLSX checks. Later release and performance gates have not run.
The implementation is committed locally; publishing waits for a complete
release gate after sufficient disk space is available.
