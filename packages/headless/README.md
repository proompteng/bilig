# @bilig/headless

[![npm: @bilig/headless](https://img.shields.io/npm/v/@bilig/headless?label=%40bilig%2Fheadless)](https://www.npmjs.com/package/@bilig/headless)
[![GitHub](https://img.shields.io/badge/GitHub-proompteng%2Fbilig-blue)](https://github.com/proompteng/bilig)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](../../LICENSE)

`@bilig/headless` is the production-targeted WorkPaper workbook facade for
`bilig`. It runs the `@bilig/core` spreadsheet engine without the browser UI and
exposes HyperFormula-style workbook workflows for services, tests, agents, and
server-side spreadsheet automation.

Use it when you need a headless spreadsheet engine for Node services, coding
agents, benchmark harnesses, persistence tests, formula-backed business logic,
or local-first workbook automation.

## Production Status

Use this package in production for documented WorkPaper/headless API workflows:
programmatic workbook creation, formula evaluation, structural edits,
persistence round trips, service-side spreadsheet automation, and agent-driven
workbook operations.

Current release posture:

- Full local CI passed after the latest headless hardening work, including unit,
  contract, fuzz, browser, clean-diff, release-budget, runtime-publish, and
  WorkPaper competitive benchmark gates.
- The checked-in competitive artifact generated on `2026-05-06T14:54:57.091Z`
  shows `46/46` comparable WorkPaper mean wins against HyperFormula-style
  workloads: `38/38` public and `8/8` holdout.
- The public benchmark evidence note explains the measured workload families,
  engine metadata, exclusions, and the current p95 nuance:
  [`docs/headless-workpaper-benchmark-evidence.md`](../../docs/headless-workpaper-benchmark-evidence.md).
- Recently fixed and hardened P1 risks are covered by regression tests:
  - `updateConfig()` now applies `useColumnIndex` correctly when a rebuild-only
    config key changes in the same update.
  - `moveCells()` now respects `maxRows` and `maxColumns` and rejects target
    ranges outside configured bounds.
  - persisted WorkPaper documents reject sparse arrays, duplicate sheet names,
    non-JSON numeric values, invalid scoped named expressions, and invalid
    persisted config before restore.

Supported scope:

- The contract is the WorkPaper/headless API exported by this package.
- Excel-file ingestion belongs to import/export pipelines before data reaches
  `WorkPaper`; this package executes the validated WorkPaper workbook model.
- Custom function plugins and callback hooks are runtime registrations. Persist
  workbook data with the helpers below, then register custom behavior in
  application code before restore.

## Requirements

- Node `24+`
- ESM imports
- `pnpm@10.32.1` for this monorepo

## Install

Published package:

```sh
pnpm add @bilig/headless
```

Repository:

- GitHub: <https://github.com/proompteng/bilig>
- npm: <https://www.npmjs.com/package/@bilig/headless>
- runnable example:
  [`examples/headless-workpaper`](../../examples/headless-workpaper)

Inside this monorepo:

```sh
pnpm install
pnpm --filter @bilig/headless build
```

## Quickstart

```ts
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
  type WorkPaperCellAddress,
} from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets(
  {
    Sheet1: [
      [10, 20, '=A1+B1'],
      [7, '=A2*3', null],
    ],
  },
  {
    maxRows: 1_000,
    maxColumns: 100,
    useColumnIndex: true,
  },
)

const sheet = workbook.getSheetId('Sheet1')
if (sheet === undefined) {
  throw new Error('Sheet1 was not created')
}

const at = (row: number, col: number): WorkPaperCellAddress => ({
  sheet,
  row,
  col,
})

console.log(workbook.getCellValue(at(0, 2))) // CellValue for 30

workbook.setCellContents(at(1, 2), '=A2+B2')
console.log(workbook.getCellFormula(at(1, 2))) // "=A2+B2"
console.log(workbook.getCellSerialized(at(1, 2))) // "=A2+B2"

const document = exportWorkPaperDocument(workbook)
const json = serializeWorkPaperDocument(document)
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(json))
const restoredSheet = restored.getSheetId('Sheet1')
if (restoredSheet === undefined) {
  throw new Error('Sheet1 was not restored')
}

console.log(restored.getCellValue({ sheet: restoredSheet, row: 1, col: 2 }))
```

## Runnable Example

The repo includes a small external-consumer project at
[`examples/headless-workpaper`](../../examples/headless-workpaper). It builds a
revenue workbook, evaluates formulas, applies an agent-style edit, persists and
restores the workbook, and verifies the final result.

```sh
cd examples/headless-workpaper
npm install
npm start
```

Repository CI also runs the same example against packed local runtime packages
through `pnpm workpaper:smoke:external`.

## Core Concepts

- `WorkPaper` is the top-level workbook object. Create it with
  `WorkPaper.buildEmpty()`, `WorkPaper.buildFromArray()`, or
  `WorkPaper.buildFromSheets()`.
- Addresses are zero-based `{ sheet, row, col }` objects. Use `getSheetId(name)`
  to resolve the numeric sheet id.
- A string beginning with `=` is a formula. Other strings are literal text.
  `null` clears a cell.
- `getCellValue()` returns a computed `CellValue`.
- `getCellFormula()` returns the formula text when the cell is a formula.
- `getCellSerialized()` returns the persisted cell input shape.
- Mutation methods return WorkPaper change arrays. Empty arrays are valid when
  evaluation is suspended or a batch defers change publication.

## Common API Recipes

Create an empty workbook:

```ts
const workbook = WorkPaper.buildEmpty({
  maxRows: 10_000,
  maxColumns: 256,
})
```

Set values, formulas, and blanks:

```ts
workbook.setCellContents(at(0, 0), 42)
workbook.setCellContents(at(0, 1), 'label')
workbook.setCellContents(at(0, 2), '=A1*2')
workbook.setCellContents(at(0, 3), null)
```

Read ranges and whole sheets:

```ts
const range = { start: at(0, 0), end: at(9, 4) }

workbook.getRangeValues(range)
workbook.getRangeFormulas(range)
workbook.getRangeSerialized(range)
workbook.getSheetValues(sheet)
workbook.getSheetSerialized(sheet)
```

Batch related edits:

```ts
const changes = workbook.batch(() => {
  workbook.setCellContents(at(0, 0), 10)
  workbook.setCellContents(at(0, 1), '=A1*5')
})
```

Move cells inside configured bounds:

```ts
const source = { start: at(0, 0), end: at(0, 1) }

workbook.moveCells(source, at(2, 0))
```

`moveCells()` throws when the source or target falls outside `maxRows` or
`maxColumns`.

Update runtime config:

```ts
workbook.updateConfig({
  useColumnIndex: false,
  maxRows: 2_000,
})
```

`updateConfig()` preserves workbook state. Some config keys can trigger an
internal engine rebuild; callers should treat the returned workbook object as the
same public facade and read `getConfig()` after update when they need the active
configuration.

Persist and restore:

```ts
const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))

const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved))
```

## Persistence Contract

`@bilig/headless` persists:

- ordered sheet content arrays
- sheet names
- named expressions
- the JSON-safe subset of `WorkPaperConfig` when `includeConfig` is enabled

It does not persist custom function implementations, callback hooks, or process
state. Register those in code before creating or restoring workbooks.

## Validation Commands

For a headless-only change, start with focused tests:

```sh
pnpm exec vitest run \
  packages/headless/src/__tests__/work-paper-runtime.test.ts \
  packages/headless/src/__tests__/work-paper-parity.test.ts \
  packages/headless/src/__tests__/persistence.test.ts \
  packages/headless/src/__tests__/persistence.fuzz.test.ts

pnpm --filter @bilig/headless build
```

Before publishing or claiming production readiness, run the full gates:

```sh
pnpm publish:runtime:check
pnpm workpaper:bench:competitive:check
pnpm run ci
```

Regenerate the competitive artifact only when intentionally updating benchmark
evidence:

```sh
pnpm workpaper:bench:competitive:generate
pnpm workpaper:bench:competitive:check
```

Do not change benchmark definitions, scoring, sampling, or workload sizes to hide
losses.

## Compatibility Notes

- The facade follows HyperFormula-style public workbook workflows, but it is not
  byte-for-byte compatible with HyperFormula.
- Public lookup helpers such as `getSheetId()`, `getSheetName()`,
  `simpleCellAddressFromString()`, `simpleCellRangeFromString()`, and
  named-expression reads return `undefined` on misses.
- `@bilig/headless` keeps `bilig` change arrays and also exposes richer detailed
  event payloads through `onDetailed()`, `onceDetailed()`, and `offDetailed()`.
- Stable compatibility adapters are available through `graph`, `rangeMapping`,
  `arrayMapping`, `sheetMapping`, `addressMapping`, `dependencyGraph`,
  `evaluator`, `columnSearch`, and `lazilyTransformingAstService`.

## For Coding Agents

Use this checklist when Codex, Claude Code, or another agent starts work here:

1. Read this README and the root `README.md` first.
2. Use public exports from `@bilig/headless`; do not import from `src/`,
   `dist/internal`, or `@bilig/core` unless the task explicitly requires engine
   integration work.
3. Use zero-based `{ sheet, row, col }` addresses and resolve sheet ids with
   `getSheetId()`.
4. Prefer `WorkPaper.buildFromSheets()` for fixtures and
   `exportWorkPaperDocument()` / `createWorkPaperFromDocument()` for persistence
   round trips.
5. Add or tighten regression tests before changing behavior around config
   rebuilds, range bounds, formulas, persistence, events, row/column moves, or
   sheet lifecycle.
6. Run the focused headless tests before broader gates.
7. Preserve benchmark definitions and workload sizes. Performance improvements
   belong in production engine/headless code.
8. Document unsupported behavior honestly instead of implying full Excel
   compatibility.

## Public Entry Points

The package root exports:

- `WorkPaper`
- WorkPaper address, range, config, sheet, change, event, and adapter types
- WorkPaper error classes
- persistence helpers:
  - `exportWorkPaperDocument()`
  - `createWorkPaperFromDocument()`
  - `serializeWorkPaperDocument()`
  - `parseWorkPaperDocument()`
  - `isPersistedWorkPaperDocument()`
  - `pickPersistableWorkPaperConfig()`

## Versioning

`@bilig/headless` ships as part of the aligned bilig runtime package set. Treat
documented public exports as the supported surface, keep integration tests around
your own workbook corpus, and rerun the validation gates before upgrading in
production.
