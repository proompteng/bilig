# WorkPaper Oracle Performance Design, Validated Against Current Code

Date: `2026-04-26`
Oracle thread: `https://chatgpt.com/c/69ee60bf-bb6c-83e8-9422-e258032e0df0`
Attached artifacts observed in the thread:

- `bilig2-codebase-current(2).zip`
- `workpaper-competitive-latest(1).json`
- `repo-state(1).txt`

Local copies used for validation:

- `/tmp/bilig2-atlas-oracle/bilig2-codebase-current.zip`
- `/tmp/bilig2-atlas-oracle/workpaper-competitive-latest.json`
- `/tmp/bilig2-atlas-oracle/repo-state.txt`

## Oracle Capture

The Oracle thread was readable, and the attachments were visible in ChatGPT. As of the capture, the assistant response had not produced the requested design. The only visible answer text was:

> I will verify the required files, derive priorities from the benchmark JSON, and tie each proposed patch to concrete source inspections rather than speculation.

The page still showed `Pro thinking` / `Stop streaming`. There was no completed benchmark ranking, source-backed root-cause list, patch queue, or first-PR spec to transcribe. This document therefore records the incomplete Oracle capture and replaces it with a local design validated directly against the attached benchmark JSON and the current repository source.

## Current Benchmark Truth

The attached benchmark scorecard has `34` comparable workloads: WorkPaper wins `18`, HyperFormula wins `16`.

HyperFormula wins, sorted by reported mean speedup:

| Rank | Workload                                       |    Ratio | WorkPaper median ms | HyperFormula median ms | Source-backed cause                                                                                                              |
| ---- | ---------------------------------------------- | -------: | ------------------: | ---------------------: | -------------------------------------------------------------------------------------------------------------------------------- |
| 1    | `lookup-approximate-sorted`                    | `3.635x` |             `0.089` |                `0.037` | Owner-backed approximate lookup carries uniform metadata but still binary-searches numeric owner arrays.                         |
| 2    | `aggregate-overlapping-sliding-window`         | `2.386x` |             `0.129` |                `0.054` | Source inspection shows shared-prefix and direct-delta paths already exist; remaining loss needs counters before a design claim. |
| 3    | `build-parser-cache-row-templates`             | `2.329x` |            `64.288` |               `27.294` | Parse templates are reused, but formula install still allocates/binds per formula family member and uploads full WASM state.     |
| 4    | `lookup-approximate-sorted-after-column-write` | `2.044x` |             `0.100` |                `0.048` | Same owner-backed approximate binary-search path, with refreshed owner summaries after writes.                                   |
| 5    | `build-mixed-content`                          | `2.017x` |            `17.308` |                `8.237` | Build path still pays formula installation and full WASM upload overhead.                                                        |
| 6    | `batch-edit-single-column`                     | `1.871x` |             `1.055` |                `0.708` | Mutation observers are still invoked at cell/edit granularity for hot batch paths.                                               |
| 7    | `lookup-with-column-index-after-column-write`  | `1.868x` |             `0.103` |                `0.059` | Exact lookup impact checks are narrow, but the remaining after-write overhead still needs counters before changing owner logic.  |
| 8    | `rebuild-runtime-from-snapshot`                | `1.639x` |            `54.908` |               `33.748` | Runtime image is not a full warm semantic runtime; rebuild still performs broad initialization and full WASM upload.             |
| 9    | `batch-suspended-multi-column`                 | `1.633x` |             `0.867` |                `0.632` | Batch edits need owner-scoped coalescing and cheaper changed payload construction.                                               |
| 10   | `batch-edit-multi-column`                      | `1.511x` |             `0.889` |                `0.608` | Same batch mutation coalescing issue across columns.                                                                             |
| 11   | `build-many-sheets`                            | `1.510x` |             `9.538` |                `6.365` | Sheet/runtime initialization still has per-sheet fixed overhead.                                                                 |
| 12   | `batch-edit-single-column-with-undo`           | `1.463x` |             `1.439` |                `1.014` | Undo capture and mutation fanout need coalesced owner deltas.                                                                    |
| 13   | `partial-recompute-mixed-frontier`             | `1.160x` |             `3.289` |                `3.439` | Column-owner rebuild still appears once in the mixed frontier path.                                                              |
| 14   | `lookup-with-column-index`                     | `1.152x` |             `0.077` |                `0.064` | Exact lookup is close; remaining loss is likely call overhead and owner refresh checks.                                          |
| 15   | `single-edit-chain`                            | `1.083x` |             `1.131` |                `1.135` | Near tie; do not prioritize until higher-ratio lanes move.                                                                       |
| 16   | `build-parser-cache-mixed-templates`           | `1.022x` |            `73.728` |               `71.230` | Near tie; same build-family overhead, but low priority.                                                                          |

## Validated Source Findings

### 1. First patch: uniform owner-backed approximate lookup

`packages/core/src/engine/services/lookup-column-owner.ts` computes uniform approximate summaries:

- `summarizeApproximateRange(...)`
- `detectUniformNumericStepInOwner(...)`
- `supportsNumericApproximateRange(...)`

`packages/core/src/engine/services/sorted-column-search-service.ts` preserves that summary in `prepareVectorLookup(...)` as `uniformStart` and `uniformStep`.

The non-owner fallback branch in `findPreparedVectorMatch(...)` already uses uniform arithmetic to return a 1-based approximate match position without binary search. The owner-backed branch does not; it always binary-searches `owner.numericValues` for numeric and empty lookup values. That is the cleanest validated gap in the top red workload.

Invariant: only use the arithmetic shortcut when the existing summary proves a supported numeric approximate range and the uniform step direction matches `matchMode`. Otherwise retain the existing binary search.

Expected movement: primary improvement in `lookup-approximate-sorted` and `lookup-approximate-sorted-after-column-write`, with minimal semantic risk.

### 2. Sliding aggregate state is the next measured hot path

`packages/core/src/deps/aggregate-state-store.ts` owns prefix aggregate entries and update propagation. `packages/core/src/engine/services/formula-evaluation-service.ts` already routes longer `SUM`/`AVERAGE`/`COUNT` ranges through a shared prefix start, and `packages/core/src/engine/services/operation-service.ts` already has direct numeric-delta handling for small direct-aggregate fanout. The current red workload has no nonzero engine counters, so the next change should add focused counters around region-graph dependent collection, direct-delta application, changed payload construction, and verification reads before changing the data structure.

Expected movement: `aggregate-overlapping-sliding-window`.

### 3. Build lanes are install/runtime-owner problems, not parser-only problems

The benchmark counters for row-template and mixed-template build lanes show very low `formulasParsed` counts, which means parser caching is already active. The remaining work is formula install, family ownership, dependency graph registration, runtime image warmness, and full WASM upload reduction.

Expected movement: `build-parser-cache-row-templates`, `build-mixed-content`, `rebuild-runtime-from-snapshot`, and the near-tie mixed-template lane.

### 4. Batch edit lanes need owner-scoped coalescing

The batch edit losses have no nonzero engine counters in the benchmark artifact, so they should not be guessed from counters alone. The correct next step is to inspect mutation fanout around write batching, changed-cell payload construction, lookup owner write updates, and undo capture, then add counters that prove which observer dominates.

Expected movement: `batch-edit-*` and `batch-suspended-*`.

## Patch Queue

1. Owner-backed uniform approximate lookup arithmetic.
   - Files: `sorted-column-search-service.ts`, `sorted-column-search-service.test.ts`.
   - Tests: owner-backed ascending and descending uniform prepared lookups; fallback branch still covered.
   - Benchmarks: `pnpm bench:workpaper:competitive`.
   - First result: implemented. Targeted tests pass. A post-change competitive run produced `21` WorkPaper wins and `13` HyperFormula wins. The base approximate lookup lane improved versus the attached baseline but still remained red in that sample, so lookup work is directionally valid but not complete.

2. Aggregate sliding-window instrumentation before structural changes.
   - Files: `aggregate-state-store.ts`, aggregate tests, benchmark fixture if needed.
   - Tests: overlapping SUM/AVERAGE/COUNT windows and literal write invalidation.
   - Benchmarks: `aggregate-overlapping-sliding-window`.

3. Batch mutation owner coalescing.
   - Files: mutation/write services and owner index stores after source inspection.
   - Tests: single-column, multi-column, suspended batch, undo capture.
   - Benchmarks: all `batch-edit-*` lanes.

4. Exact lookup after-write refresh trimming.
   - Files: `lookup-column-owner.ts`, `exact-column-index-service.ts`, `column-index-store.ts`.
   - Tests: exact duplicate-key update, type transitions, after-write prepared lookup reuse.
   - Benchmarks: exact lookup and exact after-write lanes.

5. Build family install compaction.
   - Files: formula binding, family store, dependency graph registration.
   - Tests: row-template families, mixed-template families, formula correctness.
   - Benchmarks: build parser-template and mixed-content lanes.

6. Runtime image warm semantic sections.
   - Files: snapshot/runtime image and formula initialization.
   - Tests: restore parity and benchmark restore.
   - Benchmarks: `rebuild-runtime-from-snapshot`.

## Do Not Do

- Do not special-case benchmark names or fixture dimensions.
- Do not move semantics into WASM before JS parity and differential tests prove the closed numeric family.
- Do not restart a broad structural-row rewrite; the current red lanes are lookup, aggregate, build, batch, and restore.
- Do not optimize parser caching first: the counters already show parser reuse is working in the main build-template losses.
- Do not accept the Oracle response as authoritative until it finishes and provides source-backed content.

## Execution Notes

Implemented and validated so far:

- `sorted-column-search-service.ts` now shares uniform numeric approximate-position arithmetic across fallback and owner-backed prepared lookup paths.
- Owner-backed prepared matching now uses the already-refreshed prepared range kind/sort certificates instead of recomputing `supports...Range` checks during every match.
- `sorted-column-search-service.test.ts` now includes a regression test that prepares owner-backed ascending and descending uniform ranges, poisons the owner numeric arrays, and confirms the uniform metadata path still returns correct approximate positions.

Observed benchmark movement:

- Targeted test command: `bun scripts/run-vitest.ts --run packages/core/src/__tests__/sorted-column-search-service.test.ts`.
- Competitive benchmark command: `pnpm bench:workpaper:competitive`.
- Best post-change scorecard sample so far: `21` WorkPaper wins, `13` HyperFormula wins.
- Remaining high-priority red lanes after that sample: approximate lookup after write, batch suspended multi-column, row-template build, base approximate lookup, exact lookup after write, sliding aggregate, mixed build, runtime restore, build-many-sheets, and a few near-tie batch/build lanes.
