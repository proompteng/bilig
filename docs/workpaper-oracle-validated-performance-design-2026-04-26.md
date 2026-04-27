# WorkPaper Oracle Performance Design, Validated Against Current Code

Date: `2026-04-26`
Oracle thread: `https://chatgpt.com/c/69ed8169-1d94-83e8-bfc2-a34c22558617`
Local validation commit before implementation: `d7cbd690c3710e15e1735c0971d4f97eda9fbf72`

## Oracle Capture

The oracle response was complete and usable. It analyzed the attached
`bilig2-codebase-current(1).zip`, `workpaper-competitive-latest.json`,
`repo-state.txt`, and `oracle-cleanroom-prompt.md`.

The oracle's key conclusion was that lookup splitting should not be the next
first patch. Current source already has exact/approximate lookup fast paths and
the missing evidence is constant-factor counters. Its proposed small patch was
sliding aggregate prefix promotion: the 32-row sliding aggregate formulas should
not repeatedly scan cell ranges when reusable direct aggregate machinery exists.

## Current Checkout Validation

The current checkout is newer than the oracle attachment. A pre-change local
sample showed the expanded benchmark now has `38` comparable workloads, not the
oracle's `34`, with a `19` WorkPaper / `19` HyperFormula split. The same high
priority red families remain: build, runtime restore, batch edit, lookup,
after-write lookup, and sliding-window aggregate.

Validated source facts:

- `packages/core/src/engine/services/formula-evaluation-service.ts` still had
  `DIRECT_AGGREGATE_SCAN_MAX_LENGTH = 64`, so `SUM(A1:A32)` and shifted
  `SUM(A2:A33)` windows scanned 32 cells per formula.
- `packages/core/src/deps/aggregate-state-store.ts` already provides reusable
  prefix buffers with incremental extension and literal-write updates.
- `packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts`
  still defines `aggregate-overlapping-sliding-window` with `window: 32`.
- The actual sliding benchmark measures a literal mutation after build:
  `packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded-additional-workloads.ts`
  calls `workbook.setCellContents(address(sheetId, 0, 0), 99)`. After adding
  direct-evaluation counters, the benchmark showed zero direct scan/prefix
  evaluations for this row. That means the measured hot path is operation-time
  direct aggregate delta handling, not formula-evaluation scans.
- `packages/core/src/engine/services/operation-service.ts` already computes
  numeric deltas for pure direct `SUM` aggregate dependents, but still calls the
  kernel-sync recalc path before applying those deltas.
- `packages/core/src/deps/region-graph.ts` materialized point-query matches
  through `Set` allocations even when only one region and one dependent match.
- `operation-service.ts` converted the returned `Uint32Array` to a JavaScript
  array before filtering direct range dependents, adding avoidable allocation in
  the single-dependent sliding case.
- `operation-service.ts` merged post-recalc direct formula changes through a
  `Set` even when the base changed set was empty and the direct aggregate delta
  produced a single changed formula cell.
- Lowering 32-row formulas into the prefix cache made build-time evaluation
  cheaper, but it also exposed a mutation cost: `aggregate-state-store.ts`
  updated every suffix prefix value after an early-row write. For the benchmark's
  `A1` edit, that turns one literal mutation into a 1,500-entry prefix update
  even though the direct aggregate formula value is already handled by a numeric
  delta.
- The previous local implementation had already worked on owner-backed uniform
  approximate lookup, so repeating the oracle's lookup discussion would not be
  the right next design target.

## Implemented Design

This document scopes the end-to-end implementation to the validated sliding
aggregate patch. The broader oracle queue remains useful background, but it is
not treated as implementation truth until each tranche is revalidated against
the current checkout and current benchmark rows.

### 1. Add Direct Aggregate Counters

Add direct evaluation counters:

- `directAggregateScanEvaluations`
- `directAggregateScanCells`
- `directAggregatePrefixEvaluations`

Add operation-time delta counters:

- `directAggregateDeltaApplications`
- `directAggregateDeltaOnlyRecalcSkips`

Purpose: make the aggregate paths counter-gated. Evaluation tests must prove
32-row formulas move from scans to prefix evaluation. The sliding mutation
benchmark must prove it used direct aggregate delta application and skipped the
unnecessary recalc path.

### 2. Preserve Tiny Window Scan Behavior

Keep direct scans for:

- formulas with scalar dependencies
- aggregate ranges with length `16` or below
- existing unsupported or semantics-sensitive fallback cases

This preserves the no-column-owner behavior for genuinely tiny aggregate
windows and keeps the simple path cheap.

### 3. Promote SUM, COUNT, and AVERAGE Windows Above 16 Rows

For direct aggregate formulas with no scalar dependencies:

- route `SUM`, `COUNT`, and `AVERAGE` ranges longer than `16` rows through the
  shared prefix path
- continue to share a lower prefix start for shifted windows so `SUM(A1:A32)`,
  `SUM(A2:A33)`, and later windows reuse one prefix buffer
- preserve the existing large-range prefix path for other aggregate kinds

Correctness invariants:

- numeric, boolean, blank, string, and error behavior must match the old scan or
  generic formula behavior
- shifted windows must return the same results as unshifted windows of the same
  values
- small windows must still scan and avoid building column owners
- benchmark verification for `aggregate-overlapping-sliding-window` must remain
  unchanged

### 4. Tests

Required tests:

- engine counters initialize, clone, merge, and reset the new aggregate counters
- a 16-row aggregate still scans, records scan counters, and avoids column-owner
  construction
- 32-row `SUM`, shifted `SUM`, `COUNT`, and `AVERAGE` formulas use prefix
  evaluation counters and produce correct values
- single-cell and generic-batch mutations against a 32-row direct `SUM`
  aggregate apply a numeric delta, skip dirty traversal, and skip recalc when
  every post-recalc direct formula is covered by a numeric delta
- region-graph point lookups preserve dependent deduplication when one formula
  subscribes to multiple matching regions
- aggregate prefix state evicts early-row large-suffix writes instead of doing a
  long in-place prefix update

### 5. Remove Avoidable Hot-Path Allocation

For the sliding benchmark's single matching range, region graph collection should
avoid building two `Set` objects. The implementation collects matching region
IDs into a small array, returns the single subscriber set directly when only one
region matches, and only performs explicit dependent deduplication when multiple
regions match.

The operation service also loops over the returned typed dependent list directly
instead of spreading it into an array and then filtering.

### 6. Evict Expensive Prefix Suffix Updates

When a literal write would require updating more than `128` prefix entries, the
aggregate state store evicts that prefix entry rather than applying an O(n)
suffix delta. The direct mutation path still applies the formula-value delta for
currently affected formulas, so correctness is preserved; future formula
evaluation rebuilds the prefix lazily from current column state.

### 7. Fast-Path Tiny Changed-Set Merges

The direct delta path commonly merges an empty recalculated set with one formula
cell. `mergeChangedCellIndices` handles empty and one-plus-one cases directly
before falling back to a `Set`.

## Broader Oracle Queue

These remain design leads, not implemented requirements in this document:

1. Batch mutation owner coalescing for `batch-edit-*` and `batch-suspended-*`.
2. Formula-family build materialization for parser-template and mixed-content
   build rows.
3. Runtime snapshot warmness v2 for `rebuild-runtime-from-snapshot`.
4. Lookup direct-eval and after-write constant-factor trimming after counters
   prove the hot branch.
5. Dirty-chain frontier counters before changing scheduler behavior.

Each item needs a fresh source and benchmark validation pass before becoming an
implementation document.

## Verification Commands

Targeted test command:

```sh
bun scripts/run-vitest.ts --run packages/core/src/__tests__/formula-evaluation-service.test.ts packages/core/src/__tests__/engine-counters.test.ts
```

Competitive benchmark command:

```sh
pnpm --silent bench:workpaper:competitive -- --sample-count 8 --warmup-count 2 > /tmp/workpaper-competitive-after-merge-fastpath.json
```

## Observed Result

The latest local sample after this implementation produced a `38` comparable
workload scorecard with `20` WorkPaper wins and `18` HyperFormula wins. The
sliding aggregate row remained red, but its WorkPaper mean improved materially
from the earlier local samples:

- pre-change local sample: about `0.226 ms`
- after direct delta recalc skip: about `0.194 ms`
- after dependent collection and prefix eviction work: best observed sample
  about `0.124 ms`
- latest `8` sample run after the tiny merge fast path: about `0.141 ms`

The row is now counter-gated: `directAggregateDeltaApplications = 1` and
`directAggregateDeltaOnlyRecalcSkips = 1` for the sliding mutation sample, with
no direct aggregate scan/prefix evaluation during the measured mutation.
