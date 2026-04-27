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

This document first scoped the end-to-end implementation to the validated
sliding aggregate patch. The current checkout has since advanced into the next
validated tranches from the broader oracle queue: formula-template build
materialization, runtime-restore warmness, and direct scalar mutation
constant-factor work. Benchmark definitions, sample counts, workload sizes, and
scoring logic remain unchanged.

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

Current status of the broader oracle queue:

1. Batch mutation owner coalescing for `batch-edit-*` and `batch-suspended-*`
   is partially implemented through direct scalar delta skips and algebraic
   delta evaluation, but the public change-materialization cost still keeps the
   family red.
2. Formula-family build materialization for parser-template and mixed-content
   build rows is partially implemented through simple direct scalar compilation,
   direct-scalar dependency binding, formula-family recent-family caching, and
   monotonic run append.
3. Runtime snapshot warmness v2 for `rebuild-runtime-from-snapshot` is partially
   implemented through trusted template fast compilation and prior runtime image
   allocation work, but the workload remains red.
4. Lookup direct-eval and after-write constant-factor trimming remains
   incomplete. The measured lookup rows are already using
   `directFormulaKernelSyncOnlyRecalcSkips` or `kernelSyncOnlyRecalcSkips`, so
   the remaining gap is public mutation overhead and lookup direct-eval
   constant factors.
5. Dirty-chain frontier counters are present, and direct scalar closure handles
   simple chains, but dirty-execution is not yet a clean win across all rows.

Each item needs a fresh source and benchmark validation pass before becoming an
implementation document.

## Additional Implemented Tranche: Build And Scalar Mutation

The following changes were validated against the current checkout after the
initial sliding-aggregate work:

- Added `tryCompileSimpleDirectScalarFormula` for common row-local formulas such
  as `A1+B1`, `C1*2`, and translated row-template variants. This avoids parser
  and generic AST translation work for simple scalar families.
- Threaded direct scalar operands into formula binding so dependency entities
  and symbolic cell bindings are materialized from one already-resolved operand
  list for same-sheet direct scalar formulas.
- Kept qualified cross-sheet scalar dependencies on the existing recalc path so
  cross-sheet rebind behavior and metrics stay compatible.
- Avoided repeated pending-WASM-sync writes while a formula initialization batch
  is already open.
- Added formula-family recent-template caching and a monotonic row-run append
  fast path for repeated template families down a column.
- Collapsed initial `buildFromSheets` validation and formula-presence scanning
  into a single inspection pass, then routed known literal sheets directly into
  the literal loader and known mixed sheets directly into the mixed loader.
- Added algebraic direct scalar delta calculation for simple scalar descriptors
  before falling back to generic old/new formula evaluation.
- Reused a process-level column-label cache when materializing tracked public
  changes, reducing repeated A1-label construction for wide fanout events.

## Additional Implemented Tranche: Current Scalar And Small Events

The `2026-04-27` heartbeat pass added two production-only changes and rejected
two measured dead ends:

- Direct scalar formulas in the post-recalc direct-formula loop can now be
  reapplied from current cell-store operands without routing through the generic
  direct formula evaluator when no numeric delta is available.
- The no-visibility tracked-event path now materializes tiny one-to-four-cell
  events directly from cell indices instead of first building a generic tracked
  event change payload. This targets lookup and sliding aggregate edits without
  changing event semantics.
- A broader exact/approximate lookup evaluator bypass was tried and removed
  because the approximate rows became slower in focused probes and the full
  suite did not hold the improvement.
- An initial-load numeric write/index deferral was tried and removed because it
  moved cost into the measured sliding aggregate mutation path.

## Verification Commands

Targeted test command:

```sh
bun scripts/run-vitest.ts --run packages/core/src/__tests__/operation-service.test.ts packages/core/src/__tests__/formula-binding-service.test.ts packages/core/src/__tests__/formula-evaluation-service.test.ts packages/core/src/__tests__/engine.test.ts packages/core/src/__tests__/formula-family-store.test.ts packages/headless/src/__tests__/initial-sheet-load.test.ts packages/headless/src/__tests__/work-paper-runtime.test.ts
```

Competitive benchmark command:

```sh
pnpm --silent bench:workpaper:competitive > /tmp/workpaper-competitive-after-delta-and-tracking.json
pnpm --silent bench:workpaper:competitive > /tmp/workpaper-competitive-after-scalar-small-event-reduced.json
```

## Observed Result

The latest local sample on the reduced current tree produced a `38` comparable
workload scorecard with `21` WorkPaper wins and `17` HyperFormula wins. This is
not complete and does not satisfy the SOTA target, but it is a real improvement
over the prior committed `17` / `21` scorecard. Latest red rows include:

- `aggregate-overlapping-sliding-window`: `0.120 ms` WorkPaper mean, still red
  at `3.368x` slower than HyperFormula.
- `lookup-with-column-index`: `0.136 ms` WorkPaper mean, still red at `3.071x`
  slower.
- `lookup-with-column-index-after-column-write`: `0.066 ms` WorkPaper mean,
  still red at `2.208x` slower.
- `lookup-approximate-sorted`: `0.060 ms` WorkPaper mean, still red at `2.186x`
  slower.
- `batch-edit-recalc`: `1.144 ms` WorkPaper mean, still red at `1.912x` slower.
- `build-parser-cache-row-templates`: `45.845 ms` WorkPaper mean, still red at
  `1.828x` slower.
- `rebuild-runtime-from-snapshot`: `41.769 ms` WorkPaper mean, still red at
  `1.386x` slower.

The sliding aggregate row remains red, but its WorkPaper mean improved
materially from the earlier local samples:

- pre-change local sample: about `0.226 ms`
- after direct delta recalc skip: about `0.194 ms`
- after dependent collection and prefix eviction work: best observed sample
  about `0.124 ms`
- latest `8` sample run after the tiny merge fast path: about `0.141 ms`
- latest unchanged competitive harness sample after the build/scalar tranche:
  about `0.124 ms`
- latest reduced current-tree sample after direct scalar current-value and
  small-event materialization: about `0.120 ms`

The row is now counter-gated: `directAggregateDeltaApplications = 1` and
`directAggregateDeltaOnlyRecalcSkips = 1` for the sliding mutation sample, with
no direct aggregate scan/prefix evaluation during the measured mutation.

The next required implementation work is still structural, not benchmark
tuning: reduce public change-materialization overhead for batch/fanout rows,
make runtime restore avoid remaining per-formula binding churn, and continue
collapsing formula-family/template build cost.
