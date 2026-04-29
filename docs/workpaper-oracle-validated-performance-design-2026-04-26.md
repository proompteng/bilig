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

## Additional Implemented Tranche: Subagent Structural Pass

The `2026-04-27` subagent pass moved several broader oracle items from
investigation into production code while keeping benchmark definitions and
sampling unchanged:

- Runtime snapshot restore now validates attached snapshot sheet dimensions
  without reading every serialized matrix cell before importing the runtime
  image. This targets `rebuild-runtime-from-snapshot` and made that row green in
  the latest full samples.
- Formula-family initial registration now has a no-return path and a fresh
  ordered-run registration path, avoiding membership return-array materialization
  and descriptor sorting for common monotonic row families.
- Direct aggregate and direct criteria mutation detection now track aggregate
  column reverse edges, so copied `SUMIF`/`SUMIFS` formulas can use direct
  numeric deltas for aggregate-column writes instead of dirty traversal.
- Post-recalc direct formula changed-cell buffering now uses typed buffers in
  the common direct-delta path, with a separate rare spill/extra-changed list.
- Lookup column owners now skip exact row-list remove/reinsert work when a write
  preserves the normalized lookup key and can defer approximate summary
  maintenance for exact-only writes.
- Aggregate prefix state now marks large early-row suffix updates stale and
  rebuilds lazily on the next prefix evaluation instead of eagerly walking long
  prefix suffixes during mutation.
- Headless tracked event capture records sorted/disjoint changed-index metadata,
  and the no-visibility runtime path can use the multi-source materializer when
  event ordering and dedupe semantics are provably safe.
- Direct scalar numeric writes now avoid generic old/new `CellValue`
  materialization for common numeric edits, and constant direct-scalar deltas can
  be applied in bulk during the post-recalc loop.
- Direct-only dirty-traversal skips no longer prepare region-query indices for
  changed input columns that have no range or aggregate subscribers.
- The simple direct scalar compiler now covers `ABS(cell)` in addition to binary
  direct scalar formulas, avoiding parser/binder churn for that mixed-template
  formula family.

This tranche is still incomplete. The current red families prove there is
remaining production work in lookup constants, sliding aggregate mutation
constants, multi-column batch direct deltas, and parser-template build cost.

## Additional Implemented Tranche: Integrated Build/Template Initialization

The integrated `2026-04-27` pass made the build/template slice production-fast
enough to clear the currently measured build rows without changing benchmark
definitions or sampling:

- `SpreadsheetEngine` exposes the already-existing synchronous formula
  initializer as `initializeCellFormulasAtNow`, and the initial mixed-sheet
  loader calls it during `WorkPaper.buildFromSheets`. This avoids routing
  initial formula hydration through the generic Effect wrapper.
- The template bank now tries row-relative scalar template keys before anchored
  aggregate template matching and compiles scalar direct formulas before
  aggregate direct formulas. This removes aggregate-probe overhead from the
  common scalar row-template families.
- Initial direct formula batches skip dynamic-range sync when the range
  registry is empty and skip region-query index preparation when the fresh
  batch can be evaluated entirely through the initial direct formula path.
- The mixed-sheet loader avoids trimming ordinary strings and direct `=...`
  formulas, keeps leading-space formula support, and uses a no-op formula
  rewrite callback when the workbook has no named expressions or function
  aliases.
- A concurrent direct-scalar batch fast-path compile blocker in
  `operation-service.ts` was fixed by narrowing the already-validated mutation
  union before reading the literal value.

This tranche is still not the overall benchmark target. It gets the eligible
build rows green in the latest full sample, but dirty edit, lookup, and sliding
aggregate rows remain red.

## Verification Commands

Targeted test command:

```sh
bun scripts/run-vitest.ts --run packages/core/src/__tests__/operation-service.test.ts packages/core/src/__tests__/formula-binding-service.test.ts packages/core/src/__tests__/formula-evaluation-service.test.ts packages/core/src/__tests__/engine.test.ts packages/core/src/__tests__/formula-family-store.test.ts packages/headless/src/__tests__/initial-sheet-load.test.ts packages/headless/src/__tests__/work-paper-runtime.test.ts
bun scripts/run-vitest.ts --run packages/core/src/__tests__/formula-family-store.test.ts packages/core/src/__tests__/lookup-column-owner.test.ts packages/core/src/__tests__/exact-column-index-service.test.ts packages/core/src/__tests__/sorted-column-search-service.test.ts packages/core/src/__tests__/aggregate-state-store.test.ts packages/core/src/__tests__/operation-service.test.ts packages/headless/src/__tests__/tracked-cell-index-changes.test.ts packages/headless/src/__tests__/tracked-engine-event-refs.test.ts packages/headless/src/__tests__/work-paper-runtime.test.ts
```

Competitive benchmark command:

```sh
pnpm --silent bench:workpaper:competitive > /tmp/workpaper-competitive-after-delta-and-tracking.json
pnpm --silent bench:workpaper:competitive > /tmp/workpaper-competitive-after-scalar-small-event-reduced.json
pnpm --silent bench:workpaper:competitive > /tmp/workpaper-competitive-after-abs-region-guard-full.json
pnpm --silent bench:workpaper:competitive > /tmp/workpaper-competitive-after-init-entrypoint-full.json
```

## Observed Result

The latest full local sample is
`/tmp/workpaper-competitive-after-init-entrypoint-full.json`. It produced a
`38` comparable workload scorecard with `29` WorkPaper wins and `9`
HyperFormula wins. This is not complete and does not satisfy the SOTA target.

Build rows in that sample:

- `build-mixed-content`: WorkPaper `8.676 ms`, HyperFormula `10.027 ms`,
  `1.156x` faster.
- `build-parser-cache-row-templates`: WorkPaper `28.108 ms`, HyperFormula
  `29.214 ms`, `1.039x` faster.
- `build-parser-cache-mixed-templates`: WorkPaper `31.225 ms`, HyperFormula
  `72.611 ms`, `2.325x` faster.

Latest red rows:

- `single-edit-recalc`: WorkPaper `1.691 ms`, HyperFormula `1.547 ms`,
  `1.093x` slower.
- `batch-edit-multi-column`: WorkPaper `0.605 ms`, HyperFormula `0.558 ms`,
  `1.085x` slower.
- `batch-edit-single-column-with-undo`: WorkPaper `1.299 ms`, HyperFormula
  `1.071 ms`, `1.213x` slower.
- `structural-delete-columns`: WorkPaper `6.125 ms`, HyperFormula `5.843 ms`,
  `1.048x` slower.
- `aggregate-overlapping-sliding-window`: WorkPaper `0.469 ms`, HyperFormula
  `0.066 ms`, `7.137x` slower.
- `lookup-with-column-index`: WorkPaper `0.114 ms`, HyperFormula `0.063 ms`,
  `1.813x` slower.
- `lookup-with-column-index-after-column-write`: WorkPaper `0.087 ms`,
  HyperFormula `0.061 ms`, `1.408x` slower.
- `lookup-approximate-sorted`: WorkPaper `0.107 ms`, HyperFormula `0.074 ms`,
  `1.441x` slower.
- `lookup-approximate-sorted-after-column-write`: WorkPaper `0.082 ms`,
  HyperFormula `0.054 ms`, `1.507x` slower.

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
tuning: finish lookup direct-eval/owner constants, collapse the remaining
sliding aggregate mutation overhead, and finish multi-column batch direct-delta
materialization.

## 2026-04-27 Continuation Status

The oracle plan is still not complete. The implementation remains constrained
to production engine/headless paths; benchmark definitions, workload sizes,
sample counts, and scoring have not been changed.

Additional production changes validated in this pass:

- Runtime-image restore now reuses sheet formula spans and index-aligned runtime
  formula values, making `rebuild-runtime-from-snapshot` green in the full
  sample recorded at `/tmp/workpaper-competitive-after-runtime-restore-spans-full.json`.
- Direct scalar batch edits use numeric result arrays for common all-number
  direct scalar batches before falling back to generic current-result objects.
- Headless tracked changes have a large sorted same-sheet lazy materialization
  path for no-visibility batch events, and the runtime now defers detaching that
  large lazy public array until the next mutation or disposal.
- Lookup column owners skip same-key/no-comparable-change row-list and summary
  maintenance, and row-list updates use binary positioning.
- Region graph single-link storage and direct aggregate reverse-map handling
  remain in place, but a local attempt to make point-impact indexing lazy was
  rejected because it broke cumulative aggregate fanout correctness.
- A local attempt to defer tiny no-listener public changes was rejected because
  it worsened the lookup and sliding rows in the unchanged full suite.

Focused validation currently passing:

```sh
pnpm exec vitest run packages/core/src/__tests__/region-graph.test.ts packages/core/src/__tests__/operation-service.test.ts
pnpm exec vitest run packages/headless/src/__tests__/tracked-cell-index-changes.test.ts packages/headless/src/__tests__/work-paper-runtime.test.ts
pnpm exec vitest run packages/headless/src/__tests__/tracked-cell-index-changes.test.ts packages/headless/src/__tests__/work-paper-runtime.test.ts packages/core/src/__tests__/region-graph.test.ts packages/core/src/__tests__/operation-service.test.ts packages/core/src/__tests__/lookup-column-owner.test.ts packages/core/src/__tests__/exact-column-index-service.test.ts
pnpm exec tsc --noEmit --pretty false -p packages/core/tsconfig.json
pnpm exec tsc -p packages/headless/tsconfig.json --noEmit --pretty false
```

Latest unchanged full competitive artifacts:

- `/tmp/workpaper-competitive-after-headless-lookup-workers.json`: `31`
  scorecard wins, `7` HyperFormula wins.
- `/tmp/workpaper-competitive-after-deferred-batch.json`: `32` scorecard wins,
  `6` HyperFormula wins.
- `/tmp/workpaper-competitive-after-tiny-deferred.json`: `32` scorecard wins,
  `6` HyperFormula wins, but the tiny-deferred experiment worsened lookup and
  sliding and was reverted.
- `/tmp/workpaper-competitive-after-heisenberg-dispatcher.json`: `32`
  scorecard wins, `6` HyperFormula wins. Exact lookup rows are now green, but
  batch, sliding aggregate, and approximate lookup near-ties remain red.

Best current full artifact to continue from:
`/tmp/workpaper-competitive-after-deferred-batch.json`.

Remaining red rows in that artifact:

- `single-edit-fanout`: WorkPaper `1.763 ms`, HyperFormula `1.397 ms`,
  `1.262x` slower; WorkPaper had one large outlier while direct scalar delta
  counters were already clean.
- `aggregate-overlapping-sliding-window`: WorkPaper `0.061 ms`, HyperFormula
  `0.057 ms`, `1.064x` slower; counters show one direct aggregate delta and
  one recalc skip.
- `lookup-with-column-index`: WorkPaper `0.060 ms`, HyperFormula `0.054 ms`,
  `1.111x` slower; counters show direct lookup current-result recalc skip.
- `lookup-with-column-index-after-column-write`: WorkPaper `0.055 ms`,
  HyperFormula `0.050 ms`, `1.093x` slower; counters show kernel-sync-only
  skip.
- `lookup-approximate-sorted`: WorkPaper `0.064 ms`, HyperFormula `0.041 ms`,
  `1.578x` slower; counters show direct lookup current-result recalc skip.
- `lookup-approximate-sorted-after-column-write`: WorkPaper `0.068 ms`,
  HyperFormula `0.037 ms`, `1.837x` slower; counters show kernel-sync-only
  skip.

The next production implementation targets are therefore:

1. Reduce direct lookup operand and after-write constant factors in core
   operation/evaluation/owner paths.
2. Collapse the remaining sliding aggregate mutation/event overhead while
   preserving cumulative aggregate fanout correctness.
3. Reduce outlier allocation/GC risk in direct scalar fanout and multi-column
   batch public-change paths.

After the core no-listener direct lookup dispatcher and same-column aggregate
version-batch shortcut, the latest full artifact
`/tmp/workpaper-competitive-after-heisenberg-dispatcher.json` has these
remaining red rows:

- `build-many-sheets`: WorkPaper `6.660 ms`, HyperFormula `6.605 ms`,
  `1.008x` slower.
- `batch-edit-recalc`: WorkPaper `0.642 ms`, HyperFormula `0.604 ms`,
  `1.062x` slower.
- `batch-edit-multi-column`: WorkPaper `0.553 ms`, HyperFormula `0.527 ms`,
  `1.049x` slower.
- `aggregate-overlapping-sliding-window`: WorkPaper `0.060 ms`, HyperFormula
  `0.058 ms`, `1.039x` slower.
- `lookup-approximate-sorted`: WorkPaper `0.051 ms`, HyperFormula `0.047 ms`,
  `1.086x` slower.
- `lookup-approximate-sorted-after-column-write`: WorkPaper `0.051 ms`,
  HyperFormula `0.049 ms`, `1.038x` slower.

## 2026-04-28 Continuation Status

The oracle plan is still not fully complete because the unchanged full
competitive suite still has at least one red comparable row. The implementation
continues to be limited to production core/headless paths; benchmark workload
definitions, sample counts, fixtures, and scoring remain untouched.

Additional production changes validated in this pass:

- Runtime-image snapshots now carry sheet dimensions and cell counts, and
  runtime-image restore uses a fresh-sheet fast attacher. This avoids coordinate
  scans during `WorkPaper.buildFromSheets` inspection and makes
  `rebuild-runtime-from-snapshot` green in repeated full samples.
- Mixed initial sheet inspection now counts formula cells in the same pass used
  for dimension/materialization checks. The mixed loader preallocates formula
  refs from that count, and formula initialization uses a target-index
  `Uint32Array` plus a smaller pending-formula bitset instead of sizing around
  full cell-store capacity.
- Headless existing-numeric edits now call the engine with
  `emitTracked: false`, a trusted already-validated numeric-literal hint, and
  the old numeric value. Core suppresses tracked event emission for that path
  and can skip duplicate coordinate/formula-map validation after headless has
  checked sheet id, row, column, formula id, overwrite flags, and numeric tag.
- Existing-numeric mutation results now have a compact scalar representation for
  one/two changed cells, including the aggregate formula's new numeric result
  when it is known. Headless uses that compact result to build direct public
  changes without allocating changed-index arrays or rereading the formula cell
  value on the common sliding aggregate edit.
- Dense single-column affine direct-scalar batches now accept both ascending and
  strict descending row order. This keeps `batch-edit-single-column-with-undo`
  on the direct-scalar skip path because inverse undo refs are recorded in
  descending order.
- Dense row-pair direct-scalar batches now have a simple direct evaluator for
  the common `A+B` and `A*B` row-local pair, making the multi-column batch row
  green in the best full samples.

Measured and rejected local experiments:

- Reusing the generic lazy public-change proxy for the two-cell
  existing-numeric path made sliding and batch rows worse under the unchanged
  full suite, so it was removed.
- A bespoke two-cell lazy public-change array avoided generic proxy machinery
  but still measured slower than direct two-object materialization in the local
  sliding micro-loop, so it was removed.
- A trusted physical numeric writer that bypassed generic workbook notification
  did not hold in the full suite and increased semantic risk around batched
  column-version updates, so it was removed.
- A no-history core operation with WorkPaper-side undo record construction also
  measured slower than the existing mutation-service wrapper, so it was
  removed.

Focused validation currently passing:

```sh
pnpm exec tsc --noEmit --pretty false -p packages/core/tsconfig.json
pnpm --filter @bilig/core build
pnpm exec tsc --noEmit --pretty false -p packages/headless/tsconfig.json
pnpm exec vitest run packages/core/src/__tests__/operation-service.test.ts packages/core/src/__tests__/mutation-service.test.ts packages/headless/src/__tests__/work-paper-runtime.test.ts
```

Latest unchanged full competitive artifacts from this pass:

- `/tmp/workpaper-competitive-after-affine-batch-runtime-dimensions.json`:
  `37` WorkPaper wins, `1` HyperFormula win. Remaining red:
  `aggregate-overlapping-sliding-window`.
- `/tmp/workpaper-competitive-after-trusted-existing-numeric-direct.json`:
  `36` WorkPaper wins, `2` HyperFormula wins. Remaining red:
  `build-mixed-content` and `aggregate-overlapping-sliding-window`; the mixed
  build row was median-green but mean-red from outliers.
- `/tmp/workpaper-competitive-after-descending-affine-undo.json`: `37`
  WorkPaper wins, `1` HyperFormula win. Remaining red:
  `aggregate-overlapping-sliding-window`.
- `/tmp/workpaper-competitive-after-reverted-nohistory-keep-descending.json`:
  `37` WorkPaper wins, `1` HyperFormula win. Remaining red:
  `aggregate-overlapping-sliding-window`.
- `/tmp/workpaper-competitive-after-format-final-2026-04-28.json`: `37`
  WorkPaper wins, `1` HyperFormula win. Remaining red:
  `aggregate-overlapping-sliding-window`.

2026-04-28 closeout update:

The oracle implementation is now green on the unchanged full competitive
WorkPaper vs HyperFormula benchmark. Benchmark definitions, sample counts,
fixtures, workload sizes, and scoring were not changed.

Additional production changes in the final closeout tranche:

- Mixed initial sheet load now uses a raw formula-source initializer instead of
  constructing nested `EngineCellMutationRef` formula mutation objects that the
  formula initialization service immediately unwraps. This preserves the same
  coordinates, source text, template compilation, binding, and evaluation path
  while removing object churn in `build-mixed-content`.
- Mixed sheet inspection and formula detection avoid trimming ordinary
  non-formula strings. Padded formulas are still recognized and normalized with
  the existing `trim()` behavior.
- Fresh mixed-sheet physical attachment caches the current row resident set and
  column resident sets during load, reducing repeated map lookups in the
  fresh-sheet-only path without changing logical identity semantics.
- Direct scalar dependency materialization now builds tiny local dependency
  arrays directly for direct-only scalar formulas instead of filling shared
  scratch buffers and slicing them for one/two-cell dependencies.
- Fresh direct-only formulas skip WASM program sync scheduling when the prepared
  runtime program is empty. Formula graph edges, families, direct descriptors,
  and direct evaluation remain unchanged.
- Direct uniform lookup operand edits now return compact existing-numeric
  mutation results when no tracked listener needs an emitted changed-index
  array, and trusted existing-numeric direct lookup writes use the trusted
  physical numeric writer after the same sheet/row/col/formula/tag guards pass.
- Dense physical range reads allocate row arrays with a tight loop instead of
  nested `Array.from` callbacks, reducing allocation overhead and GC exposure
  in `range-read-dense` while preserving the same block-scan value semantics.
- The trusted direct aggregate existing-numeric path increments its two
  performance counters inline. Counter values are unchanged, but the sub-0.05ms
  sliding aggregate path avoids two helper calls.

Final focused validation passing:

```sh
pnpm exec vitest run packages/headless/src/__tests__/initial-sheet-load.test.ts
pnpm exec vitest run packages/core/src/__tests__/operation-service.test.ts packages/headless/src/__tests__/initial-sheet-load.test.ts packages/headless/src/__tests__/work-paper-runtime.test.ts packages/headless/src/__tests__/tracked-cell-index-changes.test.ts
pnpm exec vitest run packages/core/src/__tests__/operation-service.test.ts --testNamePattern 'approximate|aggregate|trusted existing numeric direct scalar chains|typed changed cells|rejects trusted direct aggregate'
pnpm exec tsc --noEmit --pretty false -p packages/core/tsconfig.json
pnpm --filter @bilig/core build
pnpm exec tsc --noEmit --pretty false -p packages/headless/tsconfig.json
```

Final unchanged full competitive artifacts:

- `/tmp/workpaper-competitive-after-mixed-init-raw-sources-2026-04-28.json`:
  `36` WorkPaper wins, `2` HyperFormula wins. `build-mixed-content` was green;
  remaining reds were `aggregate-overlapping-sliding-window` and
  `lookup-approximate-sorted`.
- `/tmp/workpaper-competitive-after-lookup-compact-2026-04-28.json`: `37`
  WorkPaper wins, `1` HyperFormula win. Remaining red:
  `lookup-approximate-sorted`.
- `/tmp/workpaper-competitive-after-trusted-lookup-writer-2026-04-28.json`:
  `37` WorkPaper wins, `1` HyperFormula win. Remaining red:
  `batch-edit-multi-column`; `lookup-approximate-sorted`,
  `aggregate-overlapping-sliding-window`, and `build-mixed-content` were green.
- `/tmp/workpaper-competitive-repeat-after-trusted-lookup-writer-2026-04-28.json`:
  `38` WorkPaper wins, `0` HyperFormula wins out of `38` comparable workloads.
- `/tmp/workpaper-competitive-after-trusted-aggregate-counter-inline-2026-04-28.json`:
  `38` WorkPaper wins, `0` HyperFormula wins out of `38` comparable workloads
  on the formatted final code.

Current best scorecard is `38/0`. No comparable benchmark workloads are red in
the latest full run.
