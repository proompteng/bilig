# WorkPaper Benchmark Expansion And Performance Plan

Date: `2026-04-28`
Checkout: `37c02d7f279d21cb6e987f2ff055d9d89cf7487a` at prompt creation
Oracle request: `https://chatgpt.com/c/69f18c28-44b8-83e8-bd3a-eb2228c71843`
Source package uploaded to oracle: `/tmp/bilig3-oracle/bilig3-workpaper-source-37c02d7f.zip`

## Status

The source zip and benchmark request were submitted to the oracle on
`2026-04-28`. The first-prompt oracle response was later extracted through the
Browser Use in-app browser and saved verbatim in
`docs/workpaper-oracle-chatgpt-first-response-2026-04-29.md`.

This plan has been reconciled against the current checkout and the latest
checked benchmark artifact,
`packages/benchmarks/baselines/workpaper-vs-hyperformula.json`, generated at
`2026-04-29T14:47:16.831Z`.

The active implementation target is no longer the older red-list from the first
expansion run. `sheet-rename-dependencies`, `named-expression-change`,
`build-parser-cache-unique-formulas`, and `lookup-approximate-duplicates` are
green in the current artifact and must be preserved. The next production work is
focused on `build-mixed-content`, `structural-delete-rows`, and the
`lookup-text-exact` p95 tail.

## Non-Negotiable Constraints

- Do not reduce existing workload sizes, warmups, samples, verification, or
  scoring to hide losses.
- Do not remove comparable workloads or reclassify comparable workloads as
  leadership-only unless HyperFormula cannot compute an equivalent result.
- Every comparable workload must keep exact WorkPaper vs HyperFormula
  verification equality before timing is accepted.
- Performance wins must come from production engine/headless paths, not
  benchmark-specific branches.
- Scorecards must show losses directly through directional ratios where values
  above `1.0` mean WorkPaper is slower.

## Benchmark Expansion

The expanded suite now covers these additional workloads:

- Build: `build-parser-cache-unique-formulas`
- Sheet lifecycle: `sheet-rename-dependencies`
- Named expressions: `named-expression-change`
- Aggregation: `aggregate-2d-ranges`
- Conditional aggregation: `conditional-aggregation-shared-criteria`,
  `conditional-aggregation-mixed-criteria`
- Approximate lookup: `lookup-approximate-descending`,
  `lookup-approximate-duplicates`
- Unsupported capability leadership: `lookup-reverse-search`,
  `dynamic-array-sort`, `dynamic-array-unique`

`lookup-reverse-search` is leadership-only because HyperFormula 3.2.0 returned
`#NAME?` for the `XMATCH(...,0,-1)` fixture during local equivalence validation.
The dynamic-array additions follow the existing `dynamic-array-filter` rule:
they are capability evidence, not scorecard inputs.

## Reporting Integrity

The benchmark report now records:

- per-engine `standardDeviation`, `relativeStandardDeviation`, `standardError`,
  and `confidence95`
- per-workload directional mean, median, and p95 ratios
- per-workload max relative noise and confidence interval overlap
- per-family directional mean and p95 geomeans
- per-family worst mean and p95 ratios with workload names
- scorecard lanes: `overall`, `public`, and `holdout`

The generated artifact includes the enriched `families` and `scorecard`
surfaces, and the checked baseline now points to the current `bilig3`
`packages/headless` path rather than the old `bilig2` path.

## Current Scorecard After Expansion

Generated with `pnpm workpaper:bench:competitive:generate` on `2026-04-29` after
the benchmark expansion and the retained production engine/headless changes:

- Total workloads: `51`
- Comparable workloads: `47`
- Leadership-only workloads: `4`
- Scorecard-eligible comparable workloads: `46`
- Overall scorecard: WorkPaper `44`, HyperFormula `2`
- Public lane: WorkPaper `36`, HyperFormula `2`
- Holdout lane: WorkPaper `8`, HyperFormula `0`
- Worst mean ratio: `build-mixed-content`, `1.0362639565590437`
- Worst p95 ratio: `lookup-text-exact`, `2.27208263805424`

Current HyperFormula mean rows:

- `build-mixed-content`: mean ratio `1.0362639565590437`, median ratio
  `1.0069852963334736`, p95 ratio `1.156165042556`,
  `confidenceIntervalOverlaps: true`.
- `structural-delete-rows`: mean ratio `1.0234049542127845`, median ratio
  `0.8750303474565914`, p95 ratio `1.267650293785557`,
  `confidenceIntervalOverlaps: true`.

Current preservation rows from the original oracle red-list:

- `build-parser-cache-unique-formulas`: WorkPaper wins in the holdout lane.
- `sheet-rename-dependencies`: WorkPaper wins in the holdout lane.
- `named-expression-change`: WorkPaper wins in the holdout lane.
- `lookup-approximate-duplicates`: WorkPaper wins in the holdout lane.
- `aggregate-overlapping-sliding-window`: green in the latest official artifact
  after showing noise in earlier runs.

The public lane is close but not complete: `38` comparable workloads, `36`
WorkPaper wins, `2` HyperFormula wins. The holdout lane is complete in the
current artifact: `8` comparable workloads, `8` WorkPaper wins.

## Implementation Status

Completed in this tranche:

- Expanded the competitive benchmark to `51` workloads with public/holdout
  lanes, directional ratios, noise fields, family scorecards, and leadership
  classification for unsupported HyperFormula capabilities only.
- Added mixed computed-criteria operands for direct conditional aggregation,
  covering criteria such as `">="&E1` without changing the workload.
- Added a listener-free named-expression change fast path in headless, avoiding
  full visibility and named-value snapshots for the common mutation path.
- Added a listener-free sheet rename headless path and removed unnecessary core
  sheet-delete/structural invalidation treatment for metadata-only renames.
- Added rectangular direct aggregate compilation, binding, initial evaluation,
  prefix evaluation, region subscriptions, and mutation touch detection for
  ranges such as `SUM(A1:B1500)`.

Still open:

- `build-mixed-content` still needs production cold-build hardening. The useful
  work is reducing duplicated initialization and allocation across literal
  loading, formula source registration, binding, and initial evaluation.
- `structural-delete-rows` still needs production structural-row hardening. The
  useful work is narrowing row-delete metadata updates and avoiding full-sheet
  dependency/result collection when tracked evidence proves it is unnecessary.
- `lookup-text-exact` has the worst p95 ratio and should be treated as a
  tail-latency hardening issue around normalization, index reuse, invalidation,
  and allocation spikes.
- The previous oracle response is saved and validated. A new oracle consultation
  is only justified if profiling exposes a real architecture blocker or a new
  decisive non-overlap HyperFormula row appears.

## Current Production Implementation Plan

### 1. `build-mixed-content` cold-build hardening

Files:

- `packages/headless/src/initial-sheet-load.ts`
- `packages/headless/src/work-paper-runtime.ts`
- `packages/core/src/literal-sheet-loader.ts`
- `packages/core/src/engine/services/formula-initialization-service.ts`
- `packages/core/src/engine/services/formula-binding-service.ts`

Problem: `build-mixed-content` is the worst current mean-ratio row, but only by
`1.036x` and with overlapping confidence intervals. It is close enough that
small general-purpose allocation and duplicated-initialization costs decide the
row.

Plan:

- Profile cold mixed-sheet construction by phase: literal grid load, formula
  source registration, formula initialization, direct descriptor binding,
  initial evaluation, changed-cell metadata, and inspect/dimension collection.
- Remove duplicated work only in production paths used by normal headless build
  APIs.
- Keep the reverted fresh-formula changed-scratch deferral out unless a corrected
  version wins the official workload and preserves focused tests.
- Preserve all formulas, dependency descriptors, result values, source text, and
  verification keys.

### 2. `structural-delete-rows` row-delete hardening

Files:

- `packages/core/src/sheet-grid.ts`
- `packages/core/src/workbook-store.ts`
- `packages/core/src/engine/services/mutation-service.ts`
- `packages/core/src/engine/services/operation-service.ts`
- `packages/headless/src/work-paper-runtime.ts`

Problem: `structural-delete-rows` is the other current HyperFormula mean row.
It is median-green but mean-red, which points to row-delete tail overhead and
unnecessary result collection rather than a missing correctness feature.

Plan:

- Profile logical/physical row remapping, dependency/index metadata updates,
  formula-public-text rewrites, undo records, and headless changed-cell payload
  construction.
- Narrow touched metadata to deleted intervals and dependent formulas whose
  ranges actually intersect or shift.
- Avoid full-sheet recomputation and broad result materialization when the
  tracked structural plan proves the delete is local.
- Preserve logical row/column identity, formula rewrite semantics, undo/redo,
  and verification equality.

### 3. `lookup-text-exact` p95 hardening

Files:

- `packages/core/src/engine/services/lookup-column-owner.ts`
- `packages/core/src/engine/services/sorted-column-search-service.ts`
- `packages/core/src/engine/services/operation-service.ts`
- `packages/core/src/engine/services/formula-evaluation-service.ts`

Problem: `lookup-text-exact` is not a current HyperFormula mean row, but it has
the worst p95 ratio. That makes it a stability target before claiming a durable
benchmark win.

Plan:

- Measure text-key normalization, exact lookup index reuse, post-write
  invalidation, and result allocation.
- Reduce tail allocation and cache churn without changing comparison semantics.
- Keep exact text lookup verification and workload sampling unchanged.

### 4. Preservation checks

The active implementation must preserve all green rows that were red earlier:

- `sheet-rename-dependencies`
- `named-expression-change`
- `build-parser-cache-unique-formulas`
- `lookup-approximate-duplicates`
- `aggregate-overlapping-sliding-window`
- all `8/8` holdout wins

## Validation Commands

Use these in order:

```bash
pnpm --filter @bilig/benchmarks build
bun scripts/run-vitest.ts --run packages/benchmarks/src/__tests__/stats.test.ts packages/benchmarks/src/__tests__/expanded-workloads.test.ts
pnpm workpaper:bench:competitive:check
pnpm bench:workpaper:competitive
pnpm run ci
```

For a changed engine tranche, rerun:

```bash
pnpm workpaper:bench:competitive:generate
pnpm workpaper:bench:competitive:check
```

## Cheating Red Flags

- Increasing only warmups or samples until a noisy sub-ms row flips green.
- Making `lookup-reverse-search` comparable without proving HyperFormula returns
  an equivalent `XMATCH` result.
- Treating leadership-only dynamic arrays as scorecard wins.
- Dropping the holdout lane from scorecard reporting.
- Optimizing benchmark fixtures by detecting workload names or exact formulas.
- Removing verification keys from workloads that currently expose semantic
  differences.

## Implementation Log - 2026-04-29T12:27Z

This chronological log entry records an intermediate run and does not override
the current scorecard above. It implemented production-only cold-load and
direct-lookup improvements without changing benchmark definitions, scoring,
sampling, or workload sizes.

Changes validated in this cycle:

- `SheetGrid.createRowMajorSetter()` now avoids repeated block map lookups during
  fresh row-major sheet loads.
- `loadLiteralSheetIntoEmptySheet()` and headless mixed initial sheet loading use
  the row-major grid setter.
- `operation-service` now evaluates non-uniform approximate direct lookup operand
  edits through the prepared sorted lookup service and applies the compact
  result directly. This covers duplicate-key approximate `MATCH` edits without
  falling back to generic direct-formula evaluation.
- `work-paper-runtime.test.ts` now covers duplicate approximate `MATCH` operand
  edits and asserts the compact direct path counters.

Focused post-warmup timing on the three red rows was green:

- `build-mixed-content`: WorkPaper/HyperFormula mean ratio `0.889`.
- `aggregate-overlapping-sliding-window`: mean ratio `0.509`.
- `lookup-approximate-duplicates`: mean ratio `0.742`.

Intermediate regenerated artifact:

- Generated at `2026-04-29T12:27:03.667Z`.
- Overall scorecard: `42` WorkPaper wins, `4` HyperFormula wins, `46`
  comparable workloads.
- Remaining HyperFormula rows:
  - `build-mixed-content`: mean ratio `1.2034`, overlap `false`.
  - `aggregate-overlapping-sliding-window`: mean ratio `1.1361`, overlap
    `true`.
  - `build-from-sheets`: mean ratio `1.0087`, overlap `true`.
  - `build-dense-literals`: mean ratio `1.0903`, overlap `true`.

Oracle status at that moment:

- Browser Use attach was unavailable in that cycle with
  `Browser turn does not belong to this IAB pipe`.
- Computer Use app-state request for ChatGPT Atlas timed out.
- At that moment the oracle response had not yet been captured. It was later
  captured and saved in `docs/workpaper-oracle-chatgpt-first-response-2026-04-29.md`.

## Implementation Log - 2026-04-29T13:05Z

This chronological log entry records an intermediate run and does not override
the current scorecard above. It continued production-only implementation against
the same validated design. This cycle did not change workload definitions,
scoring, sampling, or workload sizes.

Additional implementation:

- `WorkPaper.buildFromSheets()` now keeps static-sheet construction on one
  sheet-entry vector instead of repeatedly allocating `Object.keys()` /
  `Object.entries()` results and `Map` lookups across inspect, create, load, and
  dimension-cache phases.
- `SpreadsheetEngine.createSheetForInitialization()` now returns the initialized
  sheet id so headless build can avoid a second name lookup.
- `WorkbookStore.createLogicalAxisIdEnsurer()` lets fresh literal and mixed
  loaders cache the sheet lookup for row/column logical-id creation while still
  creating only the axis ids needed by materialized cells.

Validation:

- `bun scripts/run-vitest.ts --run packages/core/src/__tests__/literal-sheet-loader.test.ts packages/headless/src/__tests__/initial-sheet-load.test.ts packages/headless/src/__tests__/work-paper-runtime.test.ts`
  passed.
- `pnpm --filter @bilig/core build`, `pnpm --filter @bilig/headless build`, and
  `pnpm --filter @bilig/benchmarks build` passed.
- `pnpm lint` passed.
- `pnpm workpaper:bench:competitive:generate` regenerated
  `packages/benchmarks/baselines/workpaper-vs-hyperformula.json` at
  `2026-04-29T13:05:50.243Z`.
- `pnpm workpaper:bench:competitive:check` passed.

Updated official scorecard:

- Overall: `44` WorkPaper wins, `2` HyperFormula wins, `46` comparable.
- Public: `36` WorkPaper wins, `2` HyperFormula wins.
- Holdout: `8` WorkPaper wins, `0` HyperFormula wins.
- Remaining HyperFormula rows are overlapping/noisy: `build-mixed-content`
  mean ratio `1.0166`, overlap `true`; `aggregate-overlapping-sliding-window`
  mean ratio `1.1127`, overlap `true`.

Oracle status at that moment:

- Browser Use retry was unavailable in that cycle after a Node REPL reset with
  no available in-app browser backend.
- At that moment the oracle response had not yet been captured. It was later
  captured and saved in `docs/workpaper-oracle-chatgpt-first-response-2026-04-29.md`.

## Implementation Log - 2026-04-29T13:55Z

This chronological log entry records an intermediate run and does not override
the current scorecard above. The cycle continued implementation in the
production WorkPaper engine/headless paths only. Existing benchmark definitions,
scoring, sampling, and workload sizes were preserved.

Production work completed:

- Added a no-volatile runtime guard to initial formula initialization.
- Avoided fresh empty region-graph subscription replacement for direct scalar
  formulas.
- Narrowed direct lookup operand mutations to numeric terminal result writes
  where the prepared lookup result is numeric.
- Added compact second-cell row/column metadata for direct lookup and
  single-direct-aggregate mutation results.
- Replaced generic aggregate hot-path counter helper calls with direct counter
  increments.

Validation:

- Focused Vitest suites passed for formula initialization, lookup mutation,
  aggregate compilation, and WorkPaper runtime paths.
- `pnpm --filter @bilig/core build` passed.
- `pnpm --filter @bilig/headless build` passed.
- `pnpm workpaper:bench:competitive:generate` regenerated
  `packages/benchmarks/baselines/workpaper-vs-hyperformula.json` at
  `2026-04-29T13:52:59.186Z`.
- `pnpm workpaper:bench:competitive:check` passed.

Current official benchmark state:

- Overall: `44` WorkPaper wins, `2` HyperFormula wins, `46` comparable.
- Public: `36` WorkPaper wins, `2` HyperFormula wins.
- Holdout: `8` WorkPaper wins, `0` HyperFormula wins.
- Remaining HyperFormula rows are `build-mixed-content` and
  `aggregate-overlapping-sliding-window`; both are confidence-overlap noisy.

External oracle status:

- Browser Use was retried against the configured `iab` backend but no Codex
  in-app-browser backend was available, so no new oracle response was saved.

## Implementation Log - 2026-04-29T14:47Z

This chronological log entry is the latest official benchmark evidence currently
referenced by the current scorecard above. It preserved the benchmark suite and
kept changes out of workload definitions, scoring, sampling, and workload sizes.

Implementation outcome:

- Tested a production cold-build optimization that deferred changed-formula
  scratch marking for fresh initial formula loads.
- Reverted that optimization after the official competitive run worsened
  `build-mixed-content`; the degraded code was not retained.
- Regenerated the baseline from the reverted production code.

Validation:

- Focused initial-load/runtime tests passed.
- `pnpm --filter @bilig/core build` passed.
- `pnpm --filter @bilig/headless build` passed.
- `pnpm lint` passed.
- `pnpm workpaper:bench:competitive:generate` regenerated
  `packages/benchmarks/baselines/workpaper-vs-hyperformula.json` at
  `2026-04-29T14:47:16.831Z`.
- `pnpm workpaper:bench:competitive:check` passed.

Current official benchmark state:

- Overall: `44` WorkPaper wins, `2` HyperFormula wins, `46` comparable.
- Public: `36` WorkPaper wins, `2` HyperFormula wins.
- Holdout: `8` WorkPaper wins, `0` HyperFormula wins.
- Current red rows: `build-mixed-content` and `structural-delete-rows`; both
  have overlapping confidence intervals in this artifact.

CI status:

- Full CI was not rerun after the reverted experiment. Last full CI evidence
  still has all tests passing and fails only global coverage thresholds.
