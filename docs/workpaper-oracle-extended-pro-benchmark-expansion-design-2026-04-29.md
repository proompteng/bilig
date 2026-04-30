# WorkPaper Oracle Extended Pro Benchmark Expansion Design

Date: `2026-04-29`
Oracle thread: `https://chatgpt.com/c/69f18c28-44b8-83e8-bd3a-eb2228c71843`
Source zip reviewed by oracle: `/tmp/bilig3-oracle/bilig3-workpaper-source-37c02d7f.zip`
Prompt checkout: `37c02d7f279d21cb6e987f2ff055d9d89cf7487a`
Current checkout branch during validation: `main`
Current baseline artifact: `packages/benchmarks/baselines/workpaper-vs-hyperformula.json`

## Validated Oracle Plan

The first-prompt oracle response was extracted through the Browser Use in-app
browser and saved verbatim in
`docs/workpaper-oracle-chatgpt-first-response-2026-04-29.md`. That raw response
is preserved below as an appendix for provenance; it is not the active
implementation checklist because the checkout has advanced since the source zip
that oracle reviewed.

This section is the updated original oracle plan after validation against the
current checkout and
`packages/benchmarks/baselines/workpaper-vs-hyperformula.json`. It is the source
of truth for the next implementation slice.

The validated plan accepts these oracle principles:

- Existing benchmark definitions, scorecard rules, sample counts, warmups,
  workload sizes, and verification gates stay fixed.
- Remaining work is limited to production engine/headless code paths and
  focused tests for those paths.
- The current artifact already proves the expanded benchmark/reporting design is
  wired: `51` workloads, `46` comparable scorecard workloads, public and holdout
  lanes, confidence intervals, and directional ratios.
- The major decisive losses from the first oracle response are no longer active
  blockers in the latest generated artifact. `sheet-rename-dependencies`,
  `named-expression-change`, `build-parser-cache-unique-formulas`, and
  `lookup-approximate-duplicates` must be preserved as green rows, not treated as
  the next red-list.
- The next implementation slice targets the actual current evidence:
  `build-mixed-content`, `structural-delete-rows`, and the `lookup-text-exact`
  p95 tail risk.

Current benchmark evidence from
`packages/benchmarks/baselines/workpaper-vs-hyperformula.json`, generated at
`2026-04-29T14:47:16.831Z`:

- Overall: WorkPaper `44`, HyperFormula `2`, comparable `46`.
- Public: WorkPaper `36`, HyperFormula `2`, comparable `38`.
- Holdout: WorkPaper `8`, HyperFormula `0`, comparable `8`.
- Remaining HyperFormula mean rows:
  - `build-mixed-content`: mean ratio `1.0362639565590437`, median ratio
    `1.0069852963334736`, p95 ratio `1.156165042556`,
    `confidenceIntervalOverlaps: true`.
  - `structural-delete-rows`: mean ratio `1.0234049542127845`, median ratio
    `0.8750303474565914`, p95 ratio `1.267650293785557`,
    `confidenceIntervalOverlaps: true`.
- Worst p95 row is `lookup-text-exact` at p95 ratio `2.27208263805424`; this is
  a tail-risk hardening target even if its mean scorecard result is not a
  HyperFormula mean win.

Accepted production implementation sequence:

1. `build-mixed-content` hardening:
   - Profile cold mixed sheet construction across literal cells, formula
     binding, formula source registration, changed-cell metadata, and initial
     evaluation.
   - Keep the reverted fresh-formula changed-scratch deferral out unless new
     evidence proves a corrected version helps the official workload.
   - Prefer removing duplicated initialization work and unnecessary allocation
     from general build paths over adding benchmark-specific branches.

2. `structural-delete-rows` hardening:
   - Profile row deletion through sheet-grid remapping, dependency/index
     metadata updates, formula binding updates, and headless runtime result
     collection.
   - Preserve logical row/column identity semantics and formula correctness.
   - Optimize common deletion paths by narrowing touched metadata and avoiding
     full-sheet recomputation where dependency evidence proves it is unnecessary.

3. `lookup-text-exact` p95 hardening:
   - Investigate the high p95 ratio as a tail-latency issue, not as a scoring
     failure.
   - Focus on lookup key normalization, index reuse, cache invalidation after
     writes, and allocation spikes.
   - Do not change workload sampling or scoring to reduce visible variance.

4. Preservation checks:
   - Re-run focused tests for sheet rename, named expressions, parser/build,
     lookup approximate duplicates, direct aggregates, and structural edits.
   - Regenerate the competitive baseline and confirm the previously green rows
     stay green.
   - Run `pnpm workpaper:bench:competitive:check`.
   - Run `pnpm run ci` once the implementation slice is stable, with the known
     coverage threshold blocker handled honestly rather than ignored.

Completion criteria:

- No decisive non-overlap HyperFormula wins remain.
- Preferably all scorecard-eligible comparable workloads are WorkPaper mean
  wins; if a noisy overlap row remains, it must have focused high-sample evidence
  showing the row is not a stable production regression.
- Public and holdout scorecards remain visible in the committed artifact.
- No benchmark workload sizes, scoring, sampling, verification, or eligibility
  are weakened.
- Any additional oracle consultation is only needed if profiling reveals a real
  production architecture blocker or if a new decisive red row appears after the
  validated plan is exhausted.

## Historical Validation Notes

Earlier validation snapshots are retained below for auditability. They describe
intermediate benchmark states and should not override the validated plan above.

- At `2026-04-29T08:44:29.667Z`, the generated artifact still showed decisive
  HyperFormula wins in `build-parser-cache-unique-formulas`,
  `sheet-rename-dependencies`, and `named-expression-change`. Those rows are no
  longer active blockers in the latest artifact.
- At `2026-04-29T11:32:58.561Z`, the generated artifact improved to `44/46`
  WorkPaper wins, with `build-mixed-content` and
  `aggregate-overlapping-sliding-window` as confidence-overlap HyperFormula
  rows.
- Focused higher-sample checks showed `lookup-with-column-index-after-batch-write`,
  `lookup-approximate-duplicates`, and
  `aggregate-overlapping-sliding-window` were not stable production losses on
  the then-current checkout.
- Production changes already landed include direct engine initialization for
  cold sheet construction, avoiding redundant cold-load physical key writes into
  `cellKeyToIndex`, and faster sheet-inspection dimension calculation.

## Appendix: Raw First-Prompt Oracle Response

Source-of-truth check

I inspected the attached checkout. Your file-path assumptions match the source:

packages/benchmarks/src/expanded-competitive-workloads.ts is the current expanded workload union.

packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts is the expanded benchmark runner.

packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded-additional-workloads.ts and packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded-support.ts hold additional workload and support code.

packages/benchmarks/src/report-competitive-families.ts is the family scorecard code.

packages/core/src/perf/engine-counters.ts is the engine-counter surface.

The correctness gate is real: comparable runs serialize WorkPaper and HyperFormula verification JSON and reject the result if they differ. The sampled verification also rejects intra-engine verification drift across samples. Default competitive sampling is 2 warmups and 5 samples.

Important mismatch: the checked-in baseline artifact does not match the local rerun state you gave. packages/benchmarks/baselines/workpaper-vs-hyperformula.json is generated on 2026-04-26T19:45:31.057Z, has 40 results, 39 comparable, 1 leadership-only, but only 23 WorkPaper wins and 16 HyperFormula wins. It also does not include the current family/scorecard report shape. Treat that artifact as stale relative to your 2026-04-28 reruns, not as proof of the current engine state.

Current comparable workload count in source is 39, with one genuine leadership workload: dynamic-array-filter.

1. Benchmark expansion plan

Keep the existing 40 workloads exactly as-is. Append new workloads. Do not resize, remove, retag, or rescore any existing workload.

The new suite should add workloads that stress capabilities already implied by existing benchmark fixtures but not yet exercised. The point is to make cheating harder by broadening shape coverage: different lookup modes, 2D aggregates, criteria reuse, structural dependency rewrites, named expressions, and dynamic-array variants.

A. Build and parser-cache family

Add:

build-parser-cache-unique-formulas

Fixture: buildParserCacheUniqueFormulaSheet in packages/benchmarks/src/workpaper-benchmark-fixtures.ts.

Purpose: current parser-cache workloads mostly reward repeated row templates. This adds high-cardinality formula shapes so parser/template optimizations must be general, not just repeated-template shortcuts.

Comparable: yes.

Verification: dimensions, several fixed cells across the sheet, terminal formula cell, and checksum over sampled formula outputs.

Implementation files:

packages/benchmarks/src/expanded-competitive-workloads.ts

packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded-additional-workloads.ts

packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts

build-named-expression-formulas

Fixture: buildNamedExpressionBenchSheet.

Purpose: tests build/evaluation through named expressions rather than only direct cell/range references.

Comparable: yes only after the benchmark helper creates equivalent named expressions in both WorkPaper and HyperFormula. Do not mark leadership-only merely because the helper is missing.

Verification: named expression values, formula outputs using names, formula text if both engines expose equivalent formula text.

Implementation files:

packages/benchmarks/src/workpaper-benchmark-fixtures.ts

packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded-additional-workloads.ts

packages/headless/src/work-paper-runtime.ts

B. Rebuild and runtime-restore family

Add:

rebuild-unique-formulas

Fixture: buildParserCacheUniqueFormulaSheet.

Purpose: validates rebuild performance when formulas do not collapse to a small template bank.

Comparable: yes.

Verification: same as build-parser-cache-unique-formulas, but after rebuild.

rebuild-runtime-from-named-expressions

Fixture: buildNamedExpressionBenchSheet.

Purpose: named-expression persistence, restore, and recalc.

Comparable: yes after API parity is implemented.

Verification: named expression registry plus dependent formula outputs after restore.

C. Dirty execution / single-edit family

Add:

single-edit-unique-formula-frontier

Fixture: buildParserCacheUniqueFormulaSheet.

Operation: edit one upstream input cell that invalidates a meaningful but bounded frontier.

Purpose: prevents winning only on repeated formula templates and exposes dirty-graph precision.

Comparable: yes.

Verification: changed cells, terminal sampled cells, unchanged sampled cells.

single-edit-shared-criteria-threshold

Fixture: buildConditionalAggregationSharedCriteriaSheet.

Operation: edit a criteria/threshold cell shared by many formulas.

Purpose: tests whether WorkPaper handles criteria-cell invalidation with production cache invalidation, not broad full-sheet recompute.

Comparable: yes.

Verification: every dependent conditional aggregate family value must match HyperFormula.

D. Batch edit family

Add:

batch-edit-mixed-criteria-inputs

Fixture: buildConditionalAggregationMixedSheet.

Operation: batch edit a mix of aggregate values, criteria columns, and criteria cells.

Purpose: forces correct incremental invalidation across SUMIF/SUMIFS/COUNTIF/COUNTIFS-like shapes.

Comparable: yes.

Verification: sampled edited rows, criteria outputs, aggregate outputs, and non-edited sentinels.

batch-edit-unique-formula-inputs

Fixture: buildParserCacheUniqueFormulaSheet.

Operation: batch edit scattered input cells.

Purpose: tests dirty-frontier batching without relying on identical formula shape.

Comparable: yes.

E. Structural rows family

Add:

structural-delete-rows-lookup-index

Fixture: buildLookupSheet.

Operation: delete rows in the lookup table, then read lookup outputs.

Purpose: verifies lookup-index invalidation and retargeting after structural mutation.

Comparable: yes.

Verification: lookup outputs before/after delete, dimensions, and formula results.

structural-move-rows-conditional-aggregation

Fixture: buildConditionalAggregationSharedCriteriaSheet.

Operation: move a range of source rows.

Purpose: validates direct criteria descriptor retargeting and range-registry correctness.

Comparable: yes.

Verification: formula outputs plus dimensions.

F. Structural columns family

Add:

structural-insert-columns-conditional-aggregation

Fixture: buildConditionalAggregationMixedSheet.

Operation: insert columns inside or adjacent to criteria/aggregate ranges.

Purpose: tests formula retargeting and direct-criteria descriptor survival.

Comparable: yes.

structural-delete-columns-conditional-aggregation

Fixture: buildConditionalAggregationMixedSheet.

Operation: delete non-trivial columns and verify formulas update or error equivalently.

Purpose: catches benchmark-only fast paths that assume stable column layout.

Comparable: yes.

G. Cross-sheet structure family

Add a new reporting family: cross-sheet-structure.

Add:

rename-sheet-with-dependencies

Fixture: buildRenameDependencySheets.

Operation: rename a referenced sheet.

Purpose: tests dependency graph, formula rewrite, and cross-sheet invalidation.

Comparable: yes if both engines support sheet rename in the benchmark helper. HyperFormula does, so this should not be leadership-only.

Verification: dependent formula values, rewritten formula text where both engines expose it, and sheet names.

rename-sheet-many-dependent-formulas

Fixture: add a larger variant next to buildRenameDependencySheets, for example buildRenameDependencyManySheets(rowCount, dependentSheetCount).

Purpose: prevents a tiny rename workload from being dominated by fixed overhead.

Comparable: yes.

H. Range read and aggregate family

Add:

aggregate-2d-prefix-sum

Fixture: build2dAggregateSheet.

Purpose: current direct aggregate paths are likely strongest on simple one-dimensional ranges. This adds rectangular ranges.

Comparable: yes.

Verification: sampled 2D aggregate outputs and terminal values.

aggregate-2d-after-cell-write

Fixture: build2dAggregateSheet.

Operation: edit one or more source cells inside rectangular aggregate ranges.

Purpose: tests whether rectangular aggregate caches/deltas invalidate correctly.

Comparable: yes.

I. Conditional aggregation family

Add:

conditional-aggregation-shared-criteria

Fixture: buildConditionalAggregationSharedCriteriaSheet.

Purpose: tests repeated criteria-range reuse.

Comparable: yes.

conditional-aggregation-mixed-countifs-sumifs

Fixture: buildConditionalAggregationMixedSheet.

Purpose: tests mixed conditional aggregation formulas, not only one formula class.

Comparable: yes.

conditional-aggregation-mixed-criteria-cell-edit

Fixture: buildConditionalAggregationMixedSheet.

Operation: edit criteria cells, not only source values.

Purpose: criteria edits are harder to optimize honestly than aggregate-value edits.

Comparable: yes.

J. Lookup exact family

Add:

lookup-exact-reverse-search

Fixture: buildLookupSearchModeReverseSheet.

Purpose: tests reverse exact lookup/search-mode behavior.

Comparable: yes.

Verification: duplicate-key behavior must match HyperFormula exactly.

lookup-exact-reverse-after-column-write

Fixture: buildLookupSearchModeReverseSheet.

Operation: mutate lookup-column values and rerun reverse lookup.

Purpose: tests exact-index invalidation and last-match semantics.

Comparable: yes.

lookup-exact-repeated-shared-index

Fixture: add buildRepeatedLookupFormulaSheet(rowCount, formulaCopies).

Purpose: many formulas use the same lookup vector. WorkPaper should build one production index and share it.

Comparable: yes.

K. Lookup approximate family

Add:

lookup-approximate-descending

Fixture: buildApproxLookupDescendingSheet.

Purpose: tests descending approximate lookup, not only ascending.

Comparable: yes.

lookup-approximate-descending-after-column-write

Fixture: buildApproxLookupDescendingSheet.

Operation: mutate tail/interior values and verify lookup outputs.

Purpose: catches invalid uniform-vector assumptions.

Comparable: yes.

lookup-approximate-duplicates

Fixture: buildApproxLookupDuplicateSheet.

Purpose: duplicate keys are where approximate lookup shortcuts often become wrong.

Comparable: yes.

Verification must include boundary cases: before first key, exact duplicate key, between duplicate groups, and after last key.

L. Dynamic-array family

Keep existing:

dynamic-array-filter

Add:

dynamic-array-sort

Fixture: buildDynamicArraySortSheet.

Comparable: leadership-only only while HyperFormula cannot execute equivalent dynamic-array semantics in this benchmark harness.

dynamic-array-unique

Fixture: buildDynamicArrayUniqueSheet.

Comparable: leadership-only only while HyperFormula cannot execute equivalent dynamic-array semantics.

Do not include leadership-only dynamic-array rows in the WorkPaper-vs-HyperFormula public scorecard. They can be reported as capability evidence in a separate section.

2. Reporting and scoring integrity plan

The current report is too mean-oriented for sub-ms rows. Keep the existing fields for backward compatibility, but add robust metrics and scorecards.

Result-level metrics to add

In packages/benchmarks/src/stats.ts, extend the numeric summary or add a companion type:

median

p95

mean

min

max

mad

cv

sampleCount

bootstrap95Low

bootstrap95High

For each comparable row, add directional ratios:

workpaperMeanRatio = workpaper.mean / hyperformula.mean

workpaperMedianRatio = workpaper.median / hyperformula.median

workpaperP95Ratio = workpaper.p95 / hyperformula.p95

Use ratio direction consistently: below 1.0 means WorkPaper is faster.

Add classification:

workpaper-win

hyperformula-win

tie-noisy

unsupported-leadership

Recommended public classification rule:

WorkPaper win if workpaperMedianRatio < 0.95 and the confidence interval does not cross 1.0.

HyperFormula win if workpaperMedianRatio > 1.05 and the confidence interval does not cross 1.0.

Otherwise tie-noisy.

Do not convert noisy rows into wins. The two locally observed red rows moving between reruns are evidence that this is needed.

Family metrics to add

In packages/benchmarks/src/report-competitive-families.ts, replace or supplement the current family meanSpeedupGeomean with directional ratios:

familyGeomeanWorkPaperRatio

familyWorstCaseWorkPaperRatio

familyWorstCaseWorkloadId

familyMedianWorkPaperRatio

familyP95WorkPaperRatio

workpaperWins

hyperformulaWins

tieNoisy

unsupported

The family worst-case ratio is critical. A geomean can look good while one common operation is bad.

Use family worst-case as a release gate:

Public target: every comparable family has familyWorstCaseWorkPaperRatio < 1.0, or any exception is explicitly labeled as an active regression.

Stronger target: every comparable family has familyWorstCaseWorkPaperRatio <= 0.95 on the high-confidence run.

Three scorecards

Add these scorecards to the generated report artifact.

current

Exactly the existing comparable workloads. No new workloads. No changed sizes. This preserves continuity with prior claims.

public

Existing workloads plus the new checked-in comparable workloads. This is the main external scorecard.

Exclude only genuinely non-comparable leadership workloads, such as dynamic arrays while HyperFormula cannot run equivalent semantics in the harness.

holdout

A deterministic, audited anti-overfitting set. It should be generated from fixed seeds and formula families, not hand-tuned after seeing results.

Suggested file:

packages/benchmarks/src/competitive-holdout-workloads.ts

Holdout rules:

Use fixed seeds committed in source.

Keep workload families public, but vary sizes, row offsets, duplicate distributions, and edit locations.

Verify outputs exactly like public workloads.

Do not use holdout rows in marketing claims unless the definitions are published.

Periodically rotate by publishing old holdouts and adding new seeded variants.

Files to edit for scoring

packages/benchmarks/src/stats.ts

packages/benchmarks/src/report-competitive-families.ts

packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts

packages/benchmarks/src/expanded-competitive-workloads.ts

scripts/gen-workpaper-vs-hyperformula-benchmark.ts

packages/benchmarks/src/__tests__/expanded-workloads.test.ts

Add tests that fail if:

an existing workload disappears;

an existing workload changes its comparable/leadership status without explicit fixture evidence;

a comparable workload lacks verification;

a workload has no family mapping;

a report has no current/public/holdout scorecard;

a public comparable row is classified only from mean when median disagrees.

3. Production implementation plan to make WorkPaper win honestly

The benchmark expansion will likely expose four real engine needs: lookup correctness/performance, criteria-cache deltas, 2D aggregate fast paths, and structural retargeting. Implement these as production capabilities, not benchmark-specific branches.

A. Lookup engine improvements

Primary files:

packages/core/src/engine/services/exact-column-index-service.ts

packages/core/src/engine/services/sorted-column-search-service.ts

packages/core/src/engine/services/lookup-column-owner.ts

packages/core/src/engine/direct-vector-lookup.ts

packages/core/src/engine/services/formula-binding-service.ts

packages/core/src/engine/services/formula-evaluation-service.ts

packages/core/src/engine/services/operation-service.ts

Targets:

Reverse exact search should use a production exact-column index with first/last position maps. XMATCH(..., 0, -1) and equivalent reverse exact forms should read last-match positions directly.

Approximate descending lookup should use the same sorted-vector service as ascending lookup, with explicit descending comparator semantics. Do not special-case only monotonic sequential numbers.

Duplicate approximate lookup must implement exact spreadsheet semantics. For ascending approximate, choose the last value less than or equal to the target. For descending approximate, choose the last value greater than or equal to the target under descending order. Add binary-search helpers that work with duplicate runs.

After-write lookup workloads should update or invalidate lookup indexes precisely. Tail append/update can use a patch path if monotonicity is preserved. Interior writes should either update the sorted metadata correctly or invalidate and rebuild. Never assume the benchmark writes only tail cells.

Add counters in packages/core/src/perf/engine-counters.ts:

exactLookupIndexHits

exactLookupIndexBuilds

approxLookupPreparedHits

approxLookupPreparedBuilds

approxLookupTailPatches

lookupIndexInvalidations

Tests:

packages/core/src/__tests__/exact-column-index-service.test.ts

packages/core/src/__tests__/sorted-column-search-service.test.ts

packages/core/src/__tests__/direct-vector-lookup.test.ts

packages/core/src/__tests__/lookup-service.test.ts

packages/core/src/__tests__/lookup-column-owner.test.ts

packages/core/src/__tests__/operation-service.test.ts

Expected benchmark impact:

lookup-with-column-index

lookup-with-column-index-after-column-write

lookup-approximate-sorted

lookup-approximate-sorted-after-column-write

new reverse/descending/duplicate lookup workloads.

B. Conditional aggregation improvements

Primary files:

packages/core/src/engine/services/criterion-range-cache-service.ts

packages/core/src/engine/services/range-aggregate-cache-service.ts

packages/core/src/engine/services/formula-binding-service.ts

packages/core/src/engine/services/formula-evaluation-service.ts

packages/core/src/engine/services/operation-service.ts

packages/core/src/engine/runtime-state.ts

Targets:

Build reusable criteria masks keyed by sheet, range, criterion AST/value, and version. Shared criteria formulas should not each rescan the same range.

Precompile criteria matchers. Avoid rebuilding match functions for every formula and every row during an edit.

Support delta paths for:

SUMIF/SUMIFS aggregate-value edits;

COUNTIF/COUNTIFS source edits;

criteria-range edits;

criteria-cell edits;

mixed SUMIFS/COUNTIFS families sharing criteria ranges.

For criteria-cell edits, compute old-match and new-match sets through cache-backed masks, then update dependent formulas. Do not full-recalculate the workbook unless the formula shape is unsupported.

Cache invalidation must be versioned by range and sheet mutation. Structural row/column changes must invalidate affected criteria masks.

Add counters:

criteriaCacheBuilds

criteriaCacheHits

criteriaMatcherBuilds

criteriaDeltaApplications

criteriaCacheInvalidations

Tests:

packages/core/src/__tests__/criterion-range-cache-service.test.ts

packages/core/src/__tests__/range-aggregate-cache-service.test.ts

packages/core/src/__tests__/operation-service.test.ts

packages/core/src/__tests__/formula-evaluation-service.test.ts

Expected benchmark impact:

conditional-aggregation-reused-ranges

conditional-aggregation-criteria-cell-edit

new shared/mixed criteria workloads.

C. 2D aggregate fast paths

Primary files:

packages/core/src/formula/simple-direct-aggregate-compile.ts

packages/core/src/engine/runtime-state.ts

packages/core/src/engine/services/formula-binding-service.ts

packages/core/src/engine/services/formula-evaluation-service.ts

packages/core/src/engine/services/operation-service.ts

packages/core/src/engine/services/range-aggregate-cache-service.ts

Targets:

Extend direct aggregate descriptors from one-dimensional ranges to rectangular ranges.

Support SUM/COUNT/MIN/MAX over rectangular ranges using production range caches.

For rectangular SUM/COUNT, use a 2D prefix/summed-area cache or a column-partitioned aggregate cache. Column-partitioned is simpler and safer initially; 2D prefix gives stronger performance for repeated rectangular ranges.

On source-cell edit inside a rectangular aggregate range, apply a delta when safe. Fall back to invalidation when unsupported.

Retarget rectangular descriptors on row/column insert/delete/move.

Tests:

packages/core/src/__tests__/range-aggregate-cache-service.test.ts

packages/core/src/__tests__/operation-service.test.ts

packages/core/src/__tests__/formula-evaluation-service.test.ts

Expected benchmark impact:

existing overlapping aggregate workloads;

new aggregate-2d-prefix-sum;

new aggregate-2d-after-cell-write.

D. Parser/template and initial-load improvements

Primary files:

packages/core/src/formula/template-bank.ts

packages/formula/src/formula-template-key.ts

packages/core/src/engine/services/formula-binding-service.ts

packages/headless/src/initial-sheet-load.ts

packages/headless/src/work-paper-runtime.ts

Targets:

Add parameterized formula-template caching, not just repeated literal formula-string caching. Unique formulas with the same AST shape but different numeric/string literals should share parse/lowering structure with literal slots.

Keep this semantic and general. Do not key on benchmark row counts, fixture names, or exact formula text from the benchmark.

Improve initial mixed sheet load by doing one content inspection pass and one write/bind/evaluate pass. Avoid separate scans for literal cells, formula cells, named expressions, and dependency setup when loading a full sheet.

Batch workbook/sheet creation in WorkPaper.buildFromSheets so build-many-sheets is not dominated by per-sheet setup overhead.

Add counters:

formulaTemplateHits

formulaTemplateMisses

parameterizedTemplateHits

initialLoadFormulaCells

initialLoadLiteralCells

Tests:

packages/core/src/__tests__/template-bank.test.ts

packages/core/src/__tests__/formula-template-normalization-service.test.ts

packages/core/src/__tests__/compiled-plan-service.test.ts

packages/headless/src/__tests__/work-paper-runtime.test.ts

Expected benchmark impact:

build-parser-cache-row-templates

build-parser-cache-mixed-templates

build-mixed-content

build-many-sheets

new unique-formula build/rebuild workloads.

E. Structural mutation and sheet rename improvements

Primary files:

packages/core/src/engine/services/formula-binding-service.ts

packages/core/src/engine/services/structure-service.ts

packages/core/src/workbook-store.ts

packages/headless/src/work-paper-runtime.ts

packages/core/src/engine/services/operation-service.ts

Targets:

Use dependency metadata to rewrite only formulas that actually reference a renamed sheet. Do not scan every formula in every sheet if the binding service already knows cross-sheet references.

Retarget direct aggregate, direct criteria, and direct lookup descriptors after row/column insert/delete/move. If precise retargeting is hard, invalidate only affected descriptors, not the whole engine.

Preserve lookup and criteria caches when a structural operation does not affect their ranges.

Tests:

packages/core/src/__tests__/structural-retargeting.test.ts

packages/core/src/__tests__/operation-service.test.ts

packages/headless/src/__tests__/work-paper-runtime.test.ts

Expected benchmark impact:

existing structural row/column workloads;

new lookup-index structural workloads;

new conditional-aggregation structural workloads;

new rename-sheet workloads.

F. Dynamic-array production improvements

Primary files:

packages/formula/src/builtins/lookup-sort-filter-builtins.ts

packages/core/src/engine/services/formula-evaluation-service.ts

packages/core/src/workbook-store.ts

packages/headless/src/work-paper-runtime.ts

Targets:

Implement stable spill metadata and spill-shape diffing.

For FILTER/SORT/UNIQUE, reuse prior spill allocations when shape is unchanged.

Invalidate dependent cells precisely when spill shape changes.

These workloads should remain outside the comparable scorecard until HyperFormula-equivalent verification exists.

Expected benchmark impact:

dynamic-array-filter

new dynamic-array-sort

new dynamic-array-unique.

4. Prioritized code-change sequence
Phase 0: Protect current evidence

Edit:

packages/benchmarks/src/__tests__/expanded-workloads.test.ts

scripts/gen-workpaper-vs-hyperformula-benchmark.ts

Add guard tests before adding new workloads:

existing workload IDs are unchanged;

existing workload leadershipOnly status is unchanged;

existing sample/warmup defaults remain 2/5;

every comparable workload has both WorkPaper and HyperFormula verification;

the artifact records verification mismatch as a hard failure, not a warning.

This prevents accidental benchmark manipulation while adding new rows.

Phase 1: Add robust reporting without changing benchmark behavior

Edit:

packages/benchmarks/src/stats.ts

packages/benchmarks/src/report-competitive-families.ts

packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts

scripts/gen-workpaper-vs-hyperformula-benchmark.ts

Add robust ratios, noise classification, family worst-case ratio, and current/public/holdout scorecards.

Do not remove the current mean-based fields. They are useful for compatibility but should stop being the only basis for “winner” claims.

Phase 2: Wire existing unwired fixtures

Edit:

packages/benchmarks/src/expanded-competitive-workloads.ts

packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded-additional-workloads.ts

packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts

packages/benchmarks/src/report-competitive-families.ts

Wire the fixtures already present:

buildRenameDependencySheets

buildNamedExpressionBenchSheet

buildLookupSearchModeReverseSheet

buildApproxLookupDescendingSheet

buildApproxLookupDuplicateSheet

buildConditionalAggregationSharedCriteriaSheet

buildConditionalAggregationMixedSheet

buildParserCacheUniqueFormulaSheet

build2dAggregateSheet

buildDynamicArraySortSheet

buildDynamicArrayUniqueSheet

This gives the expanded suite immediate coverage without inventing synthetic benchmark-only fixtures.

Phase 3: Lookup wins

Edit:

packages/core/src/engine/services/exact-column-index-service.ts

packages/core/src/engine/services/sorted-column-search-service.ts

packages/core/src/engine/direct-vector-lookup.ts

packages/core/src/engine/services/formula-binding-service.ts

packages/core/src/engine/services/formula-evaluation-service.ts

packages/core/src/engine/services/operation-service.ts

packages/core/src/perf/engine-counters.ts

Goal: make WorkPaper robustly win exact indexed, reverse exact, ascending approximate, descending approximate, duplicate approximate, and after-write lookup workloads.

This should be the first engine phase because your local red rows include lookup-approximate-sorted, and the new lookup workloads will otherwise expose correctness/perf gaps.

Phase 4: Conditional aggregation wins

Edit:

packages/core/src/engine/services/criterion-range-cache-service.ts

packages/core/src/engine/services/range-aggregate-cache-service.ts

packages/core/src/engine/services/formula-binding-service.ts

packages/core/src/engine/services/formula-evaluation-service.ts

packages/core/src/engine/services/operation-service.ts

packages/core/src/perf/engine-counters.ts

Goal: shared criteria caches, criteria-cell edit deltas, mixed formula support.

Phase 5: 2D aggregate wins

Edit:

packages/core/src/formula/simple-direct-aggregate-compile.ts

packages/core/src/engine/runtime-state.ts

packages/core/src/engine/services/formula-binding-service.ts

packages/core/src/engine/services/formula-evaluation-service.ts

packages/core/src/engine/services/operation-service.ts

packages/core/src/engine/services/range-aggregate-cache-service.ts

Goal: rectangular aggregate descriptors and safe delta/invalidation behavior.

Phase 6: Build/parser wins

Edit:

packages/core/src/formula/template-bank.ts

packages/formula/src/formula-template-key.ts

packages/core/src/engine/services/formula-binding-service.ts

packages/headless/src/initial-sheet-load.ts

packages/headless/src/work-paper-runtime.ts

Goal: reduce overhead on mixed-content, many-sheets, and unique-formula build/rebuild without workload-specific hacks.

Phase 7: Structural and rename wins

Edit:

packages/core/src/engine/services/formula-binding-service.ts

packages/core/src/engine/services/structure-service.ts

packages/core/src/workbook-store.ts

packages/headless/src/work-paper-runtime.ts

Goal: sheet rename and structural row/column edits retarget only affected formulas/descriptors.

Phase 8: Dynamic-array leadership polish

Edit:

packages/formula/src/builtins/lookup-sort-filter-builtins.ts

packages/core/src/engine/services/formula-evaluation-service.ts

packages/core/src/workbook-store.ts

packages/headless/src/work-paper-runtime.ts

Goal: improve WorkPaper-only capability workloads, but keep them separate from comparable scorecards.

5. Tests and validation commands

Use the root scripts already present.

Baseline correctness

Run:

Bash
pnpm typecheck
pnpm test
pnpm test:correctness:core
pnpm test:correctness:formula
pnpm workpaper:parity:check

If those are too broad during development, run targeted Vitest files first:

Bash
pnpm exec vitest --run packages/benchmarks/src/__tests__/expanded-workloads.test.ts
pnpm exec vitest --run packages/core/src/__tests__/operation-service.test.ts
pnpm exec vitest --run packages/headless/src/__tests__/work-paper-runtime.test.ts
Lookup-specific validation
Bash
pnpm exec vitest --run \
  packages/core/src/__tests__/exact-column-index-service.test.ts \
  packages/core/src/__tests__/sorted-column-search-service.test.ts \
  packages/core/src/__tests__/direct-vector-lookup.test.ts \
  packages/core/src/__tests__/lookup-service.test.ts \
  packages/core/src/__tests__/lookup-column-owner.test.ts \
  packages/core/src/__tests__/operation-service.test.ts

Add tests for:

reverse exact lookup with duplicates;

ascending approximate duplicate groups;

descending approximate groups;

tail append preserving monotonicity;

interior write invalidating monotonic metadata;

structural delete invalidating lookup indexes.

Criteria and aggregate validation
Bash
pnpm exec vitest --run \
  packages/core/src/__tests__/criterion-range-cache-service.test.ts \
  packages/core/src/__tests__/range-aggregate-cache-service.test.ts \
  packages/core/src/__tests__/formula-evaluation-service.test.ts \
  packages/core/src/__tests__/operation-service.test.ts

Add tests for:

criteria-cell edit;

criteria-range edit;

aggregate-value edit;

mixed SUMIFS/COUNTIFS reuse;

rectangular SUM after single-cell edit;

rectangular aggregate after row/column insert/delete.

Parser/build validation
Bash
pnpm exec vitest --run \
  packages/core/src/__tests__/template-bank.test.ts \
  packages/core/src/__tests__/formula-template-normalization-service.test.ts \
  packages/core/src/__tests__/compiled-plan-service.test.ts \
  packages/headless/src/__tests__/work-paper-runtime.test.ts

Add tests that prove parameterized template caching preserves correctness with different literals and references.

Benchmark smoke
Bash
pnpm bench:smoke
pnpm bench:workpaper:competitive

The default run should remain 2 warmups and 5 samples.

High-confidence performance run

Add or use CLI flags for larger sample counts without changing defaults:

Bash
pnpm bench:workpaper:competitive -- --warmup-count 5 --sample-count 25

If the current runner does not support these flags, add them to packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts, but keep default values unchanged.

Artifact generation/check
Bash
pnpm workpaper:bench:competitive:generate
pnpm workpaper:bench:competitive:check

After adding current/public/holdout report sections, check should fail if the checked-in artifact is stale.

Full benchmark
Bash
pnpm bench

Acceptance target for the expanded public scorecard:

zero verified comparable HyperFormula wins on the high-confidence run;

no comparable family has familyWorstCaseWorkPaperRatio >= 1.0;

all noisy/tie rows are labeled as such, not counted as WorkPaper wins;

leadership-only dynamic-array rows are reported separately.

6. Red flags that would count as cheating or accidental cheating

Benchmark-specific branches in engine code, such as checking workload IDs, exact row counts, exact formula strings, or fixture-specific sheet names.

Changing existing workload sizes, edit locations, sampling defaults, or verification payloads to remove losses.

Moving a HyperFormula-supported workload to leadership-only because WorkPaper loses it.

Comparing fewer output cells after adding a harder workload.

Accepting rounded verification values where the engines differ materially.

Timing WorkPaper after caches are prebuilt but timing HyperFormula from cold state, or vice versa.

Running verification inside one engine’s timed section but outside the other’s timed section.

Using different formulas, different sheet shapes, different named-expression setup, or different config between WorkPaper and HyperFormula.

Changing scoring from mean to median only when median favors WorkPaper. The fix is to report mean, median, p95, confidence, and noise classification consistently.

Dropping noisy losses from the public scorecard. Noisy rows should be labeled tie-noisy, not hidden.

Implementing lookup shortcuts that only work for sequential integers, benchmark key strings, or exact fixture distributions.

Assuming approximate lookup vectors are sorted after arbitrary writes without validating monotonicity.

Failing to invalidate criteria/lookup/aggregate caches after structural edits, sheet rename, named-expression edits, or formula edits.

Treating dynamic-array leadership rows as WorkPaper-vs-HyperFormula wins while HyperFormula cannot run equivalent semantics.

Regenerating the baseline on a favorable machine and presenting it as universal without host metadata, sample count, and confidence intervals.

Recommended immediate next commit set

First commit: reporting integrity only.

Files:

packages/benchmarks/src/stats.ts

packages/benchmarks/src/report-competitive-families.ts

packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts

scripts/gen-workpaper-vs-hyperformula-benchmark.ts

packages/benchmarks/src/__tests__/expanded-workloads.test.ts

Second commit: wire new workloads using existing fixtures, with no engine changes.

Files:

packages/benchmarks/src/expanded-competitive-workloads.ts

packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded-additional-workloads.ts

packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts

packages/benchmarks/src/report-competitive-families.ts

packages/benchmarks/src/workpaper-benchmark-fixtures.ts only if a repeated-shared-index or larger rename fixture is needed.

Third commit: lookup production improvements.

Fourth commit: criteria-cache and 2D aggregate production improvements.

Fifth commit: parser/build and structural rename improvements.

This order makes the suite harder before tuning, preserves the current evidence trail, and forces WorkPaper wins to come from production engine behavior rather than benchmark control.

## Implementation Log - 2026-04-29T12:27Z

This chronological log entry records an intermediate run. It does not override
the validated plan at the top of this document. This cycle continued
implementation against the validated oracle design and did not change workload
definitions, scoring, sampling, or workload sizes.

Production implementation added:

- Fresh literal and mixed initial sheet loads now use a cached row-major
  `SheetGrid` setter for lower cold-build map overhead.
- Non-uniform approximate direct lookup operand edits now use prepared sorted
  lookup results inside `operation-service`, so duplicate-key approximate
  `MATCH` mutations can return a compact two-cell change without generic direct
  formula evaluation.
- Added focused headless coverage for duplicate approximate `MATCH` operand
  edits and compact direct-path counters.

Validation:

- Focused tests passed for headless runtime lookup paths and direct lookup
  services.
- `@bilig/core`, `@bilig/headless`, and `@bilig/benchmarks` builds passed.
- Focused timing after warmup showed the former red rows winning by mean:
  `build-mixed-content` `0.889x`, `aggregate-overlapping-sliding-window`
  `0.509x`, and `lookup-approximate-duplicates` `0.742x`.
- `pnpm workpaper:bench:competitive:generate` regenerated
  `packages/benchmarks/baselines/workpaper-vs-hyperformula.json` at
  `2026-04-29T12:27:03.667Z`.
- `pnpm workpaper:bench:competitive:check` passed.

Intermediate official scorecard:

- Overall: `42` WorkPaper wins, `4` HyperFormula wins, `46` comparable.
- Public: `34` WorkPaper wins, `4` HyperFormula wins.
- Holdout: `8` WorkPaper wins, `0` HyperFormula wins.
- Remaining HyperFormula rows: `build-mixed-content` is a non-overlap red row in the official 5-sample artifact; `build-from-sheets`, `build-dense-literals`, and `aggregate-overlapping-sliding-window` are confidence-overlap noisy.

Oracle access status at that moment:

- Browser Use attach was unavailable in that cycle with
  `Browser turn does not belong to this IAB pipe`.
- Computer Use timed out attempting to read ChatGPT Atlas app state.
- At that moment the ChatGPT thread content had not yet been captured. The
  first-prompt response was later captured successfully and saved in
  `docs/workpaper-oracle-chatgpt-first-response-2026-04-29.md`.

## Implementation Log - 2026-04-29T13:05Z

This chronological log entry records an intermediate run. It does not override
the validated plan at the top of this document. The cycle followed the validated
design with additional production cold-build work only. No benchmark
definitions, scorecard rules, sampling, or workload sizes were changed.

Implemented:

- Static `WorkPaper.buildFromSheets()` now avoids repeated sheet-name entry
  materialization and inspection maps, carrying one ordered sheet-entry vector
  through validation, sheet creation, load, and dimension-cache initialization.
- `SpreadsheetEngine.createSheetForInitialization()` returns the created sheet id
  for direct initialization callers.
- `WorkbookStore.createLogicalAxisIdEnsurer()` lets fresh load paths reuse the
  sheet lookup for logical row/column ids without precreating ids for sparse
  cells.

Validated:

- Focused literal/mixed-load and runtime tests passed.
- `@bilig/core`, `@bilig/headless`, and `@bilig/benchmarks` builds passed.
- `pnpm lint` passed.
- Competitive baseline regenerated at `2026-04-29T13:05:50.243Z`.
- `pnpm workpaper:bench:competitive:check` passed.

Scorecard after this tranche:

- Overall: `44` WorkPaper wins, `2` HyperFormula wins, `46` comparable.
- Public: `36` WorkPaper wins, `2` HyperFormula wins.
- Holdout: `8` WorkPaper wins, `0` HyperFormula wins.
- Remaining HyperFormula rows: `build-mixed-content` and
  `aggregate-overlapping-sliding-window`, both confidence-overlap noisy.

Oracle access at that moment:

- Browser Use retry after resetting the Node REPL reported no available
  in-app-browser backend, so no new oracle response was pulled in that cycle.
  The first-prompt response was later captured successfully and saved in
  `docs/workpaper-oracle-chatgpt-first-response-2026-04-29.md`.

## Implementation Log - 2026-04-29T13:55Z

This chronological log entry records an intermediate run. It does not override
the validated plan at the top of this document. Implementation continued against
the validated oracle design in production engine/headless code only. Benchmark
definitions, scorecard rules, sampling, and workload sizes were not changed.

Implemented:

- Initial formula loading now skips the volatile-formula scan when the runtime
  knows there are no volatile formulas.
- Fresh formula binding now skips empty region-graph subscription replacement
  for non-region direct scalar formulas.
- Direct lookup operand mutations now write numeric formula results through the
  numeric terminal writer and carry compact row/column metadata for the second
  changed cell.
- The single-direct-aggregate numeric mutation fast path now uses direct counter
  increments and returns compact row/column metadata for its formula result.

Validated:

- Focused runtime, lookup, aggregate, and initialization tests passed.
- `@bilig/core` and `@bilig/headless` builds passed.
- Competitive baseline regenerated at `2026-04-29T13:52:59.186Z`.
- `pnpm workpaper:bench:competitive:check` passed.

Current scorecard:

- Overall: `44` WorkPaper wins, `2` HyperFormula wins, `46` comparable.
- Public: `36` WorkPaper wins, `2` HyperFormula wins.
- Holdout: `8` WorkPaper wins, `0` HyperFormula wins.
- Remaining HyperFormula rows: `build-mixed-content` and
  `aggregate-overlapping-sliding-window`; both have overlapping confidence
  intervals in the latest official 5-sample artifact.

Oracle access at that moment:

- Browser Use was retried with the required `iab` backend and was unavailable in
  that cycle because no Codex in-app-browser backend was visible. The
  first-prompt response was later captured successfully and saved in
  `docs/workpaper-oracle-chatgpt-first-response-2026-04-29.md`.

## Implementation Log - 2026-04-29T14:47Z

This chronological log entry is the latest official benchmark evidence currently
referenced by the validated plan. The cycle stayed on production engine/headless
paths and did not alter benchmark definitions, scorecard rules, sampling, or
workload sizes.

Work performed:

- Investigated a fresh-initial-formula changed-scratch deferral for cold mixed
  builds. Focused tests and builds passed, but the official competitive
  artifact worsened `build-mixed-content`, so the edit was reverted and not
  retained.
- Regenerated the competitive baseline after the revert and confirmed the
  benchmark artifact with `pnpm workpaper:bench:competitive:check`.

Validation:

- Focused initial-load and runtime tests passed.
- `@bilig/core` and `@bilig/headless` builds passed.
- `pnpm lint` passed.
- Competitive baseline regenerated at `2026-04-29T14:47:16.831Z`.
- `pnpm workpaper:bench:competitive:check` passed.

Current scorecard:

- Overall: `44` WorkPaper wins, `2` HyperFormula wins, `46` comparable.
- Public: `36` WorkPaper wins, `2` HyperFormula wins.
- Holdout: `8` WorkPaper wins, `0` HyperFormula wins.
- Remaining HyperFormula rows in this artifact: `build-mixed-content` and
  `structural-delete-rows`; both are confidence-overlap rows. The prior
  `aggregate-overlapping-sliding-window` row is green in this artifact.

CI status:

- Full `pnpm run ci` was not rerun after the reverted experiment. The latest
  full CI evidence from this branch still has all tests passing and fails only
  the global coverage threshold (`87.53%` lines and `87.67%` statements versus
  `91%`).

## Browser Use Re-Capture - 2026-04-29

The ChatGPT thread was re-opened through the Browser Use in-app browser backend
at `https://chatgpt.com/c/69f18c28-44b8-83e8-bd3a-eb2228c71843`.

The first-prompt oracle response is the long assistant message in the thread,
with extracted length `34979` characters. It was saved verbatim to
`docs/workpaper-oracle-chatgpt-first-response-2026-04-29.md`.

Validation result: that raw capture is already included byte-for-byte in this
document's `## Appendix: Raw First-Prompt Oracle Response` section. The short
latest assistant message in the thread is only a follow-up acknowledgement, not
the first-prompt oracle plan.
