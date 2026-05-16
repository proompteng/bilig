# Headless WorkPaper Benchmark Evidence

Status: public evidence note for `@bilig/headless`

This note keeps the public performance claim auditable from checked-in repo
artifacts instead of README copy alone.

## Current Artifact

The primary workbook-wide decision artifact is
`packages/benchmarks/baselines/workpaper-vs-hyperformula.json`.

The additional scalar formula-engine comparison artifact is
`packages/benchmarks/baselines/workpaper-vs-truecalc.json`.

The additional limited workbook-wide comparison artifact is
`packages/benchmarks/baselines/workpaper-vs-xlsx-calc.json`.

The goal-tracking scorecard for broad headless-engine performance leadership is
`packages/benchmarks/baselines/headless-performance-leadership-scorecard.json`.
It intentionally stays `active-not-achieved` until the checked-in evidence covers
at least two workbook-wide direct headless spreadsheet engines across the full
eligible workload set and every comparable workload wins both mean and p95
latency. Scalar formula-engine lanes and partial workbook-wide lanes are tracked
as useful evidence, but they do not satisfy broad coverage alone.

Current checked-in metadata:

- generated at `2026-05-16T04:11:59.799Z`
- host: macOS `arm64`, Node `v24.3.0`
- benchmark sampling: `5` measured samples after `2` warmup samples
- WorkPaper package: `@bilig/headless` `0.16.0`
- comparison engine: HyperFormula `3.2.0`, local checkout commit
  `9a510a2acb97c3d3490f9e3b9e961a1c4a98b9ad`, GPL-v3 license key
- scalar formula comparison engine: TrueCalc `0.6.4`, `7` comparable scalar
  workloads via `@truecalc/core`
- limited workbook-wide comparison engine: xlsx-calc `0.9.2`, `4` comparable
  recalculation workloads covering aggregate, exact lookup, approximate lookup,
  and formula-chain families

## What The Claim Is

The current scorecard is not a blanket performance-leadership claim. A fresh
checked-in run shows WorkPaper leading HyperFormula on most, but not all,
directly comparable workbook-wide headless spreadsheet-engine workloads. The
current checked-in artifact records `42/57` mean-latency wins:

| Lane    | Comparable Workloads | WorkPaper Mean Wins | HyperFormula Mean Wins |
| ------- | -------------------: | ------------------: | ---------------------: |
| Overall |                 `57` |                `42` |                   `15` |
| Public  |                 `40` |                `31` |                    `9` |
| Holdout |                 `17` |                 `11` |                    `6` |

The overall directional mean-ratio geomean is `0.7216546733829703`. The overall
directional p95-ratio geomean is `0.7402574840907257`. Ratios below `1.0` mean
WorkPaper is faster for that metric.

The current worst mean row is `cross-sheet-dashboard-recalc`, with a mean
ratio of `3.209829169368815`. The current worst p95 row is
`cross-sheet-dashboard-recalc`, with a p95 ratio of `3.2809559202634913`. The
headless leadership scorecard currently records `41/57` workloads winning both
mean and p95 against HyperFormula.

It is also not a blanket "fastest against every formula evaluator" claim. The
TrueCalc scalar lane currently reports `0/7` WorkPaper mean+p95 wins, with a
directional mean-ratio geomean of `6.222935520223555`. That lane is intentionally
kept in the leadership scorecard as a blocker map, not as marketing copy.

The xlsx-calc lane is a direct workbook-wide recalculation comparison for the
formula families both engines can evaluate equivalently. It currently reports
`4/4` WorkPaper mean+p95 wins with a directional mean-ratio geomean of
`0.057749710674702796`, but it covers only `4/57` eligible workload rows, so the
scorecard treats it as partial coverage rather than proof of blanket leadership.

## How To Read The p95 Caveat

The `42/57` count is about mean latency: for each winning comparable workload
row, WorkPaper's average measured time is lower than HyperFormula's average
measured time. Mean wins are useful because they summarize the normal cost of
each workload, but they do not prove every slower tail sample has been
eliminated, and the current scorecard does not yet win every mean row.

Each p95 row asks a different question: "near the slow end of this workload's
sample set, which engine was faster?" A single row can lose on p95 even when its
mean wins, because a small number of slower samples can move the tail without
moving the average enough to flip the mean result.

The p95 geomean is an aggregate across the per-workload p95 ratios. It can stay
below `1.0` while one individual p95 row is above `1.0`, because the aggregate
is balanced by the other p95 rows where WorkPaper has enough margin. Read the
current result as: WorkPaper leads the overall mean and p95 aggregate, but the
repo is not claiming "faster on every row" until the mean and p95 holdouts are
fixed.

## What Is Measured

Scorecard-eligible families cover:

- workbook build and rebuild paths
- runtime restore from snapshot
- sheet lifecycle and named-expression changes
- cross-sheet scalar and aggregate recalculation
- dirty execution after single edits, chains, fanout, mixed frontiers, and
  formula edits
- batch edits, suspended batches, and undo-including batches
- structural row and column inserts, deletes, and moves
- dense and sparse range reads
- 2D, overlapping, sliding-window, and conditional aggregation
- exact lookup, INDEX/MATCH, INDEX reference, approximate lookup, after-write
  lookup, and text lookup

The scorecard excludes the `config-toggle` control family and `dynamic-array`
leadership-only family from the directly comparable win count.

## How To Verify

Check that the committed artifact still has the expected workload coverage and
shape:

```bash
pnpm workpaper:bench:competitive:check
pnpm workpaper:bench:truecalc:check
pnpm workpaper:bench:xlsx-calc:check
pnpm headless:performance:check
```

Regenerate timing evidence only when intentionally refreshing the benchmark
artifact:

```bash
pnpm workpaper:bench:competitive:generate
pnpm workpaper:bench:competitive:check
pnpm workpaper:bench:truecalc:generate
pnpm workpaper:bench:truecalc:check
pnpm workpaper:bench:xlsx-calc:generate
pnpm workpaper:bench:xlsx-calc:check
```

Do not change workload sizes, sampling, scoring, or definitions to preserve a
claim. If a rerun moves a row red, update the artifact, update this note, and
fix the production engine path rather than hiding the loss.

If a workload family is missing, a row looks too synthetic, or the p95 wording
is still too broad, use the public benchmark critique thread:
<https://github.com/proompteng/bilig/discussions/340>.
