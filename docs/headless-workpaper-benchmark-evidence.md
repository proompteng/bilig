# Headless WorkPaper Benchmark Evidence

Status: public evidence note for `@bilig/headless`

This note keeps the public performance claim auditable from checked-in repo
artifacts instead of README copy alone.

## Current Artifact

The decision artifact is
`packages/benchmarks/baselines/workpaper-vs-hyperformula.json`.

Current checked-in metadata:

- generated at `2026-05-08T15:00:27.603Z`
- host: macOS `arm64`, Node `v24.3.0`
- benchmark sampling: `5` measured samples after `2` warmup samples
- WorkPaper package: `@bilig/headless`
- comparison engine: HyperFormula `3.2.0`, local checkout commit
  `9a510a2acb97c3d3490f9e3b9e961a1c4a98b9ad`, GPL-v3 license key

## What The Claim Is

The current scorecard claim is a mean-latency claim across directly comparable
headless spreadsheet-engine workloads:

| Lane    | Comparable Workloads | WorkPaper Mean Wins | HyperFormula Mean Wins |
| ------- | -------------------: | ------------------: | ---------------------: |
| Overall |                 `46` |                `46` |                    `0` |
| Public  |                 `38` |                `38` |                    `0` |
| Holdout |                  `8` |                 `8` |                    `0` |

The overall directional mean-ratio geomean is `0.521767150331573`. The overall
directional p95-ratio geomean is `0.5359737705859149`. Ratios below `1.0` mean
WorkPaper is faster for that metric.

The closest overall mean win is `lookup-approximate-duplicates` at
`0.9108460643406784`. The closest public-lane mean win is
`build-mixed-content` at `0.9017762124360226`.

This is not a blanket "faster on every p95 row" claim. The current worst p95
ratio is `1.043096403103571` on `lookup-approximate-duplicates`, so the honest
public claim is `46/46` mean wins with an overall p95 geomean lead and one known
p95 holdout that still needs margin work.

## How To Read The p95 Caveat

The `46/46` count is about mean latency: for each comparable workload row,
WorkPaper's average measured time is lower than HyperFormula's average measured
time. Mean wins are useful for the headline because they summarize the normal
cost of each workload, but they do not prove every slower tail sample has been
eliminated.

Each p95 row asks a different question: "near the slow end of this workload's
sample set, which engine was faster?" A single row can lose on p95 even when its
mean wins, because a small number of slower samples can move the tail without
moving the average enough to flip the mean result.

The p95 geomean is an aggregate across the per-workload p95 ratios. It can stay
below `1.0` while one individual p95 row is above `1.0`, because the aggregate
is balanced by the other p95 rows where WorkPaper has enough margin. Read the
current result as: WorkPaper wins every comparable mean row and leads the
overall p95 aggregate, but the repo is not claiming "faster on every p95 row"
until the known p95 holdout is fixed.

## What Is Measured

Scorecard-eligible families cover:

- workbook build and rebuild paths
- runtime restore from snapshot
- sheet lifecycle and named-expression changes
- dirty execution after single edits, chains, fanout, mixed frontiers, and
  formula edits
- batch edits, suspended batches, and undo-including batches
- structural row and column inserts, deletes, and moves
- range reads
- 2D, overlapping, sliding-window, and conditional aggregation
- exact lookup, approximate lookup, after-write lookup, and text lookup

The scorecard excludes the `config-toggle` control family and `dynamic-array`
leadership-only family from the directly comparable win count.

## How To Verify

Check that the committed artifact still has the expected workload coverage and
shape:

```bash
pnpm workpaper:bench:competitive:check
```

Regenerate timing evidence only when intentionally refreshing the benchmark
artifact:

```bash
pnpm workpaper:bench:competitive:generate
pnpm workpaper:bench:competitive:check
```

Do not change workload sizes, sampling, scoring, or definitions to preserve a
claim. If a rerun moves a row red, update the artifact, update this note, and
fix the production engine path rather than hiding the loss.

If a workload family is missing, a row looks too synthetic, or the p95 wording
is still too broad, use the public benchmark critique thread:
<https://github.com/proompteng/bilig/discussions/340>.
