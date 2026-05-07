# Headless WorkPaper Benchmark Evidence

Status: public evidence note for `@bilig/headless`

This note keeps the public performance claim auditable from checked-in repo
artifacts instead of README copy alone.

## Current Artifact

The decision artifact is
`packages/benchmarks/baselines/workpaper-vs-hyperformula.json`.

Current checked-in metadata:

- generated at `2026-05-06T14:54:57.091Z`
- host: macOS `arm64`, Node `v24.3.0`
- benchmark sampling: `5` measured samples after `2` warmup samples
- WorkPaper package: `@bilig/headless`
- comparison engine: HyperFormula `3.2.0`, local checkout commit
  `6de904b8876f920f287b63a95934c479acf78307`, GPL-v3 license key

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
