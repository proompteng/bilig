# Local WorkPaper Benchmark Walkthrough

This walkthrough shows how to run the local WorkPaper benchmark checks for
`@bilig/headless` and how to read the output without overstating the result.

Use it when you want to verify that the checked-in benchmark artifact is current,
smoke-test the benchmark harness on your machine, or evaluate a performance
change before publishing a benchmark claim.

## Setup From A Fresh Checkout

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig
corepack enable
pnpm install --frozen-lockfile
```

The benchmark sources live in `packages/benchmarks/src/`. The checked-in public
artifact lives at
`packages/benchmarks/baselines/workpaper-vs-hyperformula.json`.

## Check The Committed Artifact

Start with the fast shape and coverage check:

```sh
pnpm workpaper:bench:competitive:check
```

Expected successful output starts with:

```json
{
  "mode": "check",
  "outputPath": "/path/to/bilig/packages/benchmarks/baselines/workpaper-vs-hyperformula.json",
  "workloads": [
    "build-from-sheets",
    "build-dense-literals"
  ]
}
```

The `workloads` array is longer in the real output. The check is not a fresh
timing run; it verifies that the committed artifact still has the expected
benchmark schema and workload coverage.

To inspect the headline scorecard in the committed artifact, run:

```sh
jq '{
  generatedAt,
  benchmark,
  scorecard: {
    comparableCount: .scorecard.comparableCount,
    workpaperWins: .scorecard.workpaperWins,
    hyperformulaWins: .scorecard.hyperformulaWins,
    directionalMeanRatioGeomean: .scorecard.directionalMeanRatioGeomean,
    directionalP95RatioGeomean: .scorecard.directionalP95RatioGeomean,
    worstP95RatioWorkload: .scorecard.worstP95RatioWorkload
  }
}' packages/benchmarks/baselines/workpaper-vs-hyperformula.json
```

At the time of this walkthrough, that reports:

```json
{
  "generatedAt": "2026-05-16T02:12:30.841Z",
  "benchmark": {
    "sampleCount": 5,
    "warmupCount": 2
  },
  "scorecard": {
    "comparableCount": 57,
    "workpaperWins": 42,
    "hyperformulaWins": 15,
    "directionalMeanRatioGeomean": 0.7553949494105464,
    "directionalP95RatioGeomean": 0.7510834854399419,
    "worstP95RatioWorkload": "structural-append-formula-rows"
  }
}
```

Ratios below `1.0` mean WorkPaper is faster for that aggregate metric. This is
still a mean-latency benchmark claim, not a blanket "faster on every p95 row" or
"faster for every possible spreadsheet" claim.

## Smoke-Test A Local Timing Run

For a quick local harness check, run a reduced one-sample benchmark:

```sh
pnpm --silent bench:workpaper:competitive -- --sample-count 1 --warmup-count 0 > /tmp/workpaper-local-benchmark.json
```

Then summarize it:

```sh
jq '{
  suite,
  scorecard: {
    comparableCount: .scorecard.comparableCount,
    workpaperWins: .scorecard.workpaperWins,
    hyperformulaWins: .scorecard.hyperformulaWins,
    directionalMeanRatioGeomean: .scorecard.directionalMeanRatioGeomean,
    directionalP95RatioGeomean: .scorecard.directionalP95RatioGeomean
  },
  firstComparable: (
    .results[]
    | select(.workload == "build-from-sheets")
    | {
      workload,
      fasterEngine: .comparison.fasterEngine,
      meanRatio: .comparison.workpaperToHyperFormulaMeanRatio,
      p95Ratio: .comparison.workpaperToHyperFormulaP95Ratio
    }
  )
}' /tmp/workpaper-local-benchmark.json
```

A one-sample run is intentionally noisy. Use it to confirm that the benchmark
harness runs on your machine, not to replace the checked-in 5-sample artifact.

## Evaluate A Benchmark Change

If you intentionally change runtime performance, regenerate the committed
artifact with the default benchmark sampling:

```sh
pnpm workpaper:bench:competitive:generate
pnpm workpaper:bench:competitive:check
```

Then compare the diff in:

```sh
git diff -- packages/benchmarks/baselines/workpaper-vs-hyperformula.json
```

The highest-signal fields to review are:

- `benchmark.sampleCount` and `benchmark.warmupCount`: verify sampling did not
  shrink accidentally.
- `scorecard.workpaperWins` and `scorecard.hyperformulaWins`: read as
  direction-of-mean wins across comparable workloads.
- `scorecard.directionalMeanRatioGeomean`: aggregate mean ratio, where lower is
  better for WorkPaper.
- `scorecard.directionalP95RatioGeomean`: aggregate p95 ratio, also lower is
  better for WorkPaper.
- `scorecard.worstP95RatioWorkload`: fastest way to spot the current tail-risk
  row.
- `results[].comparison.workpaperToHyperFormulaMeanRatio`: per-workload mean
  ratio.
- `results[].comparison.workpaperToHyperFormulaP95Ratio`: per-workload p95
  ratio.

Do not change workload sizes, scoring, sampling, or fixture definitions to make
a claim look better. If a real rerun moves a row red, update the artifact and
the public docs, then fix the production runtime path.

## Related Reading

- [`docs/headless-workpaper-benchmark-evidence.md`](headless-workpaper-benchmark-evidence.md)
- [`docs/what-workpaper-benchmark-proves.md`](what-workpaper-benchmark-proves.md)
- [`docs/hyperformula-alternative-headless-workpaper.md`](hyperformula-alternative-headless-workpaper.md)
- [`packages/benchmarks/baselines/workpaper-vs-hyperformula.json`](../packages/benchmarks/baselines/workpaper-vs-hyperformula.json)
