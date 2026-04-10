# @bilig/benchmarks

Benchmark scenarios, metrics helpers, and workbook generators for bilig.

## Install

```bash
npm install @bilig/benchmarks
```

## Package entrypoints

- ESM: `./dist/index.js`
- Types: `./dist/index.d.ts`
- Corpus: `./dist/workbook-corpus.js`

## WorkPaper baseline

The repo tracks a checked-in WorkPaper benchmark artifact at
`packages/benchmarks/baselines/workpaper-baseline.json`.

Refresh or validate it with:

```bash
pnpm workpaper:bench:generate
pnpm workpaper:bench:check
```

## WorkPaper vs HyperFormula artifact

The repo also tracks a checked-in competitive benchmark artifact at
`packages/benchmarks/baselines/workpaper-vs-hyperformula.json`.

It records directly comparable workloads against HyperFormula `3.2.0`, plus leadership workloads
that must be labeled unsupported instead of silently omitted.

Refresh or validate it with:

```bash
pnpm workpaper:bench:competitive:generate
pnpm workpaper:bench:competitive:check
```

## Named giant-workbook corpus

The package now ships deterministic named workbook cases for giant-data restore and warm-start
contracts:

- `dense-mixed-100k`
- `dense-mixed-250k`
- `analysis-multisheet-100k`
- `analysis-multisheet-250k`

Use `buildWorkbookBenchmarkCorpus(...)` to materialize an exact-size workbook snapshot and stable
viewport metadata for CI and perf harnesses.

This package is part of the [bilig](https://github.com/proompteng/bilig) monorepo.
