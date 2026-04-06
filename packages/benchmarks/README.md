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
