# Public Adoption Kit

Status: shareable evaluation material for `@bilig/headless` and the public
`proompteng/bilig` repository.

Use this page when someone asks why they should look at bilig, how to try it
quickly, or what claims are already backed by repository artifacts. Keep the
language direct and evidence-based; do not inflate benchmark or Excel-compatibility
claims.

## One-Line Positioning

`@bilig/headless` is a headless spreadsheet engine for Node services, coding
agents, local-first workbook automation, and formula-backed business workflows.

Shorter version:

> A headless spreadsheet engine for agents and Node services.

## Who Should Care

- teams building spreadsheet-backed automation without opening a browser grid
- agent workflows that need stable workbook APIs instead of screen scraping
- services that need formulas, structural edits, persistence, and validation
- developers comparing spreadsheet-engine behavior against HyperFormula-style
  workloads
- contributors interested in formula parity, workbook correctness, local-first
  sync, browser grid behavior, or WASM acceleration

## Proof Points To Link

- GitHub: <https://github.com/proompteng/bilig>
- Website: <https://proompteng.github.io/bilig/>
- npm: <https://www.npmjs.com/package/@bilig/headless>
- Package README:
  [`packages/headless/README.md`](../packages/headless/README.md)
- Runnable external-consumer example:
  [`examples/headless-workpaper`](../examples/headless-workpaper)
- Benchmark evidence:
  [`docs/headless-workpaper-benchmark-evidence.md`](headless-workpaper-benchmark-evidence.md)
- Technical article:
  [`docs/why-agents-need-workbook-apis.md`](why-agents-need-workbook-apis.md)
- Persistence article:
  [`docs/persisting-formula-backed-workpaper-documents-in-node.md`](persisting-formula-backed-workpaper-documents-in-node.md)
- GitHub stars growth plan:
  [`docs/github-stars-growth-plan.md`](github-stars-growth-plan.md)
- Launch post draft:
  [`docs/launch-post-headless-workpaper.md`](launch-post-headless-workpaper.md)
- Launch discussion: <https://github.com/proompteng/bilig/discussions/18>
- Starter issues: [`docs/starter-issues.md`](starter-issues.md)
- Public API notes: [`docs/public-api.md`](public-api.md)
- Contributing guide: [`CONTRIBUTING.md`](../CONTRIBUTING.md)

The current checked-in benchmark evidence records WorkPaper `46/46` mean wins
on scorecard-eligible comparable workloads against HyperFormula-style workloads:
`38/38` public and `8/8` holdout. Keep the p95 nuance attached to that claim:
it is not a blanket "faster on every p95 row" statement.

## Ten-Minute Evaluation

Use the published package path when the evaluator does not want to clone the
monorepo:

```sh
mkdir bilig-headless-eval
cd bilig-headless-eval
npm init -y
npm pkg set type=module
npm install @bilig/headless
```

Create `eval.mjs`:

```js
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets({
  Revenue: [
    ['Region', 'Customers', 'ARPA', 'Revenue'],
    ['West', 20, 1200, '=B2*C2'],
    ['East', 30, 250, '=B3*C3'],
    ['Central', 18, 300, '=B4*C4'],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Total revenue', '=SUM(Revenue!D2:D4)'],
    ['West customers', '=SUMIF(Revenue!A2:A4,"West",Revenue!B2:B4)'],
  ],
})

const summary = workbook.getSheetId('Summary')
if (summary === undefined) {
  throw new Error('Summary sheet was not created')
}

const total = workbook.getCellValue({ sheet: summary, row: 1, col: 1 })
const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved))

console.log({
  total,
  sheets: restored.getSheetNames(),
})
```

Run it:

```sh
node eval.mjs
```

Use the repository path when the evaluator wants the maintained example and
local validation:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm start
```

## Shareable Copy

Short:

> `@bilig/headless` is a headless spreadsheet engine for agents and Node
> services: formulas, structural edits, persistence, and benchmark evidence in
> one TypeScript package.

With proof:

> bilig's `@bilig/headless` package exposes a WorkPaper API for programmatic
> spreadsheet automation. It has npm install docs, a runnable external example,
> persistence helpers, and checked-in benchmark evidence against
> HyperFormula-style workloads.

For contributors:

> bilig is a local-first spreadsheet engine and runtime. Good first
> contributions are formula parity fixtures, engine correctness tests,
> WorkPaper benchmark scenarios, grid accessibility fixes, and docs that turn
> architecture notes into runnable examples.

## What Not To Overclaim

- Do not call bilig a finished Excel clone.
- Do not imply full Excel formula parity.
- Do not claim every p95 benchmark row is faster than HyperFormula.
- Do not ask users to import from internal `src/` or `dist/` paths.
- Do not hide unsupported behavior. Name the gap and link the relevant issue or
  doc.

## Maintainer Checklist Before Sharing

- `gh repo view proompteng/bilig --json stargazerCount,description,repositoryTopics`
- `npm view @bilig/headless version description keywords --json`
- `pnpm workpaper:bench:competitive:check`
- `pnpm workpaper:smoke:external`

If any public claim changes, update the README, package README, benchmark
evidence note, and latest GitHub release notes together.
