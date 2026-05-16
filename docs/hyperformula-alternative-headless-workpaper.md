# HyperFormula Alternative For Node WorkPaper Automation

Status: public comparison guide for `@bilig/headless`.

This page is for developers already evaluating HyperFormula or another
headless spreadsheet engine. It is intentionally not a takedown. HyperFormula
is the established TypeScript engine in this category; its official README
positions it as a headless spreadsheet for business web apps with formula
evaluation, CRUD operations, undo/redo, clipboard support, sorting, Node.js
support, and a GPLv3 or commercial license.

`@bilig/headless` is worth evaluating when the workload is closer to a
service-side WorkPaper runtime: formula-backed business logic, structural edits,
agent writeback, persistence, restore, and auditable benchmark evidence from the
same repository.

## Short Version

Use HyperFormula when you want a mature, UI-independent formula calculation
engine with broad spreadsheet-function coverage and commercial support from the
Handsontable team.

Use `@bilig/headless` when you need a MIT-licensed Node package that treats the
workbook as a service object: build it, mutate it, read formulas and values,
persist it, restore it, and verify agent-style edits without opening a browser
grid.

For the broader engine choice, start with the
[headless spreadsheet engine use-case chooser](headless-spreadsheet-engine-comparison.md#use-case-chooser).

## Comparison Surface

| Question                     | HyperFormula                                                                         | `@bilig/headless`                                                                                                  |
| ---------------------------- | ------------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------ |
| Primary shape                | Headless spreadsheet formula engine                                                  | Headless WorkPaper workbook facade                                                                                 |
| Runtime target               | Browser or Node.js                                                                   | Node services, tests, agents, and local runtime automation                                                         |
| License posture              | GPLv3 or commercial license                                                          | MIT                                                                                                                |
| API orientation              | Spreadsheet-engine instance with formula evaluation and workbook operations          | WorkPaper object with formula evaluation, structural edits, persistence helpers, history, and readback             |
| Agent workflow fit           | Possible, but the project is not specifically packaged around agent writeback proofs | First-class evaluation path includes an agent writeback demo with persistence and restored readback                |
| Benchmark claim in this repo | External comparison target                                                           | Checked-in WorkPaper-vs-HyperFormula artifact records `47/57` mean wins on scorecard-eligible comparable workloads |
| Caveat                       | Strong default engine, with its own licensing and integration model                  | Not a finished Excel clone and not full Excel formula parity                                                       |

## What The Benchmark Says

The current checked-in artifact is:

[`packages/benchmarks/baselines/workpaper-vs-hyperformula.json`](../packages/benchmarks/baselines/workpaper-vs-hyperformula.json)

The short benchmark explainer is:

[`docs/what-workpaper-benchmark-proves.md`](what-workpaper-benchmark-proves.md)

The current public claim is narrow:

- `47/57` mean-latency wins on scorecard-eligible comparable workloads
- `34/40` public-lane mean wins
- `13/17` holdout-lane mean wins
- an overall p95 geomean lead
- one named p95 caveat that remains visible instead of hidden

The verification command is:

```sh
pnpm workpaper:bench:competitive:check
```

This does not prove that bilig is faster at every possible spreadsheet task. It
does not prove full Excel compatibility. It proves the checked-in WorkPaper
runtime claim for the current comparable headless workload scorecard.

## Try The Package

Use the published package when you want a quick local evaluation:

```sh
mkdir bilig-headless-eval
cd bilig-headless-eval
npm init -y
npm pkg set type=module
npm install @bilig/headless
```

Use the maintained example when you want an end-to-end proof:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm start
npm run agent:verify
```

The agent verifier records the assumption cells changed, checks dependent
formula readback, persists the workbook, restores it, and verifies the restored
values.

## When To Choose bilig First

- You need a workbook object in a Node service, not a browser grid.
- You need formulas plus structural edits, undo/redo, persistence, and restore.
- You are building a coding-agent or workflow-agent loop that must verify
  writes by reading formulas and values back from the same workbook model.
- You want an MIT-licensed package surface.
- You want benchmark claims tied to checked-in artifacts and local commands.

## When Not To Choose bilig First

- You need a mature commercial support channel today.
- You need broad Excel formula compatibility before adding reduced fixtures.
- You need a library already centered around a visual spreadsheet component.
- You need every XLSX feature preserved across import/export right now.

For those cases, start with HyperFormula or a full spreadsheet product, then
use bilig's compatibility notes to decide whether a narrower WorkPaper runtime
fits a later slice.

## Proof Links

- Package README:
  [`packages/headless/README.md`](../packages/headless/README.md)
- Benchmark explainer:
  [`docs/what-workpaper-benchmark-proves.md`](what-workpaper-benchmark-proves.md)
- Benchmark evidence:
  [`docs/headless-workpaper-benchmark-evidence.md`](headless-workpaper-benchmark-evidence.md)
- Compatibility boundaries:
  [`docs/where-bilig-is-not-excel-compatible-yet.md`](where-bilig-is-not-excel-compatible-yet.md)
- Starter issues:
  [`docs/starter-issues.md`](starter-issues.md)
- Official HyperFormula repository:
  <https://github.com/handsontable/hyperformula>
- Official HyperFormula site:
  <https://hyperformula.handsontable.com/>
