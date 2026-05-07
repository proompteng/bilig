# Launch Post: Headless Spreadsheet Engine For Agents

Status: ready-to-adapt public launch post for `@bilig/headless`.

Use this as the first broad technical post for Hacker News, Reddit, dev.to, a
personal blog, or a project site. Adapt tone and length per channel, but keep
the evidence and caveats intact.

## Long-Form Version

Title:

> A headless spreadsheet engine for AI agents and Node services

Post:

I built `@bilig/headless`, a TypeScript spreadsheet engine for agents, Node
services, and local-first workbook automation.

The motivating problem is simple: a lot of useful business logic still lives in
spreadsheet-shaped models, but most automation either drives a browser grid or
reimplements formulas in ad hoc code. That is brittle for coding agents and
awkward for backend services.

`@bilig/headless` exposes a WorkPaper API for programmatic spreadsheet work:

- create workbooks from arrays or named sheets
- evaluate formulas
- edit cells, rows, columns, sheets, and named expressions
- batch related edits
- persist and restore workbook documents
- inspect values, formulas, serialized inputs, ranges, and sheets
- run without opening the browser UI

The repo includes a runnable external-consumer example:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm start
```

Or install the package directly:

```sh
npm install @bilig/headless
```

Minimal example:

```js
import { WorkPaper } from "@bilig/headless";

const workbook = WorkPaper.buildFromSheets({
  Deals: [
    ["Region", "Customers", "ARPA", "Revenue"],
    ["West", 20, 1200, "=B2*C2"],
    ["East", 30, 250, "=B3*C3"],
  ],
  Summary: [
    ["Metric", "Value"],
    ["Total revenue", "=SUM(Deals!D2:D3)"],
  ],
});

const summary = workbook.getSheetId("Summary");
console.log(workbook.getCellValue({ sheet: summary, row: 1, col: 1 }));
```

Performance claims are checked in rather than left as marketing copy. The
current WorkPaper benchmark artifact records `46/46` mean wins on
scorecard-eligible comparable workloads against HyperFormula-style workloads:
`38/38` public and `8/8` holdout.

That is not a blanket "faster on every p95 row" claim. The evidence note spells
out what is measured, what is excluded, and where the p95 caveat still exists:

<https://github.com/proompteng/bilig/blob/main/docs/headless-workpaper-benchmark-evidence.md>

The project is still early infrastructure, not a finished Excel clone. Known
open areas include deeper Excel formula parity, structured references, dynamic
arrays, durable multiplayer sync, and more public release hardening.

Useful entry points:

- GitHub: <https://github.com/proompteng/bilig>
- npm: <https://www.npmjs.com/package/@bilig/headless>
- package README:
  <https://github.com/proompteng/bilig/tree/main/packages/headless#readme>
- runnable example:
  <https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper>
- adoption kit:
  <https://github.com/proompteng/bilig/blob/main/docs/public-adoption-kit.md>

Good first contributions are especially useful around formula parity fixtures,
WorkPaper recipes, benchmark explanations, runnable examples, and docs that
turn architecture notes into copy-pasteable code.

## Hacker News Version

Title:

> Show HN: bilig - a headless spreadsheet engine for agents and Node services

Text:

I built `@bilig/headless`, a TypeScript spreadsheet engine for programmatic
workbook automation. It runs formulas, structural edits, persistence round
trips, and validation without opening a browser grid.

The main use case is automation that needs spreadsheet semantics but should not
screen-scrape Excel or Google Sheets. The repo has a runnable external-consumer
example, npm package docs, and checked-in benchmark evidence against
HyperFormula-style workloads.

It is early infrastructure, not a finished Excel clone. The current public
claim is `46/46` mean wins on comparable WorkPaper benchmark workloads, with a
documented p95 caveat.

Repo: <https://github.com/proompteng/bilig>

## Reddit Short Version

I built `@bilig/headless`, a TypeScript spreadsheet engine for Node services and
coding agents.

It gives you a WorkPaper API for formulas, structural edits, persistence, and
range reads without opening a browser grid. The repo includes a runnable npm
example and checked-in benchmark evidence against HyperFormula-style workloads.

Useful if you want spreadsheet-shaped business logic in code, or if an agent
needs a real workbook API instead of screen scraping.

GitHub: <https://github.com/proompteng/bilig>
npm: <https://www.npmjs.com/package/@bilig/headless>

## Follow-Up Content Queue

1. Why agents need workbook APIs, not spreadsheet screenshots.
2. Persisting formula-backed WorkPaper documents in Node.
3. What the WorkPaper-vs-HyperFormula benchmark does and does not prove.
4. Adding formula parity fixtures without overclaiming Excel compatibility.
5. Building a revenue model with `@bilig/headless`.
