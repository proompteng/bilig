# Headless Spreadsheet Engine Comparison For Node Services And Agents

Status: public comparison guide for developers evaluating spreadsheet engines.

Research date: 2026-05-12.

This page exists because "spreadsheet engine" can mean several different
things. A library that writes XLSX files, a formula-function package, a
calculation engine, a Rust/WASM spreadsheet product, and a Node WorkPaper
runtime can all be correct choices for different jobs.

`bilig` is not trying to claim that every evaluator should choose
`@bilig/headless`. The useful claim is narrower: choose it when you need a
TypeScript WorkPaper object for Node services and coding agents, with formulas,
structural edits, persistence, restore, mutation receipts, and computed
readback in one package.

## Short Version

Use `@bilig/headless` when the job is service-side workbook automation or agent
writeback verification.

Use HyperFormula when you need a mature JavaScript formula engine with broad
built-in function coverage and a commercial support path.

Use IronCalc when you want a broader open-source spreadsheet engine ecosystem
with Rust/WASM roots, embeddable product ambitions, and language bindings.

Use ExcelJS when your main job is reading, manipulating, styling, and writing
XLSX files, especially when the calculation result can be supplied or generated
by Excel or another spreadsheet app.

Use Formula.js when you need Excel-like functions as direct JavaScript calls,
not a workbook model with dependency graph, structural edits, and persistence.

Use Hucre when the main problem is TypeScript spreadsheet file I/O across XLSX,
CSV, and ODS, especially if early-access product status fits your risk profile.

Use Formualizer when you want spreadsheet logic that can run across Rust,
Python, and JavaScript/WASM runtimes.

Use JSpreadsheet Formula Pro when you are already in the JSpreadsheet ecosystem
or need its commercial JavaScript formula plugin for browser or Node
calculations.

## Use-Case Chooser

| If your job is...                                                 | Start with...                           | Check next                                                                                                                                                                        |
| ----------------------------------------------------------------- | --------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Formula-backed calculations inside a Node service                 | `@bilig/headless`                       | Start with the [quote approval WorkPaper API proof](quote-approval-workpaper-api.md), then wire the [Node service recipe](node-service-workpaper-recipe.md).                      |
| Agent writeback that must prove the value after an edit           | `@bilig/headless`                       | Use the [agent tool-calling recipe](agent-workpaper-tool-calling-recipe.md) or the [MCP WorkPaper server](mcp-workpaper-tool-server.md).                                          |
| XLSX parsing, export, styling, images, and workbook-file metadata | SheetJS or ExcelJS                      | Read the [SheetJS and ExcelJS boundary guide](sheetjs-exceljs-alternative-formula-workbook-api.md) before mixing file I/O with formula runtime state.                             |
| A mature formula engine with broad spreadsheet-function coverage  | HyperFormula                            | Compare against the [HyperFormula alternative notes](hyperformula-alternative-headless-workpaper.md) and the [compatibility caveats](where-bilig-is-not-excel-compatible-yet.md). |
| Persisting a workbook document as JSON and restoring it later     | `@bilig/headless`                       | Follow the [WorkPaper persistence guide](persisting-formula-backed-workpaper-documents-in-node.md).                                                                               |
| Embedding a spreadsheet UI that users edit directly               | A browser grid or spreadsheet component | Use bilig only if a backend WorkPaper runtime also needs to verify calculations outside the UI.                                                                                   |

This is a chooser, not a compatibility guarantee. `bilig` does not claim full
Excel parity, broad XLSX fidelity, or blanket speed wins. Check the
[documented Excel gaps](where-bilig-is-not-excel-compatible-yet.md) and run the
benchmark or small workbook check that matches your workload.

## Library Decision Table

| Workload                                                                                    | Start With        | Why                                                                                                                                                                                            |
| ------------------------------------------------------------------------------------------- | ----------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Build an XLSX report with styles, tables, images, and supplied formula results              | ExcelJS           | It is an Excel workbook manager for reading, manipulating, and writing spreadsheet data and styles. Its README says formula results must be supplied rather than calculated by ExcelJS itself. |
| Call `SUM`, `DATE`, `XLOOKUP`-style functions directly from JavaScript code                 | Formula.js        | It implements many Excel formula functions as JavaScript functions, but it is not a workbook engine.                                                                                           |
| Embed a mature headless spreadsheet formula engine in a web app or Node process             | HyperFormula      | It is UI-independent, has extensive built-in function coverage, and documents browser and server-side installation paths.                                                                      |
| Build around a Rust/WASM open-source spreadsheet ecosystem                                  | IronCalc          | It presents itself as an open-source spreadsheet engine and ecosystem, with programmatic use from multiple languages.                                                                          |
| Read, write, and transform spreadsheet files from modern TypeScript                         | Hucre             | Its site positions it around dependency-free TypeScript spreadsheet I/O for XLSX, CSV, and ODS, with built-in formula evaluation and streaming I/O.                                            |
| Share spreadsheet logic across Rust, Python, and JavaScript/WASM                            | Formualizer       | Its docs describe an embeddable spreadsheet formula engine for apps, services, and automation pipelines across those runtimes.                                                                 |
| Add commercial formula calculation to a JSpreadsheet-backed product                         | Formula Pro       | JSpreadsheet describes Formula Pro as a JavaScript plugin for spreadsheet-like calculations in the browser or Node.js.                                                                         |
| Give an agent or Node service a workbook object it can mutate, persist, restore, and verify | `@bilig/headless` | It exposes WorkPaper operations and recipes around mutation, formula readback, persistence, and restored state.                                                                                |

## What Makes The Bilig Slice Different

Most automation failures happen after the initial "write this formula" moment.
The questions become operational:

- which cells changed?
- which formulas recalculated?
- did the formula text survive the edit?
- did persisted JSON restore into the same computed workbook state?
- can the service reject unsupported formulas with useful diagnostics?
- can a coding agent prove the write instead of narrating what it meant to do?

That is the `@bilig/headless` wedge. It is a WorkPaper runtime surface, not just
a formula parser, not just an XLSX writer, and not a browser grid.

The maintained quote approval proof demonstrates the service shape without a
repo clone:

```sh
mkdir bilig-quote-approval
cd bilig-quote-approval
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install -D tsx typescript @types/node
curl -fsSLo quote-approval-api.ts \
  https://raw.githubusercontent.com/proompteng/bilig/main/examples/serverless-workpaper-api/quote-approval-api.ts
npx tsx quote-approval-api.ts
```

It writes quote inputs, recalculates net revenue, gross margin, and decision,
serializes WorkPaper JSON, restores it, and verifies `restoredMatchesAfter:
true`.

If you only need the smallest package sanity check, use the
[Node quickstart](try-bilig-headless-in-node.md) first.

The broader repo example demonstrates the agent shape:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm start
npm run agent:verify
```

`npm run agent:verify` changes assumption cells, checks dependent formula
readback, persists the workbook, restores it, and verifies that formulas and
values survived the round trip.

## Where HyperFormula Fits

HyperFormula is the strongest default comparison for a JavaScript headless
formula engine. Its official docs describe extensive built-in function coverage,
Node/server-side setup, browser integration, and explicit licensing under GPLv3
or a proprietary license.

Start with HyperFormula when the core need is formula calculation with mature
engine behavior and a commercial option.

Evaluate `@bilig/headless` when the core need is a Node WorkPaper object with
agent-oriented writeback verification, persistence helpers, restored readback,
history, and a narrow benchmark artifact tied to repository commands.

Do not treat `bilig` as a complete HyperFormula replacement. The current
tracked Office formula inventory is production-routed, but `bilig` still keeps
compatibility boundaries public: committed fixtures are evidence, not a promise
that every Excel formula argument shape, locale/date edge case, volatile result,
or arbitrary third-party workbook behaves exactly like desktop Excel.

## Where IronCalc Fits

IronCalc is the strongest adjacent open-source spreadsheet-engine project to
watch on the Rust/WASM side. Its official site describes an open-source
spreadsheet engine and ecosystem with MIT/Apache 2.0 licensing, WebAssembly in
the browser, embeddable use cases, and headless calculations. Its programming
docs describe using the same computational engine from programming languages to
create spreadsheets or run inputs through a sheet and read outputs.

Start with IronCalc when you want a broader spreadsheet engine ecosystem,
standalone or embeddable spreadsheet product direction, Rust/WASM portability,
or Python/Rust/JavaScript integration around the same engine.

Evaluate `@bilig/headless` when the immediate slice is narrower: a TypeScript
Node package for service WorkPaper state, mutation receipts, formula readback,
JSON persistence, restore checks, and coding-agent workflows.

Do not frame `bilig` as "IronCalc but in TypeScript." That is inaccurate.
IronCalc is a broader spreadsheet ecosystem; `bilig` is currently strongest as a
Node/service WorkPaper and agent-verification package.

## Where ExcelJS Fits

ExcelJS is a good choice when the workbook file is the product: generating XLSX
reports, preserving workbook structure, styling cells, streaming large files, or
writing files for Excel to open.

Its formula-value documentation is the important boundary for engine
evaluators: ExcelJS can store formulas and supplied results, but the README
states that ExcelJS cannot process a formula to generate the result. That makes
it useful for XLSX file management, but not the right primitive when a service
must recalculate formulas and verify values before Excel opens the file.

Use ExcelJS with `@bilig/headless` when the architecture needs both:

1. WorkPaper calculation and verification in Node.
2. XLSX file generation or richer workbook-file handling at the boundary.

## Where Formula.js Fits

Formula.js is useful when you want Excel-like functions as ordinary JavaScript
functions. Its README and docs position it around formula-function
implementations, with browser and Node usage.

That is a different layer from a workbook engine. Formula.js does not give an
agent a workbook document, dependency graph, structural edit model, mutation
receipt, persistence round trip, or restored readback contract by itself.

Use Formula.js when the job is "call this function." Use a workbook engine when
the job is "mutate this sheet and prove the workbook state afterward."

## Where Hucre Fits

Hucre is close enough to show up in the same search session, but the center of
gravity is different. Its site leads with spreadsheet I/O for TypeScript:
reading and writing XLSX, CSV, and ODS; streaming large files; a typed API; and
formula evaluation as part of that file-oriented engine. It also says access is
currently limited to approved early-access users.

Start with Hucre when the product needs a modern TypeScript file I/O layer and
the team is comfortable with early-access onboarding.

Evaluate `@bilig/headless` when the immediate job is not file conversion, but a
Node WorkPaper runtime that a service or coding agent can mutate, recalculate,
persist as JSON, restore, and verify through computed readback.

## Where Formualizer Fits

Formualizer's docs describe spreadsheet logic that can be built once and run in
Rust, Python, and JavaScript/WASM. That is a strong fit for teams that want one
formula engine across multiple runtime languages or that already have Rust or
Python automation paths.

Start with Formualizer when cross-language runtime support is the deciding
factor.

Evaluate `@bilig/headless` when the project is already TypeScript-first and the
hard part is an operational workbook object for Node services, agent tool calls,
mutation receipts, persistence, and readback.

## Where JSpreadsheet Formula Pro Fits

JSpreadsheet Formula Pro is a commercial JavaScript formula plugin. The v4
announcement describes browser and Node.js spreadsheet-like calculations,
JSpreadsheet integration, and a broader formula set.

Start with Formula Pro when the application already depends on JSpreadsheet or
needs that commercial plugin path.

Evaluate `@bilig/headless` when there is no grid dependency and the service
needs a small package for formula-backed WorkPaper state, persisted documents,
and verifiable agent edits.

## Evaluation Checklist

Before choosing a spreadsheet engine, write down which of these must happen
inside your process:

- formula parsing and calculation
- dependency graph recalculation after edits
- structural edits such as insert, delete, move, and undo
- multiple sheets and cross-sheet references
- XLSX import and export
- formatting, tables, images, and other workbook-file metadata
- persisted document round trips
- agent writeback receipts and computed readback
- licensing constraints for proprietary products
- fixture-scoped Excel compatibility evidence

If the must-have list is mostly XLSX output and style fidelity, start with an
XLSX library. If it is mostly calculation, start with a formula engine. If it is
agent or service workbook mutation with proof, evaluate `@bilig/headless`.

## Bilig Proof Path

Quick package evaluation:

```sh
mkdir bilig-headless-eval
cd bilig-headless-eval
npm init -y
npm pkg set type=module
npm install @bilig/headless
```

Then run the quickstart from the root README. The script builds a workbook,
edits source data, persists the document, restores it, and fails if formula
readback changes.

Maintained repo example:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm start
npm run agent:verify
```

Related proof docs:

- [`docs/hyperformula-alternative-headless-workpaper.md`](hyperformula-alternative-headless-workpaper.md)
- [`docs/javascript-spreadsheet-library-headless-node.md`](javascript-spreadsheet-library-headless-node.md)
- [`docs/why-agents-need-workbook-apis.md`](why-agents-need-workbook-apis.md)
- [`docs/agent-workpaper-tool-calling-recipe.md`](agent-workpaper-tool-calling-recipe.md)
- [`docs/persisting-formula-backed-workpaper-documents-in-node.md`](persisting-formula-backed-workpaper-documents-in-node.md)
- [`docs/unsupported-formula-troubleshooting-recipe.md`](unsupported-formula-troubleshooting-recipe.md)
- [`docs/where-bilig-is-not-excel-compatible-yet.md`](where-bilig-is-not-excel-compatible-yet.md)

## Sources

- IronCalc official site:
  <https://www.ironcalc.com/>
- IronCalc programming docs:
  <https://docs.ironcalc.com/programming/about.html>
- IronCalc unsupported features:
  <https://docs.ironcalc.com/features/unsupported-features.html>
- HyperFormula official site:
  <https://hyperformula.handsontable.com/>
- HyperFormula built-in functions:
  <https://hyperformula.handsontable.com/docs/guide/built-in-functions.html>
- HyperFormula license key:
  <https://hyperformula.handsontable.com/docs/guide/license-key.html>
- HyperFormula known limitations:
  <https://hyperformula.handsontable.com/docs/guide/known-limitations.html>
- Formula.js repository:
  <https://github.com/formulajs/formulajs>
- Formula.js function docs:
  <https://formulajs.info/functions/>
- ExcelJS README formula-value section:
  <https://github.com/exceljs/exceljs#formula-value>
- Hucre official site:
  <https://hucre.dev/>
- Formualizer docs:
  <https://www.formualizer.dev/docs>
- JSpreadsheet Formula Pro v4 announcement:
  <https://jspreadsheet.com/blog/formulas-pro-v4>
