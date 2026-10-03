---
title: Workbook runtime intent API
published: true
description: Use @bilig/workbook when a product runtime, MCP server, or tool host needs transport-neutral workbook plans, checks, and proof without owning formula calculation.
tags: workbook-api, runtime-intent, model-proofs, typescript, spreadsheet
canonical_url: https://proompteng.github.io/bilig/workbook-runtime-intent-api.html
image: /assets/github-social-preview.png
---

# Workbook runtime intent API

Use `@bilig/workbook` when your product already owns calculation but needs a
stable way to describe workbook intent before anything mutates state.

The package is for model authors, adapter authors, MCP servers, and tool hosts
that need plan data, requirements, command receipts, checks, schemas, and
readback proof. It does not calculate formulas or own WorkPaper state.

Use `@bilig/workpaper` when Bilig should run the workbook. Use
`@bilig/workbook` when another runtime should run the workbook but still needs a
proof-bound contract.

## Install

```sh
npm install @bilig/workbook
```

## Use It When

- a tool host needs to inspect workbook intent before a runtime mutates state;
- a framework wants plain JSON plan data instead of callback closures;
- a runtime adapter needs to prove `planId`, revision, applied ops, resolved
  refs, command receipts, and check results;
- a product wants workbook operations to cross process or service boundaries;
- the calculation engine is not Bilig, but the handoff still needs proof.

## Do Not Use It When

- Bilig should own workbook state and formula recalculation; use
  `@bilig/workpaper`;
- the only problem is stale formulas in an `.xlsx` file; use
  `@bilig/xlsx-formula-recalc`, `@bilig/sheetjs-formula-recalc`, or
  `@bilig/exceljs-formula-recalc`;
- the workflow depends on desktop Excel features such as macros, pivots, charts,
  add-ins, or exact UI layout.

## Proof Contract

An integration should be able to answer these questions without asking a human
to inspect a spreadsheet UI:

1. Which model and action were selected?
2. Which refs did selectors bind?
3. Which commands and low-level ops were planned?
4. Did `prepareWorkbookAction` produce valid plan data?
5. Did the plan survive JSON transport?
6. Which runtime capabilities are required?
7. Did apply proof match preview proof?
8. Are command receipts bound to the planned digests?
9. Which checks passed, and what evidence proved them?

`runWorkbookPlan(planData, adapter, { strict: true })` fails closed unless the
adapter returns the required proof. That is the main distinction from a thin
"call this function and trust the result" wrapper.

## Minimal Shape

```ts
import { defineModel, describeRunResult, formula, prepareWorkbookAction, runWorkbookPlan } from '@bilig/workbook'

const model = defineModel({
  name: 'named-range-formula',
  find(workbook) {
    return {
      input: workbook.findName('input'),
      factor: workbook.findName('factor'),
      result: workbook.findName('result'),
    }
  },
  checks({ refs, workbook }) {
    return [workbook.check.exists(refs.result), workbook.check.noFormulaErrors(refs.result)]
  },
  actions: {
    calculate({ refs, workbook }) {
      const expected = formula.multiply(refs.input, refs.factor)
      workbook.writeFormula(refs.result, expected)
      workbook.check.formulaEquals(refs.result, expected)
    },
  },
})

const prepared = prepareWorkbookAction(model, 'calculate')
if (prepared.status === 'failed') {
  throw new Error(prepared.errors[0]?.message ?? 'workbook plan failed')
}

const result = await runWorkbookPlan(prepared.planData, adapter, { strict: true })
console.log(describeRunResult(result))
```

## Package Boundary

| Package                      | Owns                                                                          | Best first proof                                         |
| ---------------------------- | ----------------------------------------------------------------------------- | -------------------------------------------------------- |
| `@bilig/workbook`            | Workbook intent, plan data, requirements, checks, schemas, and runtime proof. | Package unit tests and [public API docs](public-api.md)  |
| `@bilig/workpaper`           | WorkPaper state, recalculation, JSON persistence, MCP, and service tools.     | [WorkPaper service evaluator](eval-workpaper-service.md) |
| `@bilig/xlsx-formula-recalc` | File-level XLSX formula recalculation after input edits.                      | [XLSX recalculation evaluator](eval-xlsx-recalc.md)      |

The older `workbook-agent-intent-api.html` URL remains as a compatibility alias
for existing links.
