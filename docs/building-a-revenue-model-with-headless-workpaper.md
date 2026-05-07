# Building A Revenue Model With Headless WorkPaper

Status: public adoption article for `@bilig/headless`

Most spreadsheet automation demos stop at "generate a formula." A service or
agent needs more than that. It needs to load workbook state, edit inputs,
recalculate formulas, verify outputs, and persist the document for the next
turn.

This article uses the runnable
[`examples/headless-workpaper/revenue-scenarios.mjs`](../examples/headless-workpaper/revenue-scenarios.mjs)
example to show that loop with `@bilig/headless`.

## The Model

The example builds three sheets:

- `Pipeline`: segment-level leads, conversion rates, ARPA, churn, and net MRR
- `Summary`: total net MRR, annual run rate, enterprise net MRR, and expansion
  target
- `Scenarios`: conservative, expansion, and stretch projections

The key point is that the formulas stay in the workbook:

```js
const workbook = WorkPaper.buildFromSheets({
  Pipeline: [
    ['Segment', 'Leads', 'Conversion Rate', 'Customers', 'ARPA', 'Gross MRR', 'Churn Rate', 'Net MRR'],
    ['Enterprise', 80, 0.18, '=B2*C2', 4200, '=D2*E2', 0.05, '=F2*(1-G2)'],
    ['Mid-Market', 220, 0.22, '=B3*C3', 1100, '=D3*E3', 0.08, '=F3*(1-G3)'],
    ['SMB', 900, 0.09, '=B4*C4', 180, '=D4*E4', 0.12, '=F4*(1-G4)'],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Total net MRR', '=SUM(Pipeline!H2:H4)'],
    ['Annual run rate', '=B2*12'],
    ['Enterprise net MRR', '=SUMIF(Pipeline!A2:A4,"Enterprise",Pipeline!H2:H4)'],
    ['Expansion target', '=B3*1.18'],
  ],
})
```

That means the model can be changed by code without turning the spreadsheet into
a pile of copied numbers.

## The Agent-Style Edit

The example applies a focused planning change in a single batch:

```js
workbook.batch(() => {
  workbook.setCellContents({ sheet: pipelineSheet, row: 1, col: 1 }, 92)
  workbook.setCellContents({ sheet: pipelineSheet, row: 2, col: 2 }, 0.26)
})
```

The first edit changes enterprise leads from `80` to `92`. The second changes
mid-market conversion from `0.22` to `0.26`.

After that, formulas recalculate through the model:

- total net MRR moves from `119267.2` to `136791.2`
- annual run rate moves from `1431206.4` to `1641494.4`
- enterprise net MRR moves from `57456` to `66074.4`
- stretch projected net MRR moves from `161010.72` to `184668.12`

Those are not hand-copied outputs. The script reads them back from cells and
fails if the values drift.

## The Persistence Loop

The example also serializes and restores the workbook before final readback:

```js
const serialized = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))
```

That is the operational pattern for agents and services:

1. load workbook state
2. apply a narrow edit
3. recalculate
4. persist the workbook document
5. restore it later and verify the same formulas

## Run It

From the repository:

```bash
cd examples/headless-workpaper
npm install
npm run scenarios
```

The expected output includes both the pre-edit and post-edit model:

```json
{
  "beforeEdit": {
    "totalNetMrr": 119267.2,
    "annualRunRate": 1431206.4,
    "enterpriseNetMrr": 57456,
    "expansionTarget": 1688823.55
  },
  "afterEdit": {
    "totalNetMrr": 136791.2,
    "annualRunRate": 1641494.4,
    "enterpriseNetMrr": 66074.4,
    "expansionTarget": 1936963.39
  }
}
```

The repository smoke test also runs this scenario against packed local runtime
packages:

```bash
pnpm workpaper:smoke:external
```

## Why This Matters

Spreadsheet-backed products often need the spreadsheet model without the
browser grid:

- a backend service calculates workbook-backed business logic
- an agent edits inputs and needs deterministic readback
- a workflow stores a formula-backed document between turns
- a test runner checks that a workbook change has the expected downstream
  effect

`@bilig/headless` is meant for that boundary. It does not claim to be a complete
Excel clone; it gives code a typed workbook surface with formulas, structural
edits, persistence, and benchmark evidence in the same TypeScript package.
