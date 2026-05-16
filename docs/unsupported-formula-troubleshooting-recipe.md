# Unsupported Formula Troubleshooting Recipe

Status: runnable recipe for `@bilig/headless` diagnostics

Use this when a Node service or agent tool calls a formula-backed WorkPaper
cell and gets `#VALUE!`, `#NAME?`, or another workbook error instead of a
business value.

The rule is simple: do not scrape the display string and guess. Read the
display value for the user-facing error, then read structured formula
diagnostics for the reason your service should log, reject, normalize, or route
to a compatibility issue.

## Install

```sh
mkdir bilig-unsupported-formula-eval
cd bilig-unsupported-formula-eval
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install --save-dev tsx typescript
```

Create `unsupported-formula.ts`:

```ts
import { WorkPaper } from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets(
  {
    Tax: [
      ['Metric', 'Value', 'Date serial', 'Date label'],
      ['Cash flow 0', -100000, 45292, '2024-01-01'],
      ['Cash flow 1', 25000, 45658, '2025-01-01'],
      ['Cash flow 2', 35000, 46023, '2026-01-01'],
      ['Cash flow 3', 45000, 46388, '2027-01-01'],
      ['Invalid XIRR', '=XIRR(B2:B5,D2:D5)', null, null],
      ['Valid XIRR', '=XIRR(B2:B5,C2:C5)', null, null],
    ],
  },
  { maxRows: 1000, maxColumns: 100, useColumnIndex: true },
)

const sheet = workbook.getSheetId('Tax')
if (sheet === undefined) {
  throw new Error('Tax sheet was not created')
}

const at = (row: number, col: number) => ({ sheet, row, col })
const invalid = at(5, 1)
const valid = at(6, 1)

console.log(
  JSON.stringify(
    {
      invalidDisplay: workbook.getCellDisplayValue(invalid),
      invalidDiagnostics: workbook.getCellFormulaDiagnostics(invalid),
      validDisplay: workbook.getCellDisplayValue(valid),
      validValue: workbook.getCellValue(valid),
    },
    null,
    2,
  ),
)
```

Run it:

```sh
tsx unsupported-formula.ts
```

Expected output excerpt:

```json
{
  "invalidDisplay": "#VALUE!",
  "invalidDiagnostics": [
    {
      "severity": "error",
      "sheetName": "Tax",
      "a1": "B6",
      "formula": "=XIRR(B2:B5,D2:D5)",
      "functionName": "XIRR",
      "errorText": "#VALUE!",
      "code": "financial-unsupported-date-coercion",
      "message": "XIRR date range Tax!D2:D5 contains text \"2024-01-01\" at Tax!D2. Use numeric Excel serial dates; text date coercion is not supported for headless XIRR.",
      "references": ["Tax!D2:D5", "Tax!D2"]
    }
  ],
  "validDisplay": "0.02256857579463996",
  "validValue": {
    "tag": 1,
    "value": 0.02256857579463996
  }
}
```

## Before And After

Before:

```ts
;['Date label', '2024-01-01', '2025-01-01', '2026-01-01', '2027-01-01'][('Invalid XIRR', '=XIRR(B2:B5,D2:D5)')]
```

`XIRR()` currently accepts numeric Excel serial dates in headless formulas.
Text date strings are not coerced, so the formula evaluates to `#VALUE!` and
the diagnostic code is `financial-unsupported-date-coercion`.

After:

```ts
;['Date serial', 45292, 45658, 46023, 46388][('Valid XIRR', '=XIRR(B2:B5,C2:C5)')]
```

The service should normalize dates before building the workbook, or reject the
request with the diagnostic message and the cell/range references. Do not store
the bad formula result as if it were a valid business number.

## Service Pattern

Use a small error boundary around reads that must produce numbers:

```ts
import type { WorkPaper, WorkPaperCellAddress } from '@bilig/headless'

type RequiredNumberRead =
  | number
  | {
      error: string
      code: string
      message: string
      references: readonly string[]
    }

function readRequiredNumber(workbook: WorkPaper, address: WorkPaperCellAddress): RequiredNumberRead {
  const value = workbook.getCellValue(address) as unknown
  if (
    typeof value === 'object' &&
    value !== null &&
    'tag' in value &&
    value.tag === 1 &&
    'value' in value &&
    typeof value.value === 'number'
  ) {
    return value.value
  }

  const display = workbook.getCellDisplayValue(address)
  const diagnostics = workbook.getCellFormulaDiagnostics(address)
  const firstDiagnostic = diagnostics[0]

  return {
    error: display,
    code: firstDiagnostic?.code ?? 'formula-error',
    message: firstDiagnostic?.message ?? `Formula evaluated to ${display}.`,
    references: firstDiagnostic?.references ?? [],
  }
}
```

For agent tools, return the diagnostic object to the model and require a
follow-up edit before accepting the workbook. That keeps unsupported behavior
visible and gives the agent the exact cell/range it needs to fix.

## How To Triage

- `#VALUE!` with a specific financial diagnostic usually means the formula is
  supported, but one argument shape is not accepted yet.
- `#NAME?` usually means the function or defined name is not registered in the
  current workbook/formula environment.
- Missing host-backed, cube, web, external-data, or add-in-like functions belong
  behind the external function adapter boundary described in
  [`docs/formula-language.md`](formula-language.md).
- If the workbook came from an XLSX corpus, reproduce it with the verifier in
  [`docs/xlsx-corpus-verifier-walkthrough.md`](xlsx-corpus-verifier-walkthrough.md)
  so the unsupported behavior becomes a small fixture instead of a screenshot.

## Boundaries

This recipe does not imply that tracked formula names are missing. It shows the
diagnostic workflow for formula errors caused by workbook-specific references,
external dependencies, argument shapes, locale/date edges, or unsupported host
features in `@bilig/headless`.

For the broader compatibility boundary, read
[`docs/where-bilig-is-not-excel-compatible-yet.md`](where-bilig-is-not-excel-compatible-yet.md).
For the API contract, read
[`packages/headless/README.md`](../packages/headless/README.md).
