---
title: CSV-Shaped Input To WorkPaper
published: true
description: Turn a small CSV-shaped payload into a formula-backed Bilig WorkPaper and read calculated cells from Node.
tags: typescript, node, workpaper, formulas, csv
canonical_url: https://proompteng.github.io/bilig/csv-shaped-workpaper-input-recipe.html
---

# CSV-Shaped Input To WorkPaper

This recipe shows how to turn a small tabular payload into a formula-backed
`@bilig/workpaper` WorkPaper workbook. It is for service payloads, test fixtures,
and quick evaluator scripts where the source data looks like CSV.

The parser below is intentionally tiny. It handles a simple comma-delimited
fixture with no quoted fields, embedded commas, escaped quotes, multiline cells,
locale-specific number parsing, or import/export metadata. Use a real CSV/XLSX
import pipeline when those features matter.

## Setup

```sh
mkdir bilig-csv-shaped-input
cd bilig-csv-shaped-input
npm init -y
npm pkg set type=module
npm install @bilig/workpaper
npm install --save-dev tsx typescript
```

Download the maintained TypeScript example:

```sh
curl -fsSLo csv-input.ts \
  https://raw.githubusercontent.com/proompteng/bilig/main/examples/headless-workpaper/csv-shaped-input.ts
```

The file is also kept in the repository at
[`examples/headless-workpaper/csv-shaped-input.ts`](https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/csv-shaped-input.ts).

The important shape is small: parse the payload, build a workbook with formulas,
then read calculated cells from a separate summary sheet.

```ts
import { WorkPaper } from '@bilig/workpaper'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>
type RevenueCsvRow = {
  region: string
  customers: string
  arpa: string
}

const csvInput = `
region,customers,arpa
West,20,1200
East,30,250
Central,18,300
`.trim()

const sourceRows = parseRevenueCsv(csvInput)
const revenueRows = sourceRows.map((row, index) => {
  const spreadsheetRow = index + 2
  return [
    row.region,
    readInputNumber(row.customers, `customers row ${spreadsheetRow}`),
    readInputNumber(row.arpa, `arpa row ${spreadsheetRow}`),
    `=B${spreadsheetRow}*C${spreadsheetRow}`,
  ]
})

const workbook = WorkPaper.buildFromSheets({
  Revenue: [['Region', 'Customers', 'ARPA', 'Revenue'], ...revenueRows],
  Summary: [
    ['Metric', 'Value'],
    ['Total revenue', '=SUM(Revenue!D2:D4)'],
    ['West customers', '=SUMIF(Revenue!A2:A4,"West",Revenue!B2:B4)'],
    ['Largest deal', '=MAX(Revenue!D2:D4)'],
  ],
})

const revenueSheet = requireSheet(workbook, 'Revenue')
const summarySheet = requireSheet(workbook, 'Summary')
const serializedRevenueSheet = workbook.getSheetSerialized(revenueSheet)

const output = {
  sourceRows: sourceRows.length,
  computed: {
    totalRevenue: readComputedNumber(workbook, summarySheet, 1, 1, 'total revenue'),
    westCustomers: readComputedNumber(workbook, summarySheet, 2, 1, 'West customers'),
    largestDeal: readComputedNumber(workbook, summarySheet, 3, 1, 'largest deal'),
  },
  serializedFirstDataRow: readSerializedRow(serializedRevenueSheet, 1, 'Revenue row 2'),
  verified: true,
}

assertSummary(output)
console.log(JSON.stringify(output, null, 2))

function parseRevenueCsv(input: string): RevenueCsvRow[] {
  const lines = input
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0)
  const headerLine = lines[0]
  const dataLines = lines.slice(1)
  if (headerLine === undefined) {
    throw new Error('expected CSV header row')
  }

  const headers = headerLine.split(',').map((header) => header.trim())
  const expectedHeaders = ['region', 'customers', 'arpa']
  if (JSON.stringify(headers) !== JSON.stringify(expectedHeaders)) {
    throw new Error(`expected CSV headers ${expectedHeaders.join(',')}, received ${headers.join(',')}`)
  }

  return dataLines.map((line, index) => {
    const values = line.split(',').map((value) => value.trim())
    if (values.length !== expectedHeaders.length) {
      throw new Error(`expected ${expectedHeaders.length} CSV fields on data row ${index + 2}, received ${values.length}`)
    }

    const [region, customers, arpa] = values
    if (region === undefined || customers === undefined || arpa === undefined) {
      throw new Error(`missing CSV field on data row ${index + 2}`)
    }

    return { region, customers, arpa }
  })
}

function readInputNumber(value: string, label: string): number {
  const parsed = Number(value)
  if (!Number.isFinite(parsed)) {
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(value)}`)
  }
  return parsed
}

function requireSheet(workbook: WorkPaperInstance, sheetName: string): number {
  const sheet = workbook.getSheetId(sheetName)
  if (sheet === undefined) {
    throw new Error(`missing sheet: ${sheetName}`)
  }
  return sheet
}

function readComputedNumber(workbook: WorkPaperInstance, sheet: number, row: number, col: number, label: string): number {
  const cell = workbook.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return cell.value
}

function readSerializedRow(sheet: unknown, rowIndex: number, label: string): unknown {
  if (!Array.isArray(sheet) || !Array.isArray(sheet[rowIndex])) {
    throw new Error(`Expected ${label} to be present in serialized sheet, received ${JSON.stringify(sheet)}`)
  }

  return sheet[rowIndex]
}

function assertSummary(summary: typeof output): void {
  const expected = {
    sourceRows: 3,
    computed: {
      totalRevenue: 36900,
      westCustomers: 20,
      largestDeal: 24000,
    },
    serializedFirstDataRow: ['West', 20, 1200, '=B2*C2'],
    verified: true,
  }

  if (JSON.stringify(summary) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected WorkPaper result: ${JSON.stringify(summary)}`)
  }
}
```

Run it:

```sh
npx tsx csv-input.ts
```

Expected output:

```json
{
  "sourceRows": 3,
  "computed": {
    "totalRevenue": 36900,
    "westCustomers": 20,
    "largestDeal": 24000
  },
  "serializedFirstDataRow": ["West", 20, 1200, "=B2*C2"],
  "verified": true
}
```

## Notes For Services And Agents

- Treat this as a normalization recipe, not a complete CSV importer.
- Convert untrusted input into the `WorkPaper.buildFromSheets()` array shape
  before creating the workbook.
- Store formulas as formula strings, not precomputed literals, when downstream
  summaries need recalculation.
- Use `getCellValue()` for computed values and `getSheetSerialized()` when you
  need to inspect persisted user inputs.
- Keep import errors explicit. Do not silently coerce missing or non-numeric
  cells into zero unless the business workflow intentionally wants that.

## Validation

For the standalone recipe:

```sh
npx tsc --ignoreConfig --noEmit --module NodeNext --moduleResolution NodeNext --target ES2022 --types node csv-input.ts
npx tsx csv-input.ts
```

For a documentation patch in this repository:

```sh
pnpm --dir examples/headless-workpaper run typecheck
pnpm --dir examples/headless-workpaper run csv-shaped
pnpm docs:discovery:check
pnpm run ci
```
