import { WorkPaper } from '@bilig/headless'

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

const malformedCsvInput = `
region,customers,arpa
West,20,1200
East,30
Central,18,300
`.trim()

if (process.argv.includes('--malformed-smoke')) {
  runMalformedCsvSmoke()
} else {
  const output = buildRevenueSummary(csvInput)
  assertSummary(output)
  console.log(JSON.stringify(output, null, 2))
}

function buildRevenueSummary(input: string) {
  const sourceRows = parseRevenueCsv(input)
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

  return {
    sourceRows: sourceRows.length,
    computed: {
      totalRevenue: readComputedNumber(workbook, summarySheet, 1, 1, 'total revenue'),
      westCustomers: readComputedNumber(workbook, summarySheet, 2, 1, 'West customers'),
      largestDeal: readComputedNumber(workbook, summarySheet, 3, 1, 'largest deal'),
    },
    serializedFirstDataRow: readSerializedRow(serializedRevenueSheet, 1, 'Revenue row 2'),
    verified: true,
  }
}

function runMalformedCsvSmoke(): void {
  try {
    buildRevenueSummary(malformedCsvInput)
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
    const expectedMessage = 'expected 3 CSV fields on data row 3, received 2'
    if (message !== expectedMessage) {
      throw new Error(`Unexpected malformed CSV error: ${message}`, { cause: error })
    }

    console.log(
      JSON.stringify(
        {
          malformedCsvError: message,
          verified: true,
        },
        null,
        2,
      ),
    )
    return
  }

  throw new Error('Expected malformed CSV input to fail before building a WorkPaper')
}

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

function requireSheet(workpaper: WorkPaperInstance, sheetName: string): number {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function readComputedNumber(workpaper: WorkPaperInstance, sheet: number, row: number, col: number, label: string): number {
  const cell = workpaper.getCellValue({ sheet, row, col })
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

function assertSummary(summary: ReturnType<typeof buildRevenueSummary>): void {
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
