import { readFile } from 'node:fs/promises'
import { WorkPaper } from '@bilig/workpaper'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>
type OpportunityRecord = {
  account: string
  region: string
  stage: string
  seats: number
  arpa: number
  probability: number
}

const source = 'fixtures/opportunities.json'
const opportunityRecords = parseOpportunityRecords(await readFile(new URL(source, import.meta.url), 'utf8'))

const opportunityRows = opportunityRecords.map((record, index) => {
  const spreadsheetRow = index + 2

  return [
    record.account,
    record.region,
    record.stage,
    record.seats,
    record.arpa,
    record.probability,
    `=D${spreadsheetRow}*E${spreadsheetRow}`,
    `=G${spreadsheetRow}*F${spreadsheetRow}`,
  ]
})

const workbook = WorkPaper.buildFromSheets({
  Opportunities: [['Account', 'Region', 'Stage', 'Seats', 'ARPA', 'Probability', 'Gross MRR', 'Weighted MRR'], ...opportunityRows],
  Summary: [
    ['Metric', 'Value'],
    ['Committed MRR', '=SUMIFS(Opportunities!G2:G4,Opportunities!C2:C4,"Committed")'],
    ['Weighted pipeline MRR', '=SUM(Opportunities!H2:H4)'],
    ['West seats', '=SUMIF(Opportunities!B2:B4,"West",Opportunities!D2:D4)'],
    ['Largest opportunity MRR', '=MAX(Opportunities!G2:G4)'],
  ],
})

const summarySheet = requireSheet(workbook, 'Summary')

const output = {
  verified: true,
  source,
  sourceRecords: opportunityRecords.length,
  computed: {
    committedMrr: readNumber(workbook, summarySheet, 1, 1, 'Committed MRR'),
    weightedPipelineMrr: readNumber(workbook, summarySheet, 2, 1, 'Weighted pipeline MRR'),
    westSeats: readNumber(workbook, summarySheet, 3, 1, 'West seats'),
    largestOpportunityMrr: readNumber(workbook, summarySheet, 4, 1, 'Largest opportunity MRR'),
  },
}

assertOutput(output)
console.log(JSON.stringify(output, null, 2))

function parseOpportunityRecords(serializedRecords: string): OpportunityRecord[] {
  const records = JSON.parse(serializedRecords)
  if (!Array.isArray(records)) {
    throw new Error('expected JSON array')
  }

  return records.map((record, index) => {
    if (!record || typeof record !== 'object') {
      throw new Error(`record ${index + 1} must be an object`)
    }

    const parsed = {
      account: readString(record, 'account', index),
      region: readString(record, 'region', index),
      stage: readString(record, 'stage', index),
      seats: readNumberField(record, 'seats', index),
      arpa: readNumberField(record, 'arpa', index),
      probability: readNumberField(record, 'probability', index),
    }

    if (parsed.probability < 0 || parsed.probability > 1) {
      throw new Error(`record ${index + 1} probability must be between 0 and 1`)
    }

    return parsed
  })
}

function readString(record: Record<string, unknown>, field: string, index: number): string {
  const value = record[field]
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new Error(`record ${index + 1} ${field} must be a non-empty string`)
  }
  return value
}

function readNumberField(record: Record<string, unknown>, field: string, index: number): number {
  const value = record[field]
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new Error(`record ${index + 1} ${field} must be a finite number`)
  }
  return value
}

function requireSheet(workpaper: WorkPaperInstance, sheetName: string): number {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function readNumber(workpaper: WorkPaperInstance, sheet: number, row: number, col: number, label: string): number {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 100) / 100
}

function assertOutput(actual: typeof output): void {
  const expected = {
    verified: true,
    source: 'fixtures/opportunities.json',
    sourceRecords: 3,
    computed: {
      committedMrr: 39600,
      weightedPipelineMrr: 43400,
      westSeats: 27,
      largestOpportunityMrr: 21600,
    },
  }

  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected JSON-file WorkPaper result: ${JSON.stringify(actual)}`)
  }
}
