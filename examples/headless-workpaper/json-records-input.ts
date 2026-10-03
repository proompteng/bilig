import { WorkPaper } from '@bilig/workpaper'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>

const opportunityRecords = [
  {
    account: 'Acme Manufacturing',
    region: 'West',
    stage: 'Committed',
    seats: 12,
    arpa: 1800,
    probability: 1,
  },
  {
    account: 'Beacon Health',
    region: 'East',
    stage: 'Pipeline',
    seats: 8,
    arpa: 950,
    probability: 0.5,
  },
  {
    account: 'Cobalt Finance',
    region: 'West',
    stage: 'Committed',
    seats: 15,
    arpa: 1200,
    probability: 1,
  },
]

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

const opportunitiesSheet = requireSheet(workbook, 'Opportunities')
const summarySheet = requireSheet(workbook, 'Summary')

const output = {
  sourceRecords: opportunityRecords.length,
  computed: {
    committedMrr: readNumber(workbook, summarySheet, 1, 1, 'Committed MRR'),
    weightedPipelineMrr: readNumber(workbook, summarySheet, 2, 1, 'Weighted pipeline MRR'),
    westSeats: readNumber(workbook, summarySheet, 3, 1, 'West seats'),
    largestOpportunityMrr: readNumber(workbook, summarySheet, 4, 1, 'Largest opportunity MRR'),
  },
  serializedFirstDataRow: workbook.getRangeSerialized({
    start: { sheet: opportunitiesSheet, row: 1, col: 0 },
    end: { sheet: opportunitiesSheet, row: 1, col: 7 },
  })[0],
  verified: true,
}

assertOutput(output)
console.log(JSON.stringify(output, null, 2))

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
    sourceRecords: 3,
    computed: {
      committedMrr: 39600,
      weightedPipelineMrr: 43400,
      westSeats: 27,
      largestOpportunityMrr: 21600,
    },
    serializedFirstDataRow: ['Acme Manufacturing', 'West', 'Committed', 12, 1800, 1, '=D2*E2', '=G2*F2'],
    verified: true,
  }

  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected JSON-record WorkPaper result: ${JSON.stringify(actual)}`)
  }
}
