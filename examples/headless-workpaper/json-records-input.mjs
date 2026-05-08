import { WorkPaper } from '@bilig/headless'

const records = [
  { channel: 'Partner', region: 'West', opportunities: 12, averageDealSize: 1800 },
  { channel: 'Inbound', region: 'East', opportunities: 20, averageDealSize: 950 },
  { channel: 'Outbound', region: 'West', opportunities: 8, averageDealSize: 2200 },
]

const rows = [
  ['Channel', 'Region', 'Opportunities', 'Average Deal Size', 'Pipeline Value'],
  ...records.map((record) => [
    record.channel,
    record.region,
    record.opportunities,
    record.averageDealSize,
    '=C{row}*D{row}',
  ]),
].map((row, rowIndex) =>
  row.map((cell) => (typeof cell === 'string' ? cell.replaceAll('{row}', String(rowIndex + 1)) : cell)),
)

const workbook = WorkPaper.buildFromSheets({
  Pipeline: rows,
  Summary: [
    ['Metric', 'Value'],
    ['Total pipeline', '=SUM(Pipeline!E2:E4)'],
    ['West opportunities', '=SUMIF(Pipeline!B2:B4,"West",Pipeline!C2:C4)'],
  ],
})

const summarySheet = requireSheet(workbook, 'Summary')
const output = {
  sourceRecords: records.length,
  totalPipeline: readNumber(workbook, summarySheet, 1, 1, 'total pipeline'),
  westOpportunities: readNumber(workbook, summarySheet, 2, 1, 'West opportunities'),
  verified: true,
}

assertOutput(output)
console.log(JSON.stringify(output, null, 2))

function requireSheet(workpaper, sheetName) {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function readNumber(workpaper, sheet, row, col, label) {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be a number, received ${JSON.stringify(cell)}`)
  }
  return cell.value
}

function assertOutput(actual) {
  const expected = {
    sourceRecords: 3,
    totalPipeline: 58200,
    westOpportunities: 20,
    verified: true,
  }

  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected JSON records input result: ${JSON.stringify(actual)}`)
  }
}
