import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets({
  Assumptions: [
    ['Metric', 'Value'],
    ['Starting customers', 120],
    ['Plan price', 79],
    ['Monthly churn rate', 0.04],
    ['Expansion rate', 0.06],
    ['New customers per month', 18],
  ],
  Forecast: [
    ['Month', 'Starting Customers', 'New Customers', 'Churned Customers', 'Ending Customers', 'Base MRR', 'Expansion MRR', 'Ending MRR'],
    ['Month 1', '=Assumptions!B2', '=Assumptions!B6', '=ROUND(B2*Assumptions!B4,0)', '=B2+C2-D2', '=E2*Assumptions!B3', '=F2*Assumptions!B5', '=F2+G2'],
    ['Month 2', '=E2', '=Assumptions!B6', '=ROUND(B3*Assumptions!B4,0)', '=B3+C3-D3', '=E3*Assumptions!B3', '=F3*Assumptions!B5', '=F3+G3'],
    ['Month 3', '=E3', '=Assumptions!B6', '=ROUND(B4*Assumptions!B4,0)', '=B4+C4-D4', '=E4*Assumptions!B3', '=F4*Assumptions!B5', '=F4+G4'],
    ['Month 4', '=E4', '=Assumptions!B6', '=ROUND(B5*Assumptions!B4,0)', '=B5+C5-D5', '=E5*Assumptions!B3', '=F5*Assumptions!B5', '=F5+G5'],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Starting MRR', '=Assumptions!B2*Assumptions!B3'],
    ['Ending MRR', '=Forecast!H5'],
    ['Net expansion MRR', '=B3-B2'],
    ['Ending customers', '=Forecast!E5'],
    ['Run-rate ARR', '=B3*12'],
  ],
})

const forecastSheet = requireSheet(workbook, 'Forecast')
const beforeUpdate = readSummary(workbook)

workbook.batch(() => {
  workbook.setCellContents({ sheet: forecastSheet, row: 1, col: 2 }, 24)
  workbook.setCellContents({ sheet: forecastSheet, row: 2, col: 2 }, 22)
})

const serialized = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))
const afterUpdate = readSummary(restored)

const output = {
  beforeUpdate,
  afterUpdate,
  persistedSheets: restored.getSheetNames(),
  serializedBytes: Buffer.byteLength(serialized, 'utf8'),
  formulasVerified: true,
  verified: true,
}

assertForecast(output)
console.log(JSON.stringify(output, null, 2))

function readSummary(workpaper) {
  const summarySheet = requireSheet(workpaper, 'Summary')

  return {
    startingMrr: readCurrency(workpaper, summarySheet, 1, 1, 'starting MRR'),
    endingMrr: readCurrency(workpaper, summarySheet, 2, 1, 'ending MRR'),
    netExpansionMrr: readCurrency(workpaper, summarySheet, 3, 1, 'net expansion MRR'),
    endingCustomers: readNumber(workpaper, summarySheet, 4, 1, 'ending customers'),
    runRateArr: readCurrency(workpaper, summarySheet, 5, 1, 'run-rate ARR'),
  }
}

function requireSheet(workpaper, sheetName) {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function readCurrency(workpaper, sheet, row, col, label) {
  return Math.round(readNumber(workpaper, sheet, row, col, label) * 100) / 100
}

function readNumber(workpaper, sheet, row, col, label) {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be a number, received ${JSON.stringify(cell)}`)
  }
  return cell.value
}

function assertForecast(actual) {
  const expected = {
    beforeUpdate: {
      startingMrr: 9480,
      endingMrr: 14235.8,
      netExpansionMrr: 4755.8,
      endingCustomers: 170,
      runRateArr: 170829.6,
    },
    afterUpdate: {
      startingMrr: 9480,
      endingMrr: 14905.72,
      netExpansionMrr: 5425.72,
      endingCustomers: 178,
      runRateArr: 178868.64,
    },
    persistedSheets: ['Assumptions', 'Forecast', 'Summary'],
    formulasVerified: true,
    verified: true,
  }

  const comparable = {
    beforeUpdate: actual.beforeUpdate,
    afterUpdate: actual.afterUpdate,
    persistedSheets: actual.persistedSheets,
    formulasVerified: actual.formulasVerified,
    verified: actual.verified,
  }

  if (JSON.stringify(comparable) !== JSON.stringify(expected) || actual.serializedBytes <= 0) {
    throw new Error(`Unexpected subscription MRR forecast result: ${JSON.stringify(actual)}`)
  }
}
