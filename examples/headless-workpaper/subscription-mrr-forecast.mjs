import { WorkPaper } from '@bilig/headless'

const monthlyNewCustomers = [
  {
    month: 'January',
    newCustomers: 18,
  },
  {
    month: 'February',
    newCustomers: 22,
  },
  {
    month: 'March',
    newCustomers: 16,
  },
  {
    month: 'April',
    newCustomers: 28,
  },
]

const forecastRows = monthlyNewCustomers.map((row, index) => {
  const spreadsheetRow = index + 2
  const previousRow = spreadsheetRow - 1
  const startingCustomersFormula = index === 0 ? '=Assumptions!B2' : `=E${previousRow}`

  return [
    row.month,
    row.newCustomers,
    startingCustomersFormula,
    `=C${spreadsheetRow}*Assumptions!B4`,
    `=C${spreadsheetRow}-D${spreadsheetRow}+B${spreadsheetRow}`,
    `=E${spreadsheetRow}*Assumptions!B3`,
    `=F${spreadsheetRow}*Assumptions!B5`,
    `=F${spreadsheetRow}+G${spreadsheetRow}`,
  ]
})

const workbook = WorkPaper.buildFromSheets({
  Assumptions: [
    ['Metric', 'Value'],
    ['Starting customers', 120],
    ['Plan price', 49],
    ['Monthly churn rate', 0.04],
    ['Expansion rate', 0.08],
  ],
  Forecast: [
    ['Month', 'New customers', 'Starting customers', 'Churned customers', 'Ending customers', 'Base MRR', 'Expansion MRR', 'Net MRR'],
    ...forecastRows,
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Starting MRR', '=Forecast!C2*Assumptions!B3'],
    ['Ending MRR', '=Forecast!H5'],
    ['Ending customers', '=Forecast!E5'],
    ['Net expansion MRR', '=Forecast!G5'],
    ['Four-month net MRR', '=SUM(Forecast!H2:H5)'],
    ['MRR delta', '=B3-B2'],
  ],
})

const forecastSheet = requireSheet(workbook, 'Forecast')
const summarySheet = requireSheet(workbook, 'Summary')

const output = {
  months: monthlyNewCustomers.length,
  startingMrr: readNumber(workbook, summarySheet, 1, 1, 'Starting MRR'),
  endingMrr: readNumber(workbook, summarySheet, 2, 1, 'Ending MRR'),
  endingCustomers: readNumber(workbook, summarySheet, 3, 1, 'Ending customers'),
  netExpansionMrr: readNumber(workbook, summarySheet, 4, 1, 'Net expansion MRR'),
  fourMonthNetMrr: readNumber(workbook, summarySheet, 5, 1, 'Four-month net MRR'),
  mrrDelta: readNumber(workbook, summarySheet, 6, 1, 'MRR delta'),
  firstForecastRow: workbook.getRangeSerialized({
    start: { sheet: forecastSheet, row: 1, col: 0 },
    end: { sheet: forecastSheet, row: 1, col: 7 },
  })[0],
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
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 100) / 100
}

function assertOutput(actual) {
  const expected = {
    months: 4,
    startingMrr: 5880,
    endingMrr: 9604.03,
    endingCustomers: 181.48,
    netExpansionMrr: 711.41,
    fourMonthNetMrr: 33044.9,
    mrrDelta: 3724.03,
    firstForecastRow: [
      'January',
      18,
      '=Assumptions!B2',
      '=C2*Assumptions!B4',
      '=C2-D2+B2',
      '=E2*Assumptions!B3',
      '=F2*Assumptions!B5',
      '=F2+G2',
    ],
    verified: true,
  }

  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected subscription MRR result: ${JSON.stringify(actual)}`)
  }
}
