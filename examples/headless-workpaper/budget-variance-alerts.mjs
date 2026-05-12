import { WorkPaper } from '@bilig/headless'

const budgetRows = [
  {
    department: 'Marketing',
    budget: 50000,
    actual: 57500,
  },
  {
    department: 'Support',
    budget: 30000,
    actual: 29100,
  },
  {
    department: 'Infrastructure',
    budget: 40000,
    actual: 42000,
  },
  {
    department: 'Sales',
    budget: 65000,
    actual: 68000,
  },
]

const varianceRows = budgetRows.map((row, index) => {
  const spreadsheetRow = index + 2
  return [
    row.department,
    row.budget,
    row.actual,
    `=C${spreadsheetRow}-B${spreadsheetRow}`,
    `=D${spreadsheetRow}/B${spreadsheetRow}`,
    `=IF(E${spreadsheetRow}>0.1,"Review","OK")`,
  ]
})

const workbook = WorkPaper.buildFromSheets({
  Variance: [['Department', 'Budget', 'Actual', 'Variance', 'Variance %', 'Alert'], ...varianceRows],
  Summary: [
    ['Metric', 'Value'],
    ['Total budget', '=SUM(Variance!B2:B5)'],
    ['Total actual', '=SUM(Variance!C2:C5)'],
    ['Total variance', '=B3-B2'],
    ['Largest overage', '=MAX(Variance!D2:D5)'],
    ['Largest variance %', '=MAX(Variance!E2:E5)'],
    ['Review count', '=COUNTIF(Variance!F2:F5,"Review")'],
  ],
})

const varianceSheet = requireSheet(workbook, 'Variance')
const summarySheet = requireSheet(workbook, 'Summary')
const flaggedDepartment =
  readString(workbook, varianceSheet, 1, 5, 'Marketing alert') === 'Review'
    ? readString(workbook, varianceSheet, 1, 0, 'Flagged department')
    : undefined

const output = {
  rows: budgetRows.length,
  flaggedDepartment,
  varianceAmount: readNumber(workbook, varianceSheet, 1, 3, 'Marketing variance'),
  variancePercent: readNumber(workbook, varianceSheet, 1, 4, 'Marketing variance percent'),
  summary: {
    totalBudget: readNumber(workbook, summarySheet, 1, 1, 'Total budget'),
    totalActual: readNumber(workbook, summarySheet, 2, 1, 'Total actual'),
    totalVariance: readNumber(workbook, summarySheet, 3, 1, 'Total variance'),
    largestOverage: readNumber(workbook, summarySheet, 4, 1, 'Largest overage'),
    largestVariancePercent: readNumber(workbook, summarySheet, 5, 1, 'Largest variance percent'),
    reviewCount: readNumber(workbook, summarySheet, 6, 1, 'Review count'),
  },
  firstVarianceRow: workbook.getRangeSerialized({
    start: { sheet: varianceSheet, row: 1, col: 0 },
    end: { sheet: varianceSheet, row: 1, col: 5 },
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
  return Math.round(cell.value * 10000) / 10000
}

function readString(workpaper, sheet, row, col, label) {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'string') {
    throw new Error(`Expected ${label} to be text, received ${JSON.stringify(cell)}`)
  }
  return cell.value
}

function assertOutput(actual) {
  const expected = {
    rows: 4,
    flaggedDepartment: 'Marketing',
    varianceAmount: 7500,
    variancePercent: 0.15,
    summary: {
      totalBudget: 185000,
      totalActual: 196600,
      totalVariance: 11600,
      largestOverage: 7500,
      largestVariancePercent: 0.15,
      reviewCount: 1,
    },
    firstVarianceRow: ['Marketing', 50000, 57500, '=C2-B2', '=D2/B2', '=IF(E2>0.1,"Review","OK")'],
    verified: true,
  }

  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected budget variance result: ${JSON.stringify(actual)}`)
  }
}
