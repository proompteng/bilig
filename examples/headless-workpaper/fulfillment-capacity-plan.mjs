import { WorkPaper } from '@bilig/headless'

const fulfillmentDays = [
  {
    day: 'Monday',
    forecastOrders: 420,
    unitsPerOrder: 1.8,
    unitsPerHour: 55,
    availableHours: 14,
  },
  {
    day: 'Tuesday',
    forecastOrders: 510,
    unitsPerOrder: 1.6,
    unitsPerHour: 55,
    availableHours: 14.5,
  },
  {
    day: 'Wednesday',
    forecastOrders: 470,
    unitsPerOrder: 1.9,
    unitsPerHour: 60,
    availableHours: 15.5,
  },
  {
    day: 'Thursday',
    forecastOrders: 620,
    unitsPerOrder: 1.7,
    unitsPerHour: 60,
    availableHours: 16,
  },
]

const capacityRows = fulfillmentDays.map((item, index) => {
  const spreadsheetRow = index + 2
  return [
    item.day,
    item.forecastOrders,
    item.unitsPerOrder,
    item.unitsPerHour,
    item.availableHours,
    `=B${spreadsheetRow}*C${spreadsheetRow}/D${spreadsheetRow}`,
    `=E${spreadsheetRow}-F${spreadsheetRow}`,
    `=IF(G${spreadsheetRow}<0,"Short","Ready")`,
  ]
})

const workbook = WorkPaper.buildFromSheets({
  Capacity: [
    ['Day', 'Forecast orders', 'Units/order', 'Units/hour', 'Available hours', 'Required hours', 'Capacity gap', 'Status'],
    ...capacityRows,
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Forecast orders', '=SUM(Capacity!B2:B5)'],
    ['Required hours', '=SUM(Capacity!F2:F5)'],
    ['Available hours', '=SUM(Capacity!E2:E5)'],
    ['Capacity gap', '=B4-B3'],
    ['Status', '=IF(B5<0,"Short","Ready")'],
    ['Short days', '=COUNTIF(Capacity!H2:H5,"Short")'],
    ['Largest daily shortfall', '=MIN(Capacity!G2:G5)'],
  ],
})

const capacitySheet = requireSheet(workbook, 'Capacity')
const summarySheet = requireSheet(workbook, 'Summary')
const status = readString(workbook, summarySheet, 5, 1, 'Status')
const bottleneckDay =
  readString(workbook, capacitySheet, 4, 7, 'Thursday status') === 'Short'
    ? readString(workbook, capacitySheet, 4, 0, 'Bottleneck day')
    : undefined

const output = {
  days: fulfillmentDays.length,
  forecastOrders: readNumber(workbook, summarySheet, 1, 1, 'Forecast orders'),
  requiredHours: readNumber(workbook, summarySheet, 2, 1, 'Required hours'),
  availableHours: readNumber(workbook, summarySheet, 3, 1, 'Available hours'),
  capacityGap: readNumber(workbook, summarySheet, 4, 1, 'Capacity gap'),
  status,
  shortDays: readNumber(workbook, summarySheet, 6, 1, 'Short days'),
  largestDailyShortfall: readNumber(workbook, summarySheet, 7, 1, 'Largest daily shortfall'),
  bottleneckDay,
  firstCapacityRow: workbook.getRangeSerialized({
    start: { sheet: capacitySheet, row: 1, col: 0 },
    end: { sheet: capacitySheet, row: 1, col: 7 },
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
    days: 4,
    forecastOrders: 2020,
    requiredHours: 61.0318,
    availableHours: 60,
    capacityGap: -1.0318,
    status: 'Short',
    shortDays: 2,
    largestDailyShortfall: -1.5667,
    bottleneckDay: 'Thursday',
    firstCapacityRow: ['Monday', 420, 1.8, 55, 14, '=B2*C2/D2', '=E2-F2', '=IF(G2<0,"Short","Ready")'],
    verified: true,
  }

  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected fulfillment capacity result: ${JSON.stringify(actual)}`)
  }
}
