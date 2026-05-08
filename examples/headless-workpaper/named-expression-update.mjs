import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets({
  Pipeline: [
    ['Stage', 'Revenue'],
    ['Committed', 24000],
    ['Expansion', 12000],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Base revenue', '=SUM(Pipeline!B2:B3)'],
    ['Growth-adjusted revenue', null],
  ],
})

const summarySheet = requireSheet(workbook, 'Summary')
workbook.addNamedExpression('GrowthRatePercent', 10)
workbook.setCellContents({ sheet: summarySheet, row: 2, col: 1 }, '=B2*(100+GrowthRatePercent)/100')

const before = readSummary(workbook, summarySheet)
const beforeNamedExpression = readNamedNumber(workbook, 'GrowthRatePercent')

workbook.changeNamedExpression('GrowthRatePercent', 25)

const after = readSummary(workbook, summarySheet)
const afterNamedExpression = readNamedNumber(workbook, 'GrowthRatePercent')
const serialized = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))
const restoredSummarySheet = requireSheet(restored, 'Summary')
const restoredSummary = readSummary(restored, restoredSummarySheet)
const restoredNamedExpression = readNamedNumber(restored, 'GrowthRatePercent')

const output = {
  verified: true,
  namedExpression: 'GrowthRatePercent',
  before,
  after,
  restored: restoredSummary,
  namedExpressionValues: {
    before: beforeNamedExpression,
    after: afterNamedExpression,
    restored: restoredNamedExpression,
  },
  persistedNamedExpressions: restored.listNamedExpressions(),
  restoredMatchesAfter: sameJson(restoredSummary, after) && restoredNamedExpression === afterNamedExpression,
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

function readSummary(workpaper, summarySheet) {
  return {
    baseRevenue: readNumber(workpaper, summarySheet, 1, 1, 'base revenue'),
    growthAdjustedRevenue: readNumber(workpaper, summarySheet, 2, 1, 'growth-adjusted revenue'),
  }
}

function readNumber(workpaper, sheet, row, col, label) {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be a number, received ${JSON.stringify(cell)}`)
  }
  return cell.value
}

function readNamedNumber(workpaper, name) {
  const cell = workpaper.getNamedExpressionValue(name)
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected named expression "${name}" to be a number, received ${JSON.stringify(cell)}`)
  }
  return cell.value
}

function sameJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right)
}

function assertOutput(actual) {
  const expected = {
    verified: true,
    namedExpression: 'GrowthRatePercent',
    before: {
      baseRevenue: 36000,
      growthAdjustedRevenue: 39600,
    },
    after: {
      baseRevenue: 36000,
      growthAdjustedRevenue: 45000,
    },
    restored: {
      baseRevenue: 36000,
      growthAdjustedRevenue: 45000,
    },
    namedExpressionValues: {
      before: 10,
      after: 25,
      restored: 25,
    },
    persistedNamedExpressions: ['GrowthRatePercent'],
    restoredMatchesAfter: true,
  }

  if (!sameJson(actual, expected)) {
    throw new Error(`Unexpected named expression update result: ${JSON.stringify(actual)}`)
  }
}
