import { WorkPaper, exportWorkPaperDocument, serializeWorkPaperDocument } from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets(
  {
    Revenue: [
      ['Metric', 'Value'],
      ['Starter MRR', 12000],
      ['Expansion MRR', 3000],
      ['Churn MRR', 800],
      ['Net MRR', '=B2+B3-B4'],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['Net MRR', '=Revenue!B5'],
      ['Annualized ARR', '=B2*12'],
    ],
  },
  { maxRows: 1000, maxColumns: 64, useColumnIndex: true },
)

const revenueSheet = requireSheet(workbook, 'Revenue')
const summarySheet = requireSheet(workbook, 'Summary')
const editedAddress = { sheet: revenueSheet, row: 1, col: 1 }

const beforeDocument = exportWorkPaperDocument(workbook, { includeConfig: true })
const beforeSummary = readSummary(workbook, summarySheet)
const beforeSerializedInput = readDocumentCell(beforeDocument, 'Revenue', 1, 1)

workbook.setCellContents(editedAddress, 15000)

const afterDocument = exportWorkPaperDocument(workbook, { includeConfig: true })
const afterSummary = readSummary(workbook, summarySheet)
const afterSerializedInput = readDocumentCell(afterDocument, 'Revenue', 1, 1)

const output = {
  verified: true,
  changedCell: 'Revenue!B2',
  beforeSerializedInput,
  afterSerializedInput,
  changedSummaryValues: {
    before: beforeSummary,
    after: afterSummary,
  },
  documentBytes: {
    before: Buffer.byteLength(serializeWorkPaperDocument(beforeDocument), 'utf8'),
    after: Buffer.byteLength(serializeWorkPaperDocument(afterDocument), 'utf8'),
  },
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

function readSummary(workpaper, sheet) {
  return {
    netMrr: readNumber(workpaper, sheet, 1, 1, 'net MRR'),
    annualizedArr: readNumber(workpaper, sheet, 2, 1, 'annualized ARR'),
  }
}

function readNumber(workpaper, sheet, row, col, label) {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return cell.value
}

function readDocumentCell(document, sheetName, row, col) {
  const sheet = document.sheets.find((candidate) => candidate.name === sheetName)
  if (sheet === undefined) {
    throw new Error(`Expected exported document to include sheet "${sheetName}"`)
  }

  return sheet.content[row]?.[col] ?? null
}

function assertOutput(actual) {
  const expected = {
    verified: true,
    changedCell: 'Revenue!B2',
    beforeSerializedInput: 12000,
    afterSerializedInput: 15000,
    changedSummaryValues: {
      before: {
        netMrr: 14200,
        annualizedArr: 170400,
      },
      after: {
        netMrr: 17200,
        annualizedArr: 206400,
      },
    },
  }

  const comparable = {
    verified: actual.verified,
    changedCell: actual.changedCell,
    beforeSerializedInput: actual.beforeSerializedInput,
    afterSerializedInput: actual.afterSerializedInput,
    changedSummaryValues: actual.changedSummaryValues,
  }

  if (
    JSON.stringify(comparable) !== JSON.stringify(expected) ||
    typeof actual.documentBytes.before !== 'number' ||
    actual.documentBytes.before <= 0 ||
    typeof actual.documentBytes.after !== 'number' ||
    actual.documentBytes.after <= 0
  ) {
    throw new Error(`Unexpected snapshot diff WorkPaper result: ${JSON.stringify(actual)}`)
  }
}
