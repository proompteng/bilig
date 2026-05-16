import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { WorkPaper } from '@bilig/headless'
import { exportXlsx, importXlsx } from '@bilig/headless/xlsx'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>

type QuoteSummary = {
  readonly listRevenue: number
  readonly discountAmount: number
  readonly netRevenue: number
  readonly totalCost: number
  readonly grossMargin: number
  readonly decision: string
}

const exampleDir = dirname(fileURLToPath(import.meta.url))
const outputDir = join(exampleDir, 'dist')
const sourceXlsxPath = join(outputDir, 'pricing-model-source.xlsx')
const editedXlsxPath = join(outputDir, 'pricing-model-edited.xlsx')

mkdirSync(outputDir, { recursive: true })

const sourceWorkbook = createPricingWorkbook()
writeFileSync(sourceXlsxPath, exportXlsx(sourceWorkbook.exportSnapshot()))
sourceWorkbook.dispose()

const imported = importXlsx(new Uint8Array(readFileSync(sourceXlsxPath)), 'pricing-model-source.xlsx')
const pricingWorkbook = WorkPaper.buildFromSnapshot(imported.snapshot, {
  evaluationTimeoutMs: 30_000,
  useColumnIndex: true,
})

const inputs = requireSheet(pricingWorkbook, 'Inputs')
const summary = requireSheet(pricingWorkbook, 'Summary')
const before = readQuoteSummary(pricingWorkbook, summary)

pricingWorkbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 48)
pricingWorkbook.setCellContents({ sheet: inputs, row: 2, col: 1 }, 1500)
pricingWorkbook.setCellContents({ sheet: inputs, row: 3, col: 1 }, 0.12)
pricingWorkbook.setCellContents({ sheet: inputs, row: 4, col: 1 }, 690)
pricingWorkbook.setCellContents({ sheet: inputs, row: 5, col: 1 }, 0.3)

const after = readQuoteSummary(pricingWorkbook, summary)
writeFileSync(editedXlsxPath, exportXlsx(pricingWorkbook.exportSnapshot()))
pricingWorkbook.dispose()

const roundTrip = importXlsx(new Uint8Array(readFileSync(editedXlsxPath)), 'pricing-model-edited.xlsx')
const restored = WorkPaper.buildFromSnapshot(roundTrip.snapshot, {
  evaluationTimeoutMs: 30_000,
  useColumnIndex: true,
})
const restoredSummary = requireSheet(restored, 'Summary')
const afterReimport = readQuoteSummary(restored, restoredSummary)
const formulaStillThere = restored.getCellFormula({ sheet: restoredSummary, row: 6, col: 1 }) === '=IF(B6>=Inputs!B6,"approved","review")'
restored.dispose()

const output = {
  sourceXlsx: sourceXlsxPath,
  editedXlsx: editedXlsxPath,
  before,
  after,
  afterReimport,
  checks: {
    decisionChanged: before.decision === 'review' && after.decision === 'approved',
    exportedReimportMatchesAfter: JSON.stringify(afterReimport) === JSON.stringify(after),
    formulasSurvivedXlsxRoundTrip: formulaStillThere,
    verified: false,
  },
  nextStep: 'If this is the Node/XLSX workflow you need, star or bookmark Bilig: https://github.com/proompteng/bilig/stargazers',
}
output.checks.verified =
  output.checks.decisionChanged && output.checks.exportedReimportMatchesAfter && output.checks.formulasSurvivedXlsxRoundTrip

if (!output.checks.verified) {
  throw new Error(`XLSX recalculation proof failed: ${JSON.stringify(output, null, 2)}`)
}

console.log(JSON.stringify(output, null, 2))

function createPricingWorkbook(): WorkPaperInstance {
  return WorkPaper.buildFromSheets({
    Inputs: [
      ['Metric', 'Value'],
      ['Units', 40],
      ['List price', 1200],
      ['Discount', 0.1],
      ['Unit cost', 760],
      ['Minimum margin', 0.3],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['List revenue', '=Inputs!B2*Inputs!B3'],
      ['Discount amount', '=B2*Inputs!B4'],
      ['Net revenue', '=B2-B3'],
      ['Total cost', '=Inputs!B2*Inputs!B5'],
      ['Gross margin', '=(B4-B5)/B4'],
      ['Decision', '=IF(B6>=Inputs!B6,"approved","review")'],
    ],
  })
}

function readQuoteSummary(workbook: WorkPaperInstance, sheet: number): QuoteSummary {
  return {
    listRevenue: readRoundedNumber(workbook, sheet, 1, 1),
    discountAmount: readRoundedNumber(workbook, sheet, 2, 1),
    netRevenue: readRoundedNumber(workbook, sheet, 3, 1),
    totalCost: readRoundedNumber(workbook, sheet, 4, 1),
    grossMargin: readRoundedNumber(workbook, sheet, 5, 1),
    decision: readString(workbook, sheet, 6, 1),
  }
}

function requireSheet(workbook: WorkPaperInstance, sheetName: string): number {
  const sheet = workbook.getSheetId(sheetName)
  if (sheet === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheet
}

function readRoundedNumber(workbook: WorkPaperInstance, sheet: number, row: number, col: number): number {
  return Math.round(readNumber(workbook, sheet, row, col) * 10000) / 10000
}

function readNumber(workbook: WorkPaperInstance, sheet: number, row: number, col: number): number {
  const cell: unknown = workbook.getCellValue({ sheet, row, col })
  if (isRecord(cell) && typeof cell.value === 'number') {
    return cell.value
  }
  throw new Error(`Expected numeric cell at row ${row.toString()} col ${col.toString()}, got ${JSON.stringify(cell)}`)
}

function readString(workbook: WorkPaperInstance, sheet: number, row: number, col: number): string {
  const cell: unknown = workbook.getCellValue({ sheet, row, col })
  if (isRecord(cell) && typeof cell.value === 'string') {
    return cell.value
  }
  throw new Error(`Expected string cell at row ${row.toString()} col ${col.toString()}, got ${JSON.stringify(cell)}`)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}
