import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import process from 'node:process'

import { WorkPaper } from '@bilig/headless'
import { exportXlsx, importXlsx } from '@bilig/headless/xlsx'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>

type QuoteReadback = {
  readonly listRevenue: number
  readonly discountAmount: number
  readonly netRevenue: number
  readonly totalCost: number
  readonly grossMargin: number
  readonly decision: string
}

const outputDir = join(process.cwd(), 'bilig-xlsx-proof-output')
const sourceXlsxPath = join(outputDir, 'quote-model-source.xlsx')
const editedXlsxPath = join(outputDir, 'quote-model-edited.xlsx')

void main()

async function main(): Promise<void> {
  await mkdir(outputDir, { recursive: true })

  const sourceWorkbook = createQuoteWorkbook()
  await writeFile(sourceXlsxPath, exportXlsx(sourceWorkbook.exportSnapshot()))
  sourceWorkbook.dispose()

  const imported = importXlsx(new Uint8Array(await readFile(sourceXlsxPath)), 'quote-model-source.xlsx')
  const proofWorkbook = WorkPaper.buildFromSnapshot(imported.snapshot, {
    evaluationTimeoutMs: 30_000,
    useColumnIndex: true,
  })

  const inputs = requireSheet(proofWorkbook, 'Inputs')
  const summary = requireSheet(proofWorkbook, 'Summary')
  const before = readQuote(proofWorkbook, summary)

  proofWorkbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 48)
  proofWorkbook.setCellContents({ sheet: inputs, row: 2, col: 1 }, 1500)
  proofWorkbook.setCellContents({ sheet: inputs, row: 3, col: 1 }, 0.12)
  proofWorkbook.setCellContents({ sheet: inputs, row: 4, col: 1 }, 690)
  proofWorkbook.setCellContents({ sheet: inputs, row: 5, col: 1 }, 0.3)

  const after = readQuote(proofWorkbook, summary)
  await writeFile(editedXlsxPath, exportXlsx(proofWorkbook.exportSnapshot()))
  proofWorkbook.dispose()

  const reimported = importXlsx(new Uint8Array(await readFile(editedXlsxPath)), 'quote-model-edited.xlsx')
  const restored = WorkPaper.buildFromSnapshot(reimported.snapshot, {
    evaluationTimeoutMs: 30_000,
    useColumnIndex: true,
  })
  const restoredSummary = requireSheet(restored, 'Summary')
  const afterReimport = readQuote(restored, restoredSummary)
  const approvalFormulaSurvived =
    restored.getCellFormula({ sheet: restoredSummary, row: 6, col: 1 }) === '=IF(B6>=Inputs!B6,"approved","review")'
  restored.dispose()

  const checks = {
    decisionChanged: before.decision === 'review' && after.decision === 'approved',
    recalculatedMargin: after.grossMargin > before.grossMargin,
    exportedReimportMatchesAfter: JSON.stringify(afterReimport) === JSON.stringify(after),
    formulasSurvivedXlsxRoundTrip: approvalFormulaSurvived,
  }

  const result = {
    proof: 'Bilig recalculated formula-backed XLSX state in Node.js without opening Excel.',
    sourceXlsx: sourceXlsxPath,
    editedXlsx: editedXlsxPath,
    before,
    after,
    afterReimport,
    checks: {
      ...checks,
      verified: Object.values(checks).every(Boolean),
    },
    repo: 'https://github.com/proompteng/bilig',
  }

  if (!result.checks.verified) {
    throw new Error(`Bilig XLSX recalculation proof failed:\n${JSON.stringify(result, null, 2)}`)
  }

  console.log(JSON.stringify(result, null, 2))
}

function createQuoteWorkbook(): WorkPaperInstance {
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

function readQuote(workbook: WorkPaperInstance, sheet: number): QuoteReadback {
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
