import type { FastifyInstance, FastifyReply, FastifyRequest } from 'fastify'
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>
type CellAddress = NonNullable<ReturnType<WorkPaperInstance['simpleCellAddressFromString']>>
type ForecastRequestBody = {
  sheetName?: unknown
  address?: unknown
  value?: unknown
}

const FORECAST_INPUTS_SHEET = 'Inputs'
const DEFAULT_EDIT_ADDRESS = 'B3'
const DEFAULT_EDIT_VALUE = 0.4
const EDITABLE_INPUT_ADDRESSES = new Set(['B2', 'B3', 'B4', 'B5'])

export function registerWorkPaperN8nRoutes(app: FastifyInstance): void {
  app.post('/api/workpaper/n8n/forecast', handleN8nForecastRequest)
}

function handleN8nForecastRequest(request: FastifyRequest<{ Body: ForecastRequestBody }>, reply: FastifyReply) {
  try {
    reply.header('cache-control', 'no-store')
    reply.header('content-type', 'application/json; charset=utf-8')
    return createForecastProof(request.body ?? {})
  } catch (error) {
    reply.code(400)
    return {
      verified: false,
      error: error instanceof Error ? error.message : 'Invalid n8n WorkPaper forecast request',
    }
  }
}

function createForecastProof(body: ForecastRequestBody) {
  const sheetName = readOptionalString(body.sheetName, FORECAST_INPUTS_SHEET, 'sheetName')
  if (sheetName !== FORECAST_INPUTS_SHEET) {
    throw new Error(`Only ${FORECAST_INPUTS_SHEET} cells are editable in the public n8n forecast demo`)
  }

  const addressText = readOptionalString(body.address, DEFAULT_EDIT_ADDRESS, 'address').toUpperCase()
  if (!EDITABLE_INPUT_ADDRESSES.has(addressText)) {
    throw new Error(`Editable input address must be one of ${[...EDITABLE_INPUT_ADDRESSES].join(', ')}`)
  }

  const value = body.value === undefined ? DEFAULT_EDIT_VALUE : readCellValue(body.value, 'value')
  const workbook = buildForecastWorkPaper()
  const summarySheet = requireSheet(workbook, 'Summary')
  const address = requireCellAddress(workbook, sheetName, addressText)
  const before = readForecastSummary(workbook, summarySheet)
  const previousValue = workbook.getCellSerialized(address)
  const formulaContracts = readFormulaContracts(workbook, summarySheet)

  workbook.setCellContents(address, value)

  const after = readForecastSummary(workbook, summarySheet)
  const serialized = serializeWorkbook(workbook)
  const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))
  const restoredSummary = readForecastSummary(restored, requireSheet(restored, 'Summary'))
  const restoredFormulaContracts = readFormulaContracts(restored, requireSheet(restored, 'Summary'))

  return {
    verified: true,
    editedCell: workbook.simpleCellAddressToString(address, { includeSheetName: true }),
    before,
    after,
    restored: restoredSummary,
    formulaContracts,
    checks: {
      previousValue,
      newValue: workbook.getCellSerialized(address),
      formulasPersisted: sameJson(formulaContracts, restoredFormulaContracts),
      restoredMatchesAfter: sameJson(after, restoredSummary),
      computedOutputChanged: after.expectedArr !== before.expectedArr || after.expansionArr !== before.expansionArr,
      serializedBytes: Buffer.byteLength(serialized, 'utf8'),
    },
  }
}

function buildForecastWorkPaper(): WorkPaperInstance {
  return WorkPaper.buildFromSheets({
    Inputs: [
      ['Metric', 'Value'],
      ['Qualified opportunities', 20],
      ['Win rate', 0.25],
      ['Average ARR', 12000],
      ['Expansion multiplier', 1.1],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['Expected customers', '=Inputs!B2*Inputs!B3'],
      ['Expected ARR', '=B2*Inputs!B4'],
      ['Expansion ARR', '=B3*Inputs!B5'],
      ['Target gap', '=B4-100000'],
    ],
  })
}

function requireSheet(workpaper: WorkPaperInstance, sheetName: string): number {
  const sheetId = workpaper.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function requireCellAddress(workpaper: WorkPaperInstance, sheetName: string, a1Address: string): CellAddress {
  const sheetId = requireSheet(workpaper, sheetName)
  const parsed = workpaper.simpleCellAddressFromString(a1Address, sheetId)

  if (parsed === undefined || parsed.sheet !== sheetId) {
    throw new Error(`Invalid cell address: ${sheetName}!${a1Address}`)
  }

  return parsed
}

function readForecastSummary(workpaper: WorkPaperInstance, summary: number) {
  return {
    expectedCustomers: readNumber(workpaper, summary, 1, 1, 'expected customers'),
    expectedArr: readNumber(workpaper, summary, 2, 1, 'expected ARR'),
    expansionArr: readNumber(workpaper, summary, 3, 1, 'expansion ARR'),
    targetGap: readNumber(workpaper, summary, 4, 1, 'target gap'),
  }
}

function readFormulaContracts(workpaper: WorkPaperInstance, summary: number) {
  return {
    expectedCustomers: readFormula(workpaper, summary, 1, 1, 'expected customers'),
    expectedArr: readFormula(workpaper, summary, 2, 1, 'expected ARR'),
    expansionArr: readFormula(workpaper, summary, 3, 1, 'expansion ARR'),
    targetGap: readFormula(workpaper, summary, 4, 1, 'target gap'),
  }
}

function readNumber(workpaper: WorkPaperInstance, sheet: number, row: number, col: number, label: string): number {
  const cell = workpaper.getCellValue({ sheet, row, col })
  if (!cell || typeof cell !== 'object' || !('value' in cell) || typeof cell.value !== 'number') {
    throw new Error(`Expected ${label} to be numeric, received ${JSON.stringify(cell)}`)
  }
  return Math.round(cell.value * 100) / 100
}

function readFormula(workpaper: WorkPaperInstance, sheet: number, row: number, col: number, label: string): string {
  const formula = workpaper.getCellFormula({ sheet, row, col })
  if (formula === undefined) {
    throw new Error(`Expected ${label} to be a formula`)
  }
  return formula
}

function serializeWorkbook(workpaper: WorkPaperInstance): string {
  return serializeWorkPaperDocument(
    exportWorkPaperDocument(workpaper, {
      includeConfig: true,
    }),
  )
}

function readOptionalString(value: unknown, fallback: string, label: string): string {
  if (value === undefined || value === null) {
    return fallback
  }
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new Error(`${label} must be a non-empty string`)
  }
  return value.trim()
}

function readCellValue(value: unknown, label: string): string | number | boolean | null {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') {
    return value
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value
  }
  throw new Error(`${label} must be a finite number, string, boolean, or null`)
}

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}
