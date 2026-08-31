import { createWorkPaperFromDocument, exportWorkPaperDocument, parseWorkPaperDocument, serializeWorkPaperDocument } from './persistence.js'
import type { WorkPaper } from './work-paper.js'
import type { RawCellContent } from './work-paper-types.js'

type WorkPaperInstance = ReturnType<typeof WorkPaper.buildFromSheets>
type CellAddress = NonNullable<ReturnType<WorkPaperInstance['simpleCellAddressFromString']>>

export type N8nWorkPaperEvaluationRequestBody = Record<string, unknown> & {
  document?: unknown
  edit?: unknown
  edits?: unknown
  readCells?: unknown
  includeUpdatedDocument?: unknown
}

type WorkPaperEvaluationEdit = {
  cell: string
  value: RawCellContent
}

const MAX_EDIT_COUNT = 100
const MAX_READ_CELL_COUNT = 100

export function createN8nWorkPaperEvaluationProof(body: N8nWorkPaperEvaluationRequestBody = {}) {
  const workbook = readWorkbook(body.document)
  try {
    const edits = readEdits(body)
    const readCellTexts = readReadCellTexts(body.readCells, edits)
    const readAddresses = readCellTexts.map((cell) => requireCellAddress(workbook, cell))
    const before = readCells(workbook, readAddresses)
    const previousValues = new Map<string, RawCellContent>()

    for (const edit of edits) {
      const address = requireCellAddress(workbook, edit.cell)
      const formattedCell = workbook.simpleCellAddressToString(address, { includeSheetName: true })
      previousValues.set(formattedCell, workbook.getCellSerialized(address))
      workbook.setCellContents(address, edit.value)
    }

    const after = readCells(workbook, readAddresses)
    const exportedDocument = exportWorkPaperDocument(workbook, { includeConfig: true })
    const serialized = serializeWorkPaperDocument(exportedDocument)
    const restoredWorkbook = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))
    try {
      const restored = readCells(restoredWorkbook, readAddresses)
      const editedCells = edits.map((edit) => {
        const address = requireCellAddress(workbook, edit.cell)
        const cell = workbook.simpleCellAddressToString(address, { includeSheetName: true })
        return {
          cell,
          previousValue: previousValues.get(cell),
          newValue: workbook.getCellSerialized(address),
        }
      })

      const response = {
        verified: true,
        editedCells,
        readback: {
          before,
          after,
          restored,
        },
        checks: {
          restoredMatchesAfter: sameJson(after, restored),
          formulasPersisted: sameJson(
            after.map(({ cell, formula }) => ({ cell, formula })),
            restored.map(({ cell, formula }) => ({ cell, formula })),
          ),
          computedOutputChanged: !sameJson(
            before.map(({ cell, value, displayValue }) => ({ cell, value, displayValue })),
            after.map(({ cell, value, displayValue }) => ({ cell, value, displayValue })),
          ),
          serializedBytes: new TextEncoder().encode(serialized).byteLength,
        },
      }

      if (readBoolean(body.includeUpdatedDocument, true, 'includeUpdatedDocument')) {
        return {
          ...response,
          updatedDocument: exportedDocument,
        }
      }

      return response
    } finally {
      restoredWorkbook.dispose()
    }
  } finally {
    workbook.dispose()
  }
}

function readWorkbook(document: unknown): WorkPaperInstance {
  if (document === undefined || document === null) {
    throw new Error('document is required and must be a Bilig WorkPaper JSON document')
  }
  if (typeof document === 'string') {
    return createWorkPaperFromDocument(parseWorkPaperDocument(document))
  }
  if (!isRecord(document)) {
    throw new Error('document must be a Bilig WorkPaper JSON object or JSON string')
  }
  return createWorkPaperFromDocument(parseWorkPaperDocument(JSON.stringify(document)))
}

function readEdits(body: N8nWorkPaperEvaluationRequestBody): WorkPaperEvaluationEdit[] {
  const rawEdits =
    body.edits === undefined ? (body.edit === undefined ? [] : [readJsonText(body.edit, 'edit')]) : readArray(body.edits, 'edits')
  if (rawEdits.length === 0) {
    throw new Error('at least one edit is required')
  }
  if (rawEdits.length > MAX_EDIT_COUNT) {
    throw new Error(`edits may contain at most ${MAX_EDIT_COUNT} entries`)
  }
  return rawEdits.map((entry, index) => readEdit(entry, `edits[${index}]`))
}

function readEdit(value: unknown, label: string): WorkPaperEvaluationEdit {
  value = readJsonText(value, label)
  if (!isRecord(value)) {
    throw new Error(`${label} must be an object`)
  }
  const cell = readRequiredString(value['cell'] ?? composeCell(value['sheetName'], value['address']), `${label}.cell`)
  return {
    cell,
    value: readCellValue(value['value'], `${label}.value`),
  }
}

function composeCell(sheetName: unknown, address: unknown): string | undefined {
  if (typeof sheetName !== 'string' || typeof address !== 'string') {
    return undefined
  }
  return `${sheetName}!${address}`
}

function readReadCellTexts(value: unknown, edits: WorkPaperEvaluationEdit[]): string[] {
  if (value === undefined || value === null || value === '') {
    return edits.map((edit) => edit.cell)
  }
  const cells =
    typeof value === 'string'
      ? value
          .split(',')
          .map((entry) => entry.trim())
          .filter(Boolean)
      : readArray(value, 'readCells').map((entry, index) => readRequiredString(entry, `readCells[${index}]`))

  if (cells.length === 0) {
    throw new Error('readCells must include at least one cell')
  }
  if (cells.length > MAX_READ_CELL_COUNT) {
    throw new Error(`readCells may contain at most ${MAX_READ_CELL_COUNT} cells`)
  }
  return cells
}

function readCells(workbook: WorkPaperInstance, addresses: CellAddress[]) {
  return addresses.map((address) => ({
    cell: workbook.simpleCellAddressToString(address, { includeSheetName: true }),
    serialized: workbook.getCellSerialized(address),
    formula: workbook.getCellFormula(address),
    value: workbook.getCellValue(address),
    displayValue: workbook.getCellDisplayValue(address),
  }))
}

function requireCellAddress(workbook: WorkPaperInstance, cell: string): CellAddress {
  const address = workbook.simpleCellAddressFromString(cell)
  if (address === undefined) {
    throw new Error(`Invalid cell address: ${cell}`)
  }
  return address
}

function readRequiredString(value: unknown, label: string): string {
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new Error(`${label} must be a non-empty string`)
  }
  return value.trim()
}

function readCellValue(value: unknown, label: string): RawCellContent {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') {
    return value
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value
  }
  throw new Error(`${label} must be a finite number, string, boolean, or null`)
}

function readBoolean(value: unknown, fallback: boolean, label: string): boolean {
  if (value === undefined || value === null) {
    return fallback
  }
  if (typeof value !== 'boolean') {
    throw new Error(`${label} must be a boolean`)
  }
  return value
}

function readArray(value: unknown, label: string): unknown[] {
  value = readJsonText(value, label)
  if (!Array.isArray(value)) {
    throw new Error(`${label} must be an array`)
  }
  return value
}

function readJsonText(value: unknown, label: string): unknown {
  if (typeof value !== 'string') {
    return value
  }
  const trimmed = value.trim()
  if (!trimmed.startsWith('{') && !trimmed.startsWith('[')) {
    return value
  }
  try {
    return JSON.parse(trimmed)
  } catch {
    throw new Error(`${label} must be valid JSON text`)
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function sameJson(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}
