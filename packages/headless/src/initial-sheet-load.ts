import {
  CellFlags,
  loadLiteralSheetIntoEmptySheet,
  makeLogicalCellKey,
  type EngineFormulaSourceRef,
  type LiteralSheetLoadInspection,
  type SheetRecord,
  type SpreadsheetEngine,
} from '@bilig/core'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { WorkPaperSheet } from './work-paper-types.js'

export interface InitialSheetMaterializationInspection extends LiteralSheetLoadInspection {
  readonly formulaCellCount?: number
}

export function loadInitialLiteralSheet(
  engine: SpreadsheetEngine,
  sheetId: number,
  content: WorkPaperSheet,
  inspection?: InitialSheetMaterializationInspection,
): void {
  loadLiteralSheetIntoEmptySheet(engine.workbook, engine.strings, sheetId, content, undefined, inspection)
}

export function tryLoadInitialLiteralSheet(engine: SpreadsheetEngine, sheetId: number, content: WorkPaperSheet): boolean {
  if (sheetContainsFormulaContent(content)) {
    return false
  }
  loadInitialLiteralSheet(engine, sheetId, content)
  return true
}

function sheetContainsFormulaContent(content: WorkPaperSheet): boolean {
  return content.some((row) => row.some((value) => typeof value === 'string' && readInitialFormulaSource(value) !== undefined))
}

export interface PreparedInitialMixedSheetLoad {
  formulaRefs: EngineFormulaSourceRef[]
  potentialNewCells: number
}

interface FreshInitialCellIdentity {
  readonly sheetId: number
  readonly rowId: string
  readonly colId: string
}

interface FreshInitialResidentIdentity {
  readonly rowId: string
  readonly colId: string
}

interface FreshInitialCellPageInternals {
  readonly cells?: Map<string, number>
}

interface FreshInitialCellIdentityInternals {
  readonly identities?: Map<number, FreshInitialCellIdentity>
}

interface FreshInitialResidentCellInternals {
  readonly byCell?: Map<number, FreshInitialResidentIdentity>
  readonly byRow?: Map<string, Set<number>>
  readonly byColumn?: Map<string, Set<number>>
}

interface FreshInitialLogicalSheetInternals {
  readonly cellPages?: FreshInitialCellPageInternals
  readonly cellIdentities?: FreshInitialCellIdentityInternals
  readonly residentCells?: FreshInitialResidentCellInternals
}

type FreshInitialCellAttacher = (row: number, col: number, cellIndex: number, rowId: string, colId: string) => void

function isFreshInitialLogicalSheetInternals(value: unknown): value is FreshInitialLogicalSheetInternals {
  return typeof value === 'object' && value !== null
}

export function prepareInitialMixedSheetLoad(args: {
  engine: SpreadsheetEngine
  sheetId: number
  content: WorkPaperSheet
  rewriteFormula: (formula: string, row: number, col: number) => string
  inspection?: InitialSheetMaterializationInspection
}): PreparedInitialMixedSheetLoad {
  const sheet = args.engine.workbook.getSheetById(args.sheetId)
  if (!sheet) {
    throw new Error(`Unknown sheet id: ${args.sheetId}`)
  }

  let potentialCellCount = args.inspection?.materializedCellCount ?? 0
  let maxColumnCount = args.inspection?.maxColumnCount ?? 0
  if (!args.inspection) {
    for (let rowIndex = 0; rowIndex < args.content.length; rowIndex += 1) {
      const row = args.content[rowIndex]
      const width = row?.length ?? 0
      maxColumnCount = Math.max(maxColumnCount, width)
      if (!row) {
        continue
      }
      for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
        if (row[colIndex] !== null) {
          potentialCellCount += 1
        }
      }
    }
  }

  const cellStore = args.engine.workbook.cellStore
  if (potentialCellCount > 0) {
    cellStore.ensureCapacity(cellStore.size + potentialCellCount)
  }
  const writtenColumns = new Uint8Array(maxColumnCount)
  const rowIds: string[] = []
  const colIds: string[] = []
  let writtenColumnCount = 0
  const formulaRefs: EngineFormulaSourceRef[] =
    args.inspection?.formulaCellCount !== undefined ? Array<EngineFormulaSourceRef>(args.inspection.formulaCellCount) : []
  let formulaRefCount = 0
  const attachFreshCell = createFreshInitialCellAttacher(sheet)
  const previousOnSetValue = cellStore.onSetValue
  cellStore.onSetValue = null
  try {
    args.engine.workbook.withBatchedColumnVersionUpdates(() => {
      for (let rowIndex = 0; rowIndex < args.content.length; rowIndex += 1) {
        const row = args.content[rowIndex]!
        let rowId = rowIds[rowIndex]
        for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
          const raw = row[colIndex]!
          if (typeof raw === 'string') {
            const formula = readInitialFormulaSource(raw)
            if (formula !== undefined) {
              const cellIndex = cellStore.allocateReserved(args.sheetId, rowIndex, colIndex)
              rowId ??= args.engine.workbook.ensureLogicalAxisId(args.sheetId, 'row', rowIndex)
              rowIds[rowIndex] = rowId
              const colId = (colIds[colIndex] ??= args.engine.workbook.ensureLogicalAxisId(args.sheetId, 'column', colIndex))
              attachFreshCell(rowIndex, colIndex, cellIndex, rowId, colId)
              formulaRefs[formulaRefCount] = {
                sheetId: args.sheetId,
                cellIndex,
                row: rowIndex,
                col: colIndex,
                source: args.rewriteFormula(formula, rowIndex, colIndex),
              }
              formulaRefCount += 1
              continue
            }
          }
          if (raw === null) {
            continue
          }
          const cellIndex = cellStore.allocateReserved(args.sheetId, rowIndex, colIndex)
          if (writtenColumns[colIndex] === 0) {
            writtenColumns[colIndex] = 1
            writtenColumnCount += 1
          }
          cellStore.flags[cellIndex] = CellFlags.Materialized
          cellStore.formulaIds[cellIndex] = 0
          cellStore.errors[cellIndex] = ErrorCode.None
          cellStore.versions[cellIndex] = 1
          cellStore.topoRanks[cellIndex] = 0
          cellStore.cycleGroupIds[cellIndex] = -1
          rowId ??= args.engine.workbook.ensureLogicalAxisId(args.sheetId, 'row', rowIndex)
          rowIds[rowIndex] = rowId
          const colId = (colIds[colIndex] ??= args.engine.workbook.ensureLogicalAxisId(args.sheetId, 'column', colIndex))
          attachFreshCell(rowIndex, colIndex, cellIndex, rowId, colId)
          if (typeof raw === 'number') {
            cellStore.tags[cellIndex] = ValueTag.Number
            cellStore.numbers[cellIndex] = raw
            cellStore.stringIds[cellIndex] = 0
          } else if (typeof raw === 'boolean') {
            cellStore.tags[cellIndex] = ValueTag.Boolean
            cellStore.numbers[cellIndex] = raw ? 1 : 0
            cellStore.stringIds[cellIndex] = 0
          } else {
            cellStore.tags[cellIndex] = ValueTag.String
            cellStore.numbers[cellIndex] = 0
            cellStore.stringIds[cellIndex] = args.engine.strings.intern(raw)
          }
        }
      }
      if (writtenColumnCount > 0) {
        args.engine.workbook.notifyColumnsWritten(args.sheetId, materializeWrittenColumns(writtenColumns, writtenColumnCount))
      }
    })
  } finally {
    cellStore.onSetValue = previousOnSetValue
  }

  if (formulaRefs.length !== formulaRefCount) {
    formulaRefs.length = formulaRefCount
  }
  return {
    formulaRefs,
    potentialNewCells: formulaRefs.length,
  }
}

function createFreshInitialCellAttacher(sheet: SheetRecord): FreshInitialCellAttacher {
  const logicalCandidate: unknown = sheet.logical
  const logical = isFreshInitialLogicalSheetInternals(logicalCandidate) ? logicalCandidate : undefined
  const cells = logical?.cellPages?.cells
  const identities = logical?.cellIdentities?.identities
  const residentByCell = logical?.residentCells?.byCell
  const residentByRow = logical?.residentCells?.byRow
  const residentByColumn = logical?.residentCells?.byColumn
  if (!cells || !identities || !residentByCell || !residentByRow || !residentByColumn) {
    return (row, col, cellIndex, rowId, colId) => {
      sheet.logical.setNewVisibleCellWithAxisIds(row, col, cellIndex, rowId, colId)
      sheet.grid.set(row, col, cellIndex)
    }
  }

  let lastRowId: string | undefined
  let lastRowSet: Set<number> | undefined
  const columnSets = new Map<string, Set<number>>()
  const ensureResidentColumnSet = (id: string): Set<number> => {
    const cached = columnSets.get(id)
    if (cached) {
      return cached
    }
    let stored = residentByColumn.get(id)
    if (!stored) {
      stored = new Set<number>()
      residentByColumn.set(id, stored)
    }
    columnSets.set(id, stored)
    return stored
  }
  const ensureResidentRowSet = (id: string): Set<number> => {
    if (lastRowId === id && lastRowSet) {
      return lastRowSet
    }
    let stored = residentByRow.get(id)
    if (!stored) {
      stored = new Set<number>()
      residentByRow.set(id, stored)
    }
    lastRowId = id
    lastRowSet = stored
    return stored
  }

  return (row, col, cellIndex, rowId, colId) => {
    cells.set(makeLogicalCellKey(sheet.id, rowId, colId), cellIndex)
    identities.set(cellIndex, { sheetId: sheet.id, rowId, colId })
    residentByCell.set(cellIndex, { rowId, colId })
    ensureResidentRowSet(rowId).add(cellIndex)
    ensureResidentColumnSet(colId).add(cellIndex)
    sheet.grid.set(row, col, cellIndex)
  }
}

function readInitialFormulaSource(raw: string): string | undefined {
  const first = raw.charCodeAt(0)
  if (first === 61) {
    return raw.slice(1)
  }
  if (first !== 32 && first !== 9 && first !== 10 && first !== 13) {
    return undefined
  }
  const trimmed = raw.trim()
  return trimmed.charCodeAt(0) === 61 ? trimmed.slice(1) : undefined
}

function materializeWrittenColumns(writtenColumns: Uint8Array, count: number): Uint32Array {
  const columns = new Uint32Array(count)
  let writeIndex = 0
  for (let col = 0; col < writtenColumns.length; col += 1) {
    if (writtenColumns[col] !== 0) {
      columns[writeIndex] = col
      writeIndex += 1
    }
  }
  return columns
}

export function loadInitialMixedSheet(args: {
  engine: SpreadsheetEngine
  sheetId: number
  content: WorkPaperSheet
  rewriteFormula: (formula: string, row: number, col: number) => string
  inspection?: InitialSheetMaterializationInspection
}): void {
  const prepared = prepareInitialMixedSheetLoad(args)
  if (prepared.formulaRefs.length === 0) {
    return
  }
  args.engine.initializeFormulaSourcesAtNow(prepared.formulaRefs, prepared.potentialNewCells)
}
