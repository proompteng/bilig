import {
  CellFlags,
  loadDenseLiteralSheetIntoEmptySheet,
  loadLiteralSheetIntoEmptySheet,
  type EngineFormulaSourceRef,
  type EngineFormulaSourceRefTable,
  type LiteralSheetLoadInspection,
  type SheetRecord,
  type SpreadsheetEngine,
} from '@bilig/core'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { isBlankRawCellContent } from './work-paper-runtime-helpers.js'
import type { RawCellContent, WorkPaperSheet } from './work-paper-types.js'

export interface InitialSheetMaterializationInspection extends LiteralSheetLoadInspection {
  readonly formulaCellCount?: number
}

export function loadInitialLiteralSheet(
  engine: SpreadsheetEngine,
  sheetId: number,
  content: WorkPaperSheet,
  inspection?: InitialSheetMaterializationInspection,
): void {
  if (inspection && isDenseInitialLiteralSheet(content, inspection)) {
    loadDenseLiteralSheetIntoEmptySheet(engine.workbook, engine.strings, sheetId, content, inspection)
    return
  }
  loadLiteralSheetIntoEmptySheet(engine.workbook, engine.strings, sheetId, content, undefined, inspection)
}

function isDenseInitialLiteralSheet(content: WorkPaperSheet, inspection: InitialSheetMaterializationInspection): boolean {
  return inspection.materializedCellCount === content.length * inspection.maxColumnCount
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
  formulaRefs: EngineFormulaSourceRefTable
  potentialNewCells: number
}

interface FreshInitialLogicalSheetInternals {
  readonly setFreshVisibleCellWithAxisIdsDeferred?: (row: number, col: number, cellIndex: number, rowId: string, colId: string) => void
}

type FreshInitialCellAttacher = (row: number, col: number, cellIndex: number, rowId: string, colId: string) => void

class InitialFormulaSourceRefTable implements EngineFormulaSourceRefTable {
  private sheetIds: Uint32Array
  private cellIndices: Uint32Array
  private rows: Uint32Array
  private cols: Uint32Array
  private readonly sources: string[]
  private readonly reusable: EngineFormulaSourceRef = {
    sheetId: 0,
    cellIndex: 0,
    row: 0,
    col: 0,
    source: '',
  }

  length = 0

  constructor(capacity: number) {
    const initialCapacity = Math.max(1, capacity)
    this.sheetIds = new Uint32Array(initialCapacity)
    this.cellIndices = new Uint32Array(initialCapacity)
    this.rows = new Uint32Array(initialCapacity)
    this.cols = new Uint32Array(initialCapacity)
    this.sources = Array<string>(initialCapacity)
  }

  push(sheetId: number, cellIndex: number, row: number, col: number, source: string): void {
    if (this.length === this.sheetIds.length) {
      this.grow()
    }
    const index = this.length
    this.sheetIds[index] = sheetId
    this.cellIndices[index] = cellIndex
    this.rows[index] = row
    this.cols[index] = col
    this.sources[index] = source
    this.length = index + 1
  }

  at(index: number): EngineFormulaSourceRef {
    if (index < 0 || index >= this.length) {
      throw new RangeError(`Initial formula ref index out of bounds: ${index.toString()}`)
    }
    this.reusable.sheetId = this.sheetIds[index]!
    this.reusable.cellIndex = this.cellIndices[index]!
    this.reusable.row = this.rows[index]!
    this.reusable.col = this.cols[index]!
    this.reusable.source = this.sources[index]!
    return this.reusable
  }

  private grow(): void {
    const nextCapacity = this.sheetIds.length * 2
    const nextSheetIds = new Uint32Array(nextCapacity)
    const nextCellIndices = new Uint32Array(nextCapacity)
    const nextRows = new Uint32Array(nextCapacity)
    const nextCols = new Uint32Array(nextCapacity)
    nextSheetIds.set(this.sheetIds)
    nextCellIndices.set(this.cellIndices)
    nextRows.set(this.rows)
    nextCols.set(this.cols)
    this.sheetIds = nextSheetIds
    this.cellIndices = nextCellIndices
    this.rows = nextRows
    this.cols = nextCols
    this.sources.length = nextCapacity
  }
}

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
        if (!isBlankRawCellContent(row[colIndex])) {
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
  const colIds: string[] = []
  const ensureRowId = args.engine.workbook.createLogicalAxisIdEnsurer(args.sheetId, 'row')
  const ensureColumnId = args.engine.workbook.createLogicalAxisIdEnsurer(args.sheetId, 'column')
  let writtenColumnCount = 0
  const formulaRefs = new InitialFormulaSourceRefTable(args.inspection?.formulaCellCount ?? Math.min(potentialCellCount, 1024))
  const attachFreshCell = createFreshInitialCellAttacher(sheet)
  const previousOnSetValue = cellStore.onSetValue
  cellStore.onSetValue = null
  try {
    args.engine.workbook.withBatchedColumnVersionUpdates(() => {
      if (isDenseInitialMixedSheet(args.content, potentialCellCount, maxColumnCount)) {
        for (let rowIndex = 0; rowIndex < args.content.length; rowIndex += 1) {
          const row = args.content[rowIndex]!
          const rowId = ensureRowId(rowIndex)
          for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
            const raw = row[colIndex]!
            const cellIndex = cellStore.allocateReserved(args.sheetId, rowIndex, colIndex)
            const colId = (colIds[colIndex] ??= ensureColumnId(colIndex))
            attachFreshCell(rowIndex, colIndex, cellIndex, rowId, colId)
            if (typeof raw === 'string') {
              const formula = readInitialFormulaSource(raw)
              if (formula !== undefined) {
                formulaRefs.push(args.sheetId, cellIndex, rowIndex, colIndex, args.rewriteFormula(formula, rowIndex, colIndex))
                continue
              }
            }
            if (writtenColumns[colIndex] === 0) {
              writtenColumns[colIndex] = 1
              writtenColumnCount += 1
            }
            writeInitialLiteralCell(cellStore, args.engine.strings, cellIndex, raw)
          }
        }
      } else {
        for (let rowIndex = 0; rowIndex < args.content.length; rowIndex += 1) {
          const row = args.content[rowIndex]!
          let rowId: string | undefined
          for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
            const raw = row[colIndex]!
            if (typeof raw === 'string') {
              const formula = readInitialFormulaSource(raw)
              if (formula !== undefined) {
                const cellIndex = cellStore.allocateReserved(args.sheetId, rowIndex, colIndex)
                const materializedRowId = rowId ?? ensureRowId(rowIndex)
                rowId = materializedRowId
                const colId = (colIds[colIndex] ??= ensureColumnId(colIndex))
                attachFreshCell(rowIndex, colIndex, cellIndex, materializedRowId, colId)
                formulaRefs.push(args.sheetId, cellIndex, rowIndex, colIndex, args.rewriteFormula(formula, rowIndex, colIndex))
                continue
              }
            }
            if (isBlankRawCellContent(raw)) {
              continue
            }
            const cellIndex = cellStore.allocateReserved(args.sheetId, rowIndex, colIndex)
            if (writtenColumns[colIndex] === 0) {
              writtenColumns[colIndex] = 1
              writtenColumnCount += 1
            }
            const materializedRowId = rowId ?? ensureRowId(rowIndex)
            rowId = materializedRowId
            const colId = (colIds[colIndex] ??= ensureColumnId(colIndex))
            attachFreshCell(rowIndex, colIndex, cellIndex, materializedRowId, colId)
            writeInitialLiteralCell(cellStore, args.engine.strings, cellIndex, raw)
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

  return {
    formulaRefs,
    potentialNewCells: 0,
  }
}

function createFreshInitialCellAttacher(sheet: SheetRecord): FreshInitialCellAttacher {
  const logicalCandidate: unknown = sheet.logical
  const logical = isFreshInitialLogicalSheetInternals(logicalCandidate) ? logicalCandidate : undefined
  const attachFreshVisibleCell = logical?.setFreshVisibleCellWithAxisIdsDeferred?.bind(logical)
  if (!attachFreshVisibleCell) {
    return (row, col, cellIndex, rowId, colId) => {
      sheet.logical.setNewVisibleCellWithAxisIds(row, col, cellIndex, rowId, colId)
      sheet.grid.set(row, col, cellIndex)
    }
  }

  const setGridCell = sheet.grid.createRowMajorSetter()
  return (row, col, cellIndex, rowId, colId) => {
    attachFreshVisibleCell(row, col, cellIndex, rowId, colId)
    setGridCell(row, col, cellIndex)
  }
}

function isDenseInitialMixedSheet(content: WorkPaperSheet, materializedCellCount: number, maxColumnCount: number): boolean {
  return maxColumnCount > 0 && materializedCellCount === content.length * maxColumnCount
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

function writeInitialLiteralCell(
  cellStore: SpreadsheetEngine['workbook']['cellStore'],
  strings: { intern(value: string): number },
  cellIndex: number,
  raw: RawCellContent,
): void {
  cellStore.flags[cellIndex] = CellFlags.Materialized
  cellStore.formulaIds[cellIndex] = 0
  cellStore.errors[cellIndex] = ErrorCode.None
  cellStore.versions[cellIndex] = 1
  cellStore.topoRanks[cellIndex] = 0
  cellStore.cycleGroupIds[cellIndex] = -1
  if (raw === null) {
    cellStore.flags[cellIndex] |= CellFlags.AuthoredBlank
    cellStore.tags[cellIndex] = ValueTag.Empty
    cellStore.numbers[cellIndex] = 0
    cellStore.stringIds[cellIndex] = 0
    return
  }
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
    cellStore.stringIds[cellIndex] = strings.intern(raw)
  }
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
