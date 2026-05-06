import { ErrorCode, ValueTag, type LiteralInput } from '@bilig/protocol'
import type { StringPool } from './string-pool.js'
import type { LogicalCellLocation } from './storage/cell-page-store.js'
import type { SheetRecord, WorkbookStore } from './workbook-store.js'
import { CellFlags } from './cell-store.js'

interface FreshLiteralCellPageInternals {
  readonly setDeferred?: (location: LogicalCellLocation, cellIndex: number) => void
}

interface FreshLiteralCellIdentityInternals {
  readonly setParts?: (cellIndex: number, sheetId: number, rowId: string, colId: string) => void
}

interface FreshLiteralResidentCellInternals {
  readonly addDeferred?: (cellIndex: number, identity: { readonly rowId: string; readonly colId: string }) => void
}

interface FreshLiteralLogicalSheetInternals {
  readonly cellPages?: FreshLiteralCellPageInternals
  readonly cellIdentities?: FreshLiteralCellIdentityInternals
  readonly residentCells?: FreshLiteralResidentCellInternals
}

type FreshLiteralCellAttacher = (row: number, col: number, cellIndex: number, rowId: string, colId: string) => void

function isFreshLiteralLogicalSheetInternals(value: unknown): value is FreshLiteralLogicalSheetInternals {
  return typeof value === 'object' && value !== null
}

export interface LiteralSheetLoadInspection {
  readonly materializedCellCount: number
  readonly maxColumnCount: number
}

export function loadLiteralSheetIntoEmptySheet(
  workbook: WorkbookStore,
  strings: StringPool,
  sheetId: number,
  content: readonly (readonly LiteralInput[])[],
  shouldMaterialize: (raw: LiteralInput, rowIndex: number, colIndex: number) => boolean = (raw) => raw !== null,
  inspection?: LiteralSheetLoadInspection,
): number {
  const sheet = workbook.getSheetById(sheetId)
  if (!sheet) {
    throw new Error(`Unknown sheet id: ${sheetId}`)
  }

  let potentialCellCount = inspection?.materializedCellCount ?? 0
  let maxColumnCount = inspection?.maxColumnCount ?? 0
  if (!inspection) {
    for (let rowIndex = 0; rowIndex < content.length; rowIndex += 1) {
      const row = content[rowIndex]
      const width = row?.length ?? 0
      maxColumnCount = Math.max(maxColumnCount, width)
      if (!row) {
        continue
      }
      for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
        if (shouldMaterialize(row[colIndex]!, rowIndex, colIndex)) {
          potentialCellCount += 1
        }
      }
    }
  }
  if (potentialCellCount === 0) {
    return 0
  }

  const cellStore = workbook.cellStore
  cellStore.ensureCapacity(cellStore.size + potentialCellCount)
  const writtenColumns = new Uint8Array(maxColumnCount)
  const rowIds: string[] = []
  const colIds: string[] = []
  const ensureRowId = workbook.createLogicalAxisIdEnsurer(sheetId, 'row')
  const ensureColumnId = workbook.createLogicalAxisIdEnsurer(sheetId, 'column')
  const attachFreshCell = createFreshLiteralCellAttacher(workbook, sheet)
  let writtenColumnCount = 0
  const previousOnSetValue = cellStore.onSetValue
  cellStore.onSetValue = null
  let literalCount = 0
  try {
    for (let rowIndex = 0; rowIndex < content.length; rowIndex += 1) {
      const row = content[rowIndex]!
      for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
        const raw = row[colIndex]!
        if (!shouldMaterialize(raw, rowIndex, colIndex)) {
          continue
        }
        const cellIndex = cellStore.allocateReserved(sheetId, rowIndex, colIndex)
        literalCount += 1
        if (writtenColumns[colIndex] === 0) {
          writtenColumns[colIndex] = 1
          writtenColumnCount += 1
        }
        const rowId = (rowIds[rowIndex] ??= ensureRowId(rowIndex))
        const colId = (colIds[colIndex] ??= ensureColumnId(colIndex))
        attachFreshCell(rowIndex, colIndex, cellIndex, rowId, colId)
        writeLiteralCell(cellStore, strings, cellIndex, raw)
      }
    }
    if (writtenColumnCount > 0) {
      workbook.notifyColumnsWritten(sheetId, materializeWrittenColumns(writtenColumns, writtenColumnCount))
    }
  } finally {
    cellStore.onSetValue = previousOnSetValue
  }

  return literalCount
}

export function loadDenseLiteralSheetIntoEmptySheet(
  workbook: WorkbookStore,
  strings: StringPool,
  sheetId: number,
  content: readonly (readonly LiteralInput[])[],
  inspection?: LiteralSheetLoadInspection,
): number {
  const sheet = workbook.getSheetById(sheetId)
  if (!sheet) {
    throw new Error(`Unknown sheet id: ${sheetId}`)
  }

  let potentialCellCount = inspection?.materializedCellCount ?? 0
  let maxColumnCount = inspection?.maxColumnCount ?? 0
  if (!inspection) {
    for (let rowIndex = 0; rowIndex < content.length; rowIndex += 1) {
      const row = content[rowIndex]
      maxColumnCount = Math.max(maxColumnCount, row?.length ?? 0)
      potentialCellCount += row?.length ?? 0
    }
  }
  if (potentialCellCount === 0) {
    return 0
  }

  const cellStore = workbook.cellStore
  cellStore.ensureCapacity(cellStore.size + potentialCellCount)
  const writtenColumns = new Uint8Array(maxColumnCount)
  const rowIds: string[] = []
  const colIds: string[] = []
  const ensureRowId = workbook.createLogicalAxisIdEnsurer(sheetId, 'row')
  const ensureColumnId = workbook.createLogicalAxisIdEnsurer(sheetId, 'column')
  const attachFreshCell = createFreshLiteralCellAttacher(workbook, sheet)
  let writtenColumnCount = 0
  const previousOnSetValue = cellStore.onSetValue
  cellStore.onSetValue = null
  let literalCount = 0
  try {
    for (let rowIndex = 0; rowIndex < content.length; rowIndex += 1) {
      const row = content[rowIndex]!
      const rowId = (rowIds[rowIndex] ??= ensureRowId(rowIndex))
      for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
        const raw = row[colIndex]!
        const cellIndex = cellStore.allocateReserved(sheetId, rowIndex, colIndex)
        literalCount += 1
        if (writtenColumns[colIndex] === 0) {
          writtenColumns[colIndex] = 1
          writtenColumnCount += 1
        }
        const colId = (colIds[colIndex] ??= ensureColumnId(colIndex))
        attachFreshCell(rowIndex, colIndex, cellIndex, rowId, colId)
        writeLiteralCell(cellStore, strings, cellIndex, raw)
      }
    }
    if (writtenColumnCount > 0) {
      workbook.notifyColumnsWritten(sheetId, materializeWrittenColumns(writtenColumns, writtenColumnCount))
    }
  } finally {
    cellStore.onSetValue = previousOnSetValue
  }

  return literalCount
}

function createFreshLiteralCellAttacher(workbook: WorkbookStore, sheet: SheetRecord): FreshLiteralCellAttacher {
  const logicalCandidate: unknown = sheet.logical
  const logical = isFreshLiteralLogicalSheetInternals(logicalCandidate) ? logicalCandidate : undefined
  const setDeferredCellPage = logical?.cellPages?.setDeferred?.bind(logical.cellPages)
  const setCellIdentityParts = logical?.cellIdentities?.setParts?.bind(logical.cellIdentities)
  const addDeferredResidentCell = logical?.residentCells?.addDeferred?.bind(logical.residentCells)
  if (!setDeferredCellPage || !setCellIdentityParts || !addDeferredResidentCell) {
    return (row, col, cellIndex, rowId, colId) => {
      workbook.attachAllocatedCellWithLogicalAxisIds(sheet.id, row, col, cellIndex, rowId, colId)
    }
  }

  const setGridCell = sheet.grid.createRowMajorSetter()

  return (row, col, cellIndex, rowId, colId) => {
    setDeferredCellPage({ sheetId: sheet.id, rowId, colId }, cellIndex)
    setCellIdentityParts(cellIndex, sheet.id, rowId, colId)
    addDeferredResidentCell(cellIndex, { rowId, colId })
    setGridCell(row, col, cellIndex)
  }
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

function writeLiteralCell(cellStore: WorkbookStore['cellStore'], strings: StringPool, cellIndex: number, raw: LiteralInput): void {
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
    return
  }

  if (typeof raw === 'boolean') {
    cellStore.tags[cellIndex] = ValueTag.Boolean
    cellStore.numbers[cellIndex] = raw ? 1 : 0
    cellStore.stringIds[cellIndex] = 0
    return
  }

  cellStore.tags[cellIndex] = ValueTag.String
  cellStore.numbers[cellIndex] = 0
  cellStore.stringIds[cellIndex] = strings.intern(raw)
}
