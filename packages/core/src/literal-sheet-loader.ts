import { ErrorCode, ValueTag, type LiteralInput } from '@bilig/protocol'
import type { StringPool } from './string-pool.js'
import type { WorkbookStore } from './workbook-store.js'
import { CellFlags } from './cell-store.js'

export function loadLiteralSheetIntoEmptySheet(
  workbook: WorkbookStore,
  strings: StringPool,
  sheetId: number,
  content: readonly (readonly LiteralInput[])[],
  shouldMaterialize: (raw: LiteralInput, rowIndex: number, colIndex: number) => boolean = (raw) => raw !== null,
): number {
  const sheet = workbook.getSheetById(sheetId)
  if (!sheet) {
    throw new Error(`Unknown sheet id: ${sheetId}`)
  }

  let potentialCellCount = 0
  for (let rowIndex = 0; rowIndex < content.length; rowIndex += 1) {
    potentialCellCount += content[rowIndex]?.length ?? 0
  }
  if (potentialCellCount === 0) {
    return 0
  }

  const cellStore = workbook.cellStore
  cellStore.ensureCapacity(cellStore.size + potentialCellCount)
  const previousOnSetValue = cellStore.onSetValue
  cellStore.onSetValue = null
  let literalCount = 0
  try {
    workbook.withBatchedColumnVersionUpdates(() => {
      for (let rowIndex = 0; rowIndex < content.length; rowIndex += 1) {
        const row = content[rowIndex]!
        for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
          const raw = row[colIndex]!
          if (!shouldMaterialize(raw, rowIndex, colIndex)) {
            continue
          }
          const cellIndex = cellStore.allocateReserved(sheetId, rowIndex, colIndex)
          literalCount += 1
          workbook.attachAllocatedCell(sheetId, rowIndex, colIndex, cellIndex)
          writeLiteralCell(cellStore, strings, cellIndex, raw)
          workbook.notifyCellValueWritten(cellIndex)
        }
      }
    })
  } finally {
    cellStore.onSetValue = previousOnSetValue
  }

  return literalCount
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
