import { formatAddress, parseCellAddress } from '@bilig/formula'
import { ValueTag, type CellRangeRef, type CellSnapshot } from '@bilig/protocol'
import { OPTIMISTIC_CELL_SNAPSHOT_FLAG } from './workbook-optimistic-cell-flags.js'
import { createSupersedingCellSnapshot } from './workbook-optimistic-cell.js'

export interface OptimisticViewportStore {
  forEachCellSnapshotInRange?(range: CellRangeRef, listener: (snapshot: CellSnapshot) => void): void
  getCell(sheetName: string, address: string): CellSnapshot
  setCellSnapshot(snapshot: CellSnapshot): void
}

const MAX_MATERIALIZED_OPTIMISTIC_CLEAR_CELLS = 10_000

export function normalizeCellRange(range: CellRangeRef): { startRow: number; endRow: number; startCol: number; endCol: number } {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  }
}

export function createEmptyOptimisticSnapshot(sheetName: string, address: string, version: number): CellSnapshot {
  return {
    sheetName,
    address,
    value: { tag: ValueTag.Empty },
    flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
    version,
  }
}

export function applyOptimisticClearRange(viewportStore: OptimisticViewportStore | null, range: CellRangeRef): (() => void) | null {
  if (!viewportStore) {
    return null
  }

  const bounds = normalizeCellRange(range)
  const previousSnapshots: CellSnapshot[] = []
  const nextSnapshots: CellSnapshot[] = []
  let rollbackVersion = 0
  const cellCount = (bounds.endRow - bounds.startRow + 1) * (bounds.endCol - bounds.startCol + 1)

  if (cellCount > MAX_MATERIALIZED_OPTIMISTIC_CLEAR_CELLS) {
    if (!viewportStore.forEachCellSnapshotInRange) {
      return null
    }
    viewportStore.forEachCellSnapshotInRange(range, (previous) => {
      const next = createEmptyOptimisticSnapshot(previous.sheetName, previous.address, previous.version + 1)
      previousSnapshots.push(previous)
      nextSnapshots.push(next)
      rollbackVersion = Math.max(rollbackVersion, next.version)
    })
    if (nextSnapshots.length === 0) {
      return null
    }
    nextSnapshots.forEach((snapshot) => viewportStore.setCellSnapshot(snapshot))
    return () => {
      previousSnapshots.forEach((snapshot) => {
        rollbackVersion += 1
        viewportStore.setCellSnapshot(createSupersedingCellSnapshot(snapshot, rollbackVersion))
      })
    }
  }

  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
      const address = formatAddress(row, col)
      const previous = viewportStore.getCell(range.sheetName, address)
      const next = createEmptyOptimisticSnapshot(range.sheetName, address, previous.version + 1)
      previousSnapshots.push(previous)
      nextSnapshots.push(next)
      rollbackVersion = Math.max(rollbackVersion, next.version)
    }
  }

  nextSnapshots.forEach((snapshot) => viewportStore.setCellSnapshot(snapshot))

  return () => {
    previousSnapshots.forEach((snapshot) => {
      rollbackVersion += 1
      viewportStore.setCellSnapshot(createSupersedingCellSnapshot(snapshot, rollbackVersion))
    })
  }
}
