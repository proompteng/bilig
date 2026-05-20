import { parseCellAddress } from '@bilig/formula'
import type { WorkbookPivotRecord } from '../../workbook-store.js'
import type { U32 } from '../runtime-state.js'

export function shouldRefreshPivotSource(
  pivot: WorkbookPivotRecord,
  changed: readonly number[] | U32,
  workbook: {
    readonly cellStore: {
      readonly sheetIds: ArrayLike<number | undefined>
      readonly rows: ArrayLike<number | undefined>
      readonly cols: ArrayLike<number | undefined>
    }
    readonly getSheet: (sheetName: string) => { readonly id: number } | undefined
    readonly getCellPosition: (cellIndex: number) => { readonly row: number; readonly col: number } | undefined
  },
): boolean {
  if (!pivot.source || pivot.cacheOnly) {
    return false
  }
  const ownerSheet = workbook.getSheet(pivot.source.sheetName)
  if (!ownerSheet) {
    return true
  }
  const ownerStart = parseCellAddress(pivot.source.startAddress, pivot.source.sheetName)
  const ownerEnd = parseCellAddress(pivot.source.endAddress, pivot.source.sheetName)
  for (let index = 0; index < changed.length; index += 1) {
    const cellIndex = changed[index]!
    const sheetId = workbook.cellStore.sheetIds[cellIndex]
    if (sheetId === undefined || sheetId !== ownerSheet.id) {
      continue
    }
    const position = workbook.getCellPosition(cellIndex)
    const row = position?.row ?? workbook.cellStore.rows[cellIndex] ?? -1
    const col = position?.col ?? workbook.cellStore.cols[cellIndex] ?? -1
    if (row >= ownerStart.row && row <= ownerEnd.row && col >= ownerStart.col && col <= ownerEnd.col) {
      return true
    }
  }
  return false
}

export function refreshPivotOutputsForChangedCells(input: {
  readonly changed: readonly number[] | U32
  readonly forceAll: boolean
  readonly workbook: Parameters<typeof shouldRefreshPivotSource>[2] & {
    readonly listPivots: () => readonly WorkbookPivotRecord[]
  }
  readonly materializePivot: (pivot: WorkbookPivotRecord) => readonly number[]
  readonly emptyChangedSet: () => U32
}): U32 {
  const pivots = input.workbook.listPivots()
  if (pivots.length === 0 || (!input.forceAll && input.changed.length === 0)) {
    return input.emptyChangedSet()
  }
  const changedCellIndices: number[] = []
  const changedSeen = new Set<number>()
  for (let index = 0; index < pivots.length; index += 1) {
    const pivot = pivots[index]!
    if (!input.forceAll && !shouldRefreshPivotSource(pivot, input.changed, input.workbook)) {
      continue
    }
    const pivotChanges = input.materializePivot(pivot)
    for (let changeIndex = 0; changeIndex < pivotChanges.length; changeIndex += 1) {
      const cellIndex = pivotChanges[changeIndex]!
      if (!changedSeen.has(cellIndex)) {
        changedSeen.add(cellIndex)
        changedCellIndices.push(cellIndex)
      }
    }
  }
  return changedCellIndices.length === 0 ? input.emptyChangedSet() : Uint32Array.from(changedCellIndices)
}
