import { MAX_COLS, MAX_ROWS, ValueTag, type CellValue } from '@bilig/protocol'
import { formatAddress } from '@bilig/formula'
import { orderWorkPaperCellChanges } from './change-order.js'
import { emptyValue, valuesEqual } from './work-paper-runtime-helpers.js'
import type { WorkPaperChange } from './work-paper-types.js'

const VISIBILITY_SHEET_STRIDE = MAX_ROWS * MAX_COLS

export interface WorkPaperVisibilityCellStore {
  readonly getValue: (cellIndex: number, readString: (id: number) => string) => CellValue
}

export interface WorkPaperVisibilityStringPool {
  readonly get: (id: number) => string
}

export interface WorkPaperVisibilitySheetRecord {
  readonly id: number
  readonly name: string
  readonly order: number
  readonly grid: {
    readonly forEachCellEntry: (visit: (cellIndex: number, row: number, col: number) => void) => void
  }
}

export interface SheetStateSnapshot {
  readonly sheetId: number
  readonly sheetName: string
  readonly order: number
  readonly cells: Map<number, CellValue>
}

export type VisibilitySnapshot = Map<number, SheetStateSnapshot>

export function captureWorkPaperVisibilitySnapshot(input: {
  readonly sheets: readonly WorkPaperVisibilitySheetRecord[]
  readonly cellStore: WorkPaperVisibilityCellStore
  readonly strings: WorkPaperVisibilityStringPool
}): VisibilitySnapshot {
  const snapshot: VisibilitySnapshot = new Map()
  input.sheets.forEach((sheet) => {
    const cells = new Map<number, CellValue>()
    sheet.grid.forEachCellEntry((cellIndex, row, col) => {
      const value = input.cellStore.getValue(cellIndex, (id) => input.strings.get(id))
      if (value.tag === ValueTag.Empty) {
        return
      }
      cells.set(sheetCellVisibilityKey(sheet.id, row, col), value)
    })
    snapshot.set(sheet.id, {
      sheetId: sheet.id,
      sheetName: sheet.name,
      order: sheet.order,
      cells,
    })
  })
  return snapshot
}

export function computeWorkPaperCellChangesFromVisibilitySnapshots(input: {
  readonly beforeVisibility: VisibilitySnapshot
  readonly afterVisibility: VisibilitySnapshot
  readonly sheets: readonly { readonly id: number; readonly order: number }[]
}): WorkPaperChange[] {
  const cellChanges: WorkPaperChange[] = []
  input.afterVisibility.forEach((afterSheet, sheetId) => {
    const beforeSheet = input.beforeVisibility.get(sheetId)
    const cellKeys = new Set<number>([...(beforeSheet?.cells.keys() ?? []), ...afterSheet.cells.keys()])
    ;[...cellKeys]
      .toSorted((left, right) => left - right)
      .forEach((cellKey) => {
        const beforeValue = beforeSheet?.cells.get(cellKey) ?? emptyValue()
        const afterValue = afterSheet.cells.get(cellKey) ?? emptyValue()
        if (valuesEqual(beforeValue, afterValue)) {
          return
        }
        const { row, col } = visibilityKeyPosition(afterSheet.sheetId, cellKey)
        const address = formatAddress(row, col)
        cellChanges.push({
          kind: 'cell',
          address: { sheet: sheetId, row, col },
          sheetName: afterSheet.sheetName,
          a1: address,
          newValue: afterValue,
        })
      })
  })
  return orderWorkPaperCellChanges(cellChanges, input.sheets)
}

function sheetCellVisibilityKey(sheetId: number, row: number, col: number): number {
  return sheetId * VISIBILITY_SHEET_STRIDE + row * MAX_COLS + col
}

function visibilityKeyPosition(sheetId: number, cellKey: number): { readonly row: number; readonly col: number } {
  const localKey = cellKey - sheetId * VISIBILITY_SHEET_STRIDE
  return {
    row: Math.floor(localKey / MAX_COLS),
    col: localKey % MAX_COLS,
  }
}
