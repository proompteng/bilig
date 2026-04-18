import type { AxisKind } from './axis-map.js'
import type { SheetAxisMap } from './sheet-axis-map.js'
import type { CellPageStore } from './cell-page-store.js'

export interface LogicalVisibleAxisRef {
  readonly index: number
  readonly id: string | undefined
}

export interface LogicalVisibleCellRef {
  readonly sheetId: number
  readonly row: number
  readonly col: number
  readonly rowRef: LogicalVisibleAxisRef
  readonly colRef: LogicalVisibleAxisRef
}

export class LogicalSheetStore {
  constructor(
    private sheetId: number,
    private readonly axisMap: SheetAxisMap,
    private readonly cellPages: CellPageStore,
  ) {}

  setSheetId(sheetId: number): void {
    this.sheetId = sheetId
  }

  resolveVisibleAxis(axis: AxisKind, index: number): LogicalVisibleAxisRef {
    return {
      index,
      id: this.axisMap.snapshot(axis, index, 1)[0]?.id,
    }
  }

  resolveVisibleCell(row: number, col: number): LogicalVisibleCellRef {
    return {
      sheetId: this.sheetId,
      row,
      col,
      rowRef: this.resolveVisibleAxis('row', row),
      colRef: this.resolveVisibleAxis('column', col),
    }
  }

  getVisibleCell(row: number, col: number): number | undefined {
    return this.cellPages.get({ sheetId: this.sheetId, row, col })
  }

  setVisibleCell(row: number, col: number, cellIndex: number): LogicalVisibleCellRef {
    const resolved = this.resolveVisibleCell(row, col)
    this.cellPages.set(resolved, cellIndex)
    return resolved
  }

  deleteVisibleCell(row: number, col: number): boolean {
    return this.cellPages.delete({ sheetId: this.sheetId, row, col })
  }

  moveVisibleCell(fromRow: number, fromCol: number, toRow: number, toCol: number, cellIndex: number): LogicalVisibleCellRef {
    this.deleteVisibleCell(fromRow, fromCol)
    return this.setVisibleCell(toRow, toCol, cellIndex)
  }
}
