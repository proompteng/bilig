import type { AxisKind } from './axis-map.js'
import type { CellPageStore, LogicalCellLocation } from './cell-page-store.js'
import type { SheetAxisMap } from './sheet-axis-map.js'

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

export interface LogicalAxisIdFactories {
  readonly createRowId: () => string
  readonly createColumnId: () => string
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
      id: this.axisMap.getId(axis, index),
    }
  }

  ensureVisibleAxis(axis: AxisKind, index: number, createId: () => string): LogicalVisibleAxisRef {
    return {
      index,
      id: this.axisMap.ensureId(axis, index, createId),
    }
  }

  private resolveVisibleLocation(row: number, col: number): LogicalCellLocation | undefined {
    const rowId = this.axisMap.getId('row', row)
    const colId = this.axisMap.getId('column', col)
    if (rowId === undefined || colId === undefined) {
      return undefined
    }
    return {
      sheetId: this.sheetId,
      rowId,
      colId,
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
    const location = this.resolveVisibleLocation(row, col)
    return location ? this.cellPages.get(location) : undefined
  }

  setVisibleCell(row: number, col: number, cellIndex: number, factories: LogicalAxisIdFactories): LogicalVisibleCellRef {
    const resolved: LogicalVisibleCellRef = {
      sheetId: this.sheetId,
      row,
      col,
      rowRef: this.ensureVisibleAxis('row', row, factories.createRowId),
      colRef: this.ensureVisibleAxis('column', col, factories.createColumnId),
    }
    this.cellPages.set(
      {
        sheetId: this.sheetId,
        rowId: resolved.rowRef.id!,
        colId: resolved.colRef.id!,
      },
      cellIndex,
    )
    return resolved
  }

  deleteVisibleCell(row: number, col: number): boolean {
    const location = this.resolveVisibleLocation(row, col)
    return location ? this.cellPages.delete(location) : false
  }

  deleteVisibleCellByIds(rowId: string, colId: string): boolean {
    return this.cellPages.delete({
      sheetId: this.sheetId,
      rowId,
      colId,
    })
  }

  moveVisibleCell(
    fromRow: number,
    fromCol: number,
    toRow: number,
    toCol: number,
    cellIndex: number,
    factories: LogicalAxisIdFactories,
  ): LogicalVisibleCellRef {
    this.deleteVisibleCell(fromRow, fromCol)
    return this.setVisibleCell(toRow, toCol, cellIndex, factories)
  }
}
