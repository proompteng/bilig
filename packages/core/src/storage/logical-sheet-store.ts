import type { AxisEntrySnapshot, AxisKind } from './axis-map.js'
import type { AxisResidentCellIndex } from './axis-resident-cell-index.js'
import type { CellPageStore, LogicalCellLocation } from './cell-page-store.js'
import type { CellAxisIdentity, CellAxisIdentityStore } from './cell-axis-identity-store.js'
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

export interface LogicalVisibleCellPosition {
  readonly row: number
  readonly col: number
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
    private readonly cellIdentities: CellAxisIdentityStore,
    private readonly residentCells: AxisResidentCellIndex,
  ) {}

  setSheetId(sheetId: number): void {
    const previousSheetId = this.sheetId
    this.sheetId = sheetId
    this.cellIdentities.forEach((identity, cellIndex) => {
      if (identity.sheetId === previousSheetId) {
        this.cellIdentities.set(cellIndex, { ...identity, sheetId })
      }
    })
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
    return this.setVisibleCellInternal(row, col, cellIndex, factories, false)
  }

  setNewVisibleCell(row: number, col: number, cellIndex: number, factories: LogicalAxisIdFactories): LogicalVisibleCellRef {
    return this.setVisibleCellInternal(row, col, cellIndex, factories, true)
  }

  setNewVisibleCellWithAxisIds(row: number, col: number, cellIndex: number, rowId: string, colId: string): LogicalVisibleCellRef {
    const resolved: LogicalVisibleCellRef = {
      sheetId: this.sheetId,
      row,
      col,
      rowRef: { index: row, id: rowId },
      colRef: { index: col, id: colId },
    }
    this.cellPages.set(
      {
        sheetId: this.sheetId,
        rowId,
        colId,
      },
      cellIndex,
    )
    const identity: CellAxisIdentity = {
      sheetId: this.sheetId,
      rowId,
      colId,
    }
    this.cellIdentities.set(cellIndex, identity)
    this.residentCells.add(cellIndex, identity)
    return resolved
  }

  private setVisibleCellInternal(
    row: number,
    col: number,
    cellIndex: number,
    factories: LogicalAxisIdFactories,
    knownNewCell: boolean,
  ): LogicalVisibleCellRef {
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
    const identity: CellAxisIdentity = {
      sheetId: this.sheetId,
      rowId: resolved.rowRef.id!,
      colId: resolved.colRef.id!,
    }
    this.cellIdentities.set(cellIndex, identity)
    if (knownNewCell) {
      this.residentCells.add(cellIndex, identity)
    } else {
      this.residentCells.set(cellIndex, identity)
    }
    return resolved
  }

  deleteVisibleCell(row: number, col: number): boolean {
    const location = this.resolveVisibleLocation(row, col)
    if (!location) {
      return false
    }
    const cellIndex = this.cellPages.get(location)
    const deleted = this.cellPages.delete(location)
    if (deleted && cellIndex !== undefined) {
      this.cellIdentities.delete(cellIndex)
      this.residentCells.delete(cellIndex)
    }
    return deleted
  }

  deleteVisibleCellByIds(rowId: string, colId: string): boolean {
    const location = {
      sheetId: this.sheetId,
      rowId,
      colId,
    }
    const cellIndex = this.cellPages.get(location)
    const deleted = this.cellPages.delete(location)
    if (deleted && cellIndex !== undefined) {
      this.cellIdentities.delete(cellIndex)
      this.residentCells.delete(cellIndex)
    }
    return deleted
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

  getCellVisiblePosition(cellIndex: number): LogicalVisibleCellPosition | undefined {
    const identity = this.cellIdentities.get(cellIndex)
    if (!identity || identity.sheetId !== this.sheetId) {
      return undefined
    }
    const row = this.axisMap.indexOf('row', identity.rowId)
    const col = this.axisMap.indexOf('column', identity.colId)
    if (row < 0 || col < 0) {
      return undefined
    }
    return { row, col }
  }

  getCellVisibleAxisIndex(cellIndex: number, axis: AxisKind): number | undefined {
    const identity = this.cellIdentities.get(cellIndex)
    if (!identity || identity.sheetId !== this.sheetId) {
      return undefined
    }
    const index = axis === 'row' ? this.axisMap.indexOf('row', identity.rowId) : this.axisMap.indexOf('column', identity.colId)
    return index < 0 ? undefined : index
  }

  getCellIdentity(cellIndex: number): CellAxisIdentity | undefined {
    const identity = this.cellIdentities.get(cellIndex)
    return identity?.sheetId === this.sheetId ? identity : undefined
  }

  listResidentCellIndices(axis: AxisKind, axisIds: readonly string[]): number[] {
    return axis === 'row' ? this.residentCells.cellsInRows(axisIds) : this.residentCells.cellsInColumns(axisIds)
  }

  listResidentCellIndicesUnordered(axis: AxisKind, axisIds: readonly string[]): number[] {
    return axis === 'row' ? this.residentCells.cellsInRowsUnordered(axisIds) : this.residentCells.cellsInColumnsUnordered(axisIds)
  }

  forEachResidentCellInAxisEntries(
    axis: AxisKind,
    entries: readonly AxisEntrySnapshot[],
    callback: (cellIndex: number, identity: CellAxisIdentity, axisIndex: number) => void,
  ): void {
    const visitCell = (cellIndex: number, axisIndex: number): void => {
      const identity = this.cellIdentities.get(cellIndex)
      if (identity?.sheetId === this.sheetId) {
        callback(cellIndex, identity, axisIndex)
      }
    }
    entries.forEach((entry) => {
      if (axis === 'row') {
        this.residentCells.forEachCellInRow(entry.id, (cellIndex) => {
          visitCell(cellIndex, entry.index)
        })
        return
      }
      this.residentCells.forEachCellInColumn(entry.id, (cellIndex) => {
        visitCell(cellIndex, entry.index)
      })
    })
  }

  someResidentCellInAxisScope(
    axis: AxisKind,
    scope: { readonly start: number; readonly end?: number },
    predicate: (cellIndex: number, row: number, col: number) => boolean,
  ): boolean {
    const entries = this.axisMap.list(axis)
    for (let entryIndex = 0; entryIndex < entries.length; entryIndex += 1) {
      const entry = entries[entryIndex]!
      if (entry.index < scope.start || (scope.end !== undefined && entry.index >= scope.end)) {
        continue
      }
      let found = false
      const visitCell = (cellIndex: number): void => {
        if (found) {
          return
        }
        const identity = this.cellIdentities.get(cellIndex)
        if (identity?.sheetId !== this.sheetId) {
          return
        }
        const otherAxisIndex = axis === 'row' ? this.axisMap.indexOf('column', identity.colId) : this.axisMap.indexOf('row', identity.rowId)
        if (otherAxisIndex < 0) {
          return
        }
        const row = axis === 'row' ? entry.index : otherAxisIndex
        const col = axis === 'row' ? otherAxisIndex : entry.index
        found = predicate(cellIndex, row, col)
      }
      if (axis === 'row') {
        this.residentCells.forEachCellInRow(entry.id, visitCell)
      } else {
        this.residentCells.forEachCellInColumn(entry.id, visitCell)
      }
      if (found) {
        return true
      }
    }
    return false
  }

  forEachVisibleCellEntry(callback: (cellIndex: number, row: number, col: number) => void): void {
    const entries: Array<{ cellIndex: number; row: number; col: number }> = []
    this.cellIdentities.forEach((identity, cellIndex) => {
      if (identity.sheetId !== this.sheetId) {
        return
      }
      const row = this.axisMap.indexOf('row', identity.rowId)
      const col = this.axisMap.indexOf('column', identity.colId)
      if (row < 0 || col < 0) {
        return
      }
      entries.push({ cellIndex, row, col })
    })
    entries
      .toSorted((left, right) => left.row - right.row || left.col - right.col || left.cellIndex - right.cellIndex)
      .forEach((entry) => {
        callback(entry.cellIndex, entry.row, entry.col)
      })
  }

  forEachVisibleColumnCellEntry(col: number, callback: (cellIndex: number, row: number) => void): void {
    const colId = this.axisMap.getId('column', col)
    if (colId === undefined) {
      return
    }
    const entries = this.residentCells
      .cellsInColumn(colId)
      .flatMap((cellIndex): Array<{ cellIndex: number; row: number }> => {
        const position = this.getCellVisiblePosition(cellIndex)
        return position && position.col === col ? [{ cellIndex, row: position.row }] : []
      })
      .toSorted((left, right) => left.row - right.row || left.cellIndex - right.cellIndex)
    entries.forEach((entry) => {
      callback(entry.cellIndex, entry.row)
    })
  }
}
