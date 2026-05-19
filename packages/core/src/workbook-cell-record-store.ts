import { ValueTag } from '@bilig/protocol'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import { CellFlags, type CellStore } from './cell-store.js'
import { makeCellKey } from './workbook-cell-key-index.js'
import type { SheetRecord } from './workbook-sheet-record.js'

export interface EnsuredCell {
  cellIndex: number
  created: boolean
}

export class WorkbookCellRecordStore {
  constructor(
    private readonly options: {
      readonly cellStore: CellStore
      readonly cellKeyToIndex: Map<number, number>
      readonly cellFormats: Map<number, string>
      readonly getSheet: (sheetName: string) => SheetRecord | undefined
      readonly getOrCreateSheet: (sheetName: string) => SheetRecord
      readonly getSheetById: (sheetId: number) => SheetRecord | undefined
      readonly getSheetNameById: (sheetId: number) => string
      readonly createLogicalAxisId: (axis: 'row' | 'column') => string
    },
  ) {}

  ensureCell(sheetName: string, address: string): number {
    return this.ensureCellRecord(sheetName, address).cellIndex
  }

  ensureCellRecord(sheetName: string, address: string): EnsuredCell {
    const sheet = this.options.getOrCreateSheet(sheetName)
    const parsed = parseCellAddress(address, sheetName)
    return this.ensureCellAt(sheet.id, parsed.row, parsed.col)
  }

  ensureCellAt(sheetId: number, row: number, col: number): EnsuredCell {
    const sheet = this.options.getSheetById(sheetId)
    if (!sheet) {
      throw new Error(`Unknown sheet id: ${sheetId}`)
    }
    const physicalCellIndex = sheet.structureVersion === 1 ? sheet.grid.getPhysical(row, col) : -1
    if (physicalCellIndex !== -1) {
      if (
        sheet.logical.cellIdentityMatchesVisiblePosition(physicalCellIndex, row, col) &&
        this.options.cellStore.sheetIds[physicalCellIndex] === sheetId &&
        this.options.cellStore.rows[physicalCellIndex] === row &&
        this.options.cellStore.cols[physicalCellIndex] === col
      ) {
        return { cellIndex: physicalCellIndex, created: false }
      }
    }
    const existing = sheet.logical.getVisibleCell(row, col)
    if (existing !== undefined) {
      return { cellIndex: existing, created: false }
    }
    const cellIndex = this.options.cellStore.allocate(sheet.id, row, col)
    this.attachAllocatedCell(sheet.id, row, col, cellIndex)
    return { cellIndex, created: true }
  }

  attachAllocatedCell(sheetId: number, row: number, col: number, cellIndex: number): void {
    const sheet = this.options.getSheetById(sheetId)
    if (!sheet) {
      throw new Error(`Unknown sheet id: ${sheetId}`)
    }
    sheet.logical.setNewVisibleCell(row, col, cellIndex, {
      createRowId: () => this.options.createLogicalAxisId('row'),
      createColumnId: () => this.options.createLogicalAxisId('column'),
    })
    this.options.cellKeyToIndex.set(makeCellKey(sheet.id, row, col), cellIndex)
    sheet.grid.set(row, col, cellIndex)
  }

  ensureLogicalAxisId(sheetId: number, axis: 'row' | 'column', index: number): string {
    const sheet = this.options.getSheetById(sheetId)
    if (!sheet) {
      throw new Error(`Unknown sheet id: ${sheetId}`)
    }
    return sheet.logicalAxisMap.ensureId(axis, index, () => this.options.createLogicalAxisId(axis))
  }

  createLogicalAxisIdEnsurer(sheetId: number, axis: 'row' | 'column'): (index: number) => string {
    const sheet = this.options.getSheetById(sheetId)
    if (!sheet) {
      throw new Error(`Unknown sheet id: ${sheetId}`)
    }
    return (index) => sheet.logicalAxisMap.ensureId(axis, index, () => this.options.createLogicalAxisId(axis))
  }

  attachAllocatedCellWithLogicalAxisIds(sheetId: number, row: number, col: number, cellIndex: number, rowId: string, colId: string): void {
    const sheet = this.options.getSheetById(sheetId)
    if (!sheet) {
      throw new Error(`Unknown sheet id: ${sheetId}`)
    }
    sheet.logical.setNewVisibleCellWithAxisIds(row, col, cellIndex, rowId, colId)
    this.options.cellKeyToIndex.set(makeCellKey(sheet.id, row, col), cellIndex)
    sheet.grid.set(row, col, cellIndex)
  }

  getCellIndex(sheetName: string, address: string): number | undefined {
    const sheet = this.options.getSheet(sheetName)
    if (!sheet) return undefined
    const parsed = parseCellAddress(address, sheetName)
    if (sheet.structureVersion === 1) {
      const physicalCellIndex = sheet.grid.getPhysical(parsed.row, parsed.col)
      if (
        physicalCellIndex !== -1 &&
        sheet.logical.cellIdentityMatchesVisiblePosition(physicalCellIndex, parsed.row, parsed.col) &&
        this.options.cellStore.sheetIds[physicalCellIndex] === sheet.id &&
        this.options.cellStore.rows[physicalCellIndex] === parsed.row &&
        this.options.cellStore.cols[physicalCellIndex] === parsed.col
      ) {
        return physicalCellIndex
      }
    }
    return sheet.logical.getVisibleCell(parsed.row, parsed.col)
  }

  getCellIndexAt(sheetId: number, row: number, col: number): number | undefined {
    const sheet = this.options.getSheetById(sheetId)
    if (!sheet) {
      return undefined
    }
    if (sheet.structureVersion === 1) {
      const physicalCellIndex = sheet.grid.getPhysical(row, col)
      if (
        physicalCellIndex !== -1 &&
        sheet.logical.cellIdentityMatchesVisiblePosition(physicalCellIndex, row, col) &&
        this.options.cellStore.sheetIds[physicalCellIndex] === sheetId &&
        this.options.cellStore.rows[physicalCellIndex] === row &&
        this.options.cellStore.cols[physicalCellIndex] === col
      ) {
        return physicalCellIndex
      }
    }
    return sheet.logical.getVisibleCell(row, col)
  }

  getAddress(index: number): string {
    const position = this.getCellPosition(index)
    return formatAddress(position?.row ?? this.options.cellStore.rows[index]!, position?.col ?? this.options.cellStore.cols[index]!)
  }

  getQualifiedAddress(index: number): string {
    return `${this.options.getSheetNameById(this.options.cellStore.sheetIds[index]!)}!${this.getAddress(index)}`
  }

  getCellPosition(index: number): { sheetId: number; row: number; col: number } | undefined {
    const sheetId = this.options.cellStore.sheetIds[index]
    if (sheetId === undefined || sheetId === 0) {
      return undefined
    }
    const sheet = this.options.getSheetById(sheetId)
    if (sheet?.structureVersion === 1) {
      return { sheetId, row: this.options.cellStore.rows[index]!, col: this.options.cellStore.cols[index]! }
    }
    const logicalPosition = sheet?.logical.getCellVisiblePosition(index)
    if (logicalPosition) {
      return { sheetId, row: logicalPosition.row, col: logicalPosition.col }
    }
    if (sheet?.logical.getCellIdentity(index)) {
      return undefined
    }
    const row = this.options.cellStore.rows[index]
    const col = this.options.cellStore.cols[index]
    if (row === undefined || col === undefined) {
      return undefined
    }
    return { sheetId, row, col }
  }

  getCellAxisIndex(index: number, axis: 'row' | 'column'): number | undefined {
    const sheetId = this.options.cellStore.sheetIds[index]
    if (sheetId === undefined || sheetId === 0) {
      return undefined
    }
    const sheet = this.options.getSheetById(sheetId)
    if (sheet?.structureVersion === 1) {
      return axis === 'row' ? this.options.cellStore.rows[index] : this.options.cellStore.cols[index]
    }
    const logicalIndex = sheet?.logical.getCellVisibleAxisIndex(index, axis)
    if (logicalIndex !== undefined) {
      return logicalIndex
    }
    if (sheet?.logical.getCellIdentity(index)) {
      return undefined
    }
    return axis === 'row' ? this.options.cellStore.rows[index] : this.options.cellStore.cols[index]
  }

  detachCellIndex(index: number): boolean {
    const sheetId = this.options.cellStore.sheetIds[index]
    if (!sheetId) {
      return false
    }
    const sheet = this.options.getSheetById(sheetId)
    const position = this.getCellPosition(index)
    const row = position?.row
    const col = position?.col
    if (sheet && row !== undefined && col !== undefined) {
      if (sheet.logical.getVisibleCell(row, col) === index) {
        sheet.logical.deleteVisibleCell(row, col)
      }
      const key = makeCellKey(sheet.id, row, col)
      if (this.options.cellKeyToIndex.get(key) === index) {
        this.options.cellKeyToIndex.delete(key)
      }
      const visibleGridIndex = sheet.structureVersion === 1 ? sheet.grid.getPhysical(row, col) : sheet.grid.get(row, col)
      if (visibleGridIndex === index) {
        sheet.grid.clear(row, col)
      }
    }
    this.options.cellStore.flags[index] = (this.options.cellStore.flags[index] ?? 0) & ~CellFlags.Materialized
    return true
  }

  pruneCellIfEmpty(index: number): boolean {
    const sheetId = this.options.cellStore.sheetIds[index]
    if (!sheetId) {
      return false
    }
    const sheet = this.options.getSheetById(sheetId)
    if (!sheet) {
      return false
    }
    const position = this.getCellPosition(index)
    const row = position?.row
    const col = position?.col
    if (row === undefined || col === undefined) {
      return false
    }
    const value = this.options.cellStore.getValue(index, () => '')
    const flags = this.options.cellStore.flags[index] ?? 0
    if (
      value.tag !== ValueTag.Empty ||
      this.options.cellFormats.has(index) ||
      (flags &
        (CellFlags.HasFormula | CellFlags.AuthoredBlank | CellFlags.SpillChild | CellFlags.PivotOutput | CellFlags.PendingDelete)) !==
        0
    ) {
      return false
    }
    if (sheet.grid.get(row, col) !== index) {
      return false
    }
    return this.detachCellIndex(index)
  }
}
