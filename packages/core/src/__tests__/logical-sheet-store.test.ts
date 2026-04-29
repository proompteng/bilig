import { describe, expect, it } from 'vitest'
import { makeLogicalCellKey } from '../workbook-store.js'
import { AxisResidentCellIndex } from '../storage/axis-resident-cell-index.js'
import { CellPageStore } from '../storage/cell-page-store.js'
import { CellAxisIdentityStore } from '../storage/cell-axis-identity-store.js'
import { LogicalSheetStore } from '../storage/logical-sheet-store.js'
import { SheetAxisMap } from '../storage/sheet-axis-map.js'

describe('LogicalSheetStore', () => {
  it('resolves stable visible row and column ids and keeps cell pages attached to those ids', () => {
    const cellPages = new CellPageStore(new Map<string, number>(), (location) =>
      makeLogicalCellKey(location.sheetId, location.rowId, location.colId),
    )
    const axisMap = new SheetAxisMap()
    const logical = new LogicalSheetStore(7, axisMap, cellPages, new CellAxisIdentityStore(), new AxisResidentCellIndex())

    logical.setVisibleCell(1, 1, 42, {
      createRowId: () => 'row-b',
      createColumnId: () => 'column-b',
    })

    expect(logical.resolveVisibleCell(1, 1)).toEqual({
      sheetId: 7,
      row: 1,
      col: 1,
      rowRef: { index: 1, id: 'row-b' },
      colRef: { index: 1, id: 'column-b' },
    })
    expect(logical.getVisibleCell(1, 1)).toBe(42)

    axisMap.move('row', 1, 1, 0)
    axisMap.move('column', 1, 1, 0)

    expect(logical.resolveVisibleCell(0, 0)).toEqual({
      sheetId: 7,
      row: 0,
      col: 0,
      rowRef: { index: 0, id: 'row-b' },
      colRef: { index: 0, id: 'column-b' },
    })
    expect(logical.getVisibleCell(0, 0)).toBe(42)
    expect(logical.getCellVisiblePosition(42)).toEqual({ row: 0, col: 0 })
    expect(logical.listResidentCellIndices('row', ['row-b'])).toEqual([42])
    expect(logical.getVisibleCell(1, 1)).toBeUndefined()

    const rowMatches: Array<{ cellIndex: number; axisIndex: number; rowId: string; colId: string }> = []
    logical.forEachResidentCellInAxisEntries('row', [{ id: 'row-b', index: 0 }], (cellIndex, identity, axisIndex) => {
      rowMatches.push({ cellIndex, axisIndex, rowId: identity.rowId, colId: identity.colId })
    })
    expect(rowMatches).toEqual([{ cellIndex: 42, axisIndex: 0, rowId: 'row-b', colId: 'column-b' }])

    logical.setSheetId(9)
    expect(logical.resolveVisibleCell(0, 0).sheetId).toBe(9)
  })
})
