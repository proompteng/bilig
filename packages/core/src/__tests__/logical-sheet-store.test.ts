import { describe, expect, it } from 'vitest'
import { makeLogicalCellKey } from '../workbook-store.js'
import { CellPageStore } from '../storage/cell-page-store.js'
import { LogicalSheetStore } from '../storage/logical-sheet-store.js'
import { SheetAxisMap } from '../storage/sheet-axis-map.js'

describe('LogicalSheetStore', () => {
  it('resolves stable visible row and column ids and keeps cell pages attached to those ids', () => {
    const cellPages = new CellPageStore(new Map<string, number>(), (location) =>
      makeLogicalCellKey(location.sheetId, location.rowId, location.colId),
    )
    const axisMap = new SheetAxisMap()
    const logical = new LogicalSheetStore(7, axisMap, cellPages)

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
    expect(logical.getVisibleCell(1, 1)).toBeUndefined()

    logical.setSheetId(9)
    expect(logical.resolveVisibleCell(0, 0).sheetId).toBe(9)
  })
})
