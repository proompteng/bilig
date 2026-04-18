import { describe, expect, it } from 'vitest'
import { makeCellKey } from '../workbook-store.js'
import { CellPageStore } from '../storage/cell-page-store.js'
import { LogicalSheetStore } from '../storage/logical-sheet-store.js'
import { SheetAxisMap } from '../storage/sheet-axis-map.js'

describe('LogicalSheetStore', () => {
  it('resolves visible row and column ids and keeps visible cell pages in sync', () => {
    const cellPages = new CellPageStore(new Map<number, number>(), (location) => makeCellKey(location.sheetId, location.row, location.col))
    const axisMap = new SheetAxisMap()
    axisMap.replaceRange('row', 0, [
      { id: 'row-a', index: 0 },
      { id: 'row-b', index: 1 },
    ])
    axisMap.replaceRange('column', 0, [
      { id: 'column-a', index: 0 },
      { id: 'column-b', index: 1 },
    ])

    const logical = new LogicalSheetStore(7, axisMap, cellPages)

    expect(logical.resolveVisibleCell(1, 1)).toEqual({
      sheetId: 7,
      row: 1,
      col: 1,
      rowRef: { index: 1, id: 'row-b' },
      colRef: { index: 1, id: 'column-b' },
    })

    logical.setVisibleCell(1, 1, 42)
    expect(logical.getVisibleCell(1, 1)).toBe(42)

    logical.moveVisibleCell(1, 1, 0, 0, 42)
    expect(logical.getVisibleCell(1, 1)).toBeUndefined()
    expect(logical.getVisibleCell(0, 0)).toBe(42)

    axisMap.move('row', 0, 1, 2)
    axisMap.move('column', 0, 1, 2)
    expect(logical.resolveVisibleCell(0, 0)).toEqual({
      sheetId: 7,
      row: 0,
      col: 0,
      rowRef: { index: 0, id: 'row-b' },
      colRef: { index: 0, id: 'column-b' },
    })

    logical.setSheetId(9)
    expect(logical.resolveVisibleCell(0, 0).sheetId).toBe(9)
  })
})
