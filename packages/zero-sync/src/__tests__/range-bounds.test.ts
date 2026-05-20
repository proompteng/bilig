import { describe, expect, it } from 'vitest'
import { cellCoordinatesWithinBounds, intersectRangeBounds, normalizeRangeBounds, rangeBoundsForSheet } from '../range-bounds.js'

describe('range bounds helpers', () => {
  it('normalizes reversed workbook ranges with sheet authority attached', () => {
    expect(
      normalizeRangeBounds({
        sheetName: 'Sheet1',
        startAddress: 'D5',
        endAddress: 'B2',
      }),
    ).toEqual({
      sheetName: 'Sheet1',
      rowStart: 1,
      rowEnd: 4,
      colStart: 1,
      colEnd: 3,
    })
  })

  it('treats ranges from other sheets as no authority over the current sheet', () => {
    expect(
      rangeBoundsForSheet('Sheet1', {
        sheetName: 'Other',
        startAddress: 'A1',
        endAddress: 'C3',
      }),
    ).toBeNull()
  })

  it('does not intersect bounds across sheets', () => {
    const sheetOne = normalizeRangeBounds({
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'C3',
    })
    const other = normalizeRangeBounds({
      sheetName: 'Other',
      startAddress: 'B2',
      endAddress: 'D4',
    })

    expect(intersectRangeBounds(sheetOne, other)).toBeNull()
    expect(cellCoordinatesWithinBounds(1, 1, sheetOne)).toBe(true)
    expect(cellCoordinatesWithinBounds(5, 5, sheetOne)).toBe(false)
  })
})
