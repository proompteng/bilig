import { describe, expect, test } from 'vitest'
import {
  PRODUCT_COLUMN_WIDTH,
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_HEIGHT,
  PRODUCT_ROW_MARKER_WIDTH,
  getResolvedRowHeight,
  getGridMetrics,
  getVisibleColumnBounds,
  getVisibleRowBounds,
  resolveColumnAtClientX,
  resolveRowAtClientY,
  getResolvedColumnWidth,
} from '../gridMetrics.js'

describe('gridMetrics', () => {
  test('returns the product grid contract', () => {
    expect(getGridMetrics()).toEqual({
      columnWidth: PRODUCT_COLUMN_WIDTH,
      rowHeight: PRODUCT_ROW_HEIGHT,
      headerHeight: PRODUCT_HEADER_HEIGHT,
      rowMarkerWidth: PRODUCT_ROW_MARKER_WIDTH,
    })
  })

  test('resolves visible column bounds and pointer columns with overrides', () => {
    const columnWidths = { 1: 140, 2: 88 }
    const bounds = getVisibleColumnBounds({ x: 0, width: 4 }, 46, 16384, columnWidths, PRODUCT_COLUMN_WIDTH)

    expect(bounds.map((column) => ({ index: column.index, left: column.left, width: column.width }))).toEqual([
      { index: 0, left: 46, width: PRODUCT_COLUMN_WIDTH },
      { index: 1, left: 46 + PRODUCT_COLUMN_WIDTH, width: 140 },
      { index: 2, left: 46 + PRODUCT_COLUMN_WIDTH + 140, width: 88 },
      { index: 3, left: 46 + PRODUCT_COLUMN_WIDTH + 140 + 88, width: PRODUCT_COLUMN_WIDTH },
    ])

    expect(resolveColumnAtClientX(46 + 20, { x: 0, width: 4 }, 46, 16384, columnWidths, PRODUCT_COLUMN_WIDTH)).toBe(0)
    expect(resolveColumnAtClientX(46 + PRODUCT_COLUMN_WIDTH + 30, { x: 0, width: 4 }, 46, 16384, columnWidths, PRODUCT_COLUMN_WIDTH)).toBe(
      1,
    )
    expect(
      resolveColumnAtClientX(46 + PRODUCT_COLUMN_WIDTH + 140 + 10, { x: 0, width: 4 }, 46, 16384, columnWidths, PRODUCT_COLUMN_WIDTH),
    ).toBe(2)
    expect(getResolvedColumnWidth(columnWidths, 3, PRODUCT_COLUMN_WIDTH)).toBe(PRODUCT_COLUMN_WIDTH)
  })

  test('resolves visible row bounds and pointer rows with overrides', () => {
    const rowHeights = { 1: 34, 2: 28 }
    const bounds = getVisibleRowBounds({ y: 0, height: 4 }, 57, 1_048_576, rowHeights, PRODUCT_ROW_HEIGHT)

    expect(bounds.map((row) => ({ index: row.index, top: row.top, height: row.height }))).toEqual([
      { index: 0, top: 57, height: PRODUCT_ROW_HEIGHT },
      { index: 1, top: 57 + PRODUCT_ROW_HEIGHT, height: 34 },
      { index: 2, top: 57 + PRODUCT_ROW_HEIGHT + 34, height: 28 },
      { index: 3, top: 57 + PRODUCT_ROW_HEIGHT + 34 + 28, height: PRODUCT_ROW_HEIGHT },
    ])

    expect(resolveRowAtClientY(57 + 6, { y: 0, height: 4 }, 57, 1_048_576, rowHeights, PRODUCT_ROW_HEIGHT)).toBe(0)
    expect(resolveRowAtClientY(57 + PRODUCT_ROW_HEIGHT + 12, { y: 0, height: 4 }, 57, 1_048_576, rowHeights, PRODUCT_ROW_HEIGHT)).toBe(1)
    expect(resolveRowAtClientY(57 + PRODUCT_ROW_HEIGHT + 34 + 4, { y: 0, height: 4 }, 57, 1_048_576, rowHeights, PRODUCT_ROW_HEIGHT)).toBe(
      2,
    )
    expect(getResolvedRowHeight(rowHeights, 3, PRODUCT_ROW_HEIGHT)).toBe(PRODUCT_ROW_HEIGHT)
  })

  test('skips collapsed hidden axes during pointer resolution', () => {
    expect(resolveColumnAtClientX(46 + PRODUCT_COLUMN_WIDTH, { x: 0, width: 4 }, 46, 16384, { 1: 0, 2: 0 }, PRODUCT_COLUMN_WIDTH)).toBe(3)

    expect(resolveRowAtClientY(57 + PRODUCT_ROW_HEIGHT, { y: 0, height: 4 }, 57, 1_048_576, { 1: 0, 2: 0 }, PRODUCT_ROW_HEIGHT)).toBe(3)
  })
})
