import { describe, expect, it } from 'vitest'
import { getGridMetrics } from '../gridMetrics.js'
import { resolveWorkbookHeaderCellBounds } from '../useWorkbookHeaderCellBounds.js'

describe('resolveWorkbookHeaderCellBounds', () => {
  it('resolves frozen and resident header-local cell bounds', () => {
    const gridMetrics = getGridMetrics()
    const common = {
      columnWidths: { 0: 150, 12: 140 },
      freezeCols: 1,
      freezeRows: 1,
      frozenColumnWidth: 150,
      frozenRowHeight: 30,
      gridMetrics,
      residentViewport: { colStart: 10, colEnd: 20, rowStart: 8, rowEnd: 18 },
      rowHeights: { 0: 30, 9: 40 },
      sortedColumnWidthOverrides: [
        [0, 150],
        [12, 140],
      ] as const,
      sortedRowHeightOverrides: [
        [0, 30],
        [9, 40],
      ] as const,
    }

    expect(resolveWorkbookHeaderCellBounds({ ...common, col: 0, row: 0 })).toEqual({
      height: 30,
      width: 150,
      x: gridMetrics.rowMarkerWidth,
      y: gridMetrics.headerHeight,
    })
    expect(resolveWorkbookHeaderCellBounds({ ...common, col: 12, row: 9 })).toEqual({
      height: 40,
      width: 140,
      x: gridMetrics.rowMarkerWidth + 150 + 2 * gridMetrics.columnWidth,
      y: gridMetrics.headerHeight + 30 + gridMetrics.rowHeight,
    })
    expect(resolveWorkbookHeaderCellBounds({ ...common, col: -1, row: 0 })).toBeUndefined()
  })
})
