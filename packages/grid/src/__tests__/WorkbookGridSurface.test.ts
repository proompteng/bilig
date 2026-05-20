import { describe, expect, test } from 'vitest'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import {
  hasSelectionTargetChanged,
  resolveWorkbookGridSurfaceDisplayCell,
  resolveWorkbookGridSurfaceDisplaySelection,
  resolveWorkbookGridSurfaceTextOcclusionRanges,
} from '../WorkbookGridSurface.js'
import { getGridMetrics } from '../gridMetrics.js'
import { createColumnSliceSelection, createGridSelection, createRangeSelection, createRowSliceSelection } from '../gridSelection.js'
import { resolveViewportScrollPosition } from '../workbookGridViewport.js'

describe('WorkbookGridSurface selection autoscroll', () => {
  test('autoscrolls on first selection target', () => {
    expect(
      hasSelectionTargetChanged(null, {
        sheetName: 'Sheet1',
        col: 2,
        row: 4,
      }),
    ).toBe(true)
  })

  test('does not autoscroll again when the selection target is unchanged', () => {
    expect(
      hasSelectionTargetChanged(
        {
          sheetName: 'Sheet1',
          col: 2,
          row: 4,
        },
        {
          sheetName: 'Sheet1',
          col: 2,
          row: 4,
        },
      ),
    ).toBe(false)
  })

  test('autoscrolls when the selected cell changes', () => {
    expect(
      hasSelectionTargetChanged(
        {
          sheetName: 'Sheet1',
          col: 2,
          row: 4,
        },
        {
          sheetName: 'Sheet1',
          col: 3,
          row: 4,
        },
      ),
    ).toBe(true)
  })

  test('autoscrolls when the active sheet changes', () => {
    expect(
      hasSelectionTargetChanged(
        {
          sheetName: 'Sheet1',
          col: 2,
          row: 4,
        },
        {
          sheetName: 'Sheet2',
          col: 2,
          row: 4,
        },
      ),
    ).toBe(true)
  })

  test('restores a saved viewport to the recorded top-left cell', () => {
    expect(
      resolveViewportScrollPosition({
        viewport: {
          rowStart: 14,
          colStart: 3,
        },
        sortedColumnWidthOverrides: [],
        sortedRowHeightOverrides: [],
        gridMetrics: getGridMetrics(),
      }),
    ).toEqual({
      scrollLeft: getGridMetrics().columnWidth * 3,
      scrollTop: getGridMetrics().rowHeight * 14,
    })
  })

  test('uses the committed cell when stale render selection drifts during scroll', () => {
    const committedSelection = createGridSelection(0, 0)
    const staleRenderSelection = createGridSelection(0, 24)

    expect(
      resolveWorkbookGridSurfaceDisplaySelection({
        activeHeaderDrag: null,
        committedCellSelection: committedSelection,
        isEditingCell: false,
        isFillHandleDragging: false,
        isRangeMoveDragging: false,
        renderGridSelection: staleRenderSelection,
        renderSelectionRange: staleRenderSelection.current?.range,
        selectedCell: [0, 0],
      }),
    ).toBe(committedSelection)
  })

  test('keeps live range selections when the active cell still matches the selected address', () => {
    const committedSelection = createGridSelection(0, 0)
    const rangeSelection = {
      ...createGridSelection(0, 0),
      current: {
        cell: [0, 0] as const,
        range: { x: 0, y: 0, width: 3, height: 4 },
        rangeStack: [],
      },
    }

    expect(
      resolveWorkbookGridSurfaceDisplaySelection({
        activeHeaderDrag: null,
        committedCellSelection: committedSelection,
        isEditingCell: false,
        isFillHandleDragging: false,
        isRangeMoveDragging: false,
        renderGridSelection: rangeSelection,
        renderSelectionRange: rangeSelection.current?.range,
        selectedCell: [0, 0],
      }),
    ).toBe(rangeSelection)
  })

  test('keeps restored range selections when the active cell is not the top-left cell', () => {
    const committedSelection = createGridSelection(3, 4)
    const restoredRangeSelection = {
      ...createGridSelection(1, 1),
      current: {
        cell: [3, 4] as const,
        range: { x: 1, y: 1, width: 3, height: 4 },
        rangeStack: [],
      },
    }

    expect(
      resolveWorkbookGridSurfaceDisplaySelection({
        activeHeaderDrag: null,
        committedCellSelection: committedSelection,
        isEditingCell: false,
        isFillHandleDragging: false,
        isRangeMoveDragging: false,
        renderGridSelection: restoredRangeSelection,
        renderSelectionRange: restoredRangeSelection.current?.range,
        selectedCell: [3, 4],
      }),
    ).toBe(restoredRangeSelection)
  })

  test('rejects stale render ranges that no longer contain the committed active cell', () => {
    const committedSelection = createGridSelection(6, 7)
    const staleRangeSelection = {
      ...createGridSelection(6, 7),
      current: {
        cell: [6, 7] as const,
        range: { x: 1, y: 1, width: 3, height: 4 },
        rangeStack: [],
      },
    }

    expect(
      resolveWorkbookGridSurfaceDisplaySelection({
        activeHeaderDrag: null,
        committedCellSelection: committedSelection,
        isEditingCell: false,
        isFillHandleDragging: false,
        isRangeMoveDragging: false,
        renderGridSelection: staleRangeSelection,
        renderSelectionRange: staleRangeSelection.current?.range,
        selectedCell: [6, 7],
      }),
    ).toBe(committedSelection)
  })

  test('keeps pending local range selections visible while the committed active cell catches up', () => {
    const committedSelection = createGridSelection(0, 0)
    const pendingLocalRangeSelection = {
      ...createGridSelection(3, 3),
      current: {
        cell: [3, 3] as const,
        range: { x: 1, y: 1, width: 3, height: 3 },
        rangeStack: [],
      },
    }

    expect(
      resolveWorkbookGridSurfaceDisplaySelection({
        activeHeaderDrag: null,
        committedCellSelection: committedSelection,
        hasPendingLocalSelection: true,
        isEditingCell: false,
        isFillHandleDragging: false,
        isRangeMoveDragging: false,
        renderGridSelection: pendingLocalRangeSelection,
        renderSelectionRange: pendingLocalRangeSelection.current?.range,
        selectedCell: [0, 0],
      }),
    ).toBe(pendingLocalRangeSelection)
  })

  test('uses the resolved display selection cell while local range selection is ahead of committed state', () => {
    const pendingLocalRangeSelection = {
      ...createGridSelection(3, 5),
      current: {
        cell: [3, 5] as const,
        range: { x: 1, y: 2, width: 5, height: 7 },
        rangeStack: [],
      },
    }

    expect(
      resolveWorkbookGridSurfaceDisplayCell({
        committedCell: [0, 0],
        displayGridSelection: pendingLocalRangeSelection,
      }),
    ).toEqual([3, 5])
  })

  test('rejects stale axis selections after the committed active cell changes', () => {
    const committedSelection = createGridSelection(6, 7)
    const staleColumnSelection = createColumnSliceSelection(2, 4, 0)
    const staleRowSelection = createRowSliceSelection(0, 2, 4)

    expect(
      resolveWorkbookGridSurfaceDisplaySelection({
        activeHeaderDrag: null,
        committedCellSelection: committedSelection,
        isEditingCell: false,
        isFillHandleDragging: false,
        isRangeMoveDragging: false,
        renderGridSelection: staleColumnSelection,
        renderSelectionRange: staleColumnSelection.current?.range,
        selectedCell: [6, 7],
      }),
    ).toBe(committedSelection)

    expect(
      resolveWorkbookGridSurfaceDisplaySelection({
        activeHeaderDrag: null,
        committedCellSelection: committedSelection,
        isEditingCell: false,
        isFillHandleDragging: false,
        isRangeMoveDragging: false,
        renderGridSelection: staleRowSelection,
        renderSelectionRange: staleRowSelection.current?.range,
        selectedCell: [6, 7],
      }),
    ).toBe(committedSelection)
  })

  test('keeps live axis selections when the render active cell still matches the committed active cell', () => {
    const committedSelection = createGridSelection(2, 7)
    const columnSelection = createColumnSliceSelection(2, 4, 7)
    const rowSelection = createRowSliceSelection(2, 7, 9)

    expect(
      resolveWorkbookGridSurfaceDisplaySelection({
        activeHeaderDrag: null,
        committedCellSelection: committedSelection,
        isEditingCell: false,
        isFillHandleDragging: false,
        isRangeMoveDragging: false,
        renderGridSelection: columnSelection,
        renderSelectionRange: columnSelection.current?.range,
        selectedCell: [2, 7],
      }),
    ).toBe(columnSelection)

    expect(
      resolveWorkbookGridSurfaceDisplaySelection({
        activeHeaderDrag: null,
        committedCellSelection: committedSelection,
        isEditingCell: false,
        isFillHandleDragging: false,
        isRangeMoveDragging: false,
        renderGridSelection: rowSelection,
        renderSelectionRange: rowSelection.current?.range,
        selectedCell: [2, 7],
      }),
    ).toBe(rowSelection)
  })

  test('passes the visible range to native text occlusion for regular range selections', () => {
    const selection = createRangeSelection(createGridSelection(1, 1), [1, 1], [3, 4])

    expect(
      resolveWorkbookGridSurfaceTextOcclusionRanges({
        gridSelection: selection,
        selectionRange: selection.current?.range ?? null,
      }),
    ).toEqual([{ x: 1, y: 1, width: 3, height: 4 }])
  })

  test('expands column and row selections for native text occlusion instead of only clipping the active cell slice', () => {
    expect(
      resolveWorkbookGridSurfaceTextOcclusionRanges({
        gridSelection: createColumnSliceSelection(2, 4, 1),
        selectionRange: { x: 2, y: 1, width: 3, height: 1 },
      }),
    ).toEqual([{ x: 2, y: 0, width: 3, height: MAX_ROWS }])

    expect(
      resolveWorkbookGridSurfaceTextOcclusionRanges({
        gridSelection: createRowSliceSelection(1, 4, 6),
        selectionRange: { x: 1, y: 4, width: 1, height: 3 },
      }),
    ).toEqual([{ x: 0, y: 4, width: MAX_COLS, height: 3 }])
  })
})
