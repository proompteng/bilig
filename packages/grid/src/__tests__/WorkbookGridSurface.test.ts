import { describe, expect, test } from 'vitest'
import { hasSelectionTargetChanged, resolveWorkbookGridSurfaceDisplaySelection } from '../WorkbookGridSurface.js'
import { getGridMetrics } from '../gridMetrics.js'
import { createGridSelection } from '../gridSelection.js'
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
})
