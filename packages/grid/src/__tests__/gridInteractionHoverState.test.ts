import { describe, expect, it, vi } from 'vitest'
import { createGridGeometrySnapshot } from '../gridGeometry.js'
import { resolveWorkbookGridHoverState } from '../gridInteractionHoverState.js'
import { getGridMetrics } from '../gridMetrics.js'
import type { VisibleRegionState } from '../gridPointer.js'

const gridMetrics = getGridMetrics()
const visibleRegion: VisibleRegionState = { x: 0, y: 0, width: 10, height: 10 }
const geometry = createGridGeometrySnapshot({
  sheetName: 'Sheet1',
  scrollLeft: 0,
  scrollTop: 0,
  hostWidth: 600,
  hostHeight: 400,
  dpr: 1,
  gridMetrics,
})

function baseInput() {
  return {
    clientX: 100,
    clientY: 100,
    buttons: 0,
    isFillHandleDragging: false,
    isRangeMoveDragging: false,
    hasFillPreviewRange: false,
    allowsRangeMove: false,
    selectionRange: null,
    getCellScreenBounds: vi.fn(() => null),
    getVisibleRegion: vi.fn(() => visibleRegion),
    resolvePointerGeometry: vi.fn(() => geometry),
    columnWidths: {},
    rowHeights: {},
    gridMetrics,
    resolveColumnResizeTargetAtPointer: vi.fn(() => null),
    resolveRowResizeTargetAtPointer: vi.fn(() => null),
    resolveHeaderSelectionAtPointer: vi.fn(() => null),
    resolvePointerCell: vi.fn(() => [2, 3] as const),
  }
}

describe('resolveWorkbookGridHoverState', () => {
  it('prioritizes drag states without resolving pointer geometry', () => {
    const input = baseInput()
    const state = resolveWorkbookGridHoverState({
      ...input,
      isRangeMoveDragging: true,
    })

    expect(state).toEqual({ cell: null, header: null, cursor: 'grabbing' })
    expect(input.resolvePointerGeometry).not.toHaveBeenCalled()
  })

  it('returns range-move grab when hovering the selected range border', () => {
    const state = resolveWorkbookGridHoverState({
      ...baseInput(),
      clientX: 98,
      clientY: 41,
      allowsRangeMove: true,
      selectionRange: { x: 1, y: 1, width: 2, height: 2 },
      resolvePointerCell: vi.fn(() => [1, 1] as const),
      getCellScreenBounds: vi.fn((col, row) =>
        col >= 1 && col <= 2 && row >= 1 && row <= 2
          ? {
              x: 90 + (col - 1) * 20,
              y: 40 + (row - 1) * 20,
              width: 20,
              height: 20,
            }
          : undefined,
      ),
    })

    expect(state).toEqual({ cell: null, header: null, cursor: 'grab' })
  })

  it('resolves resize handles before range-move hit testing', () => {
    const getCellScreenBounds = vi.fn(() => ({
      x: 90,
      y: 40,
      width: 20,
      height: 20,
    }))
    const state = resolveWorkbookGridHoverState({
      ...baseInput(),
      allowsRangeMove: true,
      selectionRange: { x: 1, y: 1, width: 2, height: 2 },
      getCellScreenBounds,
      resolveColumnResizeTargetAtPointer: vi.fn(() => 3),
    })

    expect(state).toEqual({ cell: null, header: { kind: 'column', index: 3 }, cursor: 'col-resize' })
    expect(getCellScreenBounds).not.toHaveBeenCalled()
  })

  it('keeps regular cell hover in the leading content lane inside an already-selected range', () => {
    const state = resolveWorkbookGridHoverState({
      ...baseInput(),
      clientX: 98,
      clientY: 50,
      allowsRangeMove: true,
      selectionRange: { x: 1, y: 1, width: 2, height: 2 },
      resolvePointerCell: vi.fn(() => [1, 1] as const),
      getCellScreenBounds: vi.fn((col, row) =>
        col >= 1 && col <= 2 && row >= 1 && row <= 2
          ? {
              x: 90 + (col - 1) * 20,
              y: 40 + (row - 1) * 20,
              width: 20,
              height: 20,
            }
          : undefined,
      ),
    })

    expect(state).toEqual({ cell: [1, 1], header: null, cursor: 'cell' })
  })

  it('keeps regular cell hover in the center of an already-selected range so body drags can select ranges', () => {
    const state = resolveWorkbookGridHoverState({
      ...baseInput(),
      clientX: 120,
      clientY: 50,
      allowsRangeMove: true,
      selectionRange: { x: 1, y: 1, width: 2, height: 2 },
      resolvePointerCell: vi.fn(() => [2, 1] as const),
      getCellScreenBounds: vi.fn((col, row) =>
        col >= 1 && col <= 2 && row >= 1 && row <= 2
          ? {
              x: 90 + (col - 1) * 20,
              y: 40 + (row - 1) * 20,
              width: 20,
              height: 20,
            }
          : undefined,
      ),
    })

    expect(state).toEqual({ cell: [2, 1], header: null, cursor: 'cell' })
  })

  it('falls through to cell hover resolution', () => {
    const state = resolveWorkbookGridHoverState(baseInput())

    expect(state).toEqual({ cell: [2, 3], header: null, cursor: 'cell' })
  })
})
