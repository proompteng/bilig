import { describe, expect, it, vi } from 'vitest'
import { handleGridPointerDown } from '../gridInteractionController.js'

function createInteractionState() {
  return {
    ignoreNextPointerSelectionRef: { current: false },
    pendingPointerCellRef: { current: null },
    dragAnchorCellRef: { current: null },
    dragPointerCellRef: { current: null },
    dragHeaderSelectionRef: { current: null },
    dragViewportRef: { current: null },
    dragGeometryRef: { current: null },
    dragDidMoveRef: { current: false },
    postDragSelectionExpiryRef: { current: 0 },
    columnResizeActiveRef: { current: false },
  }
}

describe('gridInteractionController', () => {
  it('commits the active edit before applying a body-click selection change', () => {
    const order: string[] = []
    const onCommitEdit = vi.fn(() => {
      order.push('commit')
    })
    const onSelectionChange = vi.fn(() => {
      order.push('selection')
    })

    handleGridPointerDown({
      event: {
        button: 0,
        clientX: 20,
        clientY: 30,
        shiftKey: false,
        preventDefault: vi.fn(),
        stopPropagation: vi.fn(),
      },
      columnWidths: {},
      defaultColumnWidth: 120,
      focusGrid: vi.fn(),
      interactionState: createInteractionState(),
      isEditingCell: true,
      onCommitEdit,
      onSelectionChange,
      resolvePointerGeometry: vi.fn(() => null),
      resolveColumnResizeTargetAtPointer: vi.fn(() => null),
      resolveHeaderSelectionAtPointer: vi.fn(() => null),
      resolvePointerCell: vi.fn(() => [3, 4] as const),
      selectedCell: [1, 1],
      setGridSelection: vi.fn(),
      visibleRegion: {
        firstRow: 0,
        lastRow: 20,
        firstCol: 0,
        lastCol: 10,
        topOffset: 0,
        leftOffset: 0,
      },
    })

    expect(onCommitEdit).toHaveBeenCalledTimes(1)
    expect(onSelectionChange).toHaveBeenCalledTimes(1)
    expect(order).toEqual(['commit', 'selection'])
  })
})
