import { describe, expect, it, vi } from 'vitest'
import { handleGridPointerDown, handleGridPointerMove } from '../gridInteractionController.js'

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

  it('publishes the live rectangular drag selection instead of keeping it only in local grid state', () => {
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()

    handleGridPointerMove({
      event: {
        clientX: 80,
        clientY: 120,
        buttons: 1,
      },
      dragAnchorCell: [1, 22],
      dragHeaderSelection: null,
      dragPointerCell: [1, 22],
      dragViewport: null,
      dragGeometry: null,
      interactionState: createInteractionState(),
      resolvePointerCell: vi.fn(() => [4, 31] as const),
      resolveHeaderSelectionForPointerDrag: vi.fn(),
      selectedCell: [1, 22],
      setGridSelection,
      visibleRegion: {
        firstRow: 0,
        lastRow: 40,
        firstCol: 0,
        lastCol: 20,
        topOffset: 0,
        leftOffset: 0,
      },
      onSelectionChange,
      isEditingCell: false,
      onCommitEdit: vi.fn(),
    })

    expect(setGridSelection).toHaveBeenCalledTimes(1)
    expect(onSelectionChange).toHaveBeenCalledWith(
      expect.objectContaining({
        current: expect.objectContaining({
          cell: [1, 22],
          range: { x: 1, y: 22, width: 4, height: 10 },
        }),
      }),
    )
  })
})
