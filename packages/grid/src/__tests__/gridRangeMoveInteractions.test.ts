import { describe, expect, it, vi } from 'vitest'
import { beginWorkbookGridRangeMove } from '../gridRangeMoveInteractions.js'

describe('gridRangeMoveInteractions', () => {
  it('should preview and apply a moved range', () => {
    // Arrange
    const listenerTarget = createPointerListenerTarget()
    const cleanupRef = { current: null as (() => void) | null }
    const resolvePointerCell = vi.fn(() => [5, 5] as const)
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()
    const onMoveRange = vi.fn()
    const refreshHoverState = vi.fn()
    const setIsRangeMoveDragging = vi.fn()
    const setHoverState = vi.fn()

    // Act
    beginWorkbookGridRangeMove({
      cleanupRef,
      listenerTarget,
      sourceRange: { x: 1, y: 1, width: 2, height: 2 },
      pointerCell: [2, 2],
      resolvePointerCell,
      setGridSelection,
      onSelectionChange,
      onMoveRange,
      refreshHoverState,
      setIsRangeMoveDragging,
      setHoverState,
    })
    listenerTarget.dispatch('pointermove', { clientX: 40, clientY: 50 })
    listenerTarget.dispatch('pointerup', { clientX: 41, clientY: 51 })

    // Assert
    expect(setIsRangeMoveDragging).toHaveBeenNthCalledWith(1, true)
    expect(setIsRangeMoveDragging).toHaveBeenNthCalledWith(2, false)
    expect(setHoverState).toHaveBeenCalledWith({ cell: null, header: null, cursor: 'grabbing' })
    expect(setGridSelection).toHaveBeenNthCalledWith(
      1,
      expect.objectContaining({
        current: expect.objectContaining({
          range: { x: 4, y: 4, width: 2, height: 2 },
        }),
      }),
    )
    expect(setGridSelection).toHaveBeenNthCalledWith(
      2,
      expect.objectContaining({
        current: expect.objectContaining({
          range: { x: 4, y: 4, width: 2, height: 2 },
        }),
      }),
    )
    expect(onSelectionChange).toHaveBeenCalledWith(
      expect.objectContaining({
        current: expect.objectContaining({
          range: { x: 4, y: 4, width: 2, height: 2 },
        }),
      }),
    )
    expect(onMoveRange).toHaveBeenCalledWith('B2', 'C3', 'E5', 'F6')
    expect(refreshHoverState).toHaveBeenCalledWith(41, 51, 0)
    expect(cleanupRef.current).toBeNull()
  })

  it('should treat an unchanged drop as a no-op move', () => {
    // Arrange
    const listenerTarget = createPointerListenerTarget()
    const cleanupRef = { current: null as (() => void) | null }
    const resolvePointerCell = vi.fn(() => [1, 1] as const)
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()
    const onMoveRange = vi.fn()

    // Act
    beginWorkbookGridRangeMove({
      cleanupRef,
      listenerTarget,
      sourceRange: { x: 1, y: 1, width: 2, height: 2 },
      pointerCell: [1, 1],
      resolvePointerCell,
      setGridSelection,
      onSelectionChange,
      onMoveRange,
      refreshHoverState: vi.fn(),
      setIsRangeMoveDragging: vi.fn(),
      setHoverState: vi.fn(),
    })
    listenerTarget.dispatch('pointermove', { clientX: 10, clientY: 10 })
    listenerTarget.dispatch('pointerup', { clientX: 11, clientY: 11 })

    // Assert
    expect(setGridSelection).toHaveBeenCalledTimes(1)
    expect(setGridSelection).toHaveBeenCalledWith(
      expect.objectContaining({
        current: expect.objectContaining({
          range: { x: 1, y: 1, width: 2, height: 2 },
        }),
      }),
    )
    expect(onSelectionChange).toHaveBeenCalledTimes(1)
    expect(onMoveRange).not.toHaveBeenCalled()
  })

  it('should cleanup a previous range move before starting a new one', () => {
    // Arrange
    const previousCleanup = vi.fn()
    const cleanupRef = { current: previousCleanup as (() => void) | null }

    // Act
    beginWorkbookGridRangeMove({
      cleanupRef,
      listenerTarget: createPointerListenerTarget(),
      sourceRange: { x: 0, y: 0, width: 1, height: 1 },
      pointerCell: [0, 0],
      resolvePointerCell: vi.fn(() => [0, 0] as const),
      setGridSelection: vi.fn(),
      onSelectionChange: vi.fn(),
      onMoveRange: vi.fn(),
      refreshHoverState: vi.fn(),
      setIsRangeMoveDragging: vi.fn(),
      setHoverState: vi.fn(),
    })

    // Assert
    expect(previousCleanup).toHaveBeenCalledTimes(1)
    expect(typeof cleanupRef.current).toBe('function')
  })
})

// Helpers

function createPointerListenerTarget(): {
  addEventListener(type: string, listener: (event: { clientX: number; clientY: number }) => void, useCapture?: boolean): void
  removeEventListener(type: string, listener: (event: { clientX: number; clientY: number }) => void, useCapture?: boolean): void
  dispatch(type: string, event: { clientX: number; clientY: number }): void
} {
  const listeners = new Map<string, (event: { clientX: number; clientY: number }) => void>()
  return {
    addEventListener(type, listener) {
      listeners.set(type, listener)
    },
    removeEventListener(type, listener) {
      const current = listeners.get(type)
      if (current === listener) {
        listeners.delete(type)
      }
    },
    dispatch(type, event) {
      listeners.get(type)?.(event)
    },
  }
}
