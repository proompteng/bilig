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

  it('auto-scrolls the grid edge while range move drag stays active', () => {
    // Arrange
    const listenerTarget = createPointerListenerTarget()
    const cleanupRef = { current: null as (() => void) | null }
    const frameCallbacks: FrameRequestCallback[] = []
    const setGridSelection = vi.fn()
    const scrollViewport = createScrollViewport()
    const resolvePointerCell = vi.fn((_clientX: number, _clientY: number) => [0, Math.floor(scrollViewport.scrollTop / 10)] as const)

    // Act
    beginWorkbookGridRangeMove({
      cleanupRef,
      listenerTarget,
      sourceRange: { x: 0, y: 0, width: 1, height: 1 },
      pointerCell: [0, 0],
      resolvePointerCell,
      setGridSelection,
      onSelectionChange: vi.fn(),
      onMoveRange: vi.fn(),
      refreshHoverState: vi.fn(),
      setIsRangeMoveDragging: vi.fn(),
      setHoverState: vi.fn(),
      scrollViewport,
      requestAnimationFrame: (callback) => {
        frameCallbacks.push(callback)
        return frameCallbacks.length
      },
      cancelAnimationFrame: vi.fn(),
    })
    listenerTarget.dispatch('pointermove', { clientX: 50, clientY: 98 })
    frameCallbacks.shift()?.(performance.now())

    // Assert
    expect(scrollViewport.scrollTop).toBeGreaterThan(0)
    expect(scrollViewport.dispatchEvent).toHaveBeenCalledWith(expect.any(Event))
    expect(setGridSelection).toHaveBeenLastCalledWith(
      expect.objectContaining({
        current: expect.objectContaining({
          range: { x: 0, y: expect.any(Number), width: 1, height: 1 },
        }),
      }),
    )
    expect(setGridSelection.mock.lastCall?.[0].current?.range.y).toBeGreaterThan(0)
  })

  it('applies the auto-scrolled preview range on pointer up', () => {
    // Arrange
    const listenerTarget = createPointerListenerTarget()
    const cleanupRef = { current: null as (() => void) | null }
    const frameCallbacks: FrameRequestCallback[] = []
    const scrollViewport = createScrollViewport()
    const onMoveRange = vi.fn()
    const resolvePointerCell = vi.fn((_clientX: number, _clientY: number) => [1, 1 + Math.floor(scrollViewport.scrollTop / 10)] as const)

    // Act
    beginWorkbookGridRangeMove({
      cleanupRef,
      listenerTarget,
      sourceRange: { x: 1, y: 1, width: 2, height: 1 },
      pointerCell: [1, 1],
      resolvePointerCell,
      setGridSelection: vi.fn(),
      onSelectionChange: vi.fn(),
      onMoveRange,
      refreshHoverState: vi.fn(),
      setIsRangeMoveDragging: vi.fn(),
      setHoverState: vi.fn(),
      scrollViewport,
      requestAnimationFrame: (callback) => {
        frameCallbacks.push(callback)
        return frameCallbacks.length
      },
      cancelAnimationFrame: vi.fn(),
    })
    listenerTarget.dispatch('pointermove', { clientX: 50, clientY: 98 })
    frameCallbacks.shift()?.(performance.now())
    listenerTarget.dispatch('pointerup', { clientX: 50, clientY: 98 })

    // Assert
    expect(onMoveRange).toHaveBeenCalledWith('B2', 'C2', 'B5', 'C5')
  })

  it('recomputes the drop range on pointer up after the viewport scrolls', () => {
    // Arrange
    const listenerTarget = createPointerListenerTarget()
    const cleanupRef = { current: null as (() => void) | null }
    const scrollViewport = createScrollViewport()
    const onMoveRange = vi.fn()
    const resolvePointerCell = vi.fn((_clientX: number, _clientY: number) => [1, 1 + Math.floor(scrollViewport.scrollTop / 10)] as const)

    // Act
    beginWorkbookGridRangeMove({
      cleanupRef,
      listenerTarget,
      sourceRange: { x: 1, y: 1, width: 2, height: 1 },
      pointerCell: [1, 1],
      resolvePointerCell,
      setGridSelection: vi.fn(),
      onSelectionChange: vi.fn(),
      onMoveRange,
      refreshHoverState: vi.fn(),
      setIsRangeMoveDragging: vi.fn(),
      setHoverState: vi.fn(),
      scrollViewport,
      requestAnimationFrame: (callback) => {
        callback(performance.now())
        return 1
      },
      cancelAnimationFrame: vi.fn(),
    })
    listenerTarget.dispatch('pointermove', { clientX: 50, clientY: 50 })
    scrollViewport.scrollTop = 80
    listenerTarget.dispatch('pointerup', { clientX: 50, clientY: 50 })

    // Assert
    expect(onMoveRange).toHaveBeenCalledWith('B2', 'C2', 'B10', 'C10')
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

function createScrollViewport(): {
  clientHeight: number
  clientWidth: number
  dispatchEvent: ReturnType<typeof vi.fn>
  getBoundingClientRect(): Pick<DOMRect, 'bottom' | 'left' | 'right' | 'top'>
  scrollHeight: number
  scrollLeft: number
  scrollTop: number
  scrollWidth: number
} {
  return {
    clientHeight: 100,
    clientWidth: 100,
    dispatchEvent: vi.fn(),
    getBoundingClientRect: () => ({
      bottom: 100,
      left: 0,
      right: 100,
      top: 0,
    }),
    scrollHeight: 1_000,
    scrollLeft: 0,
    scrollTop: 0,
    scrollWidth: 1_000,
  }
}
