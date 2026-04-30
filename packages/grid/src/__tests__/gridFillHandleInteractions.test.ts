import { describe, expect, it, vi } from 'vitest'
import { createRectangleSelectionFromRange } from '../gridSelection.js'
import { beginWorkbookGridFillHandleDrag } from '../gridFillHandleInteractions.js'
import type { Rectangle } from '../gridTypes.js'

interface PointerEventLike {
  readonly clientX: number
  readonly clientY: number
  readonly pointerId: number
}

class PointerTarget {
  private readonly listeners = new Map<string, Set<(event: PointerEventLike) => void>>()

  addEventListener(type: 'pointermove' | 'pointerup' | 'pointercancel', listener: (event: PointerEventLike) => void): void {
    const listeners = this.listeners.get(type) ?? new Set()
    listeners.add(listener)
    this.listeners.set(type, listeners)
  }

  removeEventListener(type: 'pointermove' | 'pointerup' | 'pointercancel', listener: (event: PointerEventLike) => void): void {
    this.listeners.get(type)?.delete(listener)
  }

  emit(type: 'pointermove' | 'pointerup' | 'pointercancel', event: PointerEventLike): void {
    for (const listener of this.listeners.get(type) ?? []) {
      listener(event)
    }
  }

  listenerCount(type: 'pointermove' | 'pointerup' | 'pointercancel'): number {
    return this.listeners.get(type)?.size ?? 0
  }
}

describe('beginWorkbookGridFillHandleDrag', () => {
  it('previews, commits fill range, and preserves active cell anchor', () => {
    const target = new PointerTarget()
    const cleanupRef = { current: null as (() => void) | null }
    const sourceRange: Rectangle = { x: 1, y: 1, width: 2, height: 2 }
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()
    const onFillRange = vi.fn()
    const setFillPreviewRange = vi.fn()
    const setFillPreviewRangeRef = vi.fn()
    const setIsFillHandleDragging = vi.fn()
    const resetHoverState = vi.fn()

    const gridSelection = createRectangleSelectionFromRange(sourceRange)
    if (gridSelection.current) {
      gridSelection.current = {
        ...gridSelection.current,
        cell: [2, 2],
      }
    }

    beginWorkbookGridFillHandleDrag({
      cleanupRef,
      listenerTarget: target,
      pointerId: 7,
      sourceRange,
      gridSelection,
      resolvePointerCell: (clientX, clientY) => [clientX, clientY],
      setGridSelection,
      onSelectionChange,
      onFillRange,
      setFillPreviewRange,
      setFillPreviewRangeRef,
      setIsFillHandleDragging,
      resetHoverState,
    })

    target.emit('pointermove', { pointerId: 7, clientX: 4, clientY: 2 })
    expect(setFillPreviewRange).toHaveBeenLastCalledWith({ x: 3, y: 1, width: 2, height: 2 })
    expect(setFillPreviewRangeRef).toHaveBeenLastCalledWith({ x: 3, y: 1, width: 2, height: 2 })

    target.emit('pointerup', { pointerId: 7, clientX: 4, clientY: 2 })

    const expectedSelection = createRectangleSelectionFromRange({ x: 1, y: 1, width: 4, height: 2 })
    if (expectedSelection.current) {
      expectedSelection.current = {
        ...expectedSelection.current,
        cell: [2, 2],
      }
    }
    expect(setGridSelection).toHaveBeenCalledWith(expectedSelection)
    expect(onSelectionChange).toHaveBeenCalledTimes(1)
    expect(onFillRange).toHaveBeenCalledWith('B2', 'C3', 'D2', 'E3')
    expect(setFillPreviewRange).toHaveBeenLastCalledWith(null)
    expect(setIsFillHandleDragging).toHaveBeenNthCalledWith(1, true)
    expect(setIsFillHandleDragging).toHaveBeenLastCalledWith(false)
    expect(target.listenerCount('pointermove')).toBe(0)
    expect(target.listenerCount('pointerup')).toBe(0)
    expect(target.listenerCount('pointercancel')).toBe(0)
  })

  it('ignores other pointer IDs and clears preview on cancellation', () => {
    const target = new PointerTarget()
    const cleanupRef = { current: null as (() => void) | null }
    const setFillPreviewRange = vi.fn()
    const setFillPreviewRangeRef = vi.fn()
    const setIsFillHandleDragging = vi.fn()

    beginWorkbookGridFillHandleDrag({
      cleanupRef,
      listenerTarget: target,
      pointerId: 7,
      sourceRange: { x: 1, y: 1, width: 1, height: 1 },
      gridSelection: createRectangleSelectionFromRange({ x: 1, y: 1, width: 1, height: 1 }),
      resolvePointerCell: (clientX, clientY) => [clientX, clientY],
      setGridSelection: vi.fn(),
      onSelectionChange: vi.fn(),
      onFillRange: vi.fn(),
      setFillPreviewRange,
      setFillPreviewRangeRef,
      setIsFillHandleDragging,
      resetHoverState: vi.fn(),
    })

    target.emit('pointermove', { pointerId: 8, clientX: 4, clientY: 1 })
    expect(setFillPreviewRange).toHaveBeenLastCalledWith(null)

    target.emit('pointercancel', { pointerId: 7, clientX: 4, clientY: 1 })
    expect(setFillPreviewRange).toHaveBeenLastCalledWith(null)
    expect(setFillPreviewRangeRef).toHaveBeenLastCalledWith(null)
    expect(setIsFillHandleDragging).toHaveBeenLastCalledWith(false)
    expect(cleanupRef.current).toBeNull()
  })
})
