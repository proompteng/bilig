import { describe, expect, it, vi } from 'vitest'
import { beginWorkbookGridColumnResize, beginWorkbookGridRowResize } from '../gridResizeInteractions.js'

describe('gridResizeInteractions', () => {
  it('should preview and commit a column resize', () => {
    // Arrange
    const listenerTarget = createPointerListenerTarget()
    const cleanupRef = {
      current: null as ((event?: { clientX: number; clientY: number }) => void) | null,
    }
    const startResize = vi.fn()
    const finishResize = vi.fn()
    const setActiveResizeColumn = vi.fn()
    const previewColumnWidth = vi.fn()
    const getPreviewColumnWidth = vi.fn(() => 140)
    const clearColumnResizePreview = vi.fn()
    const commitColumnWidth = vi.fn()
    const refreshHoverState = vi.fn()

    // Act
    beginWorkbookGridColumnResize({
      cleanupRef,
      listenerTarget,
      startResize,
      finishResize,
      refreshHoverState,
      setActiveResizeColumn,
      previewColumnWidth,
      getPreviewColumnWidth,
      clearColumnResizePreview,
      commitColumnWidth,
      columnIndex: 3,
      startClientX: 200,
      columnWidths: { 3: 120 },
      defaultColumnWidth: 96,
    })
    listenerTarget.dispatch('pointermove', { clientX: 230, clientY: 40 })
    listenerTarget.dispatch('pointerup', { clientX: 231, clientY: 41 })

    // Assert
    expect(startResize).toHaveBeenCalledTimes(1)
    expect(previewColumnWidth).toHaveBeenCalledWith(3, 150)
    expect(commitColumnWidth).toHaveBeenCalledWith(3, 140)
    expect(clearColumnResizePreview).not.toHaveBeenCalled()
    expect(setActiveResizeColumn).toHaveBeenNthCalledWith(1, 3)
    expect(setActiveResizeColumn).toHaveBeenNthCalledWith(2, null)
    expect(finishResize).toHaveBeenCalledTimes(1)
    expect(refreshHoverState).toHaveBeenCalledWith(231, 41, 0)
    expect(cleanupRef.current).toBeNull()
  })

  it('should clear row preview when the resized height does not change', () => {
    // Arrange
    const listenerTarget = createPointerListenerTarget()
    const cleanupRef = {
      current: null as ((event?: { clientX: number; clientY: number }) => void) | null,
    }
    const startResize = vi.fn()
    const finishResize = vi.fn()
    const setActiveResizeRow = vi.fn()
    const previewRowHeight = vi.fn()
    const getPreviewRowHeight = vi.fn(() => 32)
    const clearRowResizePreview = vi.fn()
    const commitRowHeight = vi.fn()
    const refreshHoverState = vi.fn()

    // Act
    beginWorkbookGridRowResize({
      cleanupRef,
      listenerTarget,
      startResize,
      finishResize,
      refreshHoverState,
      setActiveResizeRow,
      previewRowHeight,
      getPreviewRowHeight,
      clearRowResizePreview,
      commitRowHeight,
      rowIndex: 5,
      startClientY: 100,
      rowHeights: { 5: 32 },
      defaultRowHeight: 24,
    })
    listenerTarget.dispatch('pointermove', { clientX: 20, clientY: 110 })
    listenerTarget.dispatch('pointerup', { clientX: 21, clientY: 111 })

    // Assert
    expect(startResize).toHaveBeenCalledTimes(1)
    expect(previewRowHeight).toHaveBeenCalledWith(5, 42)
    expect(clearRowResizePreview).toHaveBeenCalledWith(5)
    expect(commitRowHeight).not.toHaveBeenCalled()
    expect(setActiveResizeRow).toHaveBeenNthCalledWith(1, 5)
    expect(setActiveResizeRow).toHaveBeenNthCalledWith(2, null)
    expect(finishResize).toHaveBeenCalledTimes(1)
    expect(refreshHoverState).toHaveBeenCalledWith(21, 111, 0)
    expect(cleanupRef.current).toBeNull()
  })

  it('should cleanup a previous resize before starting the next one', () => {
    // Arrange
    const previousCleanup = vi.fn()
    const listenerTarget = createPointerListenerTarget()
    const cleanupRef = {
      current: previousCleanup as ((event?: { clientX: number; clientY: number }) => void) | null,
    }

    // Act
    beginWorkbookGridColumnResize({
      cleanupRef,
      listenerTarget,
      startResize: vi.fn(),
      finishResize: vi.fn(),
      refreshHoverState: vi.fn(),
      setActiveResizeColumn: vi.fn(),
      previewColumnWidth: vi.fn(),
      getPreviewColumnWidth: vi.fn(() => 96),
      clearColumnResizePreview: vi.fn(),
      commitColumnWidth: vi.fn(),
      columnIndex: 1,
      startClientX: 10,
      columnWidths: {},
      defaultColumnWidth: 96,
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
