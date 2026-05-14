import { describe, expect, it, vi } from 'vitest'
import {
  beginWorkbookGridColumnResize,
  beginWorkbookGridRowResize,
  handleWorkbookGridColumnAutofitAtPointer,
  handleWorkbookGridResizePointerDown,
} from '../gridResizeInteractions.js'
import type { GridHoverState } from '../gridHover.js'
import type { VisibleRegionState } from '../gridPointer.js'

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
    let previewedColumnWidth: number | null = null
    const previewColumnWidth = vi.fn((_columnIndex: number, width: number) => {
      previewedColumnWidth = width
    })
    const getPreviewColumnWidth = vi.fn(() => previewedColumnWidth)
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
    expect(commitColumnWidth).toHaveBeenCalledWith(3, 150)
    expect(clearColumnResizePreview).toHaveBeenCalledWith(3)
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
    let previewedRowHeight: number | null = null
    const previewRowHeight = vi.fn((_rowIndex: number, height: number) => {
      previewedRowHeight = height
    })
    const getPreviewRowHeight = vi.fn(() => previewedRowHeight)
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
    listenerTarget.dispatch('pointermove', { clientX: 20, clientY: 100 })
    listenerTarget.dispatch('pointerup', { clientX: 21, clientY: 100 })

    // Assert
    expect(startResize).toHaveBeenCalledTimes(1)
    expect(previewRowHeight).toHaveBeenCalledWith(5, 32)
    expect(clearRowResizePreview).toHaveBeenCalledWith(5)
    expect(commitRowHeight).not.toHaveBeenCalled()
    expect(setActiveResizeRow).toHaveBeenNthCalledWith(1, 5)
    expect(setActiveResizeRow).toHaveBeenNthCalledWith(2, null)
    expect(finishResize).toHaveBeenCalledTimes(1)
    expect(refreshHoverState).toHaveBeenCalledWith(21, 100, 0)
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
      getPreviewColumnWidth: vi.fn(() => null),
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

  it('should clear preview and skip commit when a resize is cancelled', () => {
    const listenerTarget = createPointerListenerTarget()
    const cleanupRef = {
      current: null as ((event?: { clientX: number; clientY: number }) => void) | null,
    }
    const finishResize = vi.fn()
    const setActiveResizeColumn = vi.fn()
    let previewedColumnWidth: number | null = null
    const previewColumnWidth = vi.fn((_columnIndex: number, width: number) => {
      previewedColumnWidth = width
    })
    const getPreviewColumnWidth = vi.fn(() => previewedColumnWidth)
    const clearColumnResizePreview = vi.fn()
    const commitColumnWidth = vi.fn()
    const refreshHoverState = vi.fn()

    beginWorkbookGridColumnResize({
      cleanupRef,
      listenerTarget,
      startResize: vi.fn(),
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
    listenerTarget.dispatch('pointercancel', { clientX: 229, clientY: 41 })

    expect(commitColumnWidth).not.toHaveBeenCalled()
    expect(clearColumnResizePreview).toHaveBeenCalledWith(3)
    expect(setActiveResizeColumn).toHaveBeenNthCalledWith(1, 3)
    expect(setActiveResizeColumn).toHaveBeenNthCalledWith(2, null)
    expect(finishResize).toHaveBeenCalledTimes(1)
    expect(refreshHoverState).toHaveBeenCalledWith(229, 41, 0)
    expect(cleanupRef.current).toBeNull()
  })

  it('should apply column autofit from a resize handle hit', () => {
    const event = createResizePointerEvent({ detail: 2 })
    const resetPointerInteraction = vi.fn()
    const applyAutofitWidth = vi.fn()
    const handled = handleWorkbookGridColumnAutofitAtPointer({
      event,
      visibleRegion: createVisibleRegion(),
      pointerGeometry: null,
      columnWidths: { 4: 120 },
      defaultColumnWidth: 96,
      isEditingCell: true,
      commitActiveEdit: vi.fn(),
      computeAutofitColumnWidth: vi.fn(() => 148),
      applyAutofitWidth,
      finishResize: vi.fn(),
      resetPointerInteraction,
      setActiveResizeColumn: vi.fn(),
      resolveColumnResizeTargetAtPointer: vi.fn(() => 4),
    })

    expect(handled).toBe(true)
    expect(event.preventDefault).toHaveBeenCalledTimes(1)
    expect(event.stopPropagation).toHaveBeenCalledTimes(1)
    expect(resetPointerInteraction).toHaveBeenCalledTimes(1)
    expect(applyAutofitWidth).toHaveBeenCalledWith(4, 148)
  })

  it('should start column resize on first handle activation and autofit on repeated activation', () => {
    const lastResizeHandleActivationRef = { current: null as { columnIndex: number; at: number } | null }
    const beginColumnResize = vi.fn()
    const applyAutofitWidth = vi.fn()
    const hoverState = createHoverStateSetter()
    const baseInput = {
      visibleRegion: createVisibleRegion(),
      pointerGeometry: null,
      columnWidths: { 2: 100 },
      rowHeights: {},
      defaultColumnWidth: 96,
      defaultRowHeight: 24,
      isEditingCell: false,
      commitActiveEdit: vi.fn(),
      focusGrid: vi.fn(),
      setActiveHeaderDrag: vi.fn(),
      setHoverState: hoverState.set,
      lastResizeHandleActivationRef,
      computeAutofitColumnWidth: vi.fn(() => 132),
      applyAutofitWidth,
      finishResize: vi.fn(),
      resetPointerInteraction: vi.fn(),
      setActiveResizeColumn: vi.fn(),
      beginColumnResize,
      beginRowResize: vi.fn(),
      resolveColumnResizeTargetAtPointer: vi.fn(() => 2),
      resolveRowResizeTargetAtPointer: vi.fn(() => null),
    }

    expect(
      handleWorkbookGridResizePointerDown({
        ...baseInput,
        event: createResizePointerEvent({ detail: 1 }),
        now: () => 1000,
      }),
    ).toBe(true)
    expect(beginColumnResize).toHaveBeenCalledWith(2, 200)
    expect(hoverState.current).toEqual({ cell: null, header: { kind: 'column', index: 2 }, cursor: 'col-resize' })
    expect(applyAutofitWidth).not.toHaveBeenCalled()

    expect(
      handleWorkbookGridResizePointerDown({
        ...baseInput,
        event: createResizePointerEvent({ detail: 1 }),
        now: () => 1200,
      }),
    ).toBe(true)
    expect(applyAutofitWidth).toHaveBeenCalledWith(2, 132)
    expect(lastResizeHandleActivationRef.current).toBeNull()
  })

  it('should start row resize when the pointer hits a row handle', () => {
    const beginRowResize = vi.fn()
    const hoverState = createHoverStateSetter()

    const handled = handleWorkbookGridResizePointerDown({
      event: createResizePointerEvent({ clientY: 90 }),
      visibleRegion: createVisibleRegion(),
      pointerGeometry: null,
      columnWidths: {},
      rowHeights: { 7: 30 },
      defaultColumnWidth: 96,
      defaultRowHeight: 24,
      isEditingCell: true,
      commitActiveEdit: vi.fn(),
      focusGrid: vi.fn(),
      setActiveHeaderDrag: vi.fn(),
      setHoverState: hoverState.set,
      lastResizeHandleActivationRef: { current: null },
      now: () => 1000,
      computeAutofitColumnWidth: vi.fn(() => 96),
      applyAutofitWidth: vi.fn(),
      finishResize: vi.fn(),
      resetPointerInteraction: vi.fn(),
      setActiveResizeColumn: vi.fn(),
      beginColumnResize: vi.fn(),
      beginRowResize,
      resolveColumnResizeTargetAtPointer: vi.fn(() => null),
      resolveRowResizeTargetAtPointer: vi.fn(() => 7),
    })

    expect(handled).toBe(true)
    expect(beginRowResize).toHaveBeenCalledWith(7, 90)
    expect(hoverState.current).toEqual({ cell: null, header: { kind: 'row', index: 7 }, cursor: 'row-resize' })
  })
})

// Helpers

function createVisibleRegion(): VisibleRegionState {
  return {
    range: { x: 0, y: 0, width: 10, height: 20 },
    tx: 0,
    ty: 0,
  }
}

function createResizePointerEvent(options: { clientX?: number; clientY?: number; detail?: number } = {}) {
  return {
    clientX: options.clientX ?? 200,
    clientY: options.clientY ?? 40,
    detail: options.detail ?? 1,
    preventDefault: vi.fn(),
    stopPropagation: vi.fn(),
  }
}

function createHoverStateSetter(): {
  current: GridHoverState
  readonly set: (updater: (current: GridHoverState) => GridHoverState) => void
} {
  const state = {
    current: { cell: null, header: null, cursor: 'default' } satisfies GridHoverState,
    set: (updater: (current: GridHoverState) => GridHoverState): void => {
      state.current = updater(state.current)
    },
  }
  return state
}

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
