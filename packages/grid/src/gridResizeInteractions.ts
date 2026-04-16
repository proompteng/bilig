interface PointerEventLike {
  clientX: number
  clientY: number
}

interface PointerListenerTarget {
  addEventListener(type: 'pointermove' | 'pointerup', listener: (event: PointerEventLike) => void, useCapture: boolean): void
  removeEventListener(type: 'pointermove' | 'pointerup', listener: (event: PointerEventLike) => void, useCapture: boolean): void
}

type ResizeCleanup = ((event?: PointerEventLike) => void) | null

function installGridResizeLifecycle(input: {
  cleanupRef: { current: ResizeCleanup }
  listenerTarget: PointerListenerTarget
  startResize: () => void
  finishResize: () => void
  refreshHoverState: (clientX: number, clientY: number, buttons: number) => void
  activate: () => void
  deactivate: () => void
  preview: (event: PointerEventLike) => void
  commitOrClear: () => void
}): void {
  const { activate, cleanupRef, commitOrClear, deactivate, finishResize, listenerTarget, preview, refreshHoverState, startResize } = input
  cleanupRef.current?.()
  startResize()
  activate()

  const handlePointerMove = (nativeEvent: PointerEventLike) => {
    preview(nativeEvent)
  }

  const cleanup = (nativeEvent?: PointerEventLike) => {
    listenerTarget.removeEventListener('pointermove', handlePointerMove, true)
    listenerTarget.removeEventListener('pointerup', handlePointerUp, true)
    cleanupRef.current = null
    deactivate()
    finishResize()
    if (nativeEvent) {
      refreshHoverState(nativeEvent.clientX, nativeEvent.clientY, 0)
    }
  }

  const handlePointerUp = (nativeEvent: PointerEventLike) => {
    commitOrClear()
    cleanup(nativeEvent)
  }

  cleanupRef.current = cleanup
  listenerTarget.addEventListener('pointermove', handlePointerMove, true)
  listenerTarget.addEventListener('pointerup', handlePointerUp, true)
}

export function beginWorkbookGridColumnResize(input: {
  cleanupRef: { current: ResizeCleanup }
  listenerTarget: PointerListenerTarget
  startResize: () => void
  finishResize: () => void
  refreshHoverState: (clientX: number, clientY: number, buttons: number) => void
  setActiveResizeColumn: (columnIndex: number | null) => void
  previewColumnWidth: (columnIndex: number, width: number) => void
  getPreviewColumnWidth: (columnIndex: number) => number | null | undefined
  clearColumnResizePreview: (columnIndex: number) => void
  commitColumnWidth: (columnIndex: number, width: number) => void
  columnIndex: number
  startClientX: number
  columnWidths: Readonly<Record<number, number>>
  defaultColumnWidth: number
}): void {
  const {
    cleanupRef,
    clearColumnResizePreview,
    columnIndex,
    columnWidths,
    commitColumnWidth,
    defaultColumnWidth,
    finishResize,
    getPreviewColumnWidth,
    listenerTarget,
    previewColumnWidth,
    refreshHoverState,
    setActiveResizeColumn,
    startClientX,
    startResize,
  } = input
  const startWidth = columnWidths[columnIndex] ?? defaultColumnWidth

  installGridResizeLifecycle({
    cleanupRef,
    listenerTarget,
    startResize,
    finishResize,
    refreshHoverState,
    activate: () => setActiveResizeColumn(columnIndex),
    deactivate: () => setActiveResizeColumn(null),
    preview: (nativeEvent) => {
      previewColumnWidth(columnIndex, startWidth + (nativeEvent.clientX - startClientX))
    },
    commitOrClear: () => {
      const finalWidth = getPreviewColumnWidth(columnIndex) ?? startWidth
      if (finalWidth === startWidth) {
        clearColumnResizePreview(columnIndex)
      } else {
        commitColumnWidth(columnIndex, finalWidth)
      }
    },
  })
}

export function beginWorkbookGridRowResize(input: {
  cleanupRef: { current: ResizeCleanup }
  listenerTarget: PointerListenerTarget
  startResize: () => void
  finishResize: () => void
  refreshHoverState: (clientX: number, clientY: number, buttons: number) => void
  setActiveResizeRow: (rowIndex: number | null) => void
  previewRowHeight: (rowIndex: number, height: number) => void
  getPreviewRowHeight: (rowIndex: number) => number | null | undefined
  clearRowResizePreview: (rowIndex: number) => void
  commitRowHeight: (rowIndex: number, height: number) => void
  rowIndex: number
  startClientY: number
  rowHeights: Readonly<Record<number, number>>
  defaultRowHeight: number
}): void {
  const {
    cleanupRef,
    clearRowResizePreview,
    commitRowHeight,
    defaultRowHeight,
    finishResize,
    getPreviewRowHeight,
    listenerTarget,
    previewRowHeight,
    refreshHoverState,
    rowHeights,
    rowIndex,
    setActiveResizeRow,
    startClientY,
    startResize,
  } = input
  const startHeight = rowHeights[rowIndex] ?? defaultRowHeight

  installGridResizeLifecycle({
    cleanupRef,
    listenerTarget,
    startResize,
    finishResize,
    refreshHoverState,
    activate: () => setActiveResizeRow(rowIndex),
    deactivate: () => setActiveResizeRow(null),
    preview: (nativeEvent) => {
      previewRowHeight(rowIndex, startHeight + (nativeEvent.clientY - startClientY))
    },
    commitOrClear: () => {
      const finalHeight = getPreviewRowHeight(rowIndex) ?? startHeight
      if (finalHeight === startHeight) {
        clearRowResizePreview(rowIndex)
      } else {
        commitRowHeight(rowIndex, finalHeight)
      }
    },
  })
}
