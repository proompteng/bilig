import { useCallback, type MouseEvent as ReactMouseEvent, type PointerEvent as ReactPointerEvent } from 'react'
import { resolveSelectionMoveAnchorCellFromPointerCell } from './gridRangeMove.js'
import { sameGridHoverState, type GridHoverState } from './gridHover.js'
import {
  finishGridResize,
  handleGridBodyDoubleClick,
  handleGridPointerDown,
  handleGridPointerMove,
  handleGridPointerUp,
  startGridResize,
} from './gridInteractionController.js'
import { resetGridPointerInteraction } from './gridInteractionState.js'
import { resolveWorkbookGridHoverState } from './gridInteractionHoverState.js'
import {
  applyWorkbookGridColumnAutofit,
  beginWorkbookGridColumnResize,
  beginWorkbookGridRowResize,
  handleWorkbookGridColumnAutofitAtPointer,
  handleWorkbookGridResizePointerDown,
} from './gridResizeInteractions.js'
import { beginWorkbookGridRangeMove } from './gridRangeMoveInteractions.js'
import { beginWorkbookGridFillHandleDrag } from './gridFillHandleInteractions.js'
import type { GridInputController } from './runtime/gridInputController.js'
import type { GridSelection, Item, Rectangle } from './gridTypes.js'
import type { WorkbookGridSurfaceProps } from './workbookGridSurfaceTypes.js'
import type { useWorkbookGridPointerResolvers } from './useWorkbookGridPointerResolvers.js'
import type { useWorkbookGridRenderState } from './useWorkbookGridRenderState.js'

const DEFAULT_GRID_HOVER_STATE: GridHoverState = { cell: null, header: null, cursor: 'default' }

function resetGridHoverState(current: GridHoverState): GridHoverState {
  return sameGridHoverState(current, DEFAULT_GRID_HOVER_STATE) ? current : DEFAULT_GRID_HOVER_STATE
}

function toGridResizeCursorHoverState(next: GridHoverState): GridHoverState {
  return { cell: null, header: next.header, cursor: next.cursor }
}

function isFillHandleTarget(target: EventTarget | null): boolean {
  return target instanceof Element && target.closest("[data-grid-fill-handle='true']") !== null
}

export function useWorkbookGridHostPointerHandlers(input: {
  readonly activeSelectionCell: Item
  readonly allowsRangeMove: boolean
  readonly applyAutofitWidth: (columnIndex: number, width: number) => void
  readonly beginEditAt: (addr: string, seed?: string) => void
  readonly commitActiveEdit: () => void
  readonly emitSelectionChange: (nextSelection: GridSelection) => void
  readonly inputController: GridInputController
  readonly isEditingCell: boolean
  readonly onAutofitColumn: WorkbookGridSurfaceProps['onAutofitColumn']
  readonly onFillRange: WorkbookGridSurfaceProps['onFillRange']
  readonly onMoveRange: WorkbookGridSurfaceProps['onMoveRange']
  readonly pointerResolvers: ReturnType<typeof useWorkbookGridPointerResolvers>
  readonly gridSelection: GridSelection
  readonly selectionRange: Rectangle | null
  readonly renderState: ReturnType<typeof useWorkbookGridRenderState>
}) {
  const {
    activeSelectionCell,
    allowsRangeMove,
    applyAutofitWidth,
    beginEditAt,
    commitActiveEdit,
    emitSelectionChange,
    inputController,
    isEditingCell,
    onAutofitColumn,
    onFillRange,
    onMoveRange,
    pointerResolvers,
    gridSelection,
    selectionRange,
    renderState,
  } = input
  const {
    activeResizeColumn,
    activeResizeRow,
    clearColumnResizePreview,
    clearRowResizePreview,
    columnWidths,
    commitColumnWidth,
    commitRowHeight,
    computeAutofitColumnWidth,
    focusGrid,
    getCellScreenBounds,
    getPreviewColumnWidth,
    getPreviewRowHeight,
    getVisibleRegion,
    gridMetrics,
    hostElement,
    isFillHandleDragging,
    isRangeMoveDragging,
    previewColumnWidth,
    previewRowHeight,
    rowHeights,
    setActiveHeaderDrag,
    setActiveResizeColumn,
    setActiveResizeRow,
    setFillPreviewRange,
    setGridSelection,
    setHoverState,
    setIsFillHandleDragging,
    setIsRangeMoveDragging,
  } = renderState
  const {
    resolveColumnResizeTarget: resolveColumnResizeTargetAtPointer,
    resolveRowResizeTarget: resolveRowResizeTargetAtPointer,
    resolveHeaderSelectionAtPointer,
    resolveHeaderSelectionForPointerDrag,
    resolvePointerCell,
    resolvePointerGeometry,
  } = pointerResolvers
  const {
    dragAnchorCellRef,
    dragDidMoveRef,
    dragHeaderSelectionRef,
    dragPointerCellRef,
    fillHandleCleanupRef,
    fillPreviewRangeRef,
    interactionState,
    lastBodyClickCellRef,
    lastResizeHandleActivationRef,
    postDragSelectionExpiryRef,
    rangeMoveCleanupRef,
    resizeCleanupRef,
  } = inputController

  const handleFillHandlePointerDown = useCallback(
    (event: ReactPointerEvent<HTMLDivElement>) => {
      if (!selectionRange || event.button !== 0) {
        return
      }
      event.preventDefault()
      event.stopPropagation()
      focusGrid()
      beginWorkbookGridFillHandleDrag({
        cleanupRef: fillHandleCleanupRef,
        listenerTarget: window,
        pointerId: event.pointerId,
        sourceRange: selectionRange,
        gridSelection,
        resolvePointerCell,
        setGridSelection,
        onSelectionChange: emitSelectionChange,
        onFillRange,
        setFillPreviewRange,
        setFillPreviewRangeRef: (range) => {
          fillPreviewRangeRef.current = range
        },
        setIsFillHandleDragging,
        scrollViewport: renderState.scrollViewportRef.current,
        resetHoverState: () => {
          setHoverState(resetGridHoverState)
        },
      })
    },
    [
      emitSelectionChange,
      fillHandleCleanupRef,
      fillPreviewRangeRef,
      focusGrid,
      gridSelection,
      onFillRange,
      renderState.scrollViewportRef,
      resolvePointerCell,
      selectionRange,
      setFillPreviewRange,
      setGridSelection,
      setHoverState,
      setIsFillHandleDragging,
    ],
  )

  const refreshHoverState = useCallback(
    (clientX: number, clientY: number, buttons: number) => {
      const next = resolveWorkbookGridHoverState({
        clientX,
        clientY,
        buttons,
        isFillHandleDragging,
        isRangeMoveDragging,
        hasFillPreviewRange: fillPreviewRangeRef.current !== null,
        allowsRangeMove,
        selectionRange,
        getCellScreenBounds,
        getVisibleRegion,
        resolvePointerGeometry,
        columnWidths,
        gridMetrics,
        resolveColumnResizeTargetAtPointer,
        resolveHeaderSelectionAtPointer,
        resolvePointerCell,
        resolveRowResizeTargetAtPointer,
        rowHeights,
      })
      const isResizeCursorOnlyHover =
        next.cell === null && (next.header?.kind === 'column' || next.header?.kind === 'row') && next.cursor !== 'default'
      if (isResizeCursorOnlyHover) {
        const cursorOnlyHoverState = toGridResizeCursorHoverState(next)
        if (hostElement) {
          hostElement.style.cursor = cursorOnlyHoverState.cursor
        }
        setHoverState((current) => (sameGridHoverState(current, cursorOnlyHoverState) ? current : cursorOnlyHoverState))
        return
      }
      if (hostElement) {
        hostElement.style.cursor = next.cursor
      }
      setHoverState((current) => (sameGridHoverState(current, next) ? current : next))
    },
    [
      allowsRangeMove,
      columnWidths,
      fillPreviewRangeRef,
      getCellScreenBounds,
      getVisibleRegion,
      gridMetrics,
      hostElement,
      isFillHandleDragging,
      isRangeMoveDragging,
      resolveColumnResizeTargetAtPointer,
      resolveHeaderSelectionAtPointer,
      resolvePointerCell,
      resolvePointerGeometry,
      resolveRowResizeTargetAtPointer,
      rowHeights,
      selectionRange,
      setHoverState,
    ],
  )

  const beginRangeMove = useCallback(
    (pointerCell: Item) => {
      if (!selectionRange) {
        return
      }
      if (isEditingCell) {
        commitActiveEdit()
      }
      focusGrid()
      beginWorkbookGridRangeMove({
        cleanupRef: rangeMoveCleanupRef,
        listenerTarget: window,
        sourceRange: selectionRange,
        pointerCell,
        activeCell: activeSelectionCell,
        resolvePointerCell,
        setGridSelection,
        onSelectionChange: emitSelectionChange,
        onMoveRange,
        refreshHoverState,
        scrollViewport: renderState.scrollViewportRef.current,
        setIsRangeMoveDragging,
        setHoverState,
      })
    },
    [
      commitActiveEdit,
      activeSelectionCell,
      emitSelectionChange,
      focusGrid,
      isEditingCell,
      onMoveRange,
      rangeMoveCleanupRef,
      refreshHoverState,
      renderState.scrollViewportRef,
      resolvePointerCell,
      selectionRange,
      setGridSelection,
      setHoverState,
      setIsRangeMoveDragging,
    ],
  )

  const beginColumnResize = useCallback(
    (columnIndex: number, startClientX: number) => {
      beginWorkbookGridColumnResize({
        cleanupRef: resizeCleanupRef,
        listenerTarget: window,
        startResize: () => startGridResize(interactionState),
        finishResize: () => finishGridResize(interactionState),
        refreshHoverState,
        setActiveResizeColumn,
        previewColumnWidth,
        getPreviewColumnWidth,
        clearColumnResizePreview,
        commitColumnWidth,
        columnIndex,
        startClientX,
        columnWidths,
        defaultColumnWidth: gridMetrics.columnWidth,
      })
    },
    [
      clearColumnResizePreview,
      columnWidths,
      commitColumnWidth,
      gridMetrics.columnWidth,
      getPreviewColumnWidth,
      interactionState,
      previewColumnWidth,
      refreshHoverState,
      resizeCleanupRef,
      setActiveResizeColumn,
    ],
  )

  const beginRowResize = useCallback(
    (rowIndex: number, startClientY: number) => {
      beginWorkbookGridRowResize({
        cleanupRef: resizeCleanupRef,
        listenerTarget: window,
        startResize: () => startGridResize(interactionState),
        finishResize: () => finishGridResize(interactionState),
        refreshHoverState,
        setActiveResizeRow,
        previewRowHeight,
        getPreviewRowHeight,
        clearRowResizePreview,
        commitRowHeight,
        rowIndex,
        startClientY,
        rowHeights,
        defaultRowHeight: gridMetrics.rowHeight,
      })
    },
    [
      clearRowResizePreview,
      commitRowHeight,
      gridMetrics.rowHeight,
      getPreviewRowHeight,
      interactionState,
      previewRowHeight,
      refreshHoverState,
      resizeCleanupRef,
      rowHeights,
      setActiveResizeRow,
    ],
  )

  const applyColumnAutofitAtPointer = useCallback(
    (event: ReactMouseEvent<HTMLDivElement> | ReactPointerEvent<HTMLDivElement>, visibleRegion = getVisibleRegion()): boolean => {
      return handleWorkbookGridColumnAutofitAtPointer({
        event,
        visibleRegion,
        pointerGeometry: resolvePointerGeometry(visibleRegion),
        columnWidths,
        defaultColumnWidth: gridMetrics.columnWidth,
        isEditingCell,
        commitActiveEdit,
        computeAutofitColumnWidth,
        applyAutofitWidth,
        finishResize: () => finishGridResize(interactionState),
        resetPointerInteraction: () => {
          resetGridPointerInteraction(interactionState, {
            clearIgnoreNextPointerSelection: true,
          })
        },
        setActiveResizeColumn,
        resolveColumnResizeTargetAtPointer,
      })
    },
    [
      applyAutofitWidth,
      columnWidths,
      commitActiveEdit,
      computeAutofitColumnWidth,
      getVisibleRegion,
      gridMetrics.columnWidth,
      interactionState,
      isEditingCell,
      resolveColumnResizeTargetAtPointer,
      resolvePointerGeometry,
      setActiveResizeColumn,
    ],
  )

  return {
    handleFillHandlePointerDown,
    handleHostClickCapture: (event: ReactMouseEvent<HTMLDivElement>) => {
      if (event.detail < 2) {
        return
      }
      applyColumnAutofitAtPointer(event)
    },
    handleHostDoubleClickCapture: (event: ReactMouseEvent<HTMLDivElement>) => {
      const visibleRegion = getVisibleRegion()
      if (applyColumnAutofitAtPointer(event, visibleRegion)) {
        return
      }
      handleGridBodyDoubleClick({
        event,
        applyColumnWidth: commitColumnWidth,
        beginEditAt,
        columnWidths,
        computeAutofitColumnWidth,
        defaultColumnWidth: gridMetrics.columnWidth,
        interactionState,
        isEditingCell,
        lastBodyClickCell: lastBodyClickCellRef.current,
        onAutofitColumn,
        onCommitEdit: commitActiveEdit,
        onSelectionChange: emitSelectionChange,
        resolveColumnResizeTargetAtPointer,
        resolvePointerCell,
        resolvePointerGeometry,
        selectedCell: activeSelectionCell,
        setGridSelection,
        visibleRegion,
      })
    },
    handleHostPointerDownCapture: (event: ReactPointerEvent<HTMLDivElement>) => {
      if (isFillHandleTarget(event.target)) {
        return
      }
      const visibleRegion = getVisibleRegion()
      const pointerGeometry = resolvePointerGeometry(visibleRegion)
      if (
        handleWorkbookGridResizePointerDown({
          event,
          visibleRegion,
          pointerGeometry,
          columnWidths,
          rowHeights,
          defaultColumnWidth: gridMetrics.columnWidth,
          defaultRowHeight: gridMetrics.rowHeight,
          isEditingCell,
          commitActiveEdit,
          focusGrid,
          setActiveHeaderDrag,
          setHoverState,
          lastResizeHandleActivationRef,
          now: () => window.performance.now(),
          computeAutofitColumnWidth,
          applyAutofitWidth,
          finishResize: () => finishGridResize(interactionState),
          resetPointerInteraction: () => {
            resetGridPointerInteraction(interactionState, {
              clearIgnoreNextPointerSelection: true,
            })
          },
          setActiveResizeColumn,
          beginColumnResize,
          beginRowResize,
          resolveColumnResizeTargetAtPointer,
          resolveRowResizeTargetAtPointer,
        })
      ) {
        return
      }

      const headerSelection =
        pointerGeometry === null ? null : resolveHeaderSelectionAtPointer(event.clientX, event.clientY, visibleRegion, pointerGeometry)
      if (!headerSelection && allowsRangeMove) {
        const pointerCell =
          pointerGeometry === null ? null : resolvePointerCell(event.clientX, event.clientY, visibleRegion, pointerGeometry)
        const rangeMoveAnchorCell = resolveSelectionMoveAnchorCellFromPointerCell(
          event.clientX,
          event.clientY,
          selectionRange,
          pointerCell,
          getCellScreenBounds,
        )
        if (rangeMoveAnchorCell) {
          event.preventDefault()
          event.stopPropagation()
          resetGridPointerInteraction(interactionState, {
            clearIgnoreNextPointerSelection: true,
          })
          setActiveResizeColumn(null)
          setActiveResizeRow(null)
          setActiveHeaderDrag(null)
          beginRangeMove(rangeMoveAnchorCell)
          return
        }
      }

      setActiveResizeColumn(null)
      setActiveResizeRow(null)
      setActiveHeaderDrag(headerSelection)
      setHoverState(resetGridHoverState)
      handleGridPointerDown({
        columnWidths,
        defaultColumnWidth: gridMetrics.columnWidth,
        event,
        focusGrid,
        interactionState,
        isEditingCell,
        onCommitEdit: commitActiveEdit,
        onSelectionChange: emitSelectionChange,
        resolveColumnResizeTargetAtPointer,
        resolveHeaderSelectionAtPointer,
        resolvePointerCell,
        resolvePointerGeometry,
        selectedCell: activeSelectionCell,
        setGridSelection,
        visibleRegion,
      })
    },
    handleHostPointerLeave: () => {
      if (
        resizeCleanupRef.current !== null ||
        activeResizeColumn !== null ||
        activeResizeRow !== null ||
        isFillHandleDragging ||
        isRangeMoveDragging
      ) {
        return
      }
      if (hostElement) {
        hostElement.style.cursor = 'default'
      }
      setHoverState(resetGridHoverState)
    },
    handleHostPointerMoveCapture: (event: ReactPointerEvent<HTMLDivElement>) => {
      if (resizeCleanupRef.current !== null || activeResizeColumn !== null || activeResizeRow !== null) {
        return
      }
      if (isFillHandleDragging || isFillHandleTarget(event.target)) {
        return
      }
      const visibleRegion = getVisibleRegion()
      handleGridPointerMove({
        dragAnchorCell: dragAnchorCellRef.current,
        dragHeaderSelection: dragHeaderSelectionRef.current,
        dragPointerCell: dragPointerCellRef.current,
        event,
        interactionState,
        isEditingCell,
        onCommitEdit: commitActiveEdit,
        onSelectionChange: emitSelectionChange,
        resolveHeaderSelectionForPointerDrag,
        resolvePointerCell,
        selectedCell: activeSelectionCell,
        setGridSelection,
        visibleRegion,
      })
      refreshHoverState(event.clientX, event.clientY, event.buttons)
    },
    handleHostPointerUpCapture: (event: ReactPointerEvent<HTMLDivElement>) => {
      if (resizeCleanupRef.current !== null || activeResizeColumn !== null || activeResizeRow !== null) {
        return
      }
      if (isRangeMoveDragging) {
        return
      }
      const visibleRegion = getVisibleRegion()
      const pointerGeometry = resolvePointerGeometry(visibleRegion)
      const resizeTarget = resolveColumnResizeTargetAtPointer(
        event.clientX,
        event.clientY,
        visibleRegion,
        pointerGeometry,
        columnWidths,
        gridMetrics.columnWidth,
      )
      if (resizeTarget !== null && event.detail >= 2) {
        applyWorkbookGridColumnAutofit({
          columnIndex: resizeTarget,
          computeAutofitColumnWidth,
          finishResize: () => finishGridResize(interactionState),
          resetPointerInteraction: () => {
            resetGridPointerInteraction(interactionState, {
              clearIgnoreNextPointerSelection: true,
            })
          },
          setActiveResizeColumn,
          applyAutofitWidth,
        })
        return
      }
      handleGridPointerUp({
        dragAnchorCell: dragAnchorCellRef.current,
        dragDidMove: dragDidMoveRef.current,
        dragHeaderSelection: dragHeaderSelectionRef.current,
        dragPointerCell: dragPointerCellRef.current,
        event,
        interactionState,
        isEditingCell,
        lastBodyClickCellRef,
        onCommitEdit: commitActiveEdit,
        onSelectionChange: emitSelectionChange,
        postDragSelectionExpiryRef,
        resolveHeaderSelectionForPointerDrag,
        resolvePointerCell,
        selectedCell: activeSelectionCell,
        setGridSelection,
        visibleRegion,
      })
      focusGrid()
      setActiveHeaderDrag(null)
      refreshHoverState(event.clientX, event.clientY, 0)
    },
  }
}
