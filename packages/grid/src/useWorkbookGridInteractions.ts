import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  type ClipboardEvent as ReactClipboardEvent,
  type FocusEvent as ReactFocusEvent,
  type KeyboardEvent as ReactKeyboardEvent,
  type MouseEvent as ReactMouseEvent,
  type PointerEvent as ReactPointerEvent,
} from 'react'
import { formatAddress } from '@bilig/formula'
import { flushSync } from 'react-dom'
import { selectionToSnapshot, snapshotToSelection } from './gridSelection.js'
import { resolveGridSelectionPendingSync } from './gridSelectionPendingSync.js'
import { resolveSelectionMoveAnchorCell } from './gridRangeMove.js'
import type { HeaderSelection } from './gridPointer.js'
import { sameGridHoverState, type GridHoverState } from './gridHover.js'
import type { InternalClipboardRange } from './gridInternalClipboard.js'
import {
  finishGridResize,
  handleGridBodyDoubleClick,
  handleGridPointerDown,
  handleGridPointerMove,
  handleGridPointerUp,
  startGridResize,
} from './gridInteractionController.js'
import { clearGridPendingPointerActivation, resetGridPointerInteraction } from './gridInteractionState.js'
import { resolveWorkbookGridHoverState } from './gridInteractionHoverState.js'
import {
  applyGridClipboardValues,
  captureGridClipboardSelection,
  handleGridCopyCapture,
  handleGridPasteCapture,
  isGridKeyboardEditableTarget,
} from './gridClipboardKeyboardController.js'
import {
  beginWorkbookGridEdit,
  openWorkbookGridHeaderContextMenuFromKeyboard,
  selectEntireWorkbookSheet,
  toggleWorkbookGridBooleanCell,
} from './gridInteractionCommands.js'
import {
  applyWorkbookGridColumnAutofit,
  beginWorkbookGridColumnResize,
  beginWorkbookGridRowResize,
  handleWorkbookGridColumnAutofitAtPointer,
  handleWorkbookGridResizePointerDown,
} from './gridResizeInteractions.js'
import { beginWorkbookGridRangeMove } from './gridRangeMoveInteractions.js'
import { beginWorkbookGridFillHandleDrag } from './gridFillHandleInteractions.js'
import { handleWorkbookGridKeyDownCapture } from './gridKeyboardCapture.js'
import type { GridSelection, Item } from './gridTypes.js'
import type { EditMovement, EditSelectionBehavior, GridSelectionSnapshot, WorkbookGridSurfaceProps } from './workbookGridSurfaceTypes.js'
import { useWorkbookGridContextMenu } from './useWorkbookGridContextMenu.js'
import { useWorkbookGridKeyboardHandler } from './useWorkbookGridKeyboardHandler.js'
import type { useWorkbookGridRenderState } from './useWorkbookGridRenderState.js'
import { useWorkbookGridPointerResolvers } from './useWorkbookGridPointerResolvers.js'
import { useWorkbookGridSelectionSummary } from './useWorkbookGridSelectionSummary.js'

const DEFAULT_GRID_HOVER_STATE: GridHoverState = { cell: null, header: null, cursor: 'default' }

function resetGridHoverState(current: GridHoverState): GridHoverState {
  return sameGridHoverState(current, DEFAULT_GRID_HOVER_STATE) ? current : DEFAULT_GRID_HOVER_STATE
}

export function useWorkbookGridInteractions(
  input: Pick<
    WorkbookGridSurfaceProps,
    | 'editorValue'
    | 'isEditingCell'
    | 'onAutofitColumn'
    | 'onBeginEdit'
    | 'onCancelEdit'
    | 'onClearCell'
    | 'onColumnWidthChange'
    | 'onRowHeightChange'
    | 'onCommitEdit'
    | 'onCopyRange'
    | 'onEditorChange'
    | 'onFillRange'
    | 'onMoveRange'
    | 'onPaste'
    | 'hiddenColumns'
    | 'hiddenRows'
    | 'onSetColumnHidden'
    | 'onSetRowHidden'
    | 'onInsertRows'
    | 'onDeleteRows'
    | 'onInsertColumns'
    | 'onDeleteColumns'
    | 'onSetFreezePane'
    | 'onSelectionChange'
    | 'onSelectionLabelChange'
    | 'onToggleBooleanCell'
    | 'selectionSnapshot'
    | 'getCellEditorSeed'
  > & {
    engine: WorkbookGridSurfaceProps['engine']
    sheetName: string
    selectedAddr: string
    selectedCellSnapshot: WorkbookGridSurfaceProps['selectedCellSnapshot']
    renderState: ReturnType<typeof useWorkbookGridRenderState>
  },
) {
  const {
    editorValue,
    engine,
    isEditingCell,
    onAutofitColumn,
    onBeginEdit,
    onCancelEdit,
    onClearCell,
    onCommitEdit,
    onCopyRange,
    onEditorChange,
    onFillRange,
    onMoveRange,
    onPaste,
    hiddenColumns,
    hiddenRows,
    onSetColumnHidden,
    onSetRowHidden,
    onInsertRows,
    onDeleteRows,
    onInsertColumns,
    onDeleteColumns,
    onSetFreezePane,
    onSelectionChange,
    onSelectionLabelChange,
    onToggleBooleanCell,
    selectionSnapshot,
    sheetName,
    selectedAddr,
    getCellEditorSeed,
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
    fillPreviewRange,
    focusGrid,
    getCellScreenBounds,
    getVisibleRegion,
    getPreviewColumnWidth,
    getPreviewRowHeight,
    gridMetrics,
    gridSelection,
    hostRef,
    isFillHandleDragging,
    isRangeMoveDragging,
    previewColumnWidth,
    previewRowHeight,
    rowHeights,
    selectedCell,
    selectionRange,
    setActiveHeaderDrag,
    setActiveResizeColumn,
    setActiveResizeRow,
    setFillPreviewRange,
    setGridSelection,
    setHoverState,
    setIsFillHandleDragging,
    setIsRangeMoveDragging,
  } = renderState
  const activeSelectionCell = useMemo<Item>(
    () => gridSelection.current?.cell ?? [selectedCell.col, selectedCell.row],
    [gridSelection.current, selectedCell.col, selectedCell.row],
  )
  const wasEditingOverlayRef = useRef(false)
  const ignoreNextPointerSelectionRef = useRef(false)
  const pendingPointerCellRef = useRef<Item | null>(null)
  const dragAnchorCellRef = useRef<Item | null>(null)
  const dragPointerCellRef = useRef<Item | null>(null)
  const dragHeaderSelectionRef = useRef<HeaderSelection | null>(null)
  const dragDidMoveRef = useRef(false)
  const postDragSelectionExpiryRef = useRef<number>(0)
  const columnResizeActiveRef = useRef(false)
  const lastBodyClickCellRef = useRef<Item | null>(null)
  const internalClipboardRef = useRef<InternalClipboardRange | null>(null)
  const pendingClipboardCopySequenceRef = useRef(0)
  const pendingKeyboardPasteSequenceRef = useRef(0)
  const suppressNextNativePasteRef = useRef(false)
  const pendingTypeSeedRef = useRef<string | null>(null)
  const pendingLocalSelectionSnapshotRef = useRef<GridSelectionSnapshot | null>(null)
  const pendingLocalSelectionBaseSnapshotRef = useRef<GridSelectionSnapshot | null>(null)
  const lastResizeHandleActivationRef = useRef<{ columnIndex: number; at: number } | null>(null)
  const fillPreviewRangeRef = useRef(fillPreviewRange)
  const fillHandleCleanupRef = useRef<(() => void) | null>(null)
  const rangeMoveCleanupRef = useRef<(() => void) | null>(null)
  const resizeCleanupRef = useRef<(() => void) | null>(null)
  const activeSheetRef = useRef(sheetName)
  const interactionState = useMemo(
    () => ({
      ignoreNextPointerSelectionRef,
      pendingPointerCellRef,
      dragAnchorCellRef,
      dragPointerCellRef,
      dragHeaderSelectionRef,
      dragDidMoveRef,
      postDragSelectionExpiryRef,
      columnResizeActiveRef,
    }),
    [columnResizeActiveRef],
  )
  const {
    resolveColumnResizeTarget: resolveColumnResizeTargetAtPointer,
    resolveRowResizeTarget: resolveRowResizeTargetAtPointer,
    resolveHeaderSelectionAtPointer,
    resolveHeaderSelectionForPointerDrag,
    resolvePointerCell,
    resolvePointerGeometry,
  } = useWorkbookGridPointerResolvers({
    hostRef,
    selectedCell: { col: activeSelectionCell[0], row: activeSelectionCell[1] },
    gridSelection,
    getGeometrySnapshot: renderState.getLiveGeometrySnapshot,
  })
  useEffect(() => {
    fillPreviewRangeRef.current = fillPreviewRange
  }, [fillPreviewRange])
  useEffect(() => {
    const fillHandleCleanup = fillHandleCleanupRef.current
    const rangeMoveCleanup = rangeMoveCleanupRef.current
    const resizeCleanup = resizeCleanupRef.current
    return () => {
      fillHandleCleanup?.()
      rangeMoveCleanup?.()
      resizeCleanup?.()
    }
  }, [])
  useLayoutEffect(() => {
    const sheetChanged = activeSheetRef.current !== sheetName
    activeSheetRef.current = sheetName
    setGridSelection((current) => {
      const currentSnapshot = selectionToSnapshot(current, selectionSnapshot.sheetName, selectionSnapshot.address)
      const sync = resolveGridSelectionPendingSync({
        currentSnapshot,
        externalSnapshot: selectionSnapshot,
        pendingBaseSnapshot: pendingLocalSelectionBaseSnapshotRef.current,
        pendingLocalSnapshot: pendingLocalSelectionSnapshotRef.current,
        sheetChanged,
      })
      pendingLocalSelectionSnapshotRef.current = sync.pendingLocalSnapshot
      pendingLocalSelectionBaseSnapshotRef.current = sync.pendingBaseSnapshot
      if (sync.keepCurrentSelection) {
        return current
      }
      clearGridPendingPointerActivation(interactionState)
      return snapshotToSelection(selectionSnapshot)
    })
  }, [interactionState, selectionSnapshot, setGridSelection, sheetName])
  useEffect(() => {
    if (wasEditingOverlayRef.current && !isEditingCell) {
      window.requestAnimationFrame(() => {
        focusGrid()
      })
    }
    if (isEditingCell) {
      pendingTypeSeedRef.current = null
    }
    wasEditingOverlayRef.current = isEditingCell
  }, [focusGrid, isEditingCell])
  const beginSelectedEdit = useCallback(
    (seed?: string, selectionBehavior: EditSelectionBehavior = 'caret-end') => {
      const address = formatAddress(activeSelectionCell[1], activeSelectionCell[0])
      beginWorkbookGridEdit({
        engine,
        onBeginEdit,
        sheetName,
        address,
        selectedCellSnapshot: null,
        seed: seed ?? getCellEditorSeed?.(sheetName, address) ?? undefined,
        selectionBehavior,
      })
    },
    [activeSelectionCell, engine, getCellEditorSeed, onBeginEdit, sheetName],
  )
  const beginEditAt = useCallback(
    (addr: string, seed?: string, selectionBehavior: EditSelectionBehavior = 'caret-end') => {
      beginWorkbookGridEdit({
        engine,
        onBeginEdit,
        sheetName,
        address: addr,
        selectedCellSnapshot: null,
        seed: seed ?? getCellEditorSeed?.(sheetName, addr),
        selectionBehavior,
      })
    },
    [engine, getCellEditorSeed, onBeginEdit, sheetName],
  )
  const syncMountedCellEditorValue = useCallback((): string | null => {
    if (typeof document === 'undefined') {
      return null
    }
    const editor = document.querySelector<HTMLTextAreaElement>('[data-testid="cell-editor-input"]')
    if (!editor) {
      return null
    }
    if (editor.value !== editorValue) {
      flushSync(() => {
        onEditorChange(editor.value)
      })
    }
    return editor.value
  }, [editorValue, onEditorChange])
  const commitActiveEdit = useCallback(
    (movement?: EditMovement) => {
      const valueOverride = syncMountedCellEditorValue()
      onCommitEdit(movement, valueOverride ?? undefined, {
        sheetName,
        address: formatAddress(activeSelectionCell[1], activeSelectionCell[0]),
      })
    },
    [activeSelectionCell, onCommitEdit, sheetName, syncMountedCellEditorValue],
  )
  const toggleBooleanCellAt = useCallback(
    (col: number, row: number): boolean => {
      return toggleWorkbookGridBooleanCell({
        engine,
        onToggleBooleanCell,
        sheetName,
        col,
        row,
      })
    },
    [engine, onToggleBooleanCell, sheetName],
  )
  useWorkbookGridSelectionSummary({
    gridSelection,
    selectedAddr,
    onSelectionLabelChange,
  })
  const emitSelectionChange = useCallback(
    (nextSelection: GridSelection) => {
      const nextSelectionSnapshot = selectionToSnapshot(nextSelection, sheetName, selectionSnapshot.address)
      pendingLocalSelectionBaseSnapshotRef.current = selectionSnapshot
      pendingLocalSelectionSnapshotRef.current = nextSelectionSnapshot
      onSelectionChange(nextSelectionSnapshot)
    },
    [onSelectionChange, selectionSnapshot, sheetName],
  )
  const allowsRangeMove = Boolean(
    selectionRange && gridSelection.columns.length === 0 && gridSelection.rows.length === 0 && !fillPreviewRange && !isFillHandleDragging,
  )
  const isFillHandleTarget = useCallback((target: EventTarget | null): boolean => {
    return target instanceof Element && target.closest("[data-grid-fill-handle='true']") !== null
  }, [])
  const applyClipboardValues = useCallback(
    (target: Item, values: readonly (readonly string[])[]) => {
      applyGridClipboardValues({
        internalClipboardRef,
        onCopyRange,
        onPaste,
        sheetName,
        target,
        values,
      })
    },
    [onCopyRange, onPaste, sheetName],
  )
  const captureInternalClipboardSelection = useCallback(() => {
    return captureGridClipboardSelection({
      engine,
      gridSelection,
      internalClipboardRef,
      sheetName,
    })
  }, [engine, gridSelection, sheetName])
  const { handleGridKey } = useWorkbookGridKeyboardHandler({
    applyClipboardValues,
    beginSelectedEdit,
    captureInternalClipboardSelection,
    editorValue,
    engine,
    gridSelection,
    hostRef,
    internalClipboardRef,
    isEditingCell,
    onCancelEdit,
    onClearCell,
    onCommitEdit,
    onEditorChange,
    onSelectionChange: emitSelectionChange,
    pendingClipboardCopySequenceRef,
    pendingKeyboardPasteSequenceRef,
    pendingTypeSeedRef,
    selectedCell: { col: activeSelectionCell[0], row: activeSelectionCell[1] },
    setGridSelection,
    sheetName,
    suppressNextNativePasteRef,
    toggleSelectedBooleanCell: () => {
      toggleBooleanCellAt(activeSelectionCell[0], activeSelectionCell[1])
    },
  })
  const contextMenu = useWorkbookGridContextMenu({
    focusGrid,
    getVisibleRegion,
    hiddenColumnsByIndex: hiddenColumns,
    hiddenRowsByIndex: hiddenRows,
    isEditingCell,
    onCommitEdit: commitActiveEdit,
    onSelectionChange: emitSelectionChange,
    onSetColumnHidden,
    onSetRowHidden,
    onInsertRows,
    onDeleteRows,
    onInsertColumns,
    onDeleteColumns,
    onSetFreezePane,
    resolveHeaderSelectionAtPointer,
    selectedCell: activeSelectionCell,
    setGridSelection,
  })
  const openHeaderContextMenuFromKeyboard = useCallback(() => {
    return openWorkbookGridHeaderContextMenuFromKeyboard({
      hostBounds: hostRef.current?.getBoundingClientRect(),
      gridSelection,
      selectedCell: activeSelectionCell,
      getCellScreenBounds,
      gridMetrics,
      openContextMenuForTarget: contextMenu.openContextMenuForTarget,
    })
  }, [contextMenu, getCellScreenBounds, gridMetrics, gridSelection, hostRef, activeSelectionCell])

  const handleFillHandlePointerDown = useCallback(
    (event: ReactPointerEvent<HTMLButtonElement>) => {
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
        resetHoverState: () => {
          setHoverState(resetGridHoverState)
        },
      })
    },
    [
      focusGrid,
      gridSelection,
      emitSelectionChange,
      onFillRange,
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
      setHoverState((current) => (sameGridHoverState(current, next) ? current : next))
    },
    [
      allowsRangeMove,
      columnWidths,
      getCellScreenBounds,
      gridMetrics,
      isFillHandleDragging,
      isRangeMoveDragging,
      resolveColumnResizeTargetAtPointer,
      resolveHeaderSelectionAtPointer,
      resolveRowResizeTargetAtPointer,
      rowHeights,
      resolvePointerCell,
      resolvePointerGeometry,
      selectionRange,
      setHoverState,
      getVisibleRegion,
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
      focusGrid,
      isEditingCell,
      onMoveRange,
      emitSelectionChange,
      renderState.scrollViewportRef,
      refreshHoverState,
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
      getPreviewColumnWidth,
      gridMetrics.columnWidth,
      interactionState,
      previewColumnWidth,
      refreshHoverState,
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
      getPreviewRowHeight,
      gridMetrics.rowHeight,
      interactionState,
      previewRowHeight,
      refreshHoverState,
      rowHeights,
      setActiveResizeRow,
    ],
  )
  const applyAutofitWidth = useCallback(
    (columnIndex: number, width: number) => {
      flushSync(() => {
        if (onAutofitColumn) {
          previewColumnWidth(columnIndex, width)
        } else {
          commitColumnWidth(columnIndex, width)
        }
      })
      if (onAutofitColumn) {
        void Promise.resolve(onAutofitColumn(columnIndex, width))
      }
    },
    [commitColumnWidth, onAutofitColumn, previewColumnWidth],
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

  const handleSelectEntireSheet = useCallback(() => {
    selectEntireWorkbookSheet({
      isEditingCell,
      onCommitEdit: commitActiveEdit,
      setGridSelection,
      onSelectionChange: emitSelectionChange,
      focusGrid,
    })
  }, [commitActiveEdit, emitSelectionChange, focusGrid, isEditingCell, setGridSelection])

  return {
    handleFillHandlePointerDown,
    handleGridKey,
    handleHostKeyDownCapture: (event: ReactKeyboardEvent<HTMLDivElement>) => {
      if (isGridKeyboardEditableTarget(event.target)) {
        return
      }
      if ((event.nativeEvent as KeyboardEvent & { __biligGridHandled?: boolean }).__biligGridHandled === true) {
        return
      }
      const normalizedKey = event.key
      if (!isEditingCell && normalizedKey === 'F2' && !event.altKey && !event.ctrlKey && !event.metaKey) {
        event.preventDefault()
        event.stopPropagation()
        ;(event.nativeEvent as KeyboardEvent & { __biligGridHandled?: boolean }).__biligGridHandled = true
        beginSelectedEdit(undefined, 'caret-end')
        return
      }
      handleWorkbookGridKeyDownCapture({
        event,
        handleGridKey,
        openHeaderContextMenuFromKeyboard,
        resetPointerInteraction: () => {
          resetGridPointerInteraction(interactionState, {
            clearIgnoreNextPointerSelection: true,
          })
        },
      })
      if (event.defaultPrevented) {
        ;(event.nativeEvent as KeyboardEvent & { __biligGridHandled?: boolean }).__biligGridHandled = true
      }
    },
    handleHostCopyCapture: (event: ReactClipboardEvent<HTMLDivElement>) => {
      handleGridCopyCapture({
        captureInternalClipboardSelection,
        event,
        internalClipboardRef,
      })
    },
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
        editorValue,
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
    handleHostFocus: (event: ReactFocusEvent<HTMLDivElement>) => {
      if (event.target === event.currentTarget) {
        focusGrid()
      }
    },
    handleHostContextMenuCapture: contextMenu.handleHostContextMenuCapture,
    handleHostKeyDown: (event: ReactKeyboardEvent<HTMLDivElement>) => {
      if (isGridKeyboardEditableTarget(event.target)) {
        return
      }
      if ((event.nativeEvent as KeyboardEvent & { __biligGridHandled?: boolean }).__biligGridHandled === true) {
        return
      }
      handleGridKey(event)
    },
    handleHostPasteCapture: (event: ReactClipboardEvent<HTMLDivElement>) => {
      handleGridPasteCapture({
        applyClipboardValues,
        event,
        gridSelection,
        pendingKeyboardPasteSequenceRef,
        selectedCell,
        suppressNextNativePasteRef,
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

      if (allowsRangeMove) {
        const rangeMoveAnchorCell = resolveSelectionMoveAnchorCell(event.clientX, event.clientY, selectionRange, getCellScreenBounds)
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

      const headerSelection =
        pointerGeometry === null ? null : resolveHeaderSelectionAtPointer(event.clientX, event.clientY, visibleRegion, pointerGeometry)
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
      if (activeResizeColumn !== null || activeResizeRow !== null || isFillHandleDragging || isRangeMoveDragging) {
        return
      }
      setHoverState(resetGridHoverState)
    },
    handleHostPointerMoveCapture: (event: ReactPointerEvent<HTMLDivElement>) => {
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
      const clickedCell = dragDidMoveRef.current || dragHeaderSelectionRef.current ? null : resolvePointerCell(event.clientX, event.clientY)
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
        selectedCell: [selectedCell.col, selectedCell.row],
        setGridSelection,
        visibleRegion,
      })
      if (clickedCell) {
        toggleBooleanCellAt(clickedCell[0], clickedCell[1])
      }
      focusGrid()
      setActiveHeaderDrag(null)
      refreshHoverState(event.clientX, event.clientY, 0)
    },
    handleSelectEntireSheet,
    contextMenu,
  }
}
