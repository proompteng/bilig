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
import { createRectangleSelectionFromRange, rectangleToAddresses, selectionToSnapshot, snapshotToSelection } from './gridSelection.js'
import { resolveGridSelectionPendingSync } from './gridSelectionPendingSync.js'
import { resolveFillHandlePreviewRange, resolveFillHandleSelectionRange } from './gridFillHandle.js'
import { resolveSelectionMoveAnchorCell } from './gridRangeMove.js'
import type { HeaderSelection, PointerGeometry, VisibleRegionState } from './gridPointer.js'
import { sameGridHoverState } from './gridHover.js'
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
import { resolveGridInteractionHoverState } from './gridInteractionHoverState.js'
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
import { beginWorkbookGridColumnResize, beginWorkbookGridRowResize } from './gridResizeInteractions.js'
import { beginWorkbookGridRangeMove } from './gridRangeMoveInteractions.js'
import { handleWorkbookGridKeyDownCapture } from './gridKeyboardCapture.js'
import type { GridSelection, Item } from './gridTypes.js'
import type { EditMovement, EditSelectionBehavior, GridSelectionSnapshot, WorkbookGridSurfaceProps } from './workbookGridSurfaceTypes.js'
import { useWorkbookGridContextMenu } from './useWorkbookGridContextMenu.js'
import { useWorkbookGridKeyboardHandler } from './useWorkbookGridKeyboardHandler.js'
import type { useWorkbookGridRenderState } from './useWorkbookGridRenderState.js'
import { useWorkbookGridPointerResolvers } from './useWorkbookGridPointerResolvers.js'
import { useWorkbookGridSelectionSummary } from './useWorkbookGridSelectionSummary.js'

const RESIZE_HANDLE_DOUBLE_CLICK_MS = 700

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
  const dragViewportRef = useRef<VisibleRegionState | null>(null)
  const dragGeometryRef = useRef<PointerGeometry | null>(null)
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
  const fillHandlePointerIdRef = useRef<number | null>(null)
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
      dragViewportRef,
      dragGeometryRef,
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
    getVisibleRegion,
    columnWidths,
    rowHeights,
    gridMetrics,
    selectedCell: { col: activeSelectionCell[0], row: activeSelectionCell[1] },
    gridSelection,
    getCellScreenBounds,
    getGeometrySnapshot: () => renderState.gridCameraStore.getSnapshot(),
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
      dragGeometryRef.current = null
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
      fillHandleCleanupRef.current?.()
      fillPreviewRangeRef.current = null
      setFillPreviewRange(null)
      fillHandlePointerIdRef.current = event.pointerId
      setIsFillHandleDragging(true)
      setHoverState((current) =>
        sameGridHoverState(current, { cell: null, header: null, cursor: 'default' })
          ? current
          : { cell: null, header: null, cursor: 'default' },
      )

      const move = (nativeEvent: PointerEvent) => {
        if (nativeEvent.pointerId !== fillHandlePointerIdRef.current) {
          return
        }
        const pointerCell = resolvePointerCell(nativeEvent.clientX, nativeEvent.clientY)
        const nextPreviewRange = pointerCell ? resolveFillHandlePreviewRange(selectionRange, pointerCell) : null
        fillPreviewRangeRef.current = nextPreviewRange
        setFillPreviewRange(nextPreviewRange)
      }

      const cleanup = () => {
        if (fillHandleCleanupRef.current !== cleanup) {
          return
        }
        fillHandleCleanupRef.current = null
        window.removeEventListener('pointermove', move, true)
        window.removeEventListener('pointerup', up, true)
        window.removeEventListener('pointercancel', cancel, true)
        fillHandlePointerIdRef.current = null
        setIsFillHandleDragging(false)
        setHoverState((current) =>
          sameGridHoverState(current, { cell: null, header: null, cursor: 'default' })
            ? current
            : { cell: null, header: null, cursor: 'default' },
        )
      }

      const finish = () => {
        const previewRange = fillPreviewRangeRef.current
        if (previewRange) {
          const source = rectangleToAddresses(selectionRange)
          const target = rectangleToAddresses(previewRange)
          const nextSelectionRange = resolveFillHandleSelectionRange(selectionRange, previewRange)
          const nextSelection = createRectangleSelectionFromRange(nextSelectionRange)
          if (gridSelection.current?.cell && nextSelection.current) {
            nextSelection.current = {
              ...nextSelection.current,
              cell: gridSelection.current.cell,
            }
          }
          setGridSelection(nextSelection)
          if (source.startAddress !== target.startAddress || source.endAddress !== target.endAddress) {
            onFillRange(source.startAddress, source.endAddress, target.startAddress, target.endAddress)
          }
        }
        fillPreviewRangeRef.current = null
        setFillPreviewRange(null)
        cleanup()
      }

      const up = (nativeEvent: PointerEvent) => {
        if (nativeEvent.pointerId !== fillHandlePointerIdRef.current) {
          return
        }
        finish()
      }

      const cancel = (nativeEvent: PointerEvent) => {
        if (nativeEvent.pointerId !== fillHandlePointerIdRef.current) {
          return
        }
        fillPreviewRangeRef.current = null
        setFillPreviewRange(null)
        cleanup()
      }

      fillHandleCleanupRef.current = cleanup
      window.addEventListener('pointermove', move, true)
      window.addEventListener('pointerup', up, true)
      window.addEventListener('pointercancel', cancel, true)
    },
    [
      focusGrid,
      gridSelection,
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
      if (isFillHandleDragging) {
        setHoverState((current) =>
          sameGridHoverState(current, { cell: null, header: null, cursor: 'default' })
            ? current
            : { cell: null, header: null, cursor: 'default' },
        )
        return
      }
      if (isRangeMoveDragging) {
        setHoverState((current) =>
          sameGridHoverState(current, { cell: null, header: null, cursor: 'grabbing' })
            ? current
            : { cell: null, header: null, cursor: 'grabbing' },
        )
        return
      }
      if (buttons !== 0 || fillPreviewRangeRef.current) {
        setHoverState((current) =>
          sameGridHoverState(current, { cell: null, header: null, cursor: 'default' })
            ? current
            : { cell: null, header: null, cursor: 'default' },
        )
        return
      }
      const visibleRegion = getVisibleRegion()
      const geometry = resolvePointerGeometry(visibleRegion)
      if (!geometry) {
        setHoverState((current) =>
          sameGridHoverState(current, { cell: null, header: null, cursor: 'default' })
            ? current
            : { cell: null, header: null, cursor: 'default' },
        )
        return
      }
      const rangeMoveAnchorCell = allowsRangeMove
        ? resolveSelectionMoveAnchorCell(clientX, clientY, selectionRange, getCellScreenBounds)
        : null
      if (rangeMoveAnchorCell) {
        setHoverState((current) =>
          sameGridHoverState(current, { cell: null, header: null, cursor: 'grab' })
            ? current
            : { cell: null, header: null, cursor: 'grab' },
        )
        return
      }
      const next = resolveGridInteractionHoverState({
        clientX,
        clientY,
        columnWidths,
        geometry,
        gridMetrics,
        resolveColumnResizeTargetAtPointer,
        resolveHeaderSelectionAtPointer,
        resolvePointerCell,
        resolveRowResizeTargetAtPointer,
        rowHeights,
        visibleRegion,
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
      if (resizeTarget === null) {
        return
      }
      event.preventDefault()
      event.stopPropagation()
      if (isEditingCell) {
        commitActiveEdit()
      }
      const autofitWidth = computeAutofitColumnWidth(resizeTarget)
      finishGridResize(interactionState)
      resetGridPointerInteraction(interactionState, {
        clearIgnoreNextPointerSelection: true,
      })
      setActiveResizeColumn(null)
      applyAutofitWidth(resizeTarget, autofitWidth)
    },
    handleHostDoubleClickCapture: (event: ReactMouseEvent<HTMLDivElement>) => {
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
      if (resizeTarget !== null) {
        event.preventDefault()
        event.stopPropagation()
        if (isEditingCell) {
          commitActiveEdit()
        }
        const autofitWidth = computeAutofitColumnWidth(resizeTarget)
        finishGridResize(interactionState)
        resetGridPointerInteraction(interactionState, {
          clearIgnoreNextPointerSelection: true,
        })
        setActiveResizeColumn(null)
        applyAutofitWidth(resizeTarget, autofitWidth)
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
      const resizeTarget = resolveColumnResizeTargetAtPointer(
        event.clientX,
        event.clientY,
        visibleRegion,
        pointerGeometry,
        columnWidths,
        gridMetrics.columnWidth,
      )
      const rowResizeTarget = resolveRowResizeTargetAtPointer(
        event.clientX,
        event.clientY,
        visibleRegion,
        pointerGeometry,
        rowHeights,
        gridMetrics.rowHeight,
      )
      if (resizeTarget !== null) {
        event.preventDefault()
        event.stopPropagation()
        if (isEditingCell) {
          commitActiveEdit()
        }
        focusGrid()
        setActiveHeaderDrag(null)
        const now = window.performance.now()
        const lastResizeHandleActivation = lastResizeHandleActivationRef.current
        const isResizeDoubleClick =
          lastResizeHandleActivation !== null &&
          lastResizeHandleActivation.columnIndex === resizeTarget &&
          now - lastResizeHandleActivation.at <= RESIZE_HANDLE_DOUBLE_CLICK_MS
        lastResizeHandleActivationRef.current = { columnIndex: resizeTarget, at: now }
        if (isResizeDoubleClick) {
          lastResizeHandleActivationRef.current = null
          const autofitWidth = computeAutofitColumnWidth(resizeTarget)
          finishGridResize(interactionState)
          resetGridPointerInteraction(interactionState, {
            clearIgnoreNextPointerSelection: true,
          })
          setActiveResizeColumn(null)
          applyAutofitWidth(resizeTarget, autofitWidth)
          return
        }
        setHoverState((current) =>
          sameGridHoverState(current, {
            cell: null,
            header: { kind: 'column', index: resizeTarget },
            cursor: 'col-resize',
          })
            ? current
            : {
                cell: null,
                header: { kind: 'column', index: resizeTarget },
                cursor: 'col-resize',
              },
        )
        beginColumnResize(resizeTarget, event.clientX)
        return
      }
      if (rowResizeTarget !== null) {
        event.preventDefault()
        event.stopPropagation()
        if (isEditingCell) {
          commitActiveEdit()
        }
        focusGrid()
        setActiveHeaderDrag(null)
        setHoverState((current) =>
          sameGridHoverState(current, {
            cell: null,
            header: { kind: 'row', index: rowResizeTarget },
            cursor: 'row-resize',
          })
            ? current
            : {
                cell: null,
                header: { kind: 'row', index: rowResizeTarget },
                cursor: 'row-resize',
              },
        )
        beginRowResize(rowResizeTarget, event.clientY)
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
      setHoverState((current) =>
        sameGridHoverState(current, { cell: null, header: null, cursor: 'default' })
          ? current
          : { cell: null, header: null, cursor: 'default' },
      )
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
      setHoverState((current) =>
        sameGridHoverState(current, { cell: null, header: null, cursor: 'default' })
          ? current
          : { cell: null, header: null, cursor: 'default' },
      )
    },
    handleHostPointerMoveCapture: (event: ReactPointerEvent<HTMLDivElement>) => {
      if (isFillHandleDragging || isFillHandleTarget(event.target)) {
        return
      }
      const visibleRegion = getVisibleRegion()
      handleGridPointerMove({
        dragAnchorCell: dragAnchorCellRef.current,
        dragGeometry: dragGeometryRef.current,
        dragHeaderSelection: dragHeaderSelectionRef.current,
        dragPointerCell: dragPointerCellRef.current,
        dragViewport: dragViewportRef.current,
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
        const autofitWidth = computeAutofitColumnWidth(resizeTarget)
        finishGridResize(interactionState)
        resetGridPointerInteraction(interactionState, {
          clearIgnoreNextPointerSelection: true,
        })
        setActiveResizeColumn(null)
        applyAutofitWidth(resizeTarget, autofitWidth)
        return
      }
      const clickedCell = dragDidMoveRef.current || dragHeaderSelectionRef.current ? null : resolvePointerCell(event.clientX, event.clientY)
      handleGridPointerUp({
        dragAnchorCell: dragAnchorCellRef.current,
        dragDidMove: dragDidMoveRef.current,
        dragGeometry: dragGeometryRef.current,
        dragHeaderSelection: dragHeaderSelectionRef.current,
        dragPointerCell: dragPointerCellRef.current,
        dragViewport: dragViewportRef.current,
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
