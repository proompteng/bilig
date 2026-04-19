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
import { createRectangleSelectionFromRange, rectangleToAddresses, selectionToSnapshot, snapshotToSelection } from './gridSelection.js'
import { resolveFillHandlePreviewRange, resolveFillHandleSelectionRange } from './gridFillHandle.js'
import { resolveSelectionMoveAnchorCell } from './gridRangeMove.js'
import { resolveColumnResizeTarget, type HeaderSelection, type PointerGeometry, type VisibleRegionState } from './gridPointer.js'
import { resolveGridHoverState, sameGridHoverState } from './gridHover.js'
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
import {
  applyGridClipboardValues,
  captureGridClipboardSelection,
  handleGridCopyCapture,
  handleGridPasteCapture,
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
import type { EditSelectionBehavior, WorkbookGridSurfaceProps } from './workbookGridSurfaceTypes.js'
import { useWorkbookGridContextMenu } from './useWorkbookGridContextMenu.js'
import { useWorkbookGridKeyboardHandler } from './useWorkbookGridKeyboardHandler.js'
import type { useWorkbookGridRenderState } from './useWorkbookGridRenderState.js'
import { useWorkbookGridPointerResolvers } from './useWorkbookGridPointerResolvers.js'
import { useWorkbookGridSelectionSummary } from './useWorkbookGridSelectionSummary.js'

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
    selectedCellSnapshot,
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
    visibleRegion,
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
  const pendingKeyboardPasteSequenceRef = useRef(0)
  const suppressNextNativePasteRef = useRef(false)
  const pendingTypeSeedRef = useRef<string | null>(null)
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
    resolveRowResizeTarget: resolveRowResizeTargetAtPointer,
    resolveHeaderSelectionAtPointer,
    resolveHeaderSelectionForPointerDrag,
    resolvePointerCell,
    resolvePointerGeometry,
  } = useWorkbookGridPointerResolvers({
    hostRef,
    visibleRegion,
    columnWidths,
    rowHeights,
    gridMetrics,
    selectedCell,
    gridSelection,
    getCellScreenBounds,
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
      if (
        !sheetChanged &&
        currentSnapshot.sheetName === selectionSnapshot.sheetName &&
        currentSnapshot.address === selectionSnapshot.address &&
        currentSnapshot.kind === selectionSnapshot.kind &&
        currentSnapshot.range.startAddress === selectionSnapshot.range.startAddress &&
        currentSnapshot.range.endAddress === selectionSnapshot.range.endAddress
      ) {
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
      beginWorkbookGridEdit({
        engine,
        onBeginEdit,
        sheetName,
        address: selectedAddr,
        selectedCellSnapshot,
        seed,
        selectionBehavior,
      })
    },
    [engine, onBeginEdit, selectedAddr, selectedCellSnapshot, sheetName],
  )
  const beginEditAt = useCallback(
    (addr: string, seed?: string, selectionBehavior: EditSelectionBehavior = 'caret-end') => {
      beginWorkbookGridEdit({
        engine,
        onBeginEdit,
        sheetName,
        address: addr,
        selectedCellSnapshot: addr === selectedAddr ? selectedCellSnapshot : null,
        seed,
        selectionBehavior,
      })
    },
    [engine, onBeginEdit, selectedAddr, selectedCellSnapshot, sheetName],
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
      onSelectionChange(selectionToSnapshot(nextSelection, sheetName, selectionSnapshot.address))
    },
    [onSelectionChange, selectionSnapshot.address, sheetName],
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
    isEditingCell,
    onCancelEdit,
    onClearCell,
    onCommitEdit,
    onEditorChange,
    onSelectionChange: emitSelectionChange,
    pendingKeyboardPasteSequenceRef,
    pendingTypeSeedRef,
    selectedCell,
    setGridSelection,
    sheetName,
    suppressNextNativePasteRef,
    toggleSelectedBooleanCell: () => {
      toggleBooleanCellAt(activeSelectionCell[0], activeSelectionCell[1])
    },
  })
  const contextMenu = useWorkbookGridContextMenu({
    focusGrid,
    hiddenColumnsByIndex: hiddenColumns,
    hiddenRowsByIndex: hiddenRows,
    isEditingCell,
    onCommitEdit: () => onCommitEdit(),
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
    visibleRegion,
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
      const next = resolveGridHoverState({
        clientX,
        clientY,
        region: visibleRegion,
        geometry,
        columnWidths,
        rowHeights,
        defaultColumnWidth: gridMetrics.columnWidth,
        defaultRowHeight: gridMetrics.rowHeight,
        gridMetrics,
        selectedCell: activeSelectionCell,
        selectedCellBounds: getCellScreenBounds(activeSelectionCell[0], activeSelectionCell[1]) ?? null,
        selectionRange,
        hasColumnSelection: gridSelection.columns.length > 0,
        hasRowSelection: gridSelection.rows.length > 0,
      })
      setHoverState((current) => (sameGridHoverState(current, next) ? current : next))
    },
    [
      allowsRangeMove,
      columnWidths,
      getCellScreenBounds,
      gridMetrics,
      gridSelection.columns.length,
      gridSelection.rows.length,
      isFillHandleDragging,
      isRangeMoveDragging,
      rowHeights,
      resolvePointerGeometry,
      activeSelectionCell,
      selectionRange,
      setHoverState,
      visibleRegion,
    ],
  )

  const beginRangeMove = useCallback(
    (pointerCell: Item) => {
      if (!selectionRange) {
        return
      }
      if (isEditingCell) {
        onCommitEdit()
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
      focusGrid,
      isEditingCell,
      onCommitEdit,
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

  const handleSelectEntireSheet = useCallback(() => {
    selectEntireWorkbookSheet({
      isEditingCell,
      onCommitEdit,
      setGridSelection,
      onSelectionChange: emitSelectionChange,
      focusGrid,
    })
  }, [emitSelectionChange, focusGrid, isEditingCell, onCommitEdit, setGridSelection])

  return {
    handleFillHandlePointerDown,
    handleGridKey,
    handleHostKeyDownCapture: (event: ReactKeyboardEvent<HTMLDivElement>) => {
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
    },
    handleHostCopyCapture: (event: ReactClipboardEvent<HTMLDivElement>) => {
      handleGridCopyCapture({
        captureInternalClipboardSelection,
        event,
        internalClipboardRef,
      })
    },
    handleHostDoubleClickCapture: (event: ReactMouseEvent<HTMLDivElement>) => {
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
        onCommitEdit: () => onCommitEdit(),
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
      const pointerGeometry = resolvePointerGeometry(visibleRegion)
      const resizeTarget =
        pointerGeometry === null
          ? null
          : resolveColumnResizeTarget(event.clientX, event.clientY, visibleRegion, pointerGeometry, columnWidths, gridMetrics.columnWidth)
      const rowResizeTarget =
        pointerGeometry === null
          ? null
          : resolveRowResizeTargetAtPointer(event.clientX, event.clientY, visibleRegion, pointerGeometry, rowHeights, gridMetrics.rowHeight)
      if (resizeTarget !== null) {
        event.preventDefault()
        event.stopPropagation()
        if (isEditingCell) {
          onCommitEdit()
        }
        focusGrid()
        setActiveHeaderDrag(null)
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
          onCommitEdit()
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
        onCommitEdit: () => onCommitEdit(),
        onSelectionChange: emitSelectionChange,
        resolveColumnResizeTargetAtPointer: resolveColumnResizeTarget,
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
      handleGridPointerMove({
        dragAnchorCell: dragAnchorCellRef.current,
        dragGeometry: dragGeometryRef.current,
        dragHeaderSelection: dragHeaderSelectionRef.current,
        dragPointerCell: dragPointerCellRef.current,
        dragViewport: dragViewportRef.current,
        event,
        interactionState,
        isEditingCell,
        onCommitEdit: () => onCommitEdit(),
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
        onCommitEdit: () => onCommitEdit(),
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
