import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  type ClipboardEvent as ReactClipboardEvent,
  type FocusEvent as ReactFocusEvent,
  type KeyboardEvent as ReactKeyboardEvent,
} from 'react'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import { flushSync } from 'react-dom'
import { resetGridPointerInteraction } from './gridInteractionState.js'
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
import { handleWorkbookGridKeyDownCapture } from './gridKeyboardCapture.js'
import { createGridSelection } from './gridSelection.js'
import type { GridSelection, Item } from './gridTypes.js'
import type { EditMovement, EditSelectionBehavior, WorkbookGridSurfaceProps } from './workbookGridSurfaceTypes.js'
import { useWorkbookGridContextMenu } from './useWorkbookGridContextMenu.js'
import { useWorkbookGridHostPointerHandlers } from './useWorkbookGridHostPointerHandlers.js'
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
    | 'onExternalSelectionSync'
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
    onExternalSelectionSync,
    onSelectionLabelChange,
    onToggleBooleanCell,
    selectionSnapshot,
    sheetName,
    selectedAddr,
    getCellEditorSeed,
    renderState,
  } = input
  const {
    commitColumnWidth,
    fillPreviewRange,
    focusGrid,
    getCellScreenBounds,
    getVisibleRegion,
    gridMetrics,
    gridSelection,
    gridRuntimeHost,
    hostRef,
    isFillHandleDragging,
    previewColumnWidth,
    selectedCell,
    selectionRange,
    setGridSelection,
  } = renderState
  const activeSelectionCell = useMemo<Item>(
    () => gridSelection.current?.cell ?? [selectedCell.col, selectedCell.row],
    [gridSelection.current, selectedCell.col, selectedCell.row],
  )
  const inputController = gridRuntimeHost.input
  const interactionState = inputController.interactionState
  const {
    internalClipboardRef,
    pendingClipboardCopySequenceRef,
    pendingKeyboardPasteSequenceRef,
    pendingTypeSeedRef,
    suppressNextNativePasteRef,
  } = inputController
  const pointerResolvers = useWorkbookGridPointerResolvers({
    hostRef,
    selectedCell: { col: activeSelectionCell[0], row: activeSelectionCell[1] },
    gridSelection,
    getGeometrySnapshot: renderState.getLiveGeometrySnapshot,
  })
  const { resolveHeaderSelectionAtPointer } = pointerResolvers
  useEffect(() => {
    inputController.syncFillPreviewRange(fillPreviewRange)
  }, [fillPreviewRange, inputController])
  useLayoutEffect(() => {
    setGridSelection((current) => {
      const nextSelection = inputController.syncExternalSelection({
        currentSelection: current,
        externalSnapshot: selectionSnapshot,
        sheetName,
      })
      return nextSelection ?? current
    })
    onExternalSelectionSync?.(selectionSnapshot)
  }, [inputController, onExternalSelectionSync, selectionSnapshot, setGridSelection, sheetName])
  useEffect(() => {
    inputController.syncEditingState({
      focusGrid,
      isEditingCell,
      requestAnimationFrame: window.requestAnimationFrame.bind(window),
    })
  }, [focusGrid, inputController, isEditingCell])
  const commitActiveEdit = useCallback(
    (movement?: EditMovement) => {
      const valueOverride = inputController.syncMountedEditorValue({
        editorValue,
        flushSync,
        onEditorChange,
      })
      if (valueOverride === null && !isEditingCell) {
        return
      }
      onCommitEdit(movement, valueOverride ?? undefined, {
        sheetName,
        address: formatAddress(activeSelectionCell[1], activeSelectionCell[0]),
      })
    },
    [activeSelectionCell, editorValue, inputController, isEditingCell, onCommitEdit, onEditorChange, sheetName],
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
      const nextSelectionSnapshot = inputController.noteLocalSelectionChange({
        baseSnapshot: selectionSnapshot,
        nextSelection,
        sheetName,
      })
      onSelectionChange(nextSelectionSnapshot)
    },
    [inputController, onSelectionChange, selectionSnapshot, sheetName],
  )
  const collapseSelectionForEditing = useCallback(
    (cell: Item) => {
      const currentSelectionCell = gridSelection.current?.cell ?? null
      const currentSelectionRange = gridSelection.current?.range ?? null
      const isAlreadySingleCell =
        gridSelection.columns.length === 0 &&
        gridSelection.rows.length === 0 &&
        currentSelectionCell !== null &&
        currentSelectionCell[0] === cell[0] &&
        currentSelectionCell[1] === cell[1] &&
        currentSelectionRange !== null &&
        currentSelectionRange.x === cell[0] &&
        currentSelectionRange.y === cell[1] &&
        currentSelectionRange.width === 1 &&
        currentSelectionRange.height === 1
      if (isAlreadySingleCell) {
        return
      }
      const nextSelection = createGridSelection(cell[0], cell[1])
      setGridSelection(nextSelection)
      emitSelectionChange(nextSelection)
    },
    [emitSelectionChange, gridSelection, setGridSelection],
  )
  const beginSelectedEdit = useCallback(
    (seed?: string, selectionBehavior: EditSelectionBehavior = 'caret-end') => {
      const address = formatAddress(activeSelectionCell[1], activeSelectionCell[0])
      collapseSelectionForEditing(activeSelectionCell)
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
    [activeSelectionCell, collapseSelectionForEditing, engine, getCellEditorSeed, onBeginEdit, sheetName],
  )
  const beginEditAt = useCallback(
    (addr: string, seed?: string, selectionBehavior: EditSelectionBehavior = 'caret-end') => {
      const targetCell = parseCellAddress(addr, sheetName)
      collapseSelectionForEditing([targetCell.col, targetCell.row])
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
    [collapseSelectionForEditing, engine, getCellEditorSeed, onBeginEdit, sheetName],
  )
  const allowsRangeMove = Boolean(
    selectionRange && gridSelection.columns.length === 0 && gridSelection.rows.length === 0 && !fillPreviewRange && !isFillHandleDragging,
  )
  const getCurrentGridSelection = useCallback(() => gridRuntimeHost.interactionOverlays.snapshot().gridSelection, [gridRuntimeHost])
  const applyClipboardValues = useCallback(
    (target: Item, values: readonly (readonly string[])[], options?: { readonly pasteValuesOnly?: boolean | undefined }) => {
      applyGridClipboardValues({
        internalClipboardRef,
        onCopyRange,
        onMoveRange,
        onPaste,
        pasteValuesOnly: options?.pasteValuesOnly,
        sheetName,
        target,
        values,
      })
    },
    [internalClipboardRef, onCopyRange, onMoveRange, onPaste, sheetName],
  )
  const captureInternalClipboardSelection = useCallback(
    (operation?: 'copy' | 'cut') => {
      return captureGridClipboardSelection({
        engine,
        getCellEditorSeed,
        gridSelection: getCurrentGridSelection(),
        internalClipboardRef,
        operation: operation ?? 'copy',
        sheetName,
      })
    },
    [engine, getCellEditorSeed, getCurrentGridSelection, internalClipboardRef, sheetName],
  )
  const scrollActiveCellIntoView = useCallback(() => {
    gridRuntimeHost.viewportScroll.autoScrollSelectionIntoView({ cell: activeSelectionCell, force: true })
    focusGrid({ force: true })
  }, [activeSelectionCell, focusGrid, gridRuntimeHost])
  const { handleGridKey } = useWorkbookGridKeyboardHandler({
    applyClipboardValues,
    beginSelectedEdit,
    captureInternalClipboardSelection,
    editorValue,
    engine,
    getVisibleRegion,
    gridSelection,
    getGridSelection: getCurrentGridSelection,
    hostRef,
    internalClipboardRef,
    isEditingCell,
    onCancelEdit,
    onClearCell,
    onCommitEdit,
    onDeleteColumns,
    onDeleteRows,
    onEditorChange,
    onFillRange,
    onSelectionChange: emitSelectionChange,
    scrollActiveCellIntoView,
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
    getGridSelection: () => gridRuntimeHost.interactionOverlays.snapshot().gridSelection,
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
    gridSelection,
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
  const pointerHandlers = useWorkbookGridHostPointerHandlers({
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
    renderState,
  })

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
    ...pointerHandlers,
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
    handleHostCutCapture: (event: ReactClipboardEvent<HTMLDivElement>) => {
      handleGridCopyCapture({
        captureInternalClipboardSelection,
        event,
        internalClipboardRef,
        operation: 'cut',
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
        gridSelection: getCurrentGridSelection(),
        pendingKeyboardPasteSequenceRef,
        selectedCell,
        suppressNextNativePasteRef,
      })
    },
    handleSelectEntireSheet,
    contextMenu,
  }
}
