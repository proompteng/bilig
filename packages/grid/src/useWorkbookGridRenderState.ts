import { useCallback, useMemo, useRef, useState } from 'react'
import { parseCellAddress } from '@bilig/formula'
import type { CellSnapshot, Viewport } from '@bilig/protocol'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import { getGridMetrics } from './gridMetrics.js'
import type { VisibleRegionState } from './gridPointer.js'
import { getGridTheme } from './gridPresentation.js'
import type { GridEngineLike } from './grid-engine.js'
import type { SheetGridViewportSubscription } from './workbookGridSurfaceTypes.js'
import type { GridRenderTileSource } from './renderer-v3/render-tile-source.js'
import { useGridElementSize } from './useGridElementSize.js'
import { useWorkbookHeaderPanes } from './useWorkbookHeaderPanes.js'
import { useWorkbookRenderTilePanes } from './useWorkbookRenderTilePanes.js'
import { useWorkbookEditorOverlayAnchor } from './useWorkbookEditorOverlayAnchor.js'
import { useWorkbookAxisResizeState } from './useWorkbookAxisResizeState.js'
import { useWorkbookInteractionOverlayState } from './useWorkbookInteractionOverlayState.js'
import { useWorkbookColumnAutofit } from './useWorkbookColumnAutofit.js'
import { useWorkbookViewportResidencyState } from './useWorkbookViewportResidencyState.js'
import { useWorkbookViewportScrollRuntime } from './useWorkbookViewportScrollRuntime.js'
import { useWorkbookGridGeometryRuntime } from './useWorkbookGridGeometryRuntime.js'
import { useWorkbookHeaderCellBounds } from './useWorkbookHeaderCellBounds.js'

export function useWorkbookGridRenderState(input: {
  engine: GridEngineLike
  sheetName: string
  selectedAddr: string
  selectedCellSnapshot: CellSnapshot
  editorValue: string
  isEditingCell: boolean
  sheetId?: number | undefined
  renderTileSource?: GridRenderTileSource | undefined
  subscribeViewport?: SheetGridViewportSubscription | undefined
  controlledColumnWidths?: Readonly<Record<number, number>> | undefined
  controlledRowHeights?: Readonly<Record<number, number>> | undefined
  controlledHiddenColumns?: Readonly<Record<number, true>> | undefined
  controlledHiddenRows?: Readonly<Record<number, true>> | undefined
  getCellEditorSeed?: ((sheetName: string, address: string) => string | undefined) | undefined
  freezeRows?: number | undefined
  freezeCols?: number | undefined
  onVisibleViewportChange?: ((viewport: Viewport) => void) | undefined
  onColumnWidthChange?: ((columnIndex: number, newSize: number) => void) | undefined
  onRowHeightChange?: ((rowIndex: number, newSize: number) => void) | undefined
  restoreViewportTarget?:
    | {
        readonly token: number
        readonly viewport: Viewport
      }
    | undefined
}) {
  const {
    engine,
    sheetName,
    selectedAddr,
    selectedCellSnapshot,
    editorValue,
    isEditingCell,
    sheetId,
    renderTileSource,
    controlledColumnWidths,
    controlledRowHeights,
    controlledHiddenColumns,
    controlledHiddenRows,
    getCellEditorSeed,
    freezeRows: requestedFreezeRows = 0,
    freezeCols: requestedFreezeCols = 0,
    onVisibleViewportChange,
    onColumnWidthChange,
    onRowHeightChange,
    restoreViewportTarget,
  } = input
  const freezeRows = Math.max(0, Math.min(MAX_ROWS, requestedFreezeRows))
  const freezeCols = Math.max(0, Math.min(MAX_COLS, requestedFreezeCols))
  const hostRef = useRef<HTMLDivElement | null>(null)
  const focusTargetRef = useRef<HTMLDivElement | null>(null)
  const scrollViewportRef = useRef<HTMLDivElement | null>(null)
  const [hostElement, setHostElement] = useState<HTMLDivElement | null>(null)
  const [visibleRegion, setVisibleRegion] = useState<VisibleRegionState>({
    range: { x: 0, y: 0, width: 12, height: 24 },
    tx: 0,
    ty: 0,
    freezeRows,
    freezeCols,
  })
  const selectedCell = useMemo(() => parseCellAddress(selectedAddr, sheetName), [selectedAddr, sheetName])
  const selectedItem = useMemo(() => [selectedCell.col, selectedCell.row] as const, [selectedCell.col, selectedCell.row])
  const gridMetrics = useMemo(() => getGridMetrics(), [])
  const dprBucket = typeof window === 'undefined' ? 1 : Math.max(1, Math.ceil(window.devicePixelRatio || 1))
  const shouldUseRemoteRenderTileSource = renderTileSource !== undefined && sheetId !== undefined
  const gridTheme = useMemo(() => getGridTheme(), [])
  const liveVisibleRegionRef = useRef<VisibleRegionState>(visibleRegion)
  const hostElementSize = useGridElementSize(hostElement)
  const hostClientWidth = hostElementSize.width
  const hostClientHeight = hostElementSize.height
  const {
    activeResizeColumn,
    activeResizeRow,
    clearColumnResizePreview,
    clearRowResizePreview,
    columnWidths,
    commitColumnWidth,
    commitRowHeight,
    getPreviewColumnWidth,
    getPreviewRowHeight,
    hasColumnResizePreview,
    hasRowResizePreview,
    previewColumnWidth,
    previewRowHeight,
    rowHeights,
    setActiveResizeColumn,
    setActiveResizeRow,
  } = useWorkbookAxisResizeState({
    controlledColumnWidths,
    controlledHiddenColumns,
    controlledHiddenRows,
    controlledRowHeights,
    onColumnWidthChange,
    onRowHeightChange,
    sheetName,
  })
  const {
    columnAxis,
    columnWidthOverridesAttr,
    frozenColumnWidth,
    frozenRowHeight,
    getCellLocalBounds,
    getCellScreenBounds,
    getLiveGeometrySnapshot,
    gridCameraStore,
    gridRuntimeHost,
    rowAxis,
    rowHeightOverridesAttr,
    scrollTransformRef,
    scrollTransformStore,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    syncRuntimeAxes,
    totalGridHeight,
    totalGridWidth,
  } = useWorkbookGridGeometryRuntime({
    columnWidths,
    controlledHiddenColumns,
    controlledHiddenRows,
    freezeCols,
    freezeRows,
    gridMetrics,
    hostClientHeight,
    hostClientWidth,
    hostRef,
    rowHeights,
    scrollViewportRef,
    sheetName,
  })

  const {
    activeHeaderDrag,
    fillPreviewBounds,
    fillPreviewRange,
    gridSelection,
    hoverState,
    isEntireSheetSelected,
    isFillHandleDragging,
    isRangeMoveDragging,
    requiresLiveViewportState,
    selectionRange,
    setActiveHeaderDrag,
    setFillPreviewRange,
    setGridSelection,
    setHoverState,
    setIsFillHandleDragging,
    setIsRangeMoveDragging,
  } = useWorkbookInteractionOverlayState({
    activeResizeColumn,
    activeResizeRow,
    getCellLocalBounds,
    hasColumnResizePreview,
    hasRowResizePreview,
    isEditingCell,
    selectedCol: selectedCell.col,
    selectedRow: selectedCell.row,
    visibleRange: visibleRegion.range,
  })

  const { viewport, residentViewport, renderTileViewport, residentHeaderItems, residentHeaderRegion, sceneRevision } =
    useWorkbookViewportResidencyState({
      engine,
      freezeCols,
      freezeRows,
      sheetName,
      shouldUseRemoteRenderTileSource,
      visibleRegion,
    })
  const getHeaderCellLocalBounds = useWorkbookHeaderCellBounds({
    columnWidths,
    freezeCols,
    freezeRows,
    frozenColumnWidth,
    frozenRowHeight,
    gridMetrics,
    residentViewport,
    rowHeights,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
  })
  useWorkbookViewportScrollRuntime({
    columnAxis,
    freezeCols,
    freezeRows,
    gridCameraStore,
    gridRuntimeHost,
    gridMetrics,
    hostElement,
    liveVisibleRegionRef,
    onVisibleViewportChange,
    requiresLiveViewportState,
    rowAxis,
    scrollTransformRef,
    scrollTransformStore,
    scrollViewportRef,
    selectedCell: selectedItem,
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    syncRuntimeAxes,
    viewport,
    restoreViewportTarget,
    setVisibleRegion,
  })

  const { preloadDataPanes, renderTilePanes, residentBodyPane, residentDataPanes } = useWorkbookRenderTilePanes({
    columnWidths,
    dprBucket,
    engine,
    freezeCols,
    freezeRows,
    frozenColumnWidth,
    frozenRowHeight,
    gridMetrics,
    gridRuntimeHost,
    hostClientHeight,
    hostClientWidth,
    hostElement,
    renderTileSource,
    renderTileViewport,
    residentViewport,
    rowHeights,
    sceneRevision,
    sheetId,
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    visibleViewport: viewport,
  })

  const renderHeaderPanes = useWorkbookHeaderPanes({
    columnWidths,
    freezeCols,
    freezeRows,
    frozenColumnWidth,
    frozenRowHeight,
    getHeaderCellLocalBounds,
    gridMetrics,
    hostClientHeight,
    hostClientWidth,
    hostElement,
    residentBodyPane,
    residentHeaderItems,
    residentHeaderRegion,
    residentViewport,
    rowHeights,
    sheetName,
  })

  const { editorPresentation, editorTextAlign, overlayStyle } = useWorkbookEditorOverlayAnchor({
    editorValue,
    engine,
    getCellLocalBounds,
    gridCameraStore,
    hostElement,
    isEditingCell,
    scrollTransformStore,
    selectedCellSnapshot,
    selectedCol: selectedCell.col,
    selectedRow: selectedCell.row,
  })

  const focusGrid = useCallback(() => {
    const activeElement = typeof document === 'undefined' ? null : document.activeElement
    if (
      (activeElement instanceof HTMLInputElement || activeElement instanceof HTMLTextAreaElement) &&
      activeElement.dataset['testid'] === 'cell-editor-input'
    ) {
      return
    }
    const focusTarget = focusTargetRef.current
    if (focusTarget) {
      focusTarget.focus({ preventScroll: true })
      return
    }
    hostRef.current?.focus({ preventScroll: true })
  }, [])

  const handleHostRef = useCallback((node: HTMLDivElement | null) => {
    hostRef.current = node
    setHostElement(node)
  }, [])
  const getVisibleRegion = useCallback(() => liveVisibleRegionRef.current, [])
  const computeAutofitColumnWidth = useWorkbookColumnAutofit({
    editorFontSize: gridTheme.editorFontSize,
    engine,
    freezeRows,
    getCellEditorSeed,
    getVisibleRegion,
    gridMetrics,
    headerFontStyle: gridTheme.headerFontStyle,
    selectedCell,
    selectedCellSnapshot,
    sheetName,
  })

  return {
    activeHeaderDrag,
    activeResizeColumn,
    activeResizeRow,
    clearColumnResizePreview,
    clearRowResizePreview,
    columnWidths,
    columnWidthOverridesAttr,
    commitColumnWidth,
    commitRowHeight,
    computeAutofitColumnWidth,
    fillPreviewBounds,
    fillPreviewRange,
    focusGrid,
    focusTargetRef,
    getCellLocalBounds,
    getCellScreenBounds,
    getLiveGeometrySnapshot,
    getVisibleRegion,
    getPreviewColumnWidth,
    getPreviewRowHeight,
    gridMetrics,
    gridSelection,
    gridCameraStore,
    gridTheme,
    handleHostRef,
    headerPanes: renderHeaderPanes,
    hostElement,
    hostRef,
    hoverState,
    isEntireSheetSelected,
    isFillHandleDragging,
    isRangeMoveDragging,
    overlayStyle,
    previewColumnWidth,
    previewRowHeight,
    preloadDataPanes,
    renderTilePanes,
    residentDataPanes,
    rowHeights,
    rowHeightOverridesAttr,
    scrollTransformStore,
    scrollViewportRef,
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
    totalGridHeight,
    totalGridWidth,
    visibleRegion,
    editorPresentation,
    editorTextAlign,
  }
}
