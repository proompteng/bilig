import { useMemo } from 'react'
import { parseCellAddress } from '@bilig/formula'
import type { CellSnapshot, Viewport } from '@bilig/protocol'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import { getGridMetrics } from './gridMetrics.js'
import { getGridTheme } from './gridPresentation.js'
import type { GridEngineLike } from './grid-engine.js'
import type { GridRenderTileSource } from './renderer-v3/render-tile-source.js'
import { useWorkbookGridAxisRuntime } from './useWorkbookGridAxisRuntime.js'
import { useWorkbookGridEditorRuntime } from './useWorkbookGridEditorRuntime.js'
import { useWorkbookGridDrawRuntime } from './useWorkbookGridDrawRuntime.js'
import { useWorkbookGridGeometryRuntime } from './useWorkbookGridGeometryRuntime.js'
import { useWorkbookGridHostRuntime } from './useWorkbookGridHostRuntime.js'
import { useWorkbookGridInteractionRuntime } from './useWorkbookGridInteractionRuntime.js'

export function useWorkbookGridRenderState(input: {
  engine: GridEngineLike
  sheetName: string
  selectedAddr: string
  selectedCellSnapshot: CellSnapshot
  editorValue: string
  isEditingCell: boolean
  sheetId?: number | undefined
  sheetOrdinal?: number | undefined
  renderTileSource?: GridRenderTileSource | undefined
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
    sheetOrdinal,
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
  const selectedCell = useMemo(() => parseCellAddress(selectedAddr, sheetName), [selectedAddr, sheetName])
  const selectedItem = useMemo(() => [selectedCell.col, selectedCell.row] as const, [selectedCell.col, selectedCell.row])
  const gridMetrics = useMemo(() => getGridMetrics(), [])
  const gridTheme = useMemo(() => getGridTheme(), [])
  const {
    focusTargetRef,
    getVisibleRegion,
    handleHostRef,
    hostClientHeight,
    hostClientWidth,
    hostElement,
    hostRef,
    liveVisibleRegionRef,
    scrollViewportRef,
    setVisibleRegion,
    visibleRegion,
  } = useWorkbookGridHostRuntime({ freezeCols, freezeRows })
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
  } = useWorkbookGridAxisRuntime({
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
  } = useWorkbookGridInteractionRuntime({
    activeResizeColumn,
    activeResizeRow,
    getCellLocalBounds,
    gridRuntimeHost,
    hasColumnResizePreview,
    hasRowResizePreview,
    isEditingCell,
    selectedCol: selectedCell.col,
    selectedRow: selectedCell.row,
    visibleRange: visibleRegion.range,
  })

  const { headerPanes, preloadDataPanes, renderTilePanes, residentDataPanes } = useWorkbookGridDrawRuntime({
    columnAxis,
    columnWidths,
    engine,
    freezeCols,
    freezeRows,
    frozenColumnWidth,
    frozenRowHeight,
    gridCameraStore,
    gridMetrics,
    gridRuntimeHost,
    hostClientHeight,
    hostClientWidth,
    hostElement,
    liveVisibleRegionRef,
    onVisibleViewportChange,
    renderTileSource,
    requiresLiveViewportState,
    rowAxis,
    rowHeights,
    scrollTransformRef,
    scrollTransformStore,
    scrollViewportRef,
    selectedCell: selectedItem,
    setVisibleRegion,
    sheetId,
    sheetOrdinal,
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    syncRuntimeAxes,
    visibleRegion,
    restoreViewportTarget,
  })

  const { computeAutofitColumnWidth, editorPresentation, editorTextAlign, focusGrid, overlayStyle } = useWorkbookGridEditorRuntime({
    editorFontSize: gridTheme.editorFontSize,
    editorValue,
    engine,
    focusTargetRef,
    freezeRows,
    getCellEditorSeed,
    getCellLocalBounds,
    getVisibleRegion,
    gridCameraStore,
    gridMetrics,
    headerFontStyle: gridTheme.headerFontStyle,
    hostElement,
    hostRef,
    isEditingCell,
    scrollTransformStore,
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
    headerPanes,
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
