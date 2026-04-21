import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react'
import { formatAddress, indexToColumn, parseCellAddress } from '@bilig/formula'
import type { CellSnapshot, Viewport } from '@bilig/protocol'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import { buildGridGpuScene, type GridGpuScene } from './gridGpuScene.js'
import { buildGridTextScene, type GridTextScene } from './gridTextScene.js'
import { createGridCameraSnapshot } from './gridCamera.js'
import { resolveGridTileResidencyV2, type GridTileKey } from './gridTileResidencyV2.js'
import {
  EMPTY_COLUMN_WIDTHS,
  EMPTY_ROW_HEIGHTS,
  MAX_ROW_HEIGHT,
  MAX_COLUMN_WIDTH,
  MIN_ROW_HEIGHT,
  MIN_COLUMN_WIDTH,
  getGridMetrics,
  getResolvedColumnWidth,
  getResolvedRowHeight,
  resolveRowOffset,
} from './gridMetrics.js'
import { createGridSelection, isSheetSelection } from './gridSelection.js'
import { resolveFillHandlePreviewBounds } from './gridFillHandle.js'
import { buildHeaderPaneStates } from './gridHeaderPanes.js'
import { createGridAxisWorldIndexFromRecords } from './gridAxisWorldIndex.js'
import { createGridGeometrySnapshotFromAxes } from './gridGeometry.js'
import { applyHiddenAxisSizes, resolveGridScrollSpacerSize } from './gridScrollSurface.js'
import type { HeaderSelection, VisibleRegionState } from './gridPointer.js'
import type { GridHoverState } from './gridHover.js'
import { getResolvedCellFontFamily, snapshotToRenderCell } from './gridCells.js'
import { getEditorPresentation, getEditorTextAlign, getGridTheme, getOverlayStyle } from './gridPresentation.js'
import type { GridEngineLike } from './grid-engine.js'
import type { GridSelection, Rectangle } from './gridTypes.js'
import type { SheetGridViewportSubscription } from './workbookGridSurfaceTypes.js'
import { collectViewportItems } from './gridViewportItems.js'
import { buildResidentDataPaneScenes, resolveResidentDataPaneRenderState } from './gridResidentDataLayer.js'
import type { WorkbookPaneScenePacket, WorkbookPaneSceneRequest, WorkbookRenderPaneState } from './renderer/pane-scene-types.js'
import type { GridCameraSnapshot } from './renderer/grid-render-contract.js'
import {
  hasSelectionTargetChanged,
  resolveColumnOffset,
  resolveResidentViewport,
  resolveViewportScrollPosition,
  scrollCellIntoView,
} from './workbookGridViewport.js'
import { WorkbookGridScrollStore } from './workbookGridScrollStore.js'
import { noteGridScrollInput } from './renderer/grid-render-counters.js'
import { GridCameraStore } from './renderer-v2/gridCameraStore.js'
import { visibleRegionFromCamera, viewportFromVisibleRegion } from './useGridCameraState.js'
import { sameBounds } from './useGridOverlayState.js'
import { resolveResizeGuideColumn, resolveResizeGuideRow } from './useGridResizeState.js'
import { canUseWorkerResidentPaneScenes, noteWorkerResidentPaneScenesApplied } from './useGridSceneResidency.js'
import { resolveRequiresLiveViewportState } from './useGridSelectionState.js'
import { collectViewportSubscriptions } from './useGridViewportSubscriptions.js'

function noteViewportSubscription(): void {
  if (typeof window === 'undefined') {
    return
  }
  ;(window as Window & { __biligScrollPerf?: { noteViewportSubscription?: () => void } }).__biligScrollPerf?.noteViewportSubscription?.()
}

function noteVisibleWindowChange(): void {
  if (typeof window === 'undefined') {
    return
  }
  ;(window as Window & { __biligScrollPerf?: { noteVisibleWindowChange?: () => void } }).__biligScrollPerf?.noteVisibleWindowChange?.()
}

function noteHeaderPaneBuild(): void {
  if (typeof window === 'undefined') {
    return
  }
  ;(window as Window & { __biligScrollPerf?: { noteHeaderPaneBuild?: () => void } }).__biligScrollPerf?.noteHeaderPaneBuild?.()
}

function sameVisibleRegionWindow(left: VisibleRegionState, right: VisibleRegionState): boolean {
  return (
    (left.freezeCols ?? 0) === (right.freezeCols ?? 0) &&
    (left.freezeRows ?? 0) === (right.freezeRows ?? 0) &&
    left.range.x === right.range.x &&
    left.range.y === right.range.y &&
    left.range.width === right.range.width &&
    left.range.height === right.range.height
  )
}

function sameViewportBounds(left: Viewport, right: Viewport): boolean {
  return (
    left.rowStart === right.rowStart && left.rowEnd === right.rowEnd && left.colStart === right.colStart && left.colEnd === right.colEnd
  )
}

function tileKeyToViewport(tile: GridTileKey): Viewport {
  return {
    colEnd: tile.colEnd,
    colStart: tile.colStart,
    rowEnd: tile.rowEnd,
    rowStart: tile.rowStart,
  }
}

interface ResidentPaneSceneEngine extends GridEngineLike {
  subscribeResidentPaneScenes(request: WorkbookPaneSceneRequest, listener: () => void): () => void
  peekResidentPaneScenes(request: WorkbookPaneSceneRequest): readonly WorkbookPaneScenePacket[] | null
}

function supportsResidentPaneScenes(engine: GridEngineLike): engine is ResidentPaneSceneEngine {
  return 'subscribeResidentPaneScenes' in engine && 'peekResidentPaneScenes' in engine
}

const EMPTY_PANE_SCENES: readonly WorkbookPaneScenePacket[] = Object.freeze([])

export function useWorkbookGridRenderState(input: {
  engine: GridEngineLike
  sheetName: string
  selectedAddr: string
  selectedCellSnapshot: CellSnapshot
  editorValue: string
  isEditingCell: boolean
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
    subscribeViewport,
    controlledColumnWidths,
    controlledRowHeights,
    controlledHiddenColumns,
    controlledHiddenRows,
    getCellEditorSeed,
    freezeRows: requestedFreezeRows = 0,
    freezeCols: requestedFreezeCols = 0,
    onVisibleViewportChange,
    restoreViewportTarget,
  } = input
  const freezeRows = Math.max(0, Math.min(MAX_ROWS, requestedFreezeRows))
  const freezeCols = Math.max(0, Math.min(MAX_COLS, requestedFreezeCols))
  const emptyGpuScene = useMemo<GridGpuScene>(() => ({ fillRects: [], borderRects: [] }), [])
  const emptyTextScene = useMemo<GridTextScene>(() => ({ items: [] }), [])
  const hostRef = useRef<HTMLDivElement | null>(null)
  const focusTargetRef = useRef<HTMLDivElement | null>(null)
  const scrollViewportRef = useRef<HTMLDivElement | null>(null)
  const autoScrollSelectionRef = useRef<{ sheetName: string; col: number; row: number } | null>(null)
  const restoredViewportTokenRef = useRef<number | null>(null)
  const textMeasureCanvasRef = useRef<HTMLCanvasElement | null>(null)
  const scrollSyncFrameRef = useRef<number | null>(null)
  const scrollTransformStoreRef = useRef<WorkbookGridScrollStore>(new WorkbookGridScrollStore())
  const scrollTransformRef = useRef(scrollTransformStoreRef.current.getSnapshot())
  const gridCameraStoreRef = useRef<GridCameraStore>(new GridCameraStore())
  const gridCameraRef = useRef<GridCameraSnapshot | null>(null)
  const [sceneRevision, setSceneRevision] = useState(0)
  const [fillPreviewRange, setFillPreviewRange] = useState<Rectangle | null>(null)
  const [isFillHandleDragging, setIsFillHandleDragging] = useState(false)
  const [isRangeMoveDragging, setIsRangeMoveDragging] = useState(false)
  const [hoverState, setHoverState] = useState<GridHoverState>({
    cell: null,
    header: null,
    cursor: 'default',
  })
  const [activeResizeColumn, setActiveResizeColumn] = useState<number | null>(null)
  const [activeResizeRow, setActiveResizeRow] = useState<number | null>(null)
  const [activeHeaderDrag, setActiveHeaderDrag] = useState<HeaderSelection | null>(null)
  const [hostElement, setHostElement] = useState<HTMLDivElement | null>(null)
  const [visibleRegion, setVisibleRegion] = useState<VisibleRegionState>({
    range: { x: 0, y: 0, width: 12, height: 24 },
    tx: 0,
    ty: 0,
    freezeRows,
    freezeCols,
  })
  const [overlayBounds, setOverlayBounds] = useState<Rectangle | undefined>(undefined)
  const [columnWidthsBySheet, setColumnWidthsBySheet] = useState<Record<string, Record<number, number>>>({})
  const [rowHeightsBySheet, setRowHeightsBySheet] = useState<Record<string, Record<number, number>>>({})
  const [columnResizePreview, setColumnResizePreview] = useState<{
    sheetName: string
    columnIndex: number
    width: number
  } | null>(null)
  const [rowResizePreview, setRowResizePreview] = useState<{
    sheetName: string
    rowIndex: number
    height: number
  } | null>(null)
  const selectedCell = useMemo(() => parseCellAddress(selectedAddr, sheetName), [selectedAddr, sheetName])
  const [gridSelection, setGridSelection] = useState<GridSelection>(() => createGridSelection(selectedCell.col, selectedCell.row))
  const gridMetrics = useMemo(() => getGridMetrics(), [])
  const gridTheme = useMemo(() => getGridTheme(), [])
  const columnResizePreviewRef = useRef<{
    sheetName: string
    columnIndex: number
    width: number
  } | null>(null)
  const rowResizePreviewRef = useRef<{
    sheetName: string
    rowIndex: number
    height: number
  } | null>(null)
  const liveVisibleRegionRef = useRef<VisibleRegionState>(visibleRegion)
  const hostClientWidth = hostElement?.clientWidth ?? 0
  const hostClientHeight = hostElement?.clientHeight ?? 0
  const baseColumnWidths = controlledColumnWidths ?? columnWidthsBySheet[sheetName] ?? EMPTY_COLUMN_WIDTHS
  const baseRowHeights = controlledRowHeights ?? rowHeightsBySheet[sheetName] ?? EMPTY_ROW_HEIGHTS
  const sizedColumnWidths = useMemo(() => {
    if (!columnResizePreview || columnResizePreview.sheetName !== sheetName) {
      return baseColumnWidths
    }
    if (baseColumnWidths[columnResizePreview.columnIndex] === columnResizePreview.width) {
      return baseColumnWidths
    }
    return {
      ...baseColumnWidths,
      [columnResizePreview.columnIndex]: columnResizePreview.width,
    }
  }, [baseColumnWidths, columnResizePreview, sheetName])
  const sizedRowHeights = useMemo(() => {
    if (!rowResizePreview || rowResizePreview.sheetName !== sheetName) {
      return baseRowHeights
    }
    if (baseRowHeights[rowResizePreview.rowIndex] === rowResizePreview.height) {
      return baseRowHeights
    }
    return {
      ...baseRowHeights,
      [rowResizePreview.rowIndex]: rowResizePreview.height,
    }
  }, [baseRowHeights, rowResizePreview, sheetName])
  const columnWidths = useMemo(
    () => applyHiddenAxisSizes(sizedColumnWidths, controlledHiddenColumns),
    [controlledHiddenColumns, sizedColumnWidths],
  )
  const rowHeights = useMemo(() => applyHiddenAxisSizes(sizedRowHeights, controlledHiddenRows), [controlledHiddenRows, sizedRowHeights])
  const sortedColumnWidthOverrides = useMemo(
    () =>
      Object.entries(columnWidths)
        .map(([index, width]) => [Number(index), width] as const)
        .toSorted((left, right) => left[0] - right[0]),
    [columnWidths],
  )
  const sortedRowHeightOverrides = useMemo(
    () =>
      Object.entries(rowHeights)
        .map(([index, height]) => [Number(index), height] as const)
        .toSorted((left, right) => left[0] - right[0]),
    [rowHeights],
  )
  const columnWidthOverridesAttr = useMemo(() => {
    const entries = Object.entries(columnWidths).toSorted(([left], [right]) => Number(left) - Number(right))
    return entries.length === 0 ? '{}' : JSON.stringify(Object.fromEntries(entries))
  }, [columnWidths])
  const columnAxis = useMemo(
    () =>
      createGridAxisWorldIndexFromRecords({
        axisLength: MAX_COLS,
        defaultSize: gridMetrics.columnWidth,
        hidden: controlledHiddenColumns,
        sizes: columnWidths,
      }),
    [columnWidths, controlledHiddenColumns, gridMetrics.columnWidth],
  )
  const rowAxis = useMemo(
    () =>
      createGridAxisWorldIndexFromRecords({
        axisLength: MAX_ROWS,
        defaultSize: gridMetrics.rowHeight,
        hidden: controlledHiddenRows,
        sizes: rowHeights,
      }),
    [controlledHiddenRows, gridMetrics.rowHeight, rowHeights],
  )
  const frozenColumnWidth = useMemo(() => columnAxis.span(0, freezeCols), [columnAxis, freezeCols])
  const frozenRowHeight = useMemo(() => rowAxis.span(0, freezeRows), [freezeRows, rowAxis])
  const scrollSpacerSize = useMemo(
    () =>
      resolveGridScrollSpacerSize({
        columnAxis,
        rowAxis,
        frozenColumnWidth,
        frozenRowHeight,
        hostWidth: hostClientWidth,
        hostHeight: hostClientHeight,
        gridMetrics,
      }),
    [columnAxis, frozenColumnWidth, frozenRowHeight, gridMetrics, hostClientHeight, hostClientWidth, rowAxis],
  )
  const totalGridWidth = scrollSpacerSize.width
  const totalGridHeight = scrollSpacerSize.height
  const selectionRange = gridSelection.current?.range ?? null
  const scrollTransformStore = scrollTransformStoreRef.current
  const gridCameraStore = gridCameraStoreRef.current

  const getCellLocalBounds = useCallback(
    (col: number, row: number): Rectangle | undefined => {
      if (col < 0 || col >= MAX_COLS || row < 0 || row >= MAX_ROWS) {
        return undefined
      }
      const geometryRect = gridCameraStore.getSnapshot()?.cellScreenRect(col, row)
      if (geometryRect) {
        return geometryRect
      }
      const width = columnAxis.sizeOf(col)
      const height = rowAxis.sizeOf(row)
      if (width <= 0 || height <= 0 || columnAxis.isHidden(col) || rowAxis.isHidden(row)) {
        return undefined
      }
      const scrollTransform = scrollTransformRef.current
      const scrollLeft = scrollTransform.scrollLeft ?? 0
      const scrollTop = scrollTransform.scrollTop ?? 0
      const worldX = columnAxis.offsetOf(col)
      const worldY = rowAxis.offsetOf(row)
      return {
        x: col < freezeCols ? gridMetrics.rowMarkerWidth + worldX : gridMetrics.rowMarkerWidth + worldX - scrollLeft,
        y: row < freezeRows ? gridMetrics.headerHeight + worldY : gridMetrics.headerHeight + worldY - scrollTop,
        width,
        height,
      }
    },
    [columnAxis, freezeCols, freezeRows, gridCameraStore, gridMetrics.headerHeight, gridMetrics.rowMarkerWidth, rowAxis],
  )

  const getCellScreenBounds = useCallback(
    (col: number, row: number): Rectangle | undefined => {
      const hostBounds = hostRef.current?.getBoundingClientRect()
      const localBounds = getCellLocalBounds(col, row)
      if (!hostBounds || !localBounds) {
        return undefined
      }
      return {
        x: hostBounds.left + localBounds.x,
        y: hostBounds.top + localBounds.y,
        width: localBounds.width,
        height: localBounds.height,
      }
    },
    [getCellLocalBounds],
  )

  const viewport = useMemo<Viewport>(() => viewportFromVisibleRegion(visibleRegion), [visibleRegion])
  const residentViewportRef = useRef<Viewport>(resolveResidentViewport(viewport))
  const nextResidentViewport = resolveResidentViewport(viewport)
  if (!sameViewportBounds(residentViewportRef.current, nextResidentViewport)) {
    residentViewportRef.current = nextResidentViewport
  }
  const residentViewport = residentViewportRef.current
  const warmResidentViewports = useMemo(() => {
    const camera = gridCameraRef.current
    return resolveGridTileResidencyV2({
      velocityX: camera?.velocityX ?? 0,
      velocityY: camera?.velocityY ?? 0,
      visibleViewport: viewport,
      warmNeighbors: 1,
    }).warm.map(tileKeyToViewport)
  }, [viewport])
  const getHeaderCellLocalBounds = useCallback(
    (col: number, row: number): Rectangle | undefined => {
      if (col < 0 || col >= MAX_COLS || row < 0 || row >= MAX_ROWS) {
        return undefined
      }
      return {
        x:
          col < freezeCols
            ? gridMetrics.rowMarkerWidth + resolveColumnOffset(col, sortedColumnWidthOverrides, gridMetrics.columnWidth)
            : gridMetrics.rowMarkerWidth +
              frozenColumnWidth +
              resolveColumnOffset(col, sortedColumnWidthOverrides, gridMetrics.columnWidth) -
              resolveColumnOffset(residentViewport.colStart, sortedColumnWidthOverrides, gridMetrics.columnWidth),
        y:
          row < freezeRows
            ? gridMetrics.headerHeight + resolveRowOffset(row, sortedRowHeightOverrides, gridMetrics.rowHeight)
            : gridMetrics.headerHeight +
              frozenRowHeight +
              resolveRowOffset(row, sortedRowHeightOverrides, gridMetrics.rowHeight) -
              resolveRowOffset(residentViewport.rowStart, sortedRowHeightOverrides, gridMetrics.rowHeight),
        width: getResolvedColumnWidth(columnWidths, col, gridMetrics.columnWidth),
        height: getResolvedRowHeight(rowHeights, row, gridMetrics.rowHeight),
      }
    },
    [
      columnWidths,
      freezeCols,
      freezeRows,
      frozenColumnWidth,
      frozenRowHeight,
      gridMetrics.columnWidth,
      gridMetrics.headerHeight,
      gridMetrics.rowHeight,
      gridMetrics.rowMarkerWidth,
      residentViewport.colStart,
      residentViewport.rowStart,
      rowHeights,
      sortedColumnWidthOverrides,
      sortedRowHeightOverrides,
    ],
  )
  const residentViewports = useMemo(
    () =>
      collectViewportSubscriptions({
        freezeCols,
        freezeRows,
        viewport: residentViewport,
        warmViewports: warmResidentViewports,
      }),
    [freezeCols, freezeRows, residentViewport, warmResidentViewports],
  )
  const visibleItems = useMemo(() => {
    return collectViewportItems(residentViewport, { freezeRows, freezeCols })
  }, [freezeCols, freezeRows, residentViewport])
  const residentHeaderItems = useMemo(() => {
    return collectViewportItems(residentViewport, { freezeRows, freezeCols })
  }, [freezeCols, freezeRows, residentViewport])
  const visibleAddresses = useMemo(() => visibleItems.map(([col, row]) => formatAddress(row, col)), [visibleItems])
  const residentHeaderRegion = useMemo(
    () => ({
      range: {
        x: residentViewport.colStart,
        y: residentViewport.rowStart,
        width: residentViewport.colEnd - residentViewport.colStart + 1,
        height: residentViewport.rowEnd - residentViewport.rowStart + 1,
      },
      tx: 0,
      ty: 0,
      freezeRows,
      freezeCols,
    }),
    [freezeCols, freezeRows, residentViewport.colEnd, residentViewport.colStart, residentViewport.rowEnd, residentViewport.rowStart],
  )

  const invalidateScene = useCallback(() => {
    setSceneRevision((current) => current + 1)
  }, [])

  const requiresLiveViewportState = resolveRequiresLiveViewportState({
    fillPreviewActive: fillPreviewRange !== null,
    hasActiveHeaderDrag: activeHeaderDrag !== null,
    hasActiveResizeColumn: activeResizeColumn !== null,
    hasActiveResizeRow: activeResizeRow !== null,
    hasColumnResizePreview: columnResizePreview !== null,
    hasRowResizePreview: rowResizePreview !== null,
    isEditingCell,
    isFillHandleDragging,
    isRangeMoveDragging,
  })

  const syncVisibleRegion = useCallback(() => {
    const scrollViewport = scrollViewportRef.current
    if (!scrollViewport) {
      return
    }
    const camera = createGridCameraSnapshot({
      scrollLeft: scrollViewport.scrollLeft,
      scrollTop: scrollViewport.scrollTop,
      viewportWidth: scrollViewport.clientWidth,
      viewportHeight: scrollViewport.clientHeight,
      dpr: window.devicePixelRatio || 1,
      freezeRows,
      freezeCols,
      columnWidths,
      rowHeights,
      hiddenColumns: controlledHiddenColumns,
      hiddenRows: controlledHiddenRows,
      gridMetrics,
      previous: gridCameraRef.current,
    })
    gridCameraRef.current = camera
    gridCameraStore.setSnapshot(
      createGridGeometrySnapshotFromAxes({
        columns: columnAxis,
        dpr: window.devicePixelRatio || 1,
        freezeCols,
        freezeRows,
        gridMetrics,
        hostHeight: scrollViewport.clientHeight,
        hostWidth: scrollViewport.clientWidth,
        previousCamera: gridCameraStore.getSnapshot()?.camera ?? null,
        rows: rowAxis,
        scrollLeft: scrollViewport.scrollLeft,
        scrollTop: scrollViewport.scrollTop,
        sheetName,
      }),
    )
    const next = visibleRegionFromCamera({ camera, freezeCols, freezeRows })
    const renderTx =
      resolveColumnOffset(next.range.x, sortedColumnWidthOverrides, gridMetrics.columnWidth) -
      resolveColumnOffset(viewport.colStart, sortedColumnWidthOverrides, gridMetrics.columnWidth) +
      next.tx
    const renderTy =
      resolveRowOffset(next.range.y, sortedRowHeightOverrides, gridMetrics.rowHeight) -
      resolveRowOffset(viewport.rowStart, sortedRowHeightOverrides, gridMetrics.rowHeight) +
      next.ty
    scrollTransformRef.current = {
      renderTx,
      renderTy,
      scrollLeft: scrollViewport.scrollLeft,
      scrollTop: scrollViewport.scrollTop,
      tx: next.tx,
      ty: next.ty,
    }
    liveVisibleRegionRef.current = next
    scrollTransformStore.setSnapshot(scrollTransformRef.current)
    onVisibleViewportChange?.(viewportFromVisibleRegion(next))
    setVisibleRegion((current) => {
      if (requiresLiveViewportState) {
        if (sameVisibleRegionWindow(current, next)) {
          return current
        }
        noteVisibleWindowChange()
        return next
      }
      const currentResidentViewport = resolveResidentViewport(viewportFromVisibleRegion(current))
      const targetResidentViewport = resolveResidentViewport(viewportFromVisibleRegion(next))
      if (
        current.freezeCols === next.freezeCols &&
        current.freezeRows === next.freezeRows &&
        sameViewportBounds(currentResidentViewport, targetResidentViewport)
      ) {
        return current
      }
      noteVisibleWindowChange()
      return next
    })
  }, [
    columnWidths,
    columnAxis,
    freezeCols,
    freezeRows,
    gridCameraStore,
    gridMetrics,
    onVisibleViewportChange,
    controlledHiddenColumns,
    controlledHiddenRows,
    requiresLiveViewportState,
    rowHeights,
    rowAxis,
    scrollTransformStore,
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    viewport.colStart,
    viewport.rowStart,
  ])

  useEffect(() => {
    const preview = columnResizePreviewRef.current
    if (!preview || preview.sheetName !== sheetName) {
      return
    }
    if (baseColumnWidths[preview.columnIndex] !== preview.width) {
      return
    }
    columnResizePreviewRef.current = null
    setColumnResizePreview((current) =>
      current?.sheetName === preview.sheetName && current.columnIndex === preview.columnIndex && current.width === preview.width
        ? null
        : current,
    )
  }, [baseColumnWidths, sheetName])

  useEffect(() => {
    const preview = rowResizePreviewRef.current
    if (!preview || preview.sheetName !== sheetName) {
      return
    }
    if (baseRowHeights[preview.rowIndex] !== preview.height) {
      return
    }
    rowResizePreviewRef.current = null
    setRowResizePreview((current) =>
      current?.sheetName === preview.sheetName && current.rowIndex === preview.rowIndex && current.height === preview.height
        ? null
        : current,
    )
  }, [baseRowHeights, sheetName])

  useEffect(() => {
    if (!requiresLiveViewportState) {
      return
    }
    setVisibleRegion((current) => {
      const next = liveVisibleRegionRef.current
      if (sameVisibleRegionWindow(current, next)) {
        return current
      }
      return next
    })
  }, [requiresLiveViewportState])

  useEffect(() => {
    const scrollViewport = scrollViewportRef.current
    if (!scrollViewport) {
      return
    }

    syncVisibleRegion()
    const scheduleVisibleRegionSync = () => {
      if (scrollSyncFrameRef.current !== null) {
        return
      }
      scrollSyncFrameRef.current = window.requestAnimationFrame(() => {
        scrollSyncFrameRef.current = null
        syncVisibleRegion()
      })
    }
    const handleScroll = () => {
      noteGridScrollInput()
      syncVisibleRegion()
    }
    scrollViewport.addEventListener('scroll', handleScroll, { passive: true })
    const observer = new ResizeObserver(() => {
      scheduleVisibleRegionSync()
    })
    observer.observe(scrollViewport)
    return () => {
      if (scrollSyncFrameRef.current !== null) {
        window.cancelAnimationFrame(scrollSyncFrameRef.current)
        scrollSyncFrameRef.current = null
      }
      observer.disconnect()
      scrollViewport.removeEventListener('scroll', handleScroll)
    }
  }, [hostElement, syncVisibleRegion])

  useEffect(() => {
    if (!subscribeViewport) {
      return
    }
    noteViewportSubscription()
    const cleanups = residentViewports.map((nextViewport) => subscribeViewport(sheetName, nextViewport, invalidateScene))
    return () => {
      cleanups.forEach((cleanup) => cleanup())
    }
  }, [invalidateScene, residentViewports, sheetName, subscribeViewport])

  useEffect(() => {
    if (subscribeViewport) {
      return
    }
    return engine.subscribeCells(sheetName, visibleAddresses, invalidateScene)
  }, [engine, invalidateScene, sheetName, subscribeViewport, visibleAddresses])

  const resizeGuideColumn = useMemo(
    () => resolveResizeGuideColumn({ activeResizeColumn, cursor: hoverState.cursor, header: hoverState.header }),
    [activeResizeColumn, hoverState.cursor, hoverState.header],
  )
  const resizeGuideRow = useMemo(
    () => resolveResizeGuideRow({ activeResizeRow, cursor: hoverState.cursor, header: hoverState.header }),
    [activeResizeRow, hoverState.cursor, hoverState.header],
  )
  const residentPaneSceneRequest = useMemo<WorkbookPaneSceneRequest | null>(
    () =>
      hostElement
        ? {
            sheetName,
            residentViewport,
            freezeRows,
            freezeCols,
            selectedCell: {
              col: selectedCell.col,
              row: selectedCell.row,
            },
            selectionRange,
            editingCell: isEditingCell
              ? {
                  col: selectedCell.col,
                  row: selectedCell.row,
                }
              : null,
          }
        : null,
    [freezeCols, freezeRows, hostElement, isEditingCell, residentViewport, selectedCell.col, selectedCell.row, selectionRange, sheetName],
  )
  const residentSceneEngine = supportsResidentPaneScenes(engine) ? engine : null
  const workerResidentPaneScenes = useSyncExternalStore(
    useCallback(
      (listener: () => void) =>
        residentSceneEngine && residentPaneSceneRequest
          ? residentSceneEngine.subscribeResidentPaneScenes(residentPaneSceneRequest, listener)
          : () => undefined,
      [residentPaneSceneRequest, residentSceneEngine],
    ),
    () =>
      residentSceneEngine && residentPaneSceneRequest
        ? (residentSceneEngine.peekResidentPaneScenes(residentPaneSceneRequest) ?? EMPTY_PANE_SCENES)
        : EMPTY_PANE_SCENES,
    () => EMPTY_PANE_SCENES,
  )
  const canUseWorkerResidentPaneScenesResult = canUseWorkerResidentPaneScenes({
    hasActiveHeaderDrag: activeHeaderDrag !== null,
    hasHoverState: hoverState.cell !== null || hoverState.header !== null,
    requiresLiveViewportState,
    workerResidentPaneScenes,
  })
  useEffect(() => {
    if (!canUseWorkerResidentPaneScenesResult) {
      return
    }
    noteWorkerResidentPaneScenesApplied(workerResidentPaneScenes)
  }, [canUseWorkerResidentPaneScenesResult, workerResidentPaneScenes])
  const residentDataPaneScenes = useMemo(() => {
    if (!hostElement) {
      return []
    }
    if (canUseWorkerResidentPaneScenesResult) {
      return workerResidentPaneScenes
    }
    void sceneRevision
    // Render resident panes from the projected engine state so visible body
    // content stays current even when worker scene refreshes lag behind edits.
    return buildResidentDataPaneScenes({
      residentViewport,
      engine,
      sheetName,
      columnWidths,
      rowHeights,
      freezeRows,
      freezeCols,
      frozenColumnWidth,
      frozenRowHeight,
      gridMetrics,
      sortedColumnWidthOverrides,
      sortedRowHeightOverrides,
      gridSelection,
      selectedCell: [selectedCell.col, selectedCell.row],
      selectedCellSnapshot,
      selectionRange,
      editingCell: isEditingCell ? ([selectedCell.col, selectedCell.row] as const) : null,
      hoveredCell: hoverState.cell,
      hoveredHeader: hoverState.header,
      resizeGuideColumn,
      resizeGuideRow,
      activeHeaderDrag,
    })
  }, [
    activeHeaderDrag,
    canUseWorkerResidentPaneScenesResult,
    columnWidths,
    engine,
    freezeCols,
    freezeRows,
    frozenColumnWidth,
    frozenRowHeight,
    gridMetrics,
    gridSelection,
    hostElement,
    hoverState.cell,
    hoverState.header,
    isEditingCell,
    residentViewport,
    resizeGuideColumn,
    resizeGuideRow,
    rowHeights,
    sceneRevision,
    selectedCell.col,
    selectedCell.row,
    selectedCellSnapshot,
    selectionRange,
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    workerResidentPaneScenes,
  ])
  const residentDataPanes = useMemo(
    () =>
      resolveResidentDataPaneRenderState({
        panes: residentDataPaneScenes,
        residentViewport,
        visibleViewport: viewport,
        visibleRegion: {
          tx: 0,
          ty: 0,
        },
        gridMetrics,
        sortedColumnWidthOverrides,
        sortedRowHeightOverrides,
        hostWidth: hostClientWidth,
        hostHeight: hostClientHeight,
        rowMarkerWidth: gridMetrics.rowMarkerWidth,
        headerHeight: gridMetrics.headerHeight,
        frozenColumnWidth,
        frozenRowHeight,
      }),
    [
      gridMetrics,
      frozenColumnWidth,
      frozenRowHeight,
      hostClientHeight,
      hostClientWidth,
      residentDataPaneScenes,
      residentViewport,
      sortedColumnWidthOverrides,
      sortedRowHeightOverrides,
      viewport,
    ],
  )
  const residentBodyPane = residentDataPanes.find((pane) => pane.paneId === 'body') ?? null

  const headerGpuScene = useMemo<GridGpuScene>(() => {
    if (!hostElement) {
      return emptyGpuScene
    }
    void sceneRevision
    return buildGridGpuScene({
      contentMode: 'headers',
      engine,
      columnWidths,
      rowHeights,
      gridMetrics,
      gridSelection,
      activeHeaderDrag,
      hoveredCell: hoverState.cell,
      hoveredHeader: hoverState.header,
      resizeGuideColumn,
      resizeGuideRow,
      selectedCell: [selectedCell.col, selectedCell.row],
      selectionRange,
      sheetName,
      visibleItems: residentHeaderItems,
      visibleRegion: residentHeaderRegion,
      hostBounds: {
        left: 0,
        top: 0,
      },
      getCellBounds: getHeaderCellLocalBounds,
    })
  }, [
    activeHeaderDrag,
    columnWidths,
    emptyGpuScene,
    engine,
    getHeaderCellLocalBounds,
    gridMetrics,
    gridSelection,
    hostElement,
    hoverState.cell,
    hoverState.header,
    resizeGuideColumn,
    resizeGuideRow,
    rowHeights,
    sceneRevision,
    selectedCell.col,
    selectedCell.row,
    selectionRange,
    sheetName,
    residentHeaderItems,
    residentHeaderRegion,
  ])

  const headerTextScene = useMemo<GridTextScene>(() => {
    if (!hostElement) {
      return emptyTextScene
    }
    void sceneRevision
    return buildGridTextScene({
      contentMode: 'headers',
      engine,
      columnWidths,
      rowHeights,
      editingCell: isEditingCell ? ([selectedCell.col, selectedCell.row] as const) : null,
      gridMetrics,
      activeHeaderDrag,
      hoveredHeader: hoverState.header,
      resizeGuideColumn,
      selectedCell: [selectedCell.col, selectedCell.row],
      selectedCellSnapshot,
      selectionRange,
      sheetName,
      visibleItems: residentHeaderItems,
      visibleRegion: residentHeaderRegion,
      hostBounds: {
        left: 0,
        top: 0,
        width: hostClientWidth,
        height: hostClientHeight,
      },
      getCellBounds: getHeaderCellLocalBounds,
    })
  }, [
    activeHeaderDrag,
    columnWidths,
    emptyTextScene,
    engine,
    getHeaderCellLocalBounds,
    gridMetrics,
    hostElement,
    hostClientHeight,
    hostClientWidth,
    hoverState.header,
    isEditingCell,
    resizeGuideColumn,
    rowHeights,
    sceneRevision,
    selectedCellSnapshot,
    selectedCell.col,
    selectedCell.row,
    selectionRange,
    sheetName,
    residentHeaderItems,
    residentHeaderRegion,
  ])

  const headerPanes = useMemo(() => {
    noteHeaderPaneBuild()
    return buildHeaderPaneStates({
      gpuScene: headerGpuScene,
      textScene: headerTextScene,
      sheetName,
      residentViewport,
      freezeCols,
      freezeRows,
      hostWidth: hostClientWidth,
      hostHeight: hostClientHeight,
      gridMetrics,
      frozenColumnWidth,
      frozenRowHeight,
      residentBodyWidth: residentBodyPane?.surfaceSize.width ?? 0,
      residentBodyHeight: residentBodyPane?.surfaceSize.height ?? 0,
    })
  }, [
    frozenColumnWidth,
    frozenRowHeight,
    gridMetrics,
    headerGpuScene,
    headerTextScene,
    hostClientHeight,
    hostClientWidth,
    sheetName,
    residentViewport,
    freezeCols,
    freezeRows,
    residentBodyPane?.surfaceSize.height,
    residentBodyPane?.surfaceSize.width,
  ])
  const renderHeaderPanes = useMemo(
    () =>
      headerPanes.map((pane) =>
        pane.paneId === 'top-body'
          ? { ...pane, contentOffset: { x: residentBodyPane?.contentOffset.x ?? 0, y: 0 } }
          : pane.paneId === 'left-body'
            ? { ...pane, contentOffset: { x: 0, y: residentBodyPane?.contentOffset.y ?? 0 } }
            : pane,
      ),
    [headerPanes, residentBodyPane?.contentOffset.x, residentBodyPane?.contentOffset.y],
  )
  const renderPanes = useMemo<readonly WorkbookRenderPaneState[]>(
    () => [
      ...renderHeaderPanes,
      ...residentDataPanes.map((pane) => ({
        ...pane,
        scrollAxes: {
          x: pane.paneId === 'body' || pane.paneId === 'top',
          y: pane.paneId === 'body' || pane.paneId === 'left',
        },
      })),
    ],
    [renderHeaderPanes, residentDataPanes],
  )

  const fillPreviewBounds = useMemo<Rectangle | undefined>(() => {
    if (!fillPreviewRange) {
      return undefined
    }
    return resolveFillHandlePreviewBounds({
      previewRange: fillPreviewRange,
      visibleRange: visibleRegion.range,
      hostBounds: { left: 0, top: 0 },
      getCellBounds: getCellLocalBounds,
    })
  }, [fillPreviewRange, getCellLocalBounds, visibleRegion.range])

  useLayoutEffect(() => {
    const scrollViewport = scrollViewportRef.current
    if (!scrollViewport) {
      return
    }
    const previousAutoScrollSelection = autoScrollSelectionRef.current
    const nextAutoScrollSelection = {
      sheetName,
      col: selectedCell.col,
      row: selectedCell.row,
    }
    if (!hasSelectionTargetChanged(previousAutoScrollSelection, nextAutoScrollSelection)) {
      return
    }
    autoScrollSelectionRef.current = nextAutoScrollSelection
    scrollCellIntoView({
      cell: [selectedCell.col, selectedCell.row],
      freezeRows,
      freezeCols,
      columnWidths,
      rowHeights,
      gridMetrics,
      scrollViewport,
      sortedColumnWidthOverrides,
      sortedRowHeightOverrides,
    })
  }, [
    columnWidths,
    freezeCols,
    freezeRows,
    gridMetrics,
    rowHeights,
    selectedCell.col,
    selectedCell.row,
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
  ])

  useLayoutEffect(() => {
    const scrollViewport = scrollViewportRef.current
    if (!scrollViewport || !restoreViewportTarget) {
      return
    }
    if (restoredViewportTokenRef.current === restoreViewportTarget.token) {
      return
    }
    restoredViewportTokenRef.current = restoreViewportTarget.token
    const { scrollLeft: nextScrollLeft, scrollTop: nextScrollTop } = resolveViewportScrollPosition({
      viewport: restoreViewportTarget.viewport,
      freezeRows,
      freezeCols,
      sortedColumnWidthOverrides,
      sortedRowHeightOverrides,
      gridMetrics,
    })
    scrollViewport.scrollLeft = nextScrollLeft
    scrollViewport.scrollTop = nextScrollTop
  }, [freezeCols, freezeRows, gridMetrics, restoreViewportTarget, sortedColumnWidthOverrides, sortedRowHeightOverrides])

  const refreshOverlayBounds = useCallback(() => {
    const next = getCellScreenBounds(selectedCell.col, selectedCell.row)
    setOverlayBounds((current) => {
      if (!next) {
        return current
      }
      return sameBounds(current, next) ? current : next
    })
  }, [getCellScreenBounds, selectedCell.col, selectedCell.row])

  useLayoutEffect(() => {
    if (!isEditingCell) {
      setOverlayBounds(undefined)
      return
    }

    const frame = window.requestAnimationFrame(refreshOverlayBounds)
    const unsubscribeScrollTransform = scrollTransformStore.subscribe(refreshOverlayBounds)
    return () => {
      window.cancelAnimationFrame(frame)
      unsubscribeScrollTransform()
    }
  }, [isEditingCell, refreshOverlayBounds, scrollTransformStore])

  useEffect(() => {
    if (!isEditingCell) {
      return
    }
    const handleWindowResize = () => refreshOverlayBounds()
    window.addEventListener('resize', handleWindowResize)
    return () => {
      window.removeEventListener('resize', handleWindowResize)
    }
  }, [isEditingCell, refreshOverlayBounds])

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

  const commitColumnWidth = useCallback(
    (columnIndex: number, newSize: number) => {
      const clampedSize = Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.round(newSize)))
      if (input.onColumnWidthChange) {
        input.onColumnWidthChange(columnIndex, clampedSize)
        return
      }
      setColumnWidthsBySheet((current) => {
        const nextSheetWidths = current[sheetName] ?? EMPTY_COLUMN_WIDTHS
        if (nextSheetWidths[columnIndex] === clampedSize) {
          return current
        }
        return {
          ...current,
          [sheetName]: {
            ...nextSheetWidths,
            [columnIndex]: clampedSize,
          },
        }
      })
    },
    [input, sheetName],
  )

  const commitRowHeight = useCallback(
    (rowIndex: number, newSize: number) => {
      const clampedSize = Math.max(MIN_ROW_HEIGHT, Math.min(MAX_ROW_HEIGHT, Math.round(newSize)))
      if (input.onRowHeightChange) {
        input.onRowHeightChange(rowIndex, clampedSize)
        return
      }
      setRowHeightsBySheet((current) => {
        const nextSheetHeights = current[sheetName] ?? EMPTY_ROW_HEIGHTS
        if (nextSheetHeights[rowIndex] === clampedSize) {
          return current
        }
        return {
          ...current,
          [sheetName]: {
            ...nextSheetHeights,
            [rowIndex]: clampedSize,
          },
        }
      })
    },
    [input, sheetName],
  )

  const previewColumnWidth = useCallback(
    (columnIndex: number, newSize: number): number => {
      const clampedSize = Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.round(newSize)))
      const nextPreview = { sheetName, columnIndex, width: clampedSize }
      columnResizePreviewRef.current = nextPreview
      setColumnResizePreview((current) =>
        current?.sheetName === nextPreview.sheetName &&
        current.columnIndex === nextPreview.columnIndex &&
        current.width === nextPreview.width
          ? current
          : nextPreview,
      )
      return clampedSize
    },
    [sheetName],
  )

  const previewRowHeight = useCallback(
    (rowIndex: number, newSize: number): number => {
      const clampedSize = Math.max(MIN_ROW_HEIGHT, Math.min(MAX_ROW_HEIGHT, Math.round(newSize)))
      const nextPreview = { sheetName, rowIndex, height: clampedSize }
      rowResizePreviewRef.current = nextPreview
      setRowResizePreview((current) =>
        current?.sheetName === nextPreview.sheetName && current.rowIndex === nextPreview.rowIndex && current.height === nextPreview.height
          ? current
          : nextPreview,
      )
      return clampedSize
    },
    [sheetName],
  )

  const getPreviewColumnWidth = useCallback(
    (columnIndex: number): number | null => {
      const preview = columnResizePreviewRef.current
      return preview?.sheetName === sheetName && preview.columnIndex === columnIndex ? preview.width : null
    },
    [sheetName],
  )

  const getPreviewRowHeight = useCallback(
    (rowIndex: number): number | null => {
      const preview = rowResizePreviewRef.current
      return preview?.sheetName === sheetName && preview.rowIndex === rowIndex ? preview.height : null
    },
    [sheetName],
  )

  const clearColumnResizePreview = useCallback(
    (columnIndex: number) => {
      const preview = columnResizePreviewRef.current
      if (preview?.sheetName === sheetName && preview.columnIndex === columnIndex) {
        columnResizePreviewRef.current = null
      }
      setColumnResizePreview((current) => (current?.sheetName === sheetName && current.columnIndex === columnIndex ? null : current))
    },
    [sheetName],
  )

  const clearRowResizePreview = useCallback(
    (rowIndex: number) => {
      const preview = rowResizePreviewRef.current
      if (preview?.sheetName === sheetName && preview.rowIndex === rowIndex) {
        rowResizePreviewRef.current = null
      }
      setRowResizePreview((current) => (current?.sheetName === sheetName && current.rowIndex === rowIndex ? null : current))
    },
    [sheetName],
  )

  const computeAutofitColumnWidth = useCallback(
    (columnIndex: number): number => {
      const canvas = textMeasureCanvasRef.current ?? document.createElement('canvas')
      textMeasureCanvasRef.current = canvas
      const context = canvas.getContext('2d')
      if (!context) {
        return gridMetrics.columnWidth
      }

      let measuredWidth = 0

      context.font = gridTheme.headerFontStyle
      measuredWidth = Math.max(measuredWidth, context.measureText(indexToColumn(columnIndex)).width)

      const sheet = engine.workbook.getSheet(sheetName)
      const measureCell = (row: number, col: number) => {
        const address = formatAddress(row, col)
        const optimisticSeed = getCellEditorSeed?.(sheetName, address)
        if (optimisticSeed !== undefined) {
          context.font = `400 ${gridTheme.editorFontSize} ${getResolvedCellFontFamily()}`
          measuredWidth = Math.max(measuredWidth, context.measureText(optimisticSeed).width)
          return
        }
        const snapshot =
          col === selectedCell.col &&
          row === selectedCell.row &&
          selectedCellSnapshot.sheetName === sheetName &&
          selectedCellSnapshot.address === address
            ? selectedCellSnapshot
            : engine.getCell(sheetName, address)
        const renderCell = snapshotToRenderCell(snapshot, engine.getCellStyle(snapshot.styleId))
        const displayText = renderCell.displayText || renderCell.copyText
        context.font = `400 ${gridTheme.editorFontSize} ${getResolvedCellFontFamily()}`
        measuredWidth = Math.max(measuredWidth, context.measureText(displayText).width)
      }

      if (sheet) {
        sheet.grid.forEachCellEntry((_cellIndex, row, col) => {
          if (col !== columnIndex) {
            return
          }
          measureCell(row, col)
        })
      } else {
        const liveVisibleRegion = liveVisibleRegionRef.current
        const visibleStartRow = liveVisibleRegion.range.y
        const visibleEndRow = Math.min(MAX_ROWS - 1, liveVisibleRegion.range.y + liveVisibleRegion.range.height - 1)
        for (let row = 0; row < Math.max(0, freezeRows); row += 1) {
          measureCell(row, columnIndex)
        }
        for (let row = visibleStartRow; row <= visibleEndRow; row += 1) {
          if (row < freezeRows) {
            continue
          }
          measureCell(row, columnIndex)
        }
      }

      return Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.ceil(measuredWidth + 28)))
    },
    [
      engine,
      freezeRows,
      getCellEditorSeed,
      gridMetrics.columnWidth,
      gridTheme.editorFontSize,
      gridTheme.headerFontStyle,
      selectedCell.col,
      selectedCell.row,
      selectedCellSnapshot,
      sheetName,
    ],
  )

  const overlayStyle = useMemo(() => getOverlayStyle(isEditingCell, overlayBounds), [isEditingCell, overlayBounds])
  const editorPresentation = useMemo(() => {
    const selectedCellStyle = engine.getCellStyle(selectedCellSnapshot.styleId)
    const renderCell = snapshotToRenderCell(selectedCellSnapshot, selectedCellStyle)
    return getEditorPresentation({
      renderCell,
      fillColor: selectedCellStyle?.fill?.backgroundColor,
    })
  }, [engine, selectedCellSnapshot])
  const editorTextAlign = useMemo<'left' | 'right'>(() => getEditorTextAlign(editorValue), [editorValue])

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
    isEntireSheetSelected: isSheetSelection(gridSelection),
    isFillHandleDragging,
    isRangeMoveDragging,
    overlayStyle,
    previewColumnWidth,
    previewRowHeight,
    renderPanes,
    residentDataPanes,
    rowHeights,
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
