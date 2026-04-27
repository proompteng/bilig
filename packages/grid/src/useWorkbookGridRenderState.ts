import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import { formatAddress, indexToColumn, parseCellAddress } from '@bilig/formula'
import type { CellSnapshot, Viewport } from '@bilig/protocol'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import { buildGridGpuScene, type GridGpuScene } from './gridGpuScene.js'
import { buildGridTextScene, type GridTextScene } from './gridTextScene.js'
import { createGridCameraSnapshot } from './gridCamera.js'
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
import { applyEditorOverlayBounds, resolveEditorOverlayScreenBounds } from './gridEditorOverlayGeometry.js'
import { createGridAxisWorldIndexFromRecords } from './gridAxisWorldIndex.js'
import { createGridGeometrySnapshotFromAxes } from './gridGeometry.js'
import { applyHiddenAxisSizes, resolveGridScrollSpacerSize } from './gridScrollSurface.js'
import type { HeaderSelection, VisibleRegionState } from './gridPointer.js'
import { resolveGridRenderScrollTransform, sameViewportBounds, sameVisibleRegionWindow } from './gridViewportController.js'
import type { GridHoverState } from './gridHover.js'
import { getResolvedCellFontFamily, snapshotToRenderCell } from './gridCells.js'
import { getEditorPresentation, getEditorTextAlign, getGridTheme, getOverlayStyle } from './gridPresentation.js'
import type { GridEngineLike } from './grid-engine.js'
import { CompactSelection, type GridSelection, type Item, type Rectangle } from './gridTypes.js'
import type { SheetGridViewportSubscription } from './workbookGridSurfaceTypes.js'
import { collectViewportItems } from './gridViewportItems.js'
import type { WorkbookRenderPaneState } from './renderer-v2/pane-scene-types.js'
import type { GridCameraSnapshot } from './renderer-v2/grid-render-contract.js'
import { buildLocalFixedRenderTiles } from './renderer-v3/local-render-tile-materializer.js'
import { tileKeysForViewport } from './renderer-v3/tile-key.js'
import { buildFixedRenderTileDataPaneStates } from './renderer-v3/render-tile-v2-adapter.js'
import type { GridRenderTile, GridRenderTileSource } from './renderer-v3/render-tile-source.js'
import {
  hasSelectionTargetChanged,
  resolveColumnOffset,
  resolveResidentViewport,
  resolveViewportScrollPosition,
  scrollCellIntoView,
} from './workbookGridViewport.js'
import { WorkbookGridScrollStore } from './workbookGridScrollStore.js'
import { noteGridScrollInput } from './renderer-v2/grid-render-counters.js'
import { GridCameraStore } from './renderer-v2/gridCameraStore.js'
import { visibleRegionFromCamera, viewportFromVisibleRegion } from './useGridCameraState.js'
import { useGridElementSize } from './useGridElementSize.js'
import { sameBounds } from './useGridOverlayState.js'
import { resolveRequiresLiveViewportState } from './useGridSelectionState.js'

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

const STATIC_SCENE_SELECTED_CELL: Item = Object.freeze([-1, -1] as const)
const STATIC_SCENE_GRID_SELECTION: GridSelection = Object.freeze({
  columns: CompactSelection.empty(),
  current: undefined,
  rows: CompactSelection.empty(),
})

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
  const invalidateSceneRef = useRef<() => void>(() => undefined)
  const retainedFixedRenderTileDataPanesRef = useRef<{
    readonly sheetId: number
    readonly panes: readonly WorkbookRenderPaneState[]
  } | null>(null)
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
  useLayoutEffect(() => {
    setGridSelection((current) => {
      if (
        current.columns.length > 0 ||
        current.rows.length > 0 ||
        current.current?.range.width !== 1 ||
        current.current.range.height !== 1
      ) {
        return current
      }
      if (current.current.cell[0] === selectedCell.col && current.current.cell[1] === selectedCell.row) {
        return current
      }
      return createGridSelection(selectedCell.col, selectedCell.row)
    })
  }, [selectedCell.col, selectedCell.row])
  const gridMetrics = useMemo(() => getGridMetrics(), [])
  const dprBucket = typeof window === 'undefined' ? 1 : Math.max(1, Math.ceil(window.devicePixelRatio || 1))
  const shouldUseRemoteRenderTileSource = renderTileSource !== undefined && sheetId !== undefined
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
  const hostElementSize = useGridElementSize(hostElement)
  const hostClientWidth = hostElementSize.width
  const hostClientHeight = hostElementSize.height
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
  const rowHeightOverridesAttr = useMemo(() => {
    const entries = Object.entries(rowHeights).toSorted(([left], [right]) => Number(left) - Number(right))
    return entries.length === 0 ? '{}' : JSON.stringify(Object.fromEntries(entries))
  }, [rowHeights])
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
  const getLiveGeometrySnapshot = useCallback(() => {
    const scrollViewport = scrollViewportRef.current
    if (!scrollViewport) {
      return gridCameraStore.getSnapshot()
    }
    return createGridGeometrySnapshotFromAxes({
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
    })
  }, [columnAxis, freezeCols, freezeRows, gridCameraStore, gridMetrics, rowAxis, sheetName])

  const viewport = useMemo<Viewport>(() => viewportFromVisibleRegion(visibleRegion), [visibleRegion])
  const residentViewportRef = useRef<Viewport>(resolveResidentViewport(viewport))
  const nextResidentViewport = resolveResidentViewport(viewport)
  if (!sameViewportBounds(residentViewportRef.current, nextResidentViewport)) {
    residentViewportRef.current = nextResidentViewport
  }
  const residentViewport = residentViewportRef.current
  const renderTileViewport = useMemo<Viewport>(
    () => ({
      rowStart: freezeRows > 0 ? 0 : residentViewport.rowStart,
      rowEnd: residentViewport.rowEnd,
      colStart: freezeCols > 0 ? 0 : residentViewport.colStart,
      colEnd: residentViewport.colEnd,
    }),
    [freezeCols, freezeRows, residentViewport.colEnd, residentViewport.colStart, residentViewport.rowEnd, residentViewport.rowStart],
  )
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
  useEffect(() => {
    invalidateSceneRef.current = invalidateScene
  }, [invalidateScene])

  const requiresLiveViewportState = resolveRequiresLiveViewportState({
    fillPreviewActive: fillPreviewRange !== null,
    hasActiveHeaderDrag: activeHeaderDrag !== null,
    hasActiveResizeColumn: activeResizeColumn !== null,
    hasActiveResizeRow: activeResizeRow !== null,
    hasColumnResizePreview: columnResizePreview !== null,
    hasRowResizePreview: rowResizePreview !== null,
    isEditingCell,
    isFillHandleDragging,
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
    const { renderTx, renderTy } = resolveGridRenderScrollTransform({
      nextVisibleRegion: next,
      renderViewport: viewport,
      sortedColumnWidthOverrides,
      sortedRowHeightOverrides,
      defaultColumnWidth: gridMetrics.columnWidth,
      defaultRowHeight: gridMetrics.rowHeight,
    })
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
    viewport,
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

  useLayoutEffect(() => {
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
    if (shouldUseRemoteRenderTileSource) {
      return
    }
    return engine.subscribeCells(sheetName, visibleAddresses, invalidateScene)
  }, [engine, invalidateScene, sheetName, shouldUseRemoteRenderTileSource, visibleAddresses])

  const [renderTileRevision, setRenderTileRevision] = useState(0)
  useEffect(() => {
    if (!renderTileSource || sheetId === undefined) {
      return
    }
    return renderTileSource.subscribeRenderTileDeltas(
      {
        ...renderTileViewport,
        cameraSeq: gridCameraStore.getSnapshot()?.camera.seq ?? 0,
        dprBucket,
        initialDelta: 'full',
        sheetId,
        sheetName,
      },
      () => {
        setRenderTileRevision((current) => current + 1)
      },
    )
  }, [dprBucket, gridCameraStore, renderTileSource, renderTileViewport, sheetId, sheetName])
  const fixedRenderTileDataPanes = useMemo<readonly WorkbookRenderPaneState[] | null>(() => {
    if (!hostElement) {
      return null
    }
    const tiles: GridRenderTile[] = []
    const tileSheetId = sheetId ?? 0
    if (renderTileSource && sheetId !== undefined) {
      void renderTileRevision
      const tileKeys = tileKeysForViewport({
        dprBucket,
        sheetOrdinal: sheetId,
        viewport: renderTileViewport,
      })
      for (const tileKey of tileKeys) {
        const tile = renderTileSource.peekRenderTile(tileKey)
        if (!tile || tile.coord.sheetId !== sheetId) {
          return null
        }
        tiles.push(tile)
      }
    } else {
      void sceneRevision
      tiles.push(
        ...buildLocalFixedRenderTiles({
          cameraSeq: gridCameraStore.getSnapshot()?.camera.seq ?? sceneRevision,
          columnWidths,
          dprBucket,
          engine,
          generation: sceneRevision,
          gridMetrics,
          rowHeights,
          sheetId: tileSheetId,
          sheetName,
          sortedColumnWidthOverrides,
          sortedRowHeightOverrides,
          viewport: renderTileViewport,
        }),
      )
    }
    const panes = buildFixedRenderTileDataPaneStates({
      columnWidths,
      freezeCols,
      freezeRows,
      frozenColumnWidth,
      frozenRowHeight,
      gridMetrics,
      hostHeight: hostClientHeight,
      hostWidth: hostClientWidth,
      residentViewport,
      rowHeights,
      sheetName,
      sortedColumnWidthOverrides,
      sortedRowHeightOverrides,
      tiles,
      visibleViewport: viewport,
    })
    return panes.length > 0 ? panes : null
  }, [
    columnWidths,
    dprBucket,
    engine,
    freezeCols,
    freezeRows,
    frozenColumnWidth,
    frozenRowHeight,
    gridMetrics,
    gridCameraStore,
    hostClientHeight,
    hostClientWidth,
    hostElement,
    renderTileRevision,
    renderTileSource,
    renderTileViewport,
    residentViewport,
    rowHeights,
    sceneRevision,
    sheetId,
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    viewport,
  ])
  useEffect(() => {
    if (sheetId === undefined || !fixedRenderTileDataPanes) {
      return
    }
    retainedFixedRenderTileDataPanesRef.current = {
      panes: fixedRenderTileDataPanes,
      sheetId,
    }
  }, [fixedRenderTileDataPanes, sheetId])
  const retainedFixedRenderTileDataPanes =
    fixedRenderTileDataPanes ??
    (shouldUseRemoteRenderTileSource && sheetId !== undefined && retainedFixedRenderTileDataPanesRef.current?.sheetId === sheetId
      ? retainedFixedRenderTileDataPanesRef.current.panes
      : null)
  const residentDataPanes = useMemo(() => retainedFixedRenderTileDataPanes ?? [], [retainedFixedRenderTileDataPanes])
  const residentBodyPane = residentDataPanes.find((pane) => pane.paneId === 'body') ?? null
  const preloadDataPanes = useMemo<readonly WorkbookRenderPaneState[]>(() => [], [])

  const headerGpuScene = useMemo<GridGpuScene>(() => {
    if (!hostElement) {
      return emptyGpuScene
    }
    return buildGridGpuScene({
      contentMode: 'headers',
      engine,
      columnWidths,
      rowHeights,
      gridMetrics,
      gridSelection: STATIC_SCENE_GRID_SELECTION,
      activeHeaderDrag: null,
      hoveredCell: null,
      hoveredHeader: null,
      resizeGuideColumn: null,
      resizeGuideRow: null,
      selectedCell: STATIC_SCENE_SELECTED_CELL,
      selectionRange: null,
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
    columnWidths,
    emptyGpuScene,
    engine,
    getHeaderCellLocalBounds,
    gridMetrics,
    hostElement,
    rowHeights,
    sheetName,
    residentHeaderItems,
    residentHeaderRegion,
  ])

  const headerTextScene = useMemo<GridTextScene>(() => {
    if (!hostElement) {
      return emptyTextScene
    }
    return buildGridTextScene({
      contentMode: 'headers',
      engine,
      columnWidths,
      rowHeights,
      editingCell: null,
      gridMetrics,
      activeHeaderDrag: null,
      hoveredHeader: null,
      resizeGuideColumn: null,
      selectedCell: STATIC_SCENE_SELECTED_CELL,
      selectedCellSnapshot: null,
      selectionRange: null,
      sheetName,
      visibleItems: residentHeaderItems,
      visibleRegion: residentHeaderRegion,
      hostBounds: {
        left: 0,
        top: 0,
        width: 0,
        height: 0,
      },
      getCellBounds: getHeaderCellLocalBounds,
    })
  }, [
    columnWidths,
    emptyTextScene,
    engine,
    getHeaderCellLocalBounds,
    gridMetrics,
    hostElement,
    rowHeights,
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

  const refreshOverlayBounds = useCallback(
    (options?: { readonly commitReactState?: boolean }) => {
      const next = resolveEditorOverlayScreenBounds({
        col: selectedCell.col,
        row: selectedCell.row,
        geometry: gridCameraStore.getSnapshot(),
        getCellLocalBounds,
        hostElement: hostRef.current,
      })
      if (!next) {
        return
      }
      applyEditorOverlayBounds(next)
      if (options?.commitReactState === false) {
        return
      }
      setOverlayBounds((current) => {
        return sameBounds(current, next) ? current : next
      })
    },
    [getCellLocalBounds, gridCameraStore, selectedCell.col, selectedCell.row],
  )

  useLayoutEffect(() => {
    if (!isEditingCell) {
      setOverlayBounds(undefined)
      return
    }

    refreshOverlayBounds({ commitReactState: true })
    const frame = window.requestAnimationFrame(() => refreshOverlayBounds({ commitReactState: true }))
    const unsubscribeScrollTransform = scrollTransformStore.subscribe(() => refreshOverlayBounds({ commitReactState: false }))
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
    isEntireSheetSelected: isSheetSelection(gridSelection),
    isFillHandleDragging,
    isRangeMoveDragging,
    overlayStyle,
    previewColumnWidth,
    previewRowHeight,
    preloadDataPanes,
    renderPanes,
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
