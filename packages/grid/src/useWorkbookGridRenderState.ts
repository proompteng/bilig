import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { CellSnapshot, Viewport } from '@bilig/protocol'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import { getGridMetrics, getResolvedColumnWidth, getResolvedRowHeight, resolveRowOffset } from './gridMetrics.js'
import { createGridAxisWorldIndexFromRecords } from './gridAxisWorldIndex.js'
import { createGridGeometrySnapshotFromAxes } from './gridGeometry.js'
import { resolveGridScrollSpacerSize } from './gridScrollSurface.js'
import type { VisibleRegionState } from './gridPointer.js'
import { resolveGridRenderScrollTransform, sameViewportBounds, sameVisibleRegionWindow } from './gridViewportController.js'
import { getGridTheme } from './gridPresentation.js'
import type { GridEngineLike } from './grid-engine.js'
import type { Rectangle } from './gridTypes.js'
import type { SheetGridViewportSubscription } from './workbookGridSurfaceTypes.js'
import { collectViewportItems } from './gridViewportItems.js'
import type { GridRenderTileSource } from './renderer-v3/render-tile-source.js'
import { hasSelectionTargetChanged, resolveColumnOffset, resolveResidentViewport } from './workbookGridViewport.js'
import { WorkbookGridScrollStore } from './workbookGridScrollStore.js'
import { noteGridScrollInput } from './grid-render-counters.js'
import { GridCameraStore } from './runtime/gridCameraStore.js'
import { viewportFromVisibleRegion } from './useGridCameraState.js'
import { useGridElementSize } from './useGridElementSize.js'
import { GridRuntimeHost } from './runtime/gridRuntimeHost.js'
import {
  axisOverridesFromSortedSizes,
  createGridRuntimeAxisOverrideCache,
  syncGridRuntimeAxisOverrides,
} from './runtime/gridRuntimeAxisAdapters.js'
import { useWorkbookHeaderPanes } from './useWorkbookHeaderPanes.js'
import { useWorkbookRenderTilePanes } from './useWorkbookRenderTilePanes.js'
import { useWorkbookEditorOverlayAnchor } from './useWorkbookEditorOverlayAnchor.js'
import { useWorkbookAxisResizeState } from './useWorkbookAxisResizeState.js'
import { useWorkbookInteractionOverlayState } from './useWorkbookInteractionOverlayState.js'
import { useWorkbookColumnAutofit } from './useWorkbookColumnAutofit.js'

function noteVisibleWindowChange(): void {
  if (typeof window === 'undefined') {
    return
  }
  ;(window as Window & { __biligScrollPerf?: { noteVisibleWindowChange?: () => void } }).__biligScrollPerf?.noteVisibleWindowChange?.()
}

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
  const autoScrollSelectionRef = useRef<{ sheetName: string; col: number; row: number } | null>(null)
  const restoredViewportTokenRef = useRef<number | null>(null)
  const scrollSyncFrameRef = useRef<number | null>(null)
  const scrollTransformStoreRef = useRef<WorkbookGridScrollStore>(new WorkbookGridScrollStore())
  const scrollTransformRef = useRef(scrollTransformStoreRef.current.getSnapshot())
  const gridCameraStoreRef = useRef<GridCameraStore>(new GridCameraStore())
  const gridRuntimeHostRef = useRef<GridRuntimeHost | null>(null)
  const gridRuntimeAxisCacheRef = useRef(createGridRuntimeAxisOverrideCache())
  const invalidateSceneRef = useRef<() => void>(() => undefined)
  const [sceneRevision, setSceneRevision] = useState(0)
  const [hostElement, setHostElement] = useState<HTMLDivElement | null>(null)
  const [visibleRegion, setVisibleRegion] = useState<VisibleRegionState>({
    range: { x: 0, y: 0, width: 12, height: 24 },
    tx: 0,
    ty: 0,
    freezeRows,
    freezeCols,
  })
  const selectedCell = useMemo(() => parseCellAddress(selectedAddr, sheetName), [selectedAddr, sheetName])
  const gridMetrics = useMemo(() => getGridMetrics(), [])
  const dprBucket = typeof window === 'undefined' ? 1 : Math.max(1, Math.ceil(window.devicePixelRatio || 1))
  const shouldUseRemoteRenderTileSource = renderTileSource !== undefined && sheetId !== undefined
  const gridTheme = useMemo(() => getGridTheme(), [])
  const liveVisibleRegionRef = useRef<VisibleRegionState>(visibleRegion)
  const hostElementSize = useGridElementSize(hostElement)
  const hostClientWidth = hostElementSize.width
  const hostClientHeight = hostElementSize.height
  const gridRuntimeHost = useMemo(() => {
    const existing = gridRuntimeHostRef.current
    if (existing) {
      return existing
    }
    const host = new GridRuntimeHost({
      columnCount: MAX_COLS,
      defaultColumnWidth: gridMetrics.columnWidth,
      defaultRowHeight: gridMetrics.rowHeight,
      freezeCols,
      freezeRows,
      gridMetrics,
      rowCount: MAX_ROWS,
      viewportHeight: hostClientHeight,
      viewportWidth: hostClientWidth,
    })
    gridRuntimeHostRef.current = host
    return host
  }, [freezeCols, freezeRows, gridMetrics, hostClientHeight, hostClientWidth])
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
  const runtimeColumnAxisOverrides = useMemo(() => axisOverridesFromSortedSizes(sortedColumnWidthOverrides), [sortedColumnWidthOverrides])
  const runtimeRowAxisOverrides = useMemo(() => axisOverridesFromSortedSizes(sortedRowHeightOverrides), [sortedRowHeightOverrides])
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
  const syncRuntimeAxes = useCallback(() => {
    syncGridRuntimeAxisOverrides(gridRuntimeHost, gridRuntimeAxisCacheRef.current, {
      columnOverrides: runtimeColumnAxisOverrides,
      columnSeq: columnAxis.version,
      rowOverrides: runtimeRowAxisOverrides,
      rowSeq: rowAxis.version,
    })
  }, [columnAxis.version, gridRuntimeHost, rowAxis.version, runtimeColumnAxisOverrides, runtimeRowAxisOverrides])
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

  const syncVisibleRegion = useCallback(() => {
    const scrollViewport = scrollViewportRef.current
    if (!scrollViewport) {
      return
    }
    syncRuntimeAxes()
    const camera = gridRuntimeHost.updateCamera({
      scrollLeft: scrollViewport.scrollLeft,
      scrollTop: scrollViewport.scrollTop,
      viewportWidth: scrollViewport.clientWidth,
      viewportHeight: scrollViewport.clientHeight,
      dpr: window.devicePixelRatio || 1,
      freezeRows,
      freezeCols,
      gridMetrics,
    })
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
        seq: camera.seq,
        sheetName,
      }),
    )
    const next = camera.visibleRegion
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
    columnAxis,
    freezeCols,
    freezeRows,
    gridRuntimeHost,
    gridCameraStore,
    gridMetrics,
    onVisibleViewportChange,
    requiresLiveViewportState,
    rowAxis,
    scrollTransformStore,
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    syncRuntimeAxes,
    viewport,
  ])

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
    engine,
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
    syncRuntimeAxes()
    const nextScrollPosition = gridRuntimeHost.resolveScrollForCellIntoView({
      cell: [selectedCell.col, selectedCell.row],
      freezeRows,
      freezeCols,
      gridMetrics,
      scrollLeft: scrollViewport.scrollLeft,
      scrollTop: scrollViewport.scrollTop,
      viewportHeight: scrollViewport.clientHeight,
      viewportWidth: scrollViewport.clientWidth,
    })
    scrollViewport.scrollLeft = nextScrollPosition.scrollLeft
    scrollViewport.scrollTop = nextScrollPosition.scrollTop
  }, [freezeCols, freezeRows, gridRuntimeHost, gridMetrics, selectedCell.col, selectedCell.row, sheetName, syncRuntimeAxes])

  useLayoutEffect(() => {
    const scrollViewport = scrollViewportRef.current
    if (!scrollViewport || !restoreViewportTarget) {
      return
    }
    if (restoredViewportTokenRef.current === restoreViewportTarget.token) {
      return
    }
    restoredViewportTokenRef.current = restoreViewportTarget.token
    syncRuntimeAxes()
    const { scrollLeft: nextScrollLeft, scrollTop: nextScrollTop } = gridRuntimeHost.resolveScrollPositionForViewport({
      viewport: restoreViewportTarget.viewport,
      freezeRows,
      freezeCols,
    })
    scrollViewport.scrollLeft = nextScrollLeft
    scrollViewport.scrollTop = nextScrollTop
  }, [freezeCols, freezeRows, gridRuntimeHost, restoreViewportTarget, syncRuntimeAxes])

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
