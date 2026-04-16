import { useDeferredValue, useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import { formatAddress, indexToColumn, parseCellAddress } from '@bilig/formula'
import type { CellSnapshot, Viewport } from '@bilig/protocol'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import { buildGridGpuScene, type GridGpuScene } from './gridGpuScene.js'
import { buildGridTextScene, type GridTextScene } from './gridTextScene.js'
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
import { resolveFillHandleOverlayBounds, resolveFillHandlePreviewBounds, type FillHandleOverlayBounds } from './gridFillHandle.js'
import type { HeaderSelection, VisibleRegionState } from './gridPointer.js'
import type { GridHoverState } from './gridHover.js'
import { getResolvedCellFontFamily, snapshotToRenderCell } from './gridCells.js'
import { getEditorPresentation, getEditorTextAlign, getGridTheme, getOverlayStyle } from './gridPresentation.js'
import type { GridEngineLike } from './grid-engine.js'
import type { GridSelection, Rectangle } from './gridTypes.js'
import type { SheetGridViewportSubscription } from './workbookGridSurfaceTypes.js'
import {
  hasSelectionTargetChanged,
  resolveColumnOffset,
  resolveFrozenColumnWidth,
  resolveFrozenRowHeight,
  resolveViewportScrollPosition,
  resolveVisibleRegionFromScroll,
  scrollCellIntoView,
} from './workbookGridViewport.js'

function sameBounds(left: Rectangle | undefined, right: Rectangle | undefined): boolean {
  if (left === right) {
    return true
  }
  if (!left || !right) {
    return false
  }
  return left.x === right.x && left.y === right.y && left.width === right.width && left.height === right.height
}

function sameVisibleRegion(left: VisibleRegionState, right: VisibleRegionState): boolean {
  return (
    (left.freezeCols ?? 0) === (right.freezeCols ?? 0) &&
    (left.freezeRows ?? 0) === (right.freezeRows ?? 0) &&
    left.tx === right.tx &&
    left.ty === right.ty &&
    left.range.x === right.range.x &&
    left.range.y === right.range.y &&
    left.range.width === right.range.width &&
    left.range.height === right.range.height
  )
}

function collectVisibleItems(options: { visibleRegion: VisibleRegionState }): [number, number][] {
  const { visibleRegion } = options
  const freezeCols = Math.max(0, Math.min(MAX_COLS, visibleRegion.freezeCols ?? 0))
  const freezeRows = Math.max(0, Math.min(MAX_ROWS, visibleRegion.freezeRows ?? 0))
  const rowEnd = Math.min(MAX_ROWS - 1, visibleRegion.range.y + visibleRegion.range.height - 1)
  const colEnd = Math.min(MAX_COLS - 1, visibleRegion.range.x + visibleRegion.range.width - 1)
  const items: [number, number][] = []
  const seen = new Set<string>()

  const pushItem = (col: number, row: number) => {
    if (col < 0 || col >= MAX_COLS || row < 0 || row >= MAX_ROWS) {
      return
    }
    const key = `${col}:${row}`
    if (seen.has(key)) {
      return
    }
    seen.add(key)
    items.push([col, row])
  }

  for (let row = 0; row < freezeRows; row += 1) {
    for (let col = 0; col < freezeCols; col += 1) {
      pushItem(col, row)
    }
    for (let col = visibleRegion.range.x; col <= colEnd; col += 1) {
      pushItem(col, row)
    }
  }

  for (let row = visibleRegion.range.y; row <= rowEnd; row += 1) {
    for (let col = 0; col < freezeCols; col += 1) {
      pushItem(col, row)
    }
    for (let col = visibleRegion.range.x; col <= colEnd; col += 1) {
      pushItem(col, row)
    }
  }

  return items
}

function collectViewportSubscriptions(viewport: Viewport, freezeRows: number, freezeCols: number): Viewport[] {
  const viewports: Viewport[] = [viewport]
  if (freezeRows > 0) {
    viewports.push({
      rowStart: 0,
      rowEnd: freezeRows - 1,
      colStart: viewport.colStart,
      colEnd: viewport.colEnd,
    })
  }
  if (freezeCols > 0) {
    viewports.push({
      rowStart: viewport.rowStart,
      rowEnd: viewport.rowEnd,
      colStart: 0,
      colEnd: freezeCols - 1,
    })
  }
  if (freezeRows > 0 && freezeCols > 0) {
    viewports.push({
      rowStart: 0,
      rowEnd: freezeRows - 1,
      colStart: 0,
      colEnd: freezeCols - 1,
    })
  }
  return [...new Map(viewports.map((entry) => [`${entry.rowStart}:${entry.rowEnd}:${entry.colStart}:${entry.colEnd}`, entry])).values()]
}

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
  const [isWebGpuActive, setIsWebGpuActive] = useState(false)
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
  const baseColumnWidths = controlledColumnWidths ?? columnWidthsBySheet[sheetName] ?? EMPTY_COLUMN_WIDTHS
  const baseRowHeights = controlledRowHeights ?? rowHeightsBySheet[sheetName] ?? EMPTY_ROW_HEIGHTS
  const columnWidths = useMemo(() => {
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
  const rowHeights = useMemo(() => {
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
  const totalGridWidth = useMemo(
    () => gridMetrics.rowMarkerWidth + resolveColumnOffset(MAX_COLS, sortedColumnWidthOverrides, gridMetrics.columnWidth),
    [gridMetrics.columnWidth, gridMetrics.rowMarkerWidth, sortedColumnWidthOverrides],
  )
  const totalGridHeight = useMemo(
    () => gridMetrics.headerHeight + resolveRowOffset(MAX_ROWS, sortedRowHeightOverrides, gridMetrics.rowHeight),
    [gridMetrics.headerHeight, gridMetrics.rowHeight, sortedRowHeightOverrides],
  )
  const frozenColumnWidth = useMemo(
    () => resolveFrozenColumnWidth({ freezeCols, columnWidths, gridMetrics }),
    [columnWidths, freezeCols, gridMetrics],
  )
  const frozenRowHeight = useMemo(
    () => resolveFrozenRowHeight({ freezeRows, rowHeights, gridMetrics }),
    [freezeRows, gridMetrics, rowHeights],
  )
  const selectionRange = gridSelection.current?.range ?? null

  const getCellLocalBounds = useCallback(
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
              resolveColumnOffset(visibleRegion.range.x, sortedColumnWidthOverrides, gridMetrics.columnWidth) -
              visibleRegion.tx,
        y:
          row < freezeRows
            ? gridMetrics.headerHeight + resolveRowOffset(row, sortedRowHeightOverrides, gridMetrics.rowHeight)
            : gridMetrics.headerHeight +
              frozenRowHeight +
              resolveRowOffset(row, sortedRowHeightOverrides, gridMetrics.rowHeight) -
              resolveRowOffset(visibleRegion.range.y, sortedRowHeightOverrides, gridMetrics.rowHeight) -
              visibleRegion.ty,
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
      rowHeights,
      sortedColumnWidthOverrides,
      sortedRowHeightOverrides,
      visibleRegion.range.x,
      visibleRegion.range.y,
      visibleRegion.tx,
      visibleRegion.ty,
    ],
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

  const visibleItems = useMemo(() => {
    return collectVisibleItems({ visibleRegion })
  }, [visibleRegion])
  const visibleAddresses = useMemo(() => visibleItems.map(([col, row]) => formatAddress(row, col)), [visibleItems])
  const deferredVisibleRegion = useDeferredValue(visibleRegion)
  const deferredVisibleItems = useDeferredValue(visibleItems)
  const viewport = useMemo<Viewport>(
    () => ({
      rowStart: visibleRegion.range.y,
      rowEnd: Math.min(MAX_ROWS - 1, visibleRegion.range.y + visibleRegion.range.height - 1),
      colStart: visibleRegion.range.x,
      colEnd: Math.min(MAX_COLS - 1, visibleRegion.range.x + visibleRegion.range.width - 1),
    }),
    [visibleRegion.range.height, visibleRegion.range.width, visibleRegion.range.x, visibleRegion.range.y],
  )

  const invalidateScene = useCallback(() => {
    setSceneRevision((current) => current + 1)
  }, [])

  const syncVisibleRegion = useCallback(() => {
    const scrollViewport = scrollViewportRef.current
    if (!scrollViewport) {
      return
    }
    const next = resolveVisibleRegionFromScroll({
      scrollLeft: scrollViewport.scrollLeft,
      scrollTop: scrollViewport.scrollTop,
      viewportWidth: scrollViewport.clientWidth,
      viewportHeight: scrollViewport.clientHeight,
      freezeRows,
      freezeCols,
      columnWidths,
      rowHeights,
      gridMetrics,
    })
    setVisibleRegion((current) => (sameVisibleRegion(current, next) ? current : next))
  }, [columnWidths, freezeCols, freezeRows, gridMetrics, rowHeights])

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
    onVisibleViewportChange?.(viewport)
  }, [onVisibleViewportChange, viewport])

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
      scheduleVisibleRegionSync()
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
    if (subscribeViewport) {
      const viewports = collectViewportSubscriptions(viewport, freezeRows, freezeCols)
      const cleanups = viewports.map((nextViewport) => subscribeViewport(sheetName, nextViewport, invalidateScene))
      return () => {
        cleanups.forEach((cleanup) => cleanup())
      }
    }
    return engine.subscribeCells(sheetName, visibleAddresses, invalidateScene)
  }, [engine, freezeCols, freezeRows, invalidateScene, sheetName, subscribeViewport, viewport, visibleAddresses])

  const resizeGuideColumn = useMemo(
    () =>
      activeResizeColumn ?? (hoverState.cursor === 'col-resize' && hoverState.header?.kind === 'column' ? hoverState.header.index : null),
    [activeResizeColumn, hoverState.cursor, hoverState.header],
  )
  const resizeGuideRow = useMemo(
    () => activeResizeRow ?? (hoverState.cursor === 'row-resize' && hoverState.header?.kind === 'row' ? hoverState.header.index : null),
    [activeResizeRow, hoverState.cursor, hoverState.header],
  )

  const gpuScene = useMemo<GridGpuScene>(() => {
    if (!hostElement) {
      return emptyGpuScene
    }
    void sceneRevision
    return buildGridGpuScene({
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
      visibleItems: deferredVisibleItems,
      visibleRegion: deferredVisibleRegion,
      hostBounds: {
        left: 0,
        top: 0,
      },
      getCellBounds: getCellLocalBounds,
    })
  }, [
    activeHeaderDrag,
    columnWidths,
    deferredVisibleItems,
    deferredVisibleRegion,
    emptyGpuScene,
    engine,
    getCellLocalBounds,
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
  ])

  const textScene = useMemo<GridTextScene>(() => {
    if (!hostElement) {
      return emptyTextScene
    }
    void sceneRevision
    return buildGridTextScene({
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
      visibleItems: deferredVisibleItems,
      visibleRegion: deferredVisibleRegion,
      hostBounds: {
        left: 0,
        top: 0,
        width: hostElement.clientWidth,
        height: hostElement.clientHeight,
      },
      getCellBounds: getCellLocalBounds,
    })
  }, [
    activeHeaderDrag,
    columnWidths,
    deferredVisibleItems,
    deferredVisibleRegion,
    emptyTextScene,
    engine,
    getCellLocalBounds,
    gridMetrics,
    hostElement,
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
  ])

  const fillHandleBounds = useMemo<FillHandleOverlayBounds | undefined>(() => {
    if (
      !hostElement ||
      !selectionRange ||
      gridSelection.columns.length > 0 ||
      gridSelection.rows.length > 0 ||
      fillPreviewRange ||
      isRangeMoveDragging
    ) {
      return undefined
    }
    return resolveFillHandleOverlayBounds({
      sourceRange: selectionRange,
      hostBounds: {
        left: 0,
        top: 0,
        width: hostElement.clientWidth,
        height: hostElement.clientHeight,
      },
      getCellBounds: getCellLocalBounds,
      minX: gridMetrics.rowMarkerWidth,
      minY: gridMetrics.headerHeight,
    })
  }, [
    fillPreviewRange,
    getCellLocalBounds,
    gridMetrics.headerHeight,
    gridMetrics.rowMarkerWidth,
    gridSelection.columns.length,
    gridSelection.rows.length,
    hostElement,
    isRangeMoveDragging,
    selectionRange,
  ])

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
    return () => {
      window.cancelAnimationFrame(frame)
    }
  }, [isEditingCell, refreshOverlayBounds, visibleRegion.tx, visibleRegion.ty])

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
      sheet?.grid.forEachCellEntry((_cellIndex, row, col) => {
        if (col !== columnIndex) {
          return
        }
        const snapshot = engine.getCell(sheetName, formatAddress(row, col))
        const renderCell = snapshotToRenderCell(snapshot, engine.getCellStyle(snapshot.styleId))
        const displayText = renderCell.displayText || renderCell.copyText
        context.font = `400 ${gridTheme.editorFontSize} ${getResolvedCellFontFamily()}`
        measuredWidth = Math.max(measuredWidth, context.measureText(displayText).width)
      })

      return Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, Math.ceil(measuredWidth + 28)))
    },
    [engine, gridMetrics.columnWidth, gridTheme.editorFontSize, gridTheme.headerFontStyle, sheetName],
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
    fillHandleBounds,
    fillPreviewBounds,
    fillPreviewRange,
    focusGrid,
    focusTargetRef,
    getCellLocalBounds,
    getCellScreenBounds,
    getPreviewColumnWidth,
    getPreviewRowHeight,
    gridMetrics,
    gridSelection,
    gridTheme,
    gpuScene,
    handleHostRef,
    hostElement,
    hostRef,
    hoverState,
    isEntireSheetSelected: isSheetSelection(gridSelection),
    isFillHandleDragging,
    isRangeMoveDragging,
    isWebGpuActive,
    overlayStyle,
    previewColumnWidth,
    previewRowHeight,
    rowHeights,
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
    setIsWebGpuActive,
    textScene,
    totalGridHeight,
    totalGridWidth,
    visibleRegion,
    editorPresentation,
    editorTextAlign,
  }
}
