import { useCallback, useMemo, useRef } from 'react'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import { createGridAxisWorldIndexFromRecords, type GridAxisWorldIndex } from './gridAxisWorldIndex.js'
import { createGridGeometrySnapshotFromAxes, type GridGeometrySnapshot } from './gridGeometry.js'
import type { GridMetrics } from './gridMetrics.js'
import { resolveGridScrollSpacerSize } from './gridScrollSurface.js'
import type { Rectangle } from './gridTypes.js'
import { GridCameraStore } from './runtime/gridCameraStore.js'
import {
  axisOverridesFromSortedSizes,
  createGridRuntimeAxisOverrideCache,
  syncGridRuntimeAxisOverrides,
} from './runtime/gridRuntimeAxisAdapters.js'
import { GridRuntimeHost } from './runtime/gridRuntimeHost.js'
import { WorkbookGridScrollStore, type WorkbookGridScrollSnapshot } from './workbookGridScrollStore.js'

type MutableRef<T> = {
  current: T
}

export interface WorkbookGridGeometryRuntimeState {
  readonly columnAxis: GridAxisWorldIndex
  readonly columnWidthOverridesAttr: string
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly getCellLocalBounds: (col: number, row: number) => Rectangle | undefined
  readonly getCellScreenBounds: (col: number, row: number) => Rectangle | undefined
  readonly getLiveGeometrySnapshot: () => GridGeometrySnapshot | null
  readonly gridCameraStore: GridCameraStore
  readonly gridRuntimeHost: GridRuntimeHost
  readonly rowAxis: GridAxisWorldIndex
  readonly rowHeightOverridesAttr: string
  readonly scrollTransformRef: MutableRef<WorkbookGridScrollSnapshot>
  readonly scrollTransformStore: WorkbookGridScrollStore
  readonly sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  readonly sortedRowHeightOverrides: readonly (readonly [number, number])[]
  readonly syncRuntimeAxes: () => void
  readonly totalGridHeight: number
  readonly totalGridWidth: number
}

export function useWorkbookGridGeometryRuntime(input: {
  readonly columnWidths: Readonly<Record<number, number>>
  readonly controlledHiddenColumns?: Readonly<Record<number, true>> | undefined
  readonly controlledHiddenRows?: Readonly<Record<number, true>> | undefined
  readonly freezeCols: number
  readonly freezeRows: number
  readonly gridMetrics: GridMetrics
  readonly hostClientHeight: number
  readonly hostClientWidth: number
  readonly hostRef: MutableRef<HTMLDivElement | null>
  readonly rowHeights: Readonly<Record<number, number>>
  readonly scrollViewportRef: MutableRef<HTMLDivElement | null>
  readonly sheetName: string
}): WorkbookGridGeometryRuntimeState {
  const {
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
  } = input
  const scrollTransformStoreRef = useRef<WorkbookGridScrollStore>(new WorkbookGridScrollStore())
  const scrollTransformRef = useRef<WorkbookGridScrollSnapshot>(scrollTransformStoreRef.current.getSnapshot())
  const gridCameraStoreRef = useRef<GridCameraStore>(new GridCameraStore())
  const gridRuntimeHostRef = useRef<GridRuntimeHost | null>(null)
  const gridRuntimeAxisCacheRef = useRef(createGridRuntimeAxisOverrideCache())
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
    [getCellLocalBounds, hostRef],
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
  }, [columnAxis, freezeCols, freezeRows, gridCameraStore, gridMetrics, rowAxis, scrollViewportRef, sheetName])

  return {
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
    totalGridHeight: scrollSpacerSize.height,
    totalGridWidth: scrollSpacerSize.width,
  }
}
