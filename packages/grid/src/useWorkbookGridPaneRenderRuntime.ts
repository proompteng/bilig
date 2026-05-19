import { useMemo, useSyncExternalStore, type Dispatch, type SetStateAction } from 'react'
import type { CellSnapshot, Viewport } from '@bilig/protocol'
import type { GridAxisWorldIndex } from './gridAxisWorldIndex.js'
import type { GridEngineLike } from './grid-engine.js'
import type { GridMetrics } from './gridMetrics.js'
import type { VisibleRegionState } from './gridPointer.js'
import type { Item } from './gridTypes.js'
import type { GridCameraStore } from './runtime/gridCameraStore.js'
import type { GridRuntimeHost } from './runtime/gridRuntimeHost.js'
import { collectViewportItems } from './gridViewportItems.js'
import type { WorkbookGridScrollSnapshot, WorkbookGridScrollStore } from './workbookGridScrollStore.js'
import type { GridRenderTileSource } from './renderer-v3/render-tile-source.js'
import { useWorkbookGridViewportRuntime } from './useWorkbookGridViewportRuntime.js'
import { useWorkbookHeaderCellBounds } from './useWorkbookHeaderCellBounds.js'
import { useWorkbookHeaderPanes } from './useWorkbookHeaderPanes.js'
import { useWorkbookRenderTilePanes } from './useWorkbookRenderTilePanes.js'

type MutableRef<T> = {
  current: T
}

type SortedAxisOverrides = readonly (readonly [number, number])[]

export function resolveGridDrawDprBucket(source: { readonly devicePixelRatio?: number | undefined } | null | undefined): number {
  const ratio = source?.devicePixelRatio ?? 1
  return Number.isFinite(ratio) ? Math.max(1, Math.ceil(ratio || 1)) : 1
}

function getGridDrawDprBucketSnapshot(): number {
  return resolveGridDrawDprBucket(typeof window === 'undefined' ? null : window)
}

function subscribeGridDrawDprBucketChange(listener: () => void): () => void {
  if (typeof window === 'undefined') {
    return () => {}
  }

  let disposed = false
  let resolutionQuery: MediaQueryList | null = null
  const removeResolutionQuery = () => {
    if (!resolutionQuery) {
      return
    }
    if (typeof resolutionQuery.removeEventListener === 'function') {
      resolutionQuery.removeEventListener('change', handleChange)
    } else {
      const legacyQuery = resolutionQuery as MediaQueryList & { removeListener?: (listener: () => void) => void }
      legacyQuery.removeListener?.(handleChange)
    }
    resolutionQuery = null
  }
  const addResolutionQuery = () => {
    if (typeof window.matchMedia !== 'function') {
      return
    }
    resolutionQuery = window.matchMedia(`(resolution: ${Math.max(1, window.devicePixelRatio || 1)}dppx)`)
    if (typeof resolutionQuery.addEventListener === 'function') {
      resolutionQuery.addEventListener('change', handleChange)
    } else {
      const legacyQuery = resolutionQuery as MediaQueryList & { addListener?: (listener: () => void) => void }
      legacyQuery.addListener?.(handleChange)
    }
  }
  const resetResolutionQuery = () => {
    if (disposed) {
      return
    }
    removeResolutionQuery()
    addResolutionQuery()
  }
  function handleChange(): void {
    if (disposed) {
      return
    }
    resetResolutionQuery()
    listener()
  }

  addResolutionQuery()
  window.addEventListener('resize', handleChange)
  window.visualViewport?.addEventListener('resize', handleChange)
  return () => {
    disposed = true
    removeResolutionQuery()
    window.removeEventListener('resize', handleChange)
    window.visualViewport?.removeEventListener('resize', handleChange)
  }
}

export function useGridDrawDprBucket(): number {
  return useSyncExternalStore(subscribeGridDrawDprBucketChange, getGridDrawDprBucketSnapshot, () => 1)
}

export function resolveShouldUseRemoteRenderTileSource(input: {
  readonly renderTileSource?: unknown
  readonly sheetId?: number | undefined
}): boolean {
  return input.renderTileSource !== undefined && input.sheetId !== undefined
}

export function resolveHeaderPaneWindowMode(input: {
  readonly residentViewport: Viewport
  readonly viewport: Viewport
  readonly visibleDirtyTileKeys: readonly unknown[]
}): 'resident' | 'visible' {
  if (input.visibleDirtyTileKeys.length === 0) {
    return 'resident'
  }
  const residentCoversVisible =
    input.residentViewport.rowStart <= input.viewport.rowStart &&
    input.residentViewport.rowEnd >= input.viewport.rowEnd &&
    input.residentViewport.colStart <= input.viewport.colStart &&
    input.residentViewport.colEnd >= input.viewport.colEnd
  return residentCoversVisible ? 'resident' : 'visible'
}

export function useWorkbookGridPaneRenderRuntime(input: {
  readonly columnAxis: GridAxisWorldIndex
  readonly columnWidths: Readonly<Record<number, number>>
  readonly editingCell?: Item | null | undefined
  readonly engine: GridEngineLike
  readonly freezeCols: number
  readonly freezeRows: number
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly gridCameraStore: GridCameraStore
  readonly gridMetrics: GridMetrics
  readonly gridRuntimeHost: GridRuntimeHost
  readonly hostClientHeight: number
  readonly hostClientWidth: number
  readonly hostElement: HTMLDivElement | null
  readonly liveVisibleRegionRef: MutableRef<VisibleRegionState>
  readonly onVisibleViewportChange?: ((viewport: Viewport) => void) | undefined
  readonly renderTileSource?: GridRenderTileSource | undefined
  readonly requiresLiveViewportState: boolean
  readonly restoreViewportTarget?:
    | {
        readonly token: number
        readonly viewport: Viewport
      }
    | undefined
  readonly rowAxis: GridAxisWorldIndex
  readonly rowHeights: Readonly<Record<number, number>>
  readonly scrollTransformRef: MutableRef<WorkbookGridScrollSnapshot>
  readonly scrollTransformStore: WorkbookGridScrollStore
  readonly scrollViewportRef: MutableRef<HTMLDivElement | null>
  readonly selectedCell: Item
  readonly selectedCellSnapshot?: CellSnapshot | null | undefined
  readonly setVisibleRegion: Dispatch<SetStateAction<VisibleRegionState>>
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
  readonly sheetName: string
  readonly sortedColumnWidthOverrides: SortedAxisOverrides
  readonly sortedRowHeightOverrides: SortedAxisOverrides
  readonly syncRuntimeAxes: () => void
  readonly visibleRegion: VisibleRegionState
}) {
  const {
    columnAxis,
    columnWidths,
    editingCell,
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
    restoreViewportTarget,
    rowAxis,
    rowHeights,
    scrollTransformRef,
    scrollTransformStore,
    scrollViewportRef,
    selectedCell,
    selectedCellSnapshot,
    setVisibleRegion,
    sheetId,
    sheetOrdinal,
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    syncRuntimeAxes,
    visibleRegion,
  } = input
  const dprBucket = useGridDrawDprBucket()
  const shouldUseRemoteRenderTileSource = resolveShouldUseRemoteRenderTileSource({ renderTileSource, sheetId })
  const viewportResidency = useWorkbookGridViewportRuntime({
    columnAxis,
    engine,
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
    selectedCell,
    setVisibleRegion,
    sheetName,
    shouldUseRemoteRenderTileSource,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    syncRuntimeAxes,
    visibleRegion,
    restoreViewportTarget,
  })

  const { viewport, residentViewport, renderTileViewport, residentHeaderItems, residentHeaderRegion, sceneRevision, visibleAddresses } =
    viewportResidency
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
  const getVisibleHeaderCellLocalBounds = useWorkbookHeaderCellBounds({
    columnWidths,
    freezeCols,
    freezeRows,
    frozenColumnWidth,
    frozenRowHeight,
    gridMetrics,
    residentViewport: viewport,
    rowHeights,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
  })
  const renderTileState = useWorkbookRenderTilePanes({
    columnWidths,
    dprBucket,
    editingCell,
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
    selectedCell,
    selectedCellSnapshot,
    sheetId,
    sheetOrdinal,
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    visibleAddresses,
    visibleViewport: viewport,
  })
  const useVisibleHeaderPaneWindow =
    resolveHeaderPaneWindowMode({
      residentViewport,
      viewport,
      visibleDirtyTileKeys: renderTileState.tileReadiness.visibleDirtyTileKeys,
    }) === 'visible'
  const visibleHeaderItems = useMemo(
    () =>
      collectViewportItems(viewport, {
        freezeCols,
        freezeRows,
      }),
    [freezeCols, freezeRows, viewport],
  )
  const visibleHeaderRegion = useMemo(
    () => ({
      range: {
        x: viewport.colStart,
        y: viewport.rowStart,
        width: viewport.colEnd - viewport.colStart + 1,
        height: viewport.rowEnd - viewport.rowStart + 1,
      },
      tx: 0,
      ty: 0,
      freezeRows,
      freezeCols,
    }),
    [freezeCols, freezeRows, viewport],
  )
  const visibleHeaderBodyPane = useMemo(
    () => ({
      contentOffset: { x: 0, y: 0 },
      surfaceSize: {
        width: Math.max(0, hostClientWidth - gridMetrics.rowMarkerWidth - frozenColumnWidth),
        height: Math.max(0, hostClientHeight - gridMetrics.headerHeight - frozenRowHeight),
      },
    }),
    [frozenColumnWidth, frozenRowHeight, gridMetrics.headerHeight, gridMetrics.rowMarkerWidth, hostClientHeight, hostClientWidth],
  )
  const headerPanes = useWorkbookHeaderPanes({
    columnWidths,
    freezeCols,
    freezeRows,
    frozenColumnWidth,
    frozenRowHeight,
    getHeaderCellLocalBounds: useVisibleHeaderPaneWindow ? getVisibleHeaderCellLocalBounds : getHeaderCellLocalBounds,
    gridMetrics,
    gridRuntimeHost,
    hostClientHeight,
    hostClientWidth,
    hostElement,
    residentBodyPane: useVisibleHeaderPaneWindow ? visibleHeaderBodyPane : renderTileState.residentBodyPane,
    residentHeaderItems: useVisibleHeaderPaneWindow ? visibleHeaderItems : residentHeaderItems,
    residentHeaderRegion: useVisibleHeaderPaneWindow ? visibleHeaderRegion : residentHeaderRegion,
    residentViewport: useVisibleHeaderPaneWindow ? viewport : residentViewport,
    rowHeights,
    sheetName,
  })

  return {
    headerPanes,
    preloadDataPanes: renderTileState.preloadDataPanes,
    renderTilePanes: renderTileState.renderTilePanes,
    residentDataPanes: renderTileState.residentDataPanes,
    viewport,
  }
}
