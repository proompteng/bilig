import type { Dispatch, SetStateAction } from 'react'
import type { Viewport } from '@bilig/protocol'
import type { GridAxisWorldIndex } from './gridAxisWorldIndex.js'
import type { GridEngineLike } from './grid-engine.js'
import type { GridMetrics } from './gridMetrics.js'
import type { VisibleRegionState } from './gridPointer.js'
import type { Item } from './gridTypes.js'
import type { GridCameraStore } from './runtime/gridCameraStore.js'
import type { GridRuntimeHost } from './runtime/gridRuntimeHost.js'
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

export function resolveShouldUseRemoteRenderTileSource(input: {
  readonly renderTileSource?: unknown
  readonly sheetId?: number | undefined
}): boolean {
  return input.renderTileSource !== undefined && input.sheetId !== undefined
}

export function useWorkbookGridDrawRuntime(input: {
  readonly columnAxis: GridAxisWorldIndex
  readonly columnWidths: Readonly<Record<number, number>>
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
  readonly setVisibleRegion: Dispatch<SetStateAction<VisibleRegionState>>
  readonly sheetId?: number | undefined
  readonly sheetName: string
  readonly sortedColumnWidthOverrides: SortedAxisOverrides
  readonly sortedRowHeightOverrides: SortedAxisOverrides
  readonly syncRuntimeAxes: () => void
  readonly visibleRegion: VisibleRegionState
}) {
  const {
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
    restoreViewportTarget,
    rowAxis,
    rowHeights,
    scrollTransformRef,
    scrollTransformStore,
    scrollViewportRef,
    selectedCell,
    setVisibleRegion,
    sheetId,
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    syncRuntimeAxes,
    visibleRegion,
  } = input
  const dprBucket = resolveGridDrawDprBucket(typeof window === 'undefined' ? null : window)
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
  const renderTileState = useWorkbookRenderTilePanes({
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
    visibleAddresses,
    visibleViewport: viewport,
  })
  const headerPanes = useWorkbookHeaderPanes({
    columnWidths,
    freezeCols,
    freezeRows,
    frozenColumnWidth,
    frozenRowHeight,
    getHeaderCellLocalBounds,
    gridMetrics,
    gridRuntimeHost,
    hostClientHeight,
    hostClientWidth,
    hostElement,
    residentBodyPane: renderTileState.residentBodyPane,
    residentHeaderItems,
    residentHeaderRegion,
    residentViewport,
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
