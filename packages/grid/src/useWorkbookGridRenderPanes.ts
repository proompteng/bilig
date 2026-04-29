import type { GridEngineLike } from './grid-engine.js'
import type { GridMetrics } from './gridMetrics.js'
import type { GridRuntimeHost } from './runtime/gridRuntimeHost.js'
import type { SheetGridViewportSubscription } from './workbookGridSurfaceTypes.js'
import type { GridRenderTileSource } from './renderer-v3/render-tile-source.js'
import { useWorkbookHeaderCellBounds } from './useWorkbookHeaderCellBounds.js'
import { useWorkbookHeaderPanes } from './useWorkbookHeaderPanes.js'
import { useWorkbookRenderTilePanes } from './useWorkbookRenderTilePanes.js'
import type { WorkbookViewportResidencyState } from './useWorkbookViewportResidencyState.js'

type SortedAxisOverrides = readonly (readonly [number, number])[]

export function useWorkbookGridRenderPanes(input: {
  readonly columnWidths: Readonly<Record<number, number>>
  readonly dprBucket: number
  readonly engine: GridEngineLike
  readonly freezeCols: number
  readonly freezeRows: number
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly gridMetrics: GridMetrics
  readonly gridRuntimeHost: GridRuntimeHost
  readonly hostClientHeight: number
  readonly hostClientWidth: number
  readonly hostElement: HTMLDivElement | null
  readonly renderTileSource?: GridRenderTileSource | undefined
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sheetId?: number | undefined
  readonly sheetName: string
  readonly sortedColumnWidthOverrides: SortedAxisOverrides
  readonly sortedRowHeightOverrides: SortedAxisOverrides
  readonly subscribeViewport?: SheetGridViewportSubscription | undefined
  readonly viewportResidency: WorkbookViewportResidencyState
}) {
  const {
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
    rowHeights,
    sheetId,
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    subscribeViewport,
    viewportResidency,
  } = input
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
    subscribeViewport,
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
