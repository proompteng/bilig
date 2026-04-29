import { useEffect, useMemo, useState } from 'react'
import type { Viewport } from '@bilig/protocol'
import type { GridEngineLike } from './grid-engine.js'
import type { GridMetrics } from './gridMetrics.js'
import type { GridRenderTileSource } from './renderer-v3/render-tile-source.js'
import type { WorkbookRenderTilePaneState } from './renderer-v3/render-tile-pane-state.js'
import type { GridTileReadinessSnapshotV3 } from './runtime/gridTileCoordinator.js'
import type { GridRuntimeHost } from './runtime/gridRuntimeHost.js'

type SortedAxisOverrides = readonly (readonly [number, number])[]

export interface WorkbookRenderTilePanesState {
  readonly preloadDataPanes: readonly WorkbookRenderTilePaneState[]
  readonly renderTilePanes: readonly WorkbookRenderTilePaneState[]
  readonly residentBodyPane: WorkbookRenderTilePaneState | null
  readonly residentDataPanes: readonly WorkbookRenderTilePaneState[]
  readonly tileReadiness: GridTileReadinessSnapshotV3
}

export function useWorkbookRenderTilePanes(input: {
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
  readonly renderTileViewport: Viewport
  readonly residentViewport: Viewport
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sceneRevision: number
  readonly sheetId?: number | undefined
  readonly sheetName: string
  readonly sortedColumnWidthOverrides: SortedAxisOverrides
  readonly sortedRowHeightOverrides: SortedAxisOverrides
  readonly visibleAddresses: readonly string[]
  readonly visibleViewport: Viewport
}): WorkbookRenderTilePanesState {
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
    renderTileViewport,
    residentViewport,
    rowHeights,
    sceneRevision,
    sheetId,
    sheetName,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    visibleAddresses,
    visibleViewport,
  } = input
  const [forceLocalTiles, setForceLocalTiles] = useState(false)
  const [renderTileRevision, setRenderTileRevision] = useState(0)
  const [localFallbackRevision, setLocalFallbackRevision] = useState(0)

  useEffect(() => {
    return gridRuntimeHost.connectRenderTileDeltas(
      {
        dprBucket,
        renderTileSource,
        renderTileViewport,
        sheetId,
        sheetName,
      },
      () => {
        setForceLocalTiles(false)
        setRenderTileRevision((current) => current + 1)
      },
    )
  }, [dprBucket, gridRuntimeHost, renderTileSource, renderTileViewport, sheetId, sheetName])

  useEffect(() => {
    return gridRuntimeHost.connectWorkbookDeltaDamage(
      {
        dprBucket,
        renderTileSource,
        sheetId,
      },
      () => {
        setRenderTileRevision((current) => current + 1)
      },
    )
  }, [dprBucket, gridRuntimeHost, renderTileSource, sheetId])

  const state = useMemo<WorkbookRenderTilePanesState & { readonly needsLocalCellInvalidation: boolean }>(() => {
    void renderTileRevision
    void localFallbackRevision
    return gridRuntimeHost.resolveRenderTilePanes({
      columnWidths,
      dprBucket,
      engine,
      freezeCols,
      freezeRows,
      forceLocalTiles,
      frozenColumnWidth,
      frozenRowHeight,
      gridMetrics,
      hostClientHeight,
      hostClientWidth,
      hostReady: hostElement !== null,
      renderTileSource,
      renderTileViewport,
      residentViewport,
      rowHeights,
      sceneRevision,
      sheetId,
      sheetName,
      sortedColumnWidthOverrides,
      sortedRowHeightOverrides,
      visibleViewport,
    })
  }, [
    columnWidths,
    dprBucket,
    engine,
    freezeCols,
    freezeRows,
    forceLocalTiles,
    frozenColumnWidth,
    frozenRowHeight,
    gridMetrics,
    gridRuntimeHost,
    hostClientHeight,
    hostClientWidth,
    hostElement,
    localFallbackRevision,
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
    visibleViewport,
  ])

  useEffect(() => {
    gridRuntimeHost.noteRenderTileReadiness(state.tileReadiness)
  }, [gridRuntimeHost, state.tileReadiness])

  useEffect(() => {
    return gridRuntimeHost.connectLocalRenderTileCellInvalidation(
      {
        engine,
        needsLocalCellInvalidation: state.needsLocalCellInvalidation,
        sheetName,
        visibleAddresses,
      },
      () => {
        setForceLocalTiles(true)
        setLocalFallbackRevision((current) => current + 1)
      },
    )
  }, [engine, gridRuntimeHost, sheetName, state.needsLocalCellInvalidation, visibleAddresses])

  return state
}
