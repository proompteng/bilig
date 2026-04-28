import { useEffect, useMemo, useRef, useState } from 'react'
import type { Viewport } from '@bilig/protocol'
import type { GridEngineLike } from './grid-engine.js'
import type { GridMetrics } from './gridMetrics.js'
import { buildLocalFixedRenderTiles } from './renderer-v3/local-render-tile-materializer.js'
import { buildFixedRenderTilePaneStates } from './renderer-v3/render-tile-pane-builder.js'
import type { GridRenderTile, GridRenderTileSource } from './renderer-v3/render-tile-source.js'
import type { WorkbookRenderTilePaneState } from './renderer-v3/render-tile-pane-state.js'
import type { GridRuntimeHost } from './runtime/gridRuntimeHost.js'

type SortedAxisOverrides = readonly (readonly [number, number])[]

export interface WorkbookRenderTilePanesState {
  readonly preloadDataPanes: readonly WorkbookRenderTilePaneState[]
  readonly renderTilePanes: readonly WorkbookRenderTilePaneState[]
  readonly residentBodyPane: WorkbookRenderTilePaneState | null
  readonly residentDataPanes: readonly WorkbookRenderTilePaneState[]
}

function buildLocalFixedRenderTilePaneStates(input: {
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
  readonly renderTileViewport: Viewport
  readonly residentViewport: Viewport
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sceneRevision: number
  readonly sheetId: number
  readonly sheetName: string
  readonly sortedColumnWidthOverrides: SortedAxisOverrides
  readonly sortedRowHeightOverrides: SortedAxisOverrides
  readonly visibleViewport: Viewport
}): readonly WorkbookRenderTilePaneState[] | null {
  const tiles = buildLocalFixedRenderTiles({
    cameraSeq: input.gridRuntimeHost.snapshot().camera.seq,
    columnWidths: input.columnWidths,
    dprBucket: input.dprBucket,
    engine: input.engine,
    generation: input.sceneRevision,
    gridMetrics: input.gridMetrics,
    rowHeights: input.rowHeights,
    sheetId: input.sheetId,
    sheetName: input.sheetName,
    sortedColumnWidthOverrides: input.sortedColumnWidthOverrides,
    sortedRowHeightOverrides: input.sortedRowHeightOverrides,
    viewport: input.renderTileViewport,
  })
  const panes = buildFixedRenderTilePaneStates({
    freezeCols: input.freezeCols,
    freezeRows: input.freezeRows,
    frozenColumnWidth: input.frozenColumnWidth,
    frozenRowHeight: input.frozenRowHeight,
    gridMetrics: input.gridMetrics,
    hostHeight: input.hostClientHeight,
    hostWidth: input.hostClientWidth,
    residentViewport: input.residentViewport,
    sortedColumnWidthOverrides: input.sortedColumnWidthOverrides,
    sortedRowHeightOverrides: input.sortedRowHeightOverrides,
    tiles,
    visibleViewport: input.visibleViewport,
  })
  return panes.length > 0 ? panes : null
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
    visibleViewport,
  } = input
  const retainedFixedRenderTileDataPanesRef = useRef<{
    readonly sheetId: number
    readonly panes: readonly WorkbookRenderTilePaneState[]
  } | null>(null)
  const [renderTileRevision, setRenderTileRevision] = useState(0)
  const shouldUseRemoteRenderTileSource = renderTileSource !== undefined && sheetId !== undefined

  useEffect(() => {
    if (!renderTileSource || sheetId === undefined) {
      return
    }
    const tileInterest = gridRuntimeHost.buildViewportTileInterest({
      dprBucket,
      reason: 'scroll',
      sheetId,
      sheetOrdinal: sheetId,
      viewport: renderTileViewport,
    })
    return renderTileSource.subscribeRenderTileDeltas(
      {
        ...renderTileViewport,
        cameraSeq: tileInterest.cameraSeq,
        dprBucket,
        initialDelta: 'full',
        sheetId,
        sheetName,
      },
      () => {
        setRenderTileRevision((current) => current + 1)
      },
    )
  }, [dprBucket, gridRuntimeHost, renderTileSource, renderTileViewport, sheetId, sheetName])

  const fixedRenderTileDataPanes = useMemo<readonly WorkbookRenderTilePaneState[] | null>(() => {
    if (!hostElement) {
      return null
    }
    const tiles: GridRenderTile[] = []
    const tileSheetId = sheetId ?? 0
    const buildLocalPanes = () =>
      buildLocalFixedRenderTilePaneStates({
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
        renderTileViewport,
        residentViewport,
        rowHeights,
        sceneRevision,
        sheetId: tileSheetId,
        sheetName,
        sortedColumnWidthOverrides,
        sortedRowHeightOverrides,
        visibleViewport,
      })
    if (renderTileSource && sheetId !== undefined) {
      void renderTileRevision
      const tileKeys = gridRuntimeHost.viewportTileKeys({
        dprBucket,
        sheetOrdinal: sheetId,
        viewport: renderTileViewport,
      })
      for (const tileKey of tileKeys) {
        const tile = renderTileSource.peekRenderTile(tileKey)
        if (!tile || tile.coord.sheetId !== sheetId) {
          return retainedFixedRenderTileDataPanesRef.current?.sheetId === sheetId ? null : buildLocalPanes()
        }
        tiles.push(tile)
      }
    } else {
      return buildLocalPanes()
    }
    const panes = buildFixedRenderTilePaneStates({
      freezeCols,
      freezeRows,
      frozenColumnWidth,
      frozenRowHeight,
      gridMetrics,
      hostHeight: hostClientHeight,
      hostWidth: hostClientWidth,
      residentViewport,
      sortedColumnWidthOverrides,
      sortedRowHeightOverrides,
      tiles,
      visibleViewport,
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
    gridRuntimeHost,
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
    visibleViewport,
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
  const preloadDataPanes = useMemo<readonly WorkbookRenderTilePaneState[]>(() => [], [])

  return {
    preloadDataPanes,
    renderTilePanes: residentDataPanes,
    residentBodyPane,
    residentDataPanes,
  }
}
