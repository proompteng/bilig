import { useEffect, useMemo, useRef, useState } from 'react'
import type { Viewport } from '@bilig/protocol'
import type { GridEngineLike } from './grid-engine.js'
import type { GridMetrics } from './gridMetrics.js'
import type { GridRenderTileSource } from './renderer-v3/render-tile-source.js'
import type { WorkbookRenderTilePaneState } from './renderer-v3/render-tile-pane-state.js'
import { GridRenderTilePaneRuntime } from './runtime/gridRenderTilePaneRuntime.js'
import type { GridRuntimeHost } from './runtime/gridRuntimeHost.js'

type SortedAxisOverrides = readonly (readonly [number, number])[]

export interface WorkbookRenderTilePanesState {
  readonly preloadDataPanes: readonly WorkbookRenderTilePaneState[]
  readonly renderTilePanes: readonly WorkbookRenderTilePaneState[]
  readonly residentBodyPane: WorkbookRenderTilePaneState | null
  readonly residentDataPanes: readonly WorkbookRenderTilePaneState[]
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
  const tilePaneRuntimeRef = useRef<GridRenderTilePaneRuntime | null>(null)
  tilePaneRuntimeRef.current ??= new GridRenderTilePaneRuntime()
  const [renderTileRevision, setRenderTileRevision] = useState(0)

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

  return useMemo<WorkbookRenderTilePanesState>(() => {
    void renderTileRevision
    return tilePaneRuntimeRef.current!.resolve({
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
}
