import { useMemo } from 'react'
import type { Viewport } from '@bilig/protocol'
import type { GridHeaderPaneState } from './gridHeaderPanes.js'
import type { GridMetrics } from './gridMetrics.js'
import type { Item, Rectangle } from './gridTypes.js'
import type { WorkbookRenderTilePaneState } from './renderer-v3/render-tile-pane-state.js'
import type { GridRuntimeHost } from './runtime/gridRuntimeHost.js'

export function useWorkbookHeaderPanes(input: {
  readonly columnWidths: Readonly<Record<number, number>>
  readonly freezeCols: number
  readonly freezeRows: number
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly getHeaderCellLocalBounds: (col: number, row: number) => Rectangle | undefined
  readonly gridMetrics: GridMetrics
  readonly gridRuntimeHost: GridRuntimeHost
  readonly hostClientHeight: number
  readonly hostClientWidth: number
  readonly hostElement: HTMLDivElement | null
  readonly residentBodyPane: WorkbookRenderTilePaneState | null
  readonly residentHeaderItems: readonly Item[]
  readonly residentHeaderRegion: {
    readonly range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
    readonly tx: number
    readonly ty: number
    readonly freezeRows: number
    readonly freezeCols: number
  }
  readonly residentViewport: Viewport
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sheetName: string
}): readonly GridHeaderPaneState[] {
  const {
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
    residentBodyPane,
    residentHeaderItems,
    residentHeaderRegion,
    residentViewport,
    rowHeights,
    sheetName,
  } = input
  const residentBodyOffsetX = residentBodyPane?.contentOffset.x ?? 0
  const residentBodyOffsetY = residentBodyPane?.contentOffset.y ?? 0
  const residentBodyHeight = residentBodyPane?.surfaceSize.height ?? 0
  const residentBodyWidth = residentBodyPane?.surfaceSize.width ?? 0

  return useMemo<readonly GridHeaderPaneState[]>(() => {
    return gridRuntimeHost.resolveHeaderPanes({
      columnWidths,
      freezeCols,
      freezeRows,
      frozenColumnWidth,
      frozenRowHeight,
      getHeaderCellLocalBounds,
      gridMetrics,
      hostClientHeight,
      hostClientWidth,
      hostReady: hostElement !== null,
      residentBodyPane:
        residentBodyHeight > 0 || residentBodyWidth > 0
          ? {
              contentOffset: { x: residentBodyOffsetX, y: residentBodyOffsetY },
              surfaceSize: { width: residentBodyWidth, height: residentBodyHeight },
            }
          : null,
      residentHeaderItems,
      residentHeaderRegion,
      residentViewport,
      rowHeights,
      sheetName,
    })
  }, [
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
    residentBodyHeight,
    residentBodyOffsetX,
    residentBodyOffsetY,
    residentBodyWidth,
    residentHeaderItems,
    residentHeaderRegion,
    residentViewport,
    rowHeights,
    sheetName,
  ])
}
