import { useMemo } from 'react'
import type { Viewport } from '@bilig/protocol'
import type { GridHeaderPaneState } from './gridHeaderPanes.js'
import type { GridMetrics } from './gridMetrics.js'
import type { Item, Rectangle } from './gridTypes.js'
import type { WorkbookRenderTilePaneState } from './renderer-v3/render-tile-pane-state.js'
import { buildWorkbookHeaderPaneStatesV3 } from './renderer-v3/header-pane-builder.js'

function noteHeaderPaneBuild(): void {
  if (typeof window === 'undefined') {
    return
  }
  ;(window as Window & { __biligScrollPerf?: { noteHeaderPaneBuild?: () => void } }).__biligScrollPerf?.noteHeaderPaneBuild?.()
}

export function useWorkbookHeaderPanes(input: {
  readonly columnWidths: Readonly<Record<number, number>>
  readonly freezeCols: number
  readonly freezeRows: number
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly getHeaderCellLocalBounds: (col: number, row: number) => Rectangle | undefined
  readonly gridMetrics: GridMetrics
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

  const headerPanes = useMemo(() => {
    noteHeaderPaneBuild()
    if (!hostElement) {
      return []
    }
    return buildWorkbookHeaderPaneStatesV3({
      columnWidths,
      sheetName,
      residentViewport,
      freezeCols,
      freezeRows,
      gridMetrics,
      frozenColumnWidth,
      frozenRowHeight,
      getHeaderCellLocalBounds,
      hostClientHeight,
      hostClientWidth,
      residentBodyHeight: residentBodyPane?.surfaceSize.height ?? 0,
      residentBodyWidth: residentBodyPane?.surfaceSize.width ?? 0,
      residentHeaderItems,
      residentHeaderRegion,
      rowHeights,
    })
  }, [
    columnWidths,
    frozenColumnWidth,
    frozenRowHeight,
    getHeaderCellLocalBounds,
    gridMetrics,
    hostClientHeight,
    hostClientWidth,
    hostElement,
    residentHeaderItems,
    residentHeaderRegion,
    rowHeights,
    sheetName,
    residentViewport,
    freezeCols,
    freezeRows,
    residentBodyPane?.surfaceSize.height,
    residentBodyPane?.surfaceSize.width,
  ])

  return useMemo(
    () =>
      headerPanes.map((pane) =>
        pane.paneId === 'top-body'
          ? { ...pane, contentOffset: { x: residentBodyPane?.contentOffset.x ?? 0, y: 0 } }
          : pane.paneId === 'left-body'
            ? { ...pane, contentOffset: { x: 0, y: residentBodyPane?.contentOffset.y ?? 0 } }
            : pane,
      ),
    [headerPanes, residentBodyPane?.contentOffset.x, residentBodyPane?.contentOffset.y],
  )
}
