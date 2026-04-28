import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { GridMetrics } from '../gridMetrics.js'
import type { Item, Rectangle } from '../gridTypes.js'
import { buildWorkbookHeaderPaneStatesV3 } from '../renderer-v3/header-pane-builder.js'
import type { WorkbookRenderTilePaneState } from '../renderer-v3/render-tile-pane-state.js'

type ResidentBodyPanePlacement = Pick<WorkbookRenderTilePaneState, 'contentOffset' | 'surfaceSize'>

export interface GridHeaderPaneRuntimeInput {
  readonly columnWidths: Readonly<Record<number, number>>
  readonly freezeCols: number
  readonly freezeRows: number
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly getHeaderCellLocalBounds: (col: number, row: number) => Rectangle | undefined
  readonly gridMetrics: GridMetrics
  readonly hostClientHeight: number
  readonly hostClientWidth: number
  readonly hostReady: boolean
  readonly residentBodyPane: ResidentBodyPanePlacement | null
  readonly residentHeaderItems: readonly Item[]
  readonly residentHeaderRegion: {
    readonly range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
    readonly tx: number
    readonly ty: number
    readonly freezeRows: number
    readonly freezeCols: number
  }
  readonly residentViewport: {
    readonly rowStart: number
    readonly rowEnd: number
    readonly colStart: number
    readonly colEnd: number
  }
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sheetName: string
}

const EMPTY_HEADER_PANES: readonly GridHeaderPaneState[] = Object.freeze([])

function noteHeaderPaneBuild(): void {
  if (typeof window === 'undefined') {
    return
  }
  ;(window as Window & { __biligScrollPerf?: { noteHeaderPaneBuild?: () => void } }).__biligScrollPerf?.noteHeaderPaneBuild?.()
}

export class GridHeaderPaneRuntime {
  resolve(input: GridHeaderPaneRuntimeInput): readonly GridHeaderPaneState[] {
    noteHeaderPaneBuild()
    if (!input.hostReady) {
      return EMPTY_HEADER_PANES
    }
    return applyResidentBodyOffsets(
      buildWorkbookHeaderPaneStatesV3({
        columnWidths: input.columnWidths,
        freezeCols: input.freezeCols,
        freezeRows: input.freezeRows,
        frozenColumnWidth: input.frozenColumnWidth,
        frozenRowHeight: input.frozenRowHeight,
        getHeaderCellLocalBounds: input.getHeaderCellLocalBounds,
        gridMetrics: input.gridMetrics,
        hostClientHeight: input.hostClientHeight,
        hostClientWidth: input.hostClientWidth,
        residentBodyHeight: input.residentBodyPane?.surfaceSize.height ?? 0,
        residentBodyWidth: input.residentBodyPane?.surfaceSize.width ?? 0,
        residentHeaderItems: input.residentHeaderItems,
        residentHeaderRegion: input.residentHeaderRegion,
        residentViewport: input.residentViewport,
        rowHeights: input.rowHeights,
        sheetName: input.sheetName,
      }),
      input.residentBodyPane,
    )
  }
}

function applyResidentBodyOffsets(
  headerPanes: readonly GridHeaderPaneState[],
  residentBodyPane: ResidentBodyPanePlacement | null,
): readonly GridHeaderPaneState[] {
  if (headerPanes.length === 0 || !residentBodyPane) {
    return headerPanes
  }
  return headerPanes.map((pane) =>
    pane.paneId === 'top-body'
      ? { ...pane, contentOffset: { x: residentBodyPane.contentOffset.x, y: 0 } }
      : pane.paneId === 'left-body'
        ? { ...pane, contentOffset: { x: 0, y: residentBodyPane.contentOffset.y } }
        : pane,
  )
}
