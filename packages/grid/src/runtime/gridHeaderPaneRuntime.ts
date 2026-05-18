import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { GridMetrics } from '../gridMetrics.js'
import type { Item, Rectangle } from '../gridTypes.js'
import { getResolvedCellFontFamily } from '../gridCells.js'
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

interface GridHeaderPaneRuntimeCache {
  readonly key: string
  readonly panes: readonly GridHeaderPaneState[]
}

function noteHeaderPaneBuild(): void {
  if (typeof window === 'undefined') {
    return
  }
  ;(window as Window & { __biligScrollPerf?: { noteHeaderPaneBuild?: () => void } }).__biligScrollPerf?.noteHeaderPaneBuild?.()
}

export class GridHeaderPaneRuntime {
  private cache: GridHeaderPaneRuntimeCache | null = null

  resolve(input: GridHeaderPaneRuntimeInput): readonly GridHeaderPaneState[] {
    const key = getHeaderPaneRuntimeCacheKey(input)
    if (this.cache?.key === key) {
      return this.cache.panes
    }

    noteHeaderPaneBuild()
    if (!input.hostReady) {
      this.cache = { key, panes: EMPTY_HEADER_PANES }
      return EMPTY_HEADER_PANES
    }
    const panes = applyResidentBodyOffsets(
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
    this.cache = { key, panes }
    return panes
  }
}

export function getGridHeaderPaneRuntime(current: unknown): GridHeaderPaneRuntime {
  return current instanceof GridHeaderPaneRuntime ? current : new GridHeaderPaneRuntime()
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

function getHeaderPaneRuntimeCacheKey(input: GridHeaderPaneRuntimeInput): string {
  if (!input.hostReady) {
    return stableParts('not-ready', input.sheetName, input.hostClientWidth, input.hostClientHeight)
  }

  return stableParts(
    'ready',
    input.sheetName,
    input.freezeCols,
    input.freezeRows,
    input.frozenColumnWidth,
    input.frozenRowHeight,
    input.hostClientHeight,
    input.hostClientWidth,
    input.gridMetrics.columnWidth,
    input.gridMetrics.rowHeight,
    input.gridMetrics.headerHeight,
    input.gridMetrics.rowMarkerWidth,
    getResolvedCellFontFamily(),
    recordSignature(input.columnWidths),
    recordSignature(input.rowHeights),
    residentBodyPaneSignature(input.residentBodyPane),
    residentHeaderRegionSignature(input.residentHeaderRegion),
    viewportSignature(input.residentViewport),
    itemsSignature(input.residentHeaderItems),
    headerBoundsSignature(input),
  )
}

function stableParts(...parts: readonly (number | string)[]): string {
  return parts.join('|')
}

function recordSignature(record: Readonly<Record<number, number>>): string {
  const keys = Object.keys(record)
  if (keys.length === 0) {
    return ''
  }
  return keys
    .map(Number)
    .toSorted((a, b) => a - b)
    .map((key) => `${key}:${record[key]}`)
    .join(',')
}

function residentBodyPaneSignature(pane: ResidentBodyPanePlacement | null): string {
  return pane ? stableParts(pane.contentOffset.x, pane.contentOffset.y, pane.surfaceSize.width, pane.surfaceSize.height) : 'none'
}

function residentHeaderRegionSignature(region: GridHeaderPaneRuntimeInput['residentHeaderRegion']): string {
  return stableParts(
    region.range.x,
    region.range.y,
    region.range.width,
    region.range.height,
    region.tx,
    region.ty,
    region.freezeRows,
    region.freezeCols,
  )
}

function viewportSignature(viewport: GridHeaderPaneRuntimeInput['residentViewport']): string {
  return stableParts(viewport.colStart, viewport.colEnd, viewport.rowStart, viewport.rowEnd)
}

function itemsSignature(items: readonly Item[]): string {
  return items.map(([col, row]) => `${col}:${row}`).join(',')
}

function headerBoundsSignature(input: GridHeaderPaneRuntimeInput): string {
  return input.residentHeaderItems
    .map(([col, row]) => {
      const bounds = input.getHeaderCellLocalBounds(col, row)
      return bounds ? stableParts(col, row, bounds.x, bounds.y, bounds.width, bounds.height) : `${col}:${row}:missing`
    })
    .join(',')
}
