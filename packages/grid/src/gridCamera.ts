import type { Viewport } from '@bilig/protocol'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import type { GridCameraSnapshot } from './renderer/grid-render-contract.js'
import type { getGridMetrics } from './gridMetrics.js'
import { resolveResidentViewport, resolveVisibleRegionFromScroll } from './workbookGridViewport.js'

export interface GridCameraInput {
  readonly scrollLeft: number
  readonly scrollTop: number
  readonly viewportWidth: number
  readonly viewportHeight: number
  readonly dpr: number
  readonly freezeRows?: number
  readonly freezeCols?: number
  readonly columnWidths: Readonly<Record<number, number>>
  readonly rowHeights: Readonly<Record<number, number>>
  readonly gridMetrics: ReturnType<typeof getGridMetrics>
  readonly previous?: GridCameraSnapshot | null
  readonly updatedAt?: number
}

export function createGridCameraSnapshot(input: GridCameraInput): GridCameraSnapshot {
  const updatedAt = input.updatedAt ?? performance.now()
  const visibleRegion = resolveVisibleRegionFromScroll({
    columnWidths: input.columnWidths,
    ...(input.freezeCols === undefined ? {} : { freezeCols: input.freezeCols }),
    ...(input.freezeRows === undefined ? {} : { freezeRows: input.freezeRows }),
    gridMetrics: input.gridMetrics,
    rowHeights: input.rowHeights,
    scrollLeft: input.scrollLeft,
    scrollTop: input.scrollTop,
    viewportHeight: input.viewportHeight,
    viewportWidth: input.viewportWidth,
  })
  const visibleViewport = viewportFromCameraRegion(visibleRegion.range)
  const elapsedMs = input.previous ? Math.max(1, updatedAt - input.previous.updatedAt) : 1
  return {
    dpr: Math.max(1, input.dpr),
    residentViewport: resolveResidentViewport(visibleViewport),
    scrollLeft: input.scrollLeft,
    scrollTop: input.scrollTop,
    tx: visibleRegion.tx,
    ty: visibleRegion.ty,
    updatedAt,
    velocityX: input.previous ? ((input.scrollLeft - input.previous.scrollLeft) / elapsedMs) * 1000 : 0,
    velocityY: input.previous ? ((input.scrollTop - input.previous.scrollTop) / elapsedMs) * 1000 : 0,
    viewportHeight: input.viewportHeight,
    viewportWidth: input.viewportWidth,
    visibleViewport,
  }
}

function viewportFromCameraRegion(range: {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}): Viewport {
  return {
    colEnd: Math.min(MAX_COLS - 1, range.x + range.width - 1),
    colStart: Math.max(0, range.x),
    rowEnd: Math.min(MAX_ROWS - 1, range.y + range.height - 1),
    rowStart: Math.max(0, range.y),
  }
}
