import type { Viewport } from '@bilig/protocol'
import { resolveRowOffset, type GridMetrics } from '../gridMetrics.js'
import { resolveColumnOffset } from '../workbookGridViewport.js'
import { getPaneFrame, resolvePaneLayout } from './pane-layout.js'
import type { GridRenderTile } from './render-tile-source.js'
import type { WorkbookRenderTilePaneState, WorkbookTilePaneScrollAxes, WorkbookTilePaneSurfaceSize } from './render-tile-pane-state.js'

interface AxisPlacementInput {
  readonly sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  readonly sortedRowHeightOverrides: readonly (readonly [number, number])[]
  readonly gridMetrics: GridMetrics
}

export function buildFixedRenderTilePaneStates(input: {
  readonly tiles: readonly GridRenderTile[]
  readonly residentViewport: Viewport
  readonly visibleViewport: Viewport
  readonly freezeRows: number
  readonly freezeCols: number
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly hostWidth: number
  readonly hostHeight: number
  readonly gridMetrics: GridMetrics
  readonly sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  readonly sortedRowHeightOverrides: readonly (readonly [number, number])[]
}): readonly WorkbookRenderTilePaneState[] {
  const layout = resolvePaneLayout({
    frozenColumnWidth: input.frozenColumnWidth,
    frozenRowHeight: input.frozenRowHeight,
    headerHeight: input.gridMetrics.headerHeight,
    hostHeight: input.hostHeight,
    hostWidth: input.hostWidth,
    rowMarkerWidth: input.gridMetrics.rowMarkerWidth,
  })
  const bodyFrame = getPaneFrame(layout, 'body')
  if (bodyFrame.width <= 0 || bodyFrame.height <= 0) {
    return []
  }

  const bodyTiles = input.tiles.filter((tile) => intersects(tile.bounds, input.residentViewport))
  if (bodyTiles.length === 0) {
    return []
  }

  const bodyReference = resolveBodyReference(bodyTiles)
  const residentSurfaceSize = resolveViewportSurfaceSize(input.residentViewport, input)
  const panes: WorkbookRenderTilePaneState[] = []

  bodyTiles.forEach((tile, index) => {
    panes.push(
      buildPlacementPane({
        frame: bodyFrame,
        id: index === 0 ? 'body' : `body:${tile.coord.rowTile}:${tile.coord.colTile}`,
        reference: bodyReference.bounds,
        scrollAxes: { x: true, y: true },
        surfaceSize: index === 0 ? residentSurfaceSize : resolveViewportSurfaceSize(tile.bounds, input),
        tile,
        ...input,
      }),
    )
  })

  if (input.freezeRows > 0 && input.frozenRowHeight > 0) {
    const topViewport = {
      rowStart: 0,
      rowEnd: Math.max(0, input.freezeRows - 1),
      colStart: input.residentViewport.colStart,
      colEnd: input.residentViewport.colEnd,
    }
    const frame = getPaneFrame(layout, 'top')
    input.tiles
      .filter((tile) => intersects(tile.bounds, topViewport))
      .forEach((tile) => {
        panes.push(
          buildPlacementPane({
            frame,
            id: `top:${tile.coord.rowTile}:${tile.coord.colTile}`,
            reference: bodyReference.bounds,
            scrollAxes: { x: true, y: false },
            tile,
            ...input,
          }),
        )
      })
  }

  if (input.freezeCols > 0 && input.frozenColumnWidth > 0) {
    const leftViewport = {
      rowStart: input.residentViewport.rowStart,
      rowEnd: input.residentViewport.rowEnd,
      colStart: 0,
      colEnd: Math.max(0, input.freezeCols - 1),
    }
    const frame = getPaneFrame(layout, 'left')
    input.tiles
      .filter((tile) => intersects(tile.bounds, leftViewport))
      .forEach((tile) => {
        panes.push(
          buildPlacementPane({
            frame,
            id: `left:${tile.coord.rowTile}:${tile.coord.colTile}`,
            reference: bodyReference.bounds,
            scrollAxes: { x: false, y: true },
            tile,
            ...input,
          }),
        )
      })
  }

  if (input.freezeRows > 0 && input.freezeCols > 0 && input.frozenColumnWidth > 0 && input.frozenRowHeight > 0) {
    const cornerViewport = {
      rowStart: 0,
      rowEnd: Math.max(0, input.freezeRows - 1),
      colStart: 0,
      colEnd: Math.max(0, input.freezeCols - 1),
    }
    const frame = getPaneFrame(layout, 'corner')
    input.tiles
      .filter((tile) => intersects(tile.bounds, cornerViewport))
      .forEach((tile) => {
        panes.push(
          buildPlacementPane({
            frame,
            id: `corner:${tile.coord.rowTile}:${tile.coord.colTile}`,
            reference: bodyReference.bounds,
            scrollAxes: { x: false, y: false },
            tile,
            ...input,
          }),
        )
      })
  }

  return panes
}

function buildPlacementPane(
  input: AxisPlacementInput & {
    readonly frame: WorkbookRenderTilePaneState['frame']
    readonly id: string
    readonly reference: Viewport
    readonly scrollAxes: WorkbookTilePaneScrollAxes
    readonly surfaceSize?: WorkbookTilePaneSurfaceSize | undefined
    readonly tile: GridRenderTile
  },
): WorkbookRenderTilePaneState {
  const tileX = resolveColumnOffset(input.tile.bounds.colStart, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth)
  const tileY = resolveRowOffset(input.tile.bounds.rowStart, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight)
  const referenceX = resolveColumnOffset(input.reference.colStart, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth)
  const referenceY = resolveRowOffset(input.reference.rowStart, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight)
  return {
    contentOffset: {
      x: input.scrollAxes.x ? tileX - referenceX : tileX,
      y: input.scrollAxes.y ? tileY - referenceY : tileY,
    },
    frame: input.frame,
    generation: input.tile.lastBatchId,
    paneId: input.id,
    scrollAxes: input.scrollAxes,
    surfaceSize: input.surfaceSize ?? resolveViewportSurfaceSize(input.tile.bounds, input),
    tile: input.tile,
    viewport: input.tile.bounds,
  }
}

function resolveBodyReference(tiles: readonly GridRenderTile[]): GridRenderTile {
  let best = tiles[0]!
  for (let index = 1; index < tiles.length; index += 1) {
    const tile = tiles[index]!
    if (
      tile.bounds.rowStart < best.bounds.rowStart ||
      (tile.bounds.rowStart === best.bounds.rowStart && tile.bounds.colStart < best.bounds.colStart)
    ) {
      best = tile
    }
  }
  return best
}

function resolveViewportSurfaceSize(viewport: Viewport, input: AxisPlacementInput): { readonly width: number; readonly height: number } {
  return {
    width:
      resolveColumnOffset(viewport.colEnd + 1, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth) -
      resolveColumnOffset(viewport.colStart, input.sortedColumnWidthOverrides, input.gridMetrics.columnWidth),
    height:
      resolveRowOffset(viewport.rowEnd + 1, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight) -
      resolveRowOffset(viewport.rowStart, input.sortedRowHeightOverrides, input.gridMetrics.rowHeight),
  }
}

function intersects(left: Viewport, right: Viewport): boolean {
  return left.rowStart <= right.rowEnd && left.rowEnd >= right.rowStart && left.colStart <= right.colEnd && left.colEnd >= right.colStart
}
