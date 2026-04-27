import { describe, expect, it } from 'vitest'
import { getGridMetrics } from '../gridMetrics.js'
import { buildFixedRenderTilePaneStates } from '../renderer-v3/render-tile-pane-builder.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'

function createTile(input: {
  readonly tileId: number
  readonly sheetId?: number
  readonly rowTile: number
  readonly colTile: number
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
}): GridRenderTile {
  return {
    bounds: {
      rowStart: input.rowStart,
      rowEnd: input.rowEnd,
      colStart: input.colStart,
      colEnd: input.colEnd,
    },
    coord: {
      colTile: input.colTile,
      dprBucket: 1,
      paneKind: 'body',
      rowTile: input.rowTile,
      sheetId: input.sheetId ?? 7,
    },
    lastBatchId: 3,
    lastCameraSeq: 5,
    rectCount: 0,
    rectInstances: new Float32Array(20),
    textCount: 0,
    textMetrics: new Float32Array(8),
    textRuns: [],
    tileId: input.tileId,
    version: {
      axisX: 11,
      axisY: 12,
      freeze: 13,
      styles: 14,
      text: 15,
      values: 16,
    },
  }
}

describe('render tile pane builder', () => {
  it('places fixed content tiles as mounted body panes without V2 scene packets', () => {
    const gridMetrics = getGridMetrics()
    const tile = createTile({ tileId: 101, rowTile: 0, colTile: 0, rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 })
    const panes = buildFixedRenderTilePaneStates({
      freezeCols: 0,
      freezeRows: 0,
      frozenColumnWidth: 0,
      frozenRowHeight: 0,
      gridMetrics,
      hostHeight: 400,
      hostWidth: 600,
      residentViewport: { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      tiles: [tile],
      visibleViewport: { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
    })

    expect(panes).toHaveLength(1)
    const pane = panes[0]
    if (!pane) {
      throw new Error('expected one render tile pane')
    }
    expect(pane.paneId).toBe('body')
    expect(pane.tile).toBe(tile)
    expect('packedScene' in pane).toBe(false)
    expect(pane.tile.version.values).toBe(16)
  })

  it('reuses the same content tile for frozen placements', () => {
    const gridMetrics = getGridMetrics()
    const tile = createTile({ tileId: 101, rowTile: 0, colTile: 0, rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 })
    const panes = buildFixedRenderTilePaneStates({
      freezeCols: 1,
      freezeRows: 1,
      frozenColumnWidth: gridMetrics.columnWidth,
      frozenRowHeight: gridMetrics.rowHeight,
      gridMetrics,
      hostHeight: 400,
      hostWidth: 600,
      residentViewport: { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      tiles: [tile],
      visibleViewport: { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
    })

    const body = panes.find((pane) => pane.paneId === 'body')
    const top = panes.find((pane) => pane.paneId.startsWith('top:'))
    const left = panes.find((pane) => pane.paneId.startsWith('left:'))
    const corner = panes.find((pane) => pane.paneId.startsWith('corner:'))

    expect(body).toBeDefined()
    expect(top?.tile).toBe(body?.tile)
    expect(left?.tile).toBe(body?.tile)
    expect(corner?.tile).toBe(body?.tile)
  })

  it('keeps frozen placements anchored when the body reference tile is scrolled down and right', () => {
    const gridMetrics = getGridMetrics()
    const panes = buildFixedRenderTilePaneStates({
      freezeCols: 1,
      freezeRows: 1,
      frozenColumnWidth: gridMetrics.columnWidth,
      frozenRowHeight: gridMetrics.rowHeight,
      gridMetrics,
      hostHeight: 400,
      hostWidth: 600,
      residentViewport: { rowStart: 32, rowEnd: 63, colStart: 128, colEnd: 255 },
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      tiles: [
        createTile({ tileId: 101, rowTile: 0, colTile: 0, rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 }),
        createTile({ tileId: 102, rowTile: 0, colTile: 1, rowStart: 0, rowEnd: 31, colStart: 128, colEnd: 255 }),
        createTile({ tileId: 103, rowTile: 1, colTile: 0, rowStart: 32, rowEnd: 63, colStart: 0, colEnd: 127 }),
        createTile({ tileId: 104, rowTile: 1, colTile: 1, rowStart: 32, rowEnd: 63, colStart: 128, colEnd: 255 }),
      ],
      visibleViewport: { rowStart: 32, rowEnd: 63, colStart: 128, colEnd: 255 },
    })

    const body = panes.find((pane) => pane.paneId === 'body')
    const top = panes.find((pane) => pane.paneId === 'top:0:1')
    const left = panes.find((pane) => pane.paneId === 'left:1:0')
    const corner = panes.find((pane) => pane.paneId === 'corner:0:0')

    expect(body?.contentOffset).toEqual({ x: 0, y: 0 })
    expect(top?.contentOffset.y).toBe(0)
    expect(top?.contentOffset.x).toBe(0)
    expect(left?.contentOffset.x).toBe(0)
    expect(left?.contentOffset.y).toBe(0)
    expect(corner?.contentOffset).toEqual({ x: 0, y: 0 })
  })
})
