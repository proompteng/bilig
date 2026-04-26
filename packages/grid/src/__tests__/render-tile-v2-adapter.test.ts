import { describe, expect, it } from 'vitest'
import { getGridMetrics } from '../gridMetrics.js'
import { buildFixedRenderTileDataPaneStates } from '../renderer-v3/render-tile-v2-adapter.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import { validateGridScenePacketV2 } from '../renderer-v2/scene-packet-validator.js'

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

describe('render tile v2 adapter', () => {
  it('adapts fixed content tiles into valid mounted body panes', () => {
    const gridMetrics = getGridMetrics()
    const panes = buildFixedRenderTileDataPaneStates({
      columnWidths: {},
      freezeCols: 0,
      freezeRows: 0,
      frozenColumnWidth: 0,
      frozenRowHeight: 0,
      gridMetrics,
      hostHeight: 400,
      hostWidth: 600,
      residentViewport: { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
      rowHeights: {},
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      tiles: [createTile({ tileId: 101, rowTile: 0, colTile: 0, rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 })],
      visibleViewport: { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
    })

    expect(panes).toHaveLength(1)
    const pane = panes[0]
    if (!pane) {
      throw new Error('expected one render tile pane')
    }
    expect(pane.paneId).toBe('body')
    expect(validateGridScenePacketV2(pane.packedScene)).toEqual({ ok: true })
    expect(pane.packedScene.key.valueVersion).toBe(16)
  })

  it('reuses the same content packet for frozen placements', () => {
    const gridMetrics = getGridMetrics()
    const panes = buildFixedRenderTileDataPaneStates({
      columnWidths: {},
      freezeCols: 1,
      freezeRows: 1,
      frozenColumnWidth: gridMetrics.columnWidth,
      frozenRowHeight: gridMetrics.rowHeight,
      gridMetrics,
      hostHeight: 400,
      hostWidth: 600,
      residentViewport: { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
      rowHeights: {},
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      tiles: [createTile({ tileId: 101, rowTile: 0, colTile: 0, rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 })],
      visibleViewport: { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
    })

    const body = panes.find((pane) => pane.paneId === 'body')
    const top = panes.find((pane) => pane.paneId.startsWith('top:'))
    const left = panes.find((pane) => pane.paneId.startsWith('left:'))
    const corner = panes.find((pane) => pane.paneId.startsWith('corner:'))

    expect(body).toBeDefined()
    expect(top?.packedScene).toBe(body?.packedScene)
    expect(left?.packedScene).toBe(body?.packedScene)
    expect(corner?.packedScene).toBe(body?.packedScene)
  })

  it('keeps frozen placements anchored when the body reference tile is scrolled down and right', () => {
    const gridMetrics = getGridMetrics()
    const panes = buildFixedRenderTileDataPaneStates({
      columnWidths: {},
      freezeCols: 1,
      freezeRows: 1,
      frozenColumnWidth: gridMetrics.columnWidth,
      frozenRowHeight: gridMetrics.rowHeight,
      gridMetrics,
      hostHeight: 400,
      hostWidth: 600,
      residentViewport: { rowStart: 32, rowEnd: 63, colStart: 128, colEnd: 255 },
      rowHeights: {},
      sheetName: 'Sheet1',
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
