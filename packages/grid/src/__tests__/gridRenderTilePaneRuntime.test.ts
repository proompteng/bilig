import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { GridEngineLike } from '../grid-engine.js'
import { getGridMetrics } from '../gridMetrics.js'
import { GRID_RECT_INSTANCE_FLOAT_COUNT_V3 } from '../renderer-v3/rect-instance-buffer.js'
import type { GridRenderTile, GridRenderTileSource } from '../renderer-v3/render-tile-source.js'
import { GRID_TEXT_METRIC_FLOAT_COUNT_V3 } from '../renderer-v3/text-run-buffer.js'
import { GridRenderTilePaneRuntime } from '../runtime/gridRenderTilePaneRuntime.js'
import { GridRuntimeHost } from '../runtime/gridRuntimeHost.js'

const TEST_ENGINE: GridEngineLike = {
  getCell: () => {
    throw new Error('GridRenderTilePaneRuntime remote-tile tests do not read cells')
  },
  getCellStyle: () => undefined,
  subscribeCells: () => () => {},
  workbook: {
    getSheet: () => undefined,
  },
}

function createEmptyCellSnapshot(address: string): CellSnapshot {
  return {
    address,
    flags: 0,
    input: '',
    sheetName: 'Sheet1',
    value: { tag: ValueTag.Empty },
    version: 0,
  }
}

const LOCAL_EMPTY_ENGINE: GridEngineLike = {
  getCell: (_sheetName, address) => createEmptyCellSnapshot(address),
  getCellStyle: () => undefined,
  subscribeCells: () => () => {},
  workbook: {
    getSheet: () => undefined,
  },
}

function createHost(): GridRuntimeHost {
  return new GridRuntimeHost({
    columnCount: 1000,
    defaultColumnWidth: 100,
    defaultRowHeight: 20,
    gridMetrics: getGridMetrics(),
    rowCount: 1000,
    viewportHeight: 400,
    viewportWidth: 800,
  })
}

function createRenderTile(tileId: number, sheetId = 7): GridRenderTile {
  return {
    bounds: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    coord: {
      colTile: 0,
      dprBucket: 1,
      paneKind: 'body',
      rowTile: 0,
      sheetId,
    },
    lastBatchId: 1,
    lastCameraSeq: 1,
    rectCount: 0,
    rectInstances: new Float32Array(GRID_RECT_INSTANCE_FLOAT_COUNT_V3),
    textCount: 0,
    textMetrics: new Float32Array(GRID_TEXT_METRIC_FLOAT_COUNT_V3),
    textRuns: [],
    tileId,
    version: {
      axisX: 1,
      axisY: 1,
      freeze: 0,
      styles: 1,
      text: 1,
      values: 1,
    },
  }
}

function createRenderTileSource(tiles: readonly GridRenderTile[]): GridRenderTileSource {
  const byId = new Map(tiles.map((tile) => [tile.tileId, tile]))
  return {
    peekRenderTile: (tileId) => byId.get(tileId) ?? null,
    subscribeRenderTileDeltas: () => () => {},
  }
}

function createInput(
  overrides: Partial<Parameters<GridRenderTilePaneRuntime['resolve']>[0]> = {},
): Parameters<GridRenderTilePaneRuntime['resolve']>[0] {
  return {
    columnWidths: {},
    dprBucket: 1,
    engine: TEST_ENGINE,
    freezeCols: 0,
    freezeRows: 0,
    frozenColumnWidth: 0,
    frozenRowHeight: 0,
    gridMetrics: getGridMetrics(),
    gridRuntimeHost: createHost(),
    hostClientHeight: 400,
    hostClientWidth: 800,
    hostReady: true,
    renderTileSource: undefined,
    renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    residentViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    rowHeights: {},
    sceneRevision: 1,
    sheetId: 7,
    sheetName: 'Sheet1',
    sortedColumnWidthOverrides: [],
    sortedRowHeightOverrides: [],
    visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    ...overrides,
  }
}

describe('GridRenderTilePaneRuntime', () => {
  it('resolves remote render tiles into V3 pane placements', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    const state = runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileId)]),
      }),
    )

    expect(state.residentBodyPane?.tile.tileId).toBe(tileId)
    expect(state.renderTilePanes).toHaveLength(1)
    expect(state.preloadDataPanes).toHaveLength(0)
  })

  it('retains the previous same-sheet panes while a remote tile is temporarily unavailable', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    const ready = runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileId)]),
      }),
    )
    const missing = runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([]),
      }),
    )

    expect(missing.residentDataPanes).toBe(ready.residentDataPanes)
    expect(missing.residentBodyPane).toBe(ready.residentBodyPane)
  })

  it('falls back to local fixed tiles when remote tiles are unavailable before same-sheet retention exists', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        renderTileSource: createRenderTileSource([]),
      }),
    )

    expect(state.residentBodyPane?.tile.coord.sheetId).toBe(7)
    expect(state.residentDataPanes).toHaveLength(1)
  })

  it('does not retain remote panes across sheet switches or before the host is ready', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    const ready = runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileId)]),
      }),
    )

    const switched = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([]),
        sheetId: 8,
      }),
    )

    expect(switched.residentDataPanes).not.toBe(ready.residentDataPanes)
    expect(switched.residentBodyPane?.tile.coord.sheetId).toBe(8)
    expect(
      runtime.resolve(
        createInput({
          gridRuntimeHost: host,
          hostReady: false,
          renderTileSource: createRenderTileSource([]),
        }),
      ).residentDataPanes,
    ).toEqual([])
  })
})
