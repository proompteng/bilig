import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { GridEngineLike } from '../grid-engine.js'
import { getGridMetrics } from '../gridMetrics.js'
import { GRID_RECT_INSTANCE_FLOAT_COUNT_V3 } from '../renderer-v3/rect-instance-buffer.js'
import type { GridRenderTile, GridRenderTileDeltaSubscription, GridRenderTileSource } from '../renderer-v3/render-tile-source.js'
import { GRID_TEXT_METRIC_FLOAT_COUNT_V3 } from '../renderer-v3/text-run-buffer.js'
import { packTileKey53 } from '../renderer-v3/tile-key.js'
import { GridRenderTilePaneRuntime, getGridRenderTilePaneRuntime } from '../runtime/gridRenderTilePaneRuntime.js'
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

function createCapturingRenderTileSource(): {
  readonly source: GridRenderTileSource
  readonly captured: () => GridRenderTileDeltaSubscription | null
  readonly unsubscribed: () => boolean
} {
  let captured: GridRenderTileDeltaSubscription | null = null
  let unsubscribed = false
  return {
    captured: () => captured,
    source: {
      peekRenderTile: () => null,
      subscribeRenderTileDeltas: (subscription) => {
        captured = subscription
        return () => {
          unsubscribed = true
        }
      },
    },
    unsubscribed: () => unsubscribed,
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
  it('replaces stale runtime refs from live reloads', () => {
    const runtime = new GridRenderTilePaneRuntime()

    expect(getGridRenderTilePaneRuntime(runtime)).toBe(runtime)
    expect(getGridRenderTilePaneRuntime({})).toBeInstanceOf(GridRenderTilePaneRuntime)
    expect(getGridRenderTilePaneRuntime(null)).toBeInstanceOf(GridRenderTilePaneRuntime)
  })

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
    expect(state.needsLocalCellInvalidation).toBe(true)
    expect(state.renderTilePanes).toHaveLength(1)
    expect(state.preloadDataPanes).toHaveLength(0)
    expect(state.tileReadiness).toMatchObject({
      exactHits: [tileId],
      misses: [],
      staleHits: [],
      visibleDirtyTileKeys: [],
    })
    expect(host.tiles.residency.getExact(tileId)?.packet).toMatchObject({
      tileId,
      version: {
        values: 1,
      },
    })
  })

  it('refreshes host-owned tile revisions when remote tile contents update', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]

    runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileId)]),
      }),
    )
    runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([
          {
            ...createRenderTile(tileId),
            version: {
              axisX: 5,
              axisY: 6,
              freeze: 7,
              styles: 8,
              text: 9,
              values: 10,
            },
          },
        ]),
      }),
    )

    expect(host.tiles.residency.getExact(tileId)).toMatchObject({
      axisSeqX: 5,
      axisSeqY: 6,
      freezeSeq: 7,
      styleSeq: 8,
      textSeq: 9,
      valueSeq: 10,
    })
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
    expect(missing.needsLocalCellInvalidation).toBe(true)
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
    expect(state.needsLocalCellInvalidation).toBe(true)
    expect(state.residentDataPanes).toHaveLength(1)
  })

  it('owns render tile delta subscription stamping in the runtime', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const gridRuntimeHost = createHost()
    const renderTileSource = createCapturingRenderTileSource()

    const unsubscribe = runtime.connectRenderTileDeltas(
      {
        dprBucket: 2,
        gridRuntimeHost,
        renderTileSource: renderTileSource.source,
        renderTileViewport: { colEnd: 255, colStart: 0, rowEnd: 63, rowStart: 0 },
        sheetId: 7,
        sheetName: 'Sheet1',
      },
      () => {},
    )

    expect(renderTileSource.captured()).toMatchObject({
      cameraSeq: gridRuntimeHost.snapshot().camera.seq,
      colEnd: 255,
      colStart: 0,
      dprBucket: 2,
      initialDelta: 'full',
      rowEnd: 63,
      rowStart: 0,
      sheetId: 7,
      sheetName: 'Sheet1',
    })
    expect(renderTileSource.captured()?.warmTileKeys).toContain(
      packTileKey53({
        colTile: 2,
        dprBucket: 2,
        rowTile: 0,
        sheetOrdinal: 7,
      }),
    )
    expect(renderTileSource.captured()?.warmTileKeys).not.toContain(
      packTileKey53({
        colTile: 0,
        dprBucket: 2,
        rowTile: 0,
        sheetOrdinal: 7,
      }),
    )
    unsubscribe?.()
    expect(renderTileSource.unsubscribed()).toBe(true)
  })

  it('skips render tile delta subscription until a remote source and sheet id exist', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const renderTileSource = createCapturingRenderTileSource()
    const input = {
      dprBucket: 1,
      gridRuntimeHost: createHost(),
      renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
      sheetName: 'Sheet1',
    }

    expect(
      runtime.connectRenderTileDeltas(
        {
          ...input,
          renderTileSource: undefined,
          sheetId: 7,
        },
        () => {},
      ),
    ).toBeUndefined()
    expect(
      runtime.connectRenderTileDeltas(
        {
          ...input,
          renderTileSource: renderTileSource.source,
          sheetId: undefined,
        },
        () => {},
      ),
    ).toBeUndefined()
    expect(renderTileSource.captured()).toBeNull()
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
