import { describe, expect, it } from 'vitest'
import { formatAddress } from '@bilig/formula'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { GridEngineLike } from '../grid-engine.js'
import { getGridMetrics } from '../gridMetrics.js'
import { GRID_RECT_INSTANCE_FLOAT_COUNT_V3 } from '../renderer-v3/rect-instance-buffer.js'
import type {
  GridRenderTile,
  GridRenderTileDeltaSubscription,
  GridRenderTileSceneChange,
  GridRenderTileSource,
} from '../renderer-v3/render-tile-source.js'
import { GRID_TEXT_METRIC_FLOAT_COUNT_V3 } from '../renderer-v3/text-run-buffer.js'
import { DirtyMaskV3, type WorkbookDeltaBatchLikeV3 } from '../renderer-v3/tile-damage-index.js'
import { packTileKey53 } from '../renderer-v3/tile-key.js'
import { GridRenderTilePaneRuntime, getGridRenderTilePaneRuntime } from '../runtime/gridRenderTilePaneRuntime.js'
import { GridRuntimeHost } from '../runtime/gridRuntimeHost.js'

const TEST_ENGINE: GridEngineLike = {
  getCell: (_sheetName, address) => createEmptyCellSnapshot(address),
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

function createStringCellSnapshot(address: string, value: string): CellSnapshot {
  return {
    address,
    flags: 0,
    input: value,
    sheetName: 'Sheet1',
    value: { tag: ValueTag.String, value, stringId: 0 },
    version: 1,
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

function expectedGridBorderRectCount(bounds: GridRenderTile['bounds']): number {
  return bounds.rowEnd - bounds.rowStart + 1 + bounds.colEnd - bounds.colStart + 1
}

function createGridBorderRectInstances(rectCount: number): Float32Array {
  const rectInstances = new Float32Array(rectCount * GRID_RECT_INSTANCE_FLOAT_COUNT_V3)
  for (let index = 0; index < rectCount; index += 1) {
    const offset = index * GRID_RECT_INSTANCE_FLOAT_COUNT_V3
    rectInstances[offset + 2] = index % 2 === 0 ? 100 : 1
    rectInstances[offset + 3] = index % 2 === 0 ? 1 : 20
    rectInstances[offset + 11] = 1
    rectInstances[offset + 13] = 1
  }
  return rectInstances
}

function hasOpaqueGreenFillRect(tile: GridRenderTile | undefined): boolean {
  if (!tile) {
    return false
  }
  for (let index = 0; index < tile.rectCount; index += 1) {
    const offset = index * GRID_RECT_INSTANCE_FLOAT_COUNT_V3
    const red = tile.rectInstances[offset + 4] ?? 1
    const green = tile.rectInstances[offset + 5] ?? 0
    const blue = tile.rectInstances[offset + 6] ?? 1
    const alpha = tile.rectInstances[offset + 7] ?? 0
    const instanceKind = tile.rectInstances[offset + 13] ?? -1
    if (instanceKind === 0 && red < 0.05 && green > 0.95 && blue < 0.05 && alpha > 0.95) {
      return true
    }
  }
  return false
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

function createRenderTile(tileId: number, sheetId = 7, sheetOrdinal = sheetId): GridRenderTile {
  const bounds = { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 }
  const rectCount = expectedGridBorderRectCount(bounds)
  return {
    bounds,
    coord: {
      colTile: 0,
      dprBucket: 1,
      paneKind: 'body',
      rowTile: 0,
      sheetId,
      sheetOrdinal,
    },
    lastBatchId: 1,
    lastCameraSeq: 1,
    rectCount,
    rectInstances: createGridBorderRectInstances(rectCount),
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
  readonly subscribeCount: () => number
  readonly unsubscribed: () => boolean
  readonly unsubscribeCount: () => number
} {
  let captured: GridRenderTileDeltaSubscription | null = null
  let subscribeCount = 0
  let unsubscribeCount = 0
  return {
    captured: () => captured,
    source: {
      peekRenderTile: () => null,
      subscribeRenderTileDeltas: (subscription) => {
        captured = subscription
        subscribeCount += 1
        return () => {
          unsubscribeCount += 1
        }
      },
    },
    subscribeCount: () => subscribeCount,
    unsubscribed: () => unsubscribeCount > 0,
    unsubscribeCount: () => unsubscribeCount,
  }
}

function createMutableRenderTileSource(tiles: readonly GridRenderTile[] = []): {
  readonly source: GridRenderTileSource
  readonly emit: (change: {
    readonly batchId?: number | undefined
    readonly cameraSeq?: number | undefined
    readonly changedTileIds?: readonly number[] | undefined
    readonly invalidatedTileIds?: readonly number[] | undefined
  }) => void
  readonly setTile: (tile: GridRenderTile) => void
  readonly unsubscribed: () => boolean
} {
  const byId = new Map(tiles.map((tile) => [tile.tileId, tile]))
  let listener: ((change: GridRenderTileSceneChange) => void) | null = null
  let unsubscribed = false
  return {
    emit: (change) =>
      listener?.({
        batchId: change.batchId ?? 2,
        cameraSeq: change.cameraSeq ?? 3,
        changedTileIds: change.changedTileIds ?? [],
        invalidatedTileIds: change.invalidatedTileIds ?? [],
        structural: false,
      }),
    setTile: (tile) => byId.set(tile.tileId, tile),
    source: {
      peekRenderTile: (tileId) => byId.get(tileId) ?? null,
      subscribeRenderTileDeltas: (_subscription, nextListener) => {
        listener = nextListener
        return () => {
          listener = null
          unsubscribed = true
        }
      },
    },
    unsubscribed: () => unsubscribed,
  }
}

function createWorkbookDeltaSource(): {
  readonly source: GridRenderTileSource
  readonly emit: (batch: WorkbookDeltaBatchLikeV3) => void
  readonly unsubscribed: () => boolean
} {
  let listener: ((batch: WorkbookDeltaBatchLikeV3) => void) | null = null
  let unsubscribed = false
  return {
    emit: (batch) => listener?.(batch),
    source: {
      peekRenderTile: () => null,
      subscribeRenderTileDeltas: () => () => {},
      subscribeWorkbookDeltas: (nextListener) => {
        listener = nextListener
        return () => {
          listener = null
          unsubscribed = true
        }
      },
    },
    unsubscribed: () => unsubscribed,
  }
}

function createMutableWorkbookDeltaRenderTileSource(tiles: readonly GridRenderTile[] = []): {
  readonly source: GridRenderTileSource
  readonly emitWorkbookDelta: (batch: WorkbookDeltaBatchLikeV3) => void
  readonly emitRenderTileDelta: (change: {
    readonly batchId?: number | undefined
    readonly cameraSeq?: number | undefined
    readonly changedTileIds?: readonly number[] | undefined
    readonly invalidatedTileIds?: readonly number[] | undefined
  }) => void
  readonly setTile: (tile: GridRenderTile) => void
} {
  const renderTiles = createMutableRenderTileSource(tiles)
  let workbookDeltaListener: ((batch: WorkbookDeltaBatchLikeV3) => void) | null = null
  return {
    emitRenderTileDelta: (change) => renderTiles.emit(change),
    emitWorkbookDelta: (batch) => workbookDeltaListener?.(batch),
    setTile: renderTiles.setTile,
    source: {
      peekRenderTile: (tileId) => renderTiles.source.peekRenderTile(tileId),
      subscribeRenderTileDeltas: (subscription, listener) => renderTiles.source.subscribeRenderTileDeltas(subscription, listener),
      subscribeWorkbookDeltas: (listener) => {
        workbookDeltaListener = listener
        return () => {
          workbookDeltaListener = null
        }
      },
    },
  }
}

function createWorkbookDeltaBatch(overrides: Partial<WorkbookDeltaBatchLikeV3> = {}): WorkbookDeltaBatchLikeV3 {
  return {
    dirty: {
      axisX: new Uint32Array(),
      axisY: new Uint32Array(),
      cellRanges: new Uint32Array([0, 0, 0, 0, DirtyMaskV3.Value | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
    },
    seq: 1,
    sheetId: 7,
    sheetOrdinal: 7,
    ...overrides,
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

  it('publishes bridge revisions from the runtime-owned external store', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const snapshots: unknown[] = []
    const unsubscribe = runtime.subscribeBridgeState(() => {
      snapshots.push(runtime.snapshotBridgeState())
    })

    runtime.noteRenderTileDelta()
    runtime.noteWorkbookDeltaDamage()
    runtime.noteLocalFallbackInvalidation()
    unsubscribe()
    runtime.noteRenderTileDelta()

    expect(snapshots).toEqual([
      {
        forceLocalTiles: false,
        localFallbackRevision: 0,
        renderTileRevision: 1,
      },
      {
        forceLocalTiles: false,
        localFallbackRevision: 0,
        renderTileRevision: 2,
      },
      {
        forceLocalTiles: true,
        localFallbackRevision: 1,
        renderTileRevision: 2,
      },
    ])
    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 1,
      renderTileRevision: 3,
    })
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
    expect(state.needsLocalCellInvalidation).toBe(false)
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

  it('resolves available warm remote tiles into the V3 preload lane without drawing them', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const visibleTileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    const warmTileId = packTileKey53({
      colTile: 1,
      dprBucket: 1,
      rowTile: 0,
      sheetOrdinal: 7,
    })
    const baseWarmTile = createRenderTile(warmTileId)
    const warmTile: GridRenderTile = {
      ...baseWarmTile,
      bounds: { colEnd: 255, colStart: 128, rowEnd: 31, rowStart: 0 },
      coord: {
        ...baseWarmTile.coord,
        colTile: 1,
        rowTile: 0,
      },
    }

    const state = runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(visibleTileId), warmTile]),
      }),
    )

    expect(state.renderTilePanes.map((pane) => pane.tile.tileId)).toEqual([visibleTileId])
    expect(state.preloadDataPanes.map((pane) => pane.tile.tileId)).toContain(warmTileId)
    expect(state.preloadDataPanes.map((pane) => pane.tile.tileId)).not.toContain(visibleTileId)
    expect(host.tiles.residency.getExact(warmTileId)?.packet).toMatchObject({
      tileId: warmTileId,
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

    expect(missing.residentDataPanes.map((pane) => pane.tile.tileId)).toEqual([tileId])
    expect(missing.residentBodyPane?.tile).toBe(ready.residentBodyPane?.tile)
    expect(missing.needsLocalCellInvalidation).toBe(false)
  })

  it('builds a local visible tile when remote tiles are unavailable before same-sheet retention exists', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([]),
      }),
    )

    expect(state.residentBodyPane?.tile.tileId).toBe(tileId)
    expect(state.residentBodyPane?.tile.rectCount).toBeGreaterThan(0)
    expect(state.needsLocalCellInvalidation).toBe(true)
    expect(state.residentDataPanes).toHaveLength(1)
    expect(state.tileReadiness.misses).toEqual([])
  })

  it('fills visible remote tile holes with local grid tiles', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
    })
    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileIds[0])]),
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        residentViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
      }),
    )

    expect(state.needsLocalCellInvalidation).toBe(true)
    expect(state.renderTilePanes.map((pane) => pane.tile.tileId)).toEqual(tileIds)
    expect(state.renderTilePanes[1]?.tile.rectCount).toBeGreaterThan(0)
    expect(state.tileReadiness).toMatchObject({
      exactHits: tileIds,
      misses: [],
    })
  })

  it('localizes a selected-cell tile when the remote tile is missing selected text', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
    })
    const state = runtime.resolve(
      createInput({
        engine: {
          ...LOCAL_EMPTY_ENGINE,
          getCell: (_sheetName, address) =>
            address === 'D53' ? createStringCellSnapshot('D53', 'Month 1') : createEmptyCellSnapshot(address),
        },
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileIds[0])]),
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
        residentViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
        selectedCell: [3, 52],
        selectedCellSnapshot: createStringCellSnapshot('D53', 'Month 1'),
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
      }),
    )

    expect(state.needsLocalCellInvalidation).toBe(true)
    expect(state.renderTilePanes.flatMap((pane) => pane.tile.textRuns.map((run) => run.text))).toContain('Month 1')
  })

  it('localizes a selected-cell tile when the remote tile still has text for a cleared cell', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
    })
    const staleRemoteTile: GridRenderTile = {
      ...createRenderTile(tileIds[0]),
      bounds: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
      coord: {
        colTile: 0,
        dprBucket: 1,
        paneKind: 'body',
        rowTile: 1,
        sheetId: 7,
        sheetOrdinal: 7,
      },
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 100,
          clipX: 300,
          clipY: 400,
          color: '#111827',
          col: 3,
          font: '400 12px Arial',
          fontSize: 12,
          height: 20,
          row: 52,
          strike: false,
          text: 'Month 1',
          underline: false,
          width: 100,
          x: 300,
          y: 400,
        },
      ],
    }
    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([staleRemoteTile]),
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
        residentViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
        selectedCell: [3, 52],
        selectedCellSnapshot: createEmptyCellSnapshot('D53'),
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
      }),
    )

    expect(state.needsLocalCellInvalidation).toBe(true)
    expect(state.renderTilePanes.flatMap((pane) => pane.tile.textRuns.map((run) => run.text))).not.toContain('Month 1')
  })

  it('localizes an active editor cell tile so remote text cannot show under the editor', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
    })
    const staleRemoteTile: GridRenderTile = {
      ...createRenderTile(tileIds[0]),
      bounds: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
      coord: {
        colTile: 0,
        dprBucket: 1,
        paneKind: 'body',
        rowTile: 1,
        sheetId: 7,
        sheetOrdinal: 7,
      },
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 100,
          clipX: 300,
          clipY: 400,
          color: '#111827',
          col: 3,
          font: '400 12px Arial',
          fontSize: 12,
          height: 20,
          row: 52,
          strike: false,
          text: 'Month 1',
          underline: false,
          width: 100,
          x: 300,
          y: 400,
        },
      ],
    }
    const state = runtime.resolve(
      createInput({
        editingCell: [3, 52],
        engine: {
          ...LOCAL_EMPTY_ENGINE,
          getCell: (_sheetName, address) =>
            address === 'D53' ? createStringCellSnapshot('D53', 'Month 1') : createEmptyCellSnapshot(address),
        },
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([staleRemoteTile]),
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
        residentViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
        selectedCell: [3, 52],
        selectedCellSnapshot: createStringCellSnapshot('D53', 'Month 1'),
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
      }),
    )

    expect(state.needsLocalCellInvalidation).toBe(true)
    expect(state.renderTilePanes.flatMap((pane) => pane.tile.textRuns.map((run) => run.text))).not.toContain('Month 1')
  })

  it('rebuilds visible remote tiles with missing grid payloads as local grid tiles', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
    })
    const emptyRemoteTile: GridRenderTile = {
      ...createRenderTile(tileIds[1]),
      bounds: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
      coord: {
        colTile: 0,
        dprBucket: 1,
        paneKind: 'body',
        rowTile: 1,
        sheetId: 7,
        sheetOrdinal: 7,
      },
      rectCount: 0,
      rectInstances: new Float32Array(),
      textCount: 0,
      textMetrics: new Float32Array(),
    }
    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileIds[0]), emptyRemoteTile]),
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        residentViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
      }),
    )

    expect(state.renderTilePanes.map((pane) => pane.tile.tileId)).toEqual(tileIds)
    expect(state.renderTilePanes[0]?.tile.rectCount).toBe(expectedGridBorderRectCount(state.renderTilePanes[0].tile.bounds))
    expect(state.renderTilePanes[1]?.tile).not.toBe(emptyRemoteTile)
    expect(state.renderTilePanes[1]?.tile.rectCount).toBeGreaterThan(0)
    expect(state.tileReadiness.misses).toEqual([])
  })

  it('rebuilds visible remote tiles with partial gridline payloads as local grid tiles', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
    })
    const partialRectCount = 12
    const partialRemoteTile: GridRenderTile = {
      ...createRenderTile(tileIds[1]),
      bounds: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
      coord: {
        colTile: 0,
        dprBucket: 1,
        paneKind: 'body',
        rowTile: 1,
        sheetId: 7,
        sheetOrdinal: 7,
      },
      rectCount: partialRectCount,
      rectInstances: createGridBorderRectInstances(partialRectCount),
    }
    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileIds[0]), partialRemoteTile]),
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        residentViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
      }),
    )

    expect(state.renderTilePanes.map((pane) => pane.tile.tileId)).toEqual(tileIds)
    expect(state.renderTilePanes[1]?.tile).not.toBe(partialRemoteTile)
    expect(state.renderTilePanes[1]?.tile.rectCount).toBeGreaterThan(partialRectCount)
    expect(state.tileReadiness.misses).toEqual([])
  })

  it('rebuilds visible text-only remote tiles so blank cells keep gridlines', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
    })
    const textOnlyRemoteTile: GridRenderTile = {
      ...createRenderTile(tileIds[1]),
      bounds: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
      coord: {
        colTile: 0,
        dprBucket: 1,
        paneKind: 'body',
        rowTile: 1,
        sheetId: 7,
        sheetOrdinal: 7,
      },
      rectCount: 0,
      rectInstances: new Float32Array(),
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 120,
          clipX: 0,
          clipY: 0,
          color: '#000000',
          col: 0,
          font: '12px sans-serif',
          fontSize: 12,
          height: 20,
          row: 32,
          strike: false,
          text: 'remote text without grid rects',
          underline: false,
          width: 120,
          wrap: false,
          x: 0,
          y: 0,
        },
      ],
    }
    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileIds[0]), textOnlyRemoteTile]),
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        residentViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
      }),
    )

    expect(state.renderTilePanes.map((pane) => pane.tile.tileId)).toEqual(tileIds)
    expect(state.renderTilePanes[1]?.tile).not.toBe(textOnlyRemoteTile)
    expect(state.renderTilePanes[1]?.tile.rectCount).toBeGreaterThan(0)
    expect(state.tileReadiness.misses).toEqual([])
  })

  it('rebuilds visible row-tile boundary text from authoritative cache when remote text is stale', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
    })
    const staleRemoteTile: GridRenderTile = {
      ...createRenderTile(tileIds[1]),
      bounds: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
      coord: {
        colTile: 0,
        dprBucket: 1,
        paneKind: 'body',
        rowTile: 1,
        sheetId: 7,
        sheetOrdinal: 7,
      },
      textCount: 0,
      textMetrics: new Float32Array(GRID_TEXT_METRIC_FLOAT_COUNT_V3),
      textRuns: [],
    }
    const engine: GridEngineLike = {
      ...LOCAL_EMPTY_ENGINE,
      getCell: (_sheetName, address) =>
        address === 'C33' ? createStringCellSnapshot(address, 'Annual software subscription') : createEmptyCellSnapshot(address),
    }
    const state = runtime.resolve(
      createInput({
        engine,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileIds[0]), staleRemoteTile]),
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        residentViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 42, rowStart: 11 },
      }),
    )

    const boundaryTile = state.renderTilePanes.find((pane) => pane.tile.bounds.rowStart === 32)?.tile
    expect(boundaryTile).toBeDefined()
    expect(boundaryTile).not.toBe(staleRemoteTile)
    expect(boundaryTile?.textRuns.some((run) => run.row === 32 && run.col === 2 && run.text === 'Annual software subscription')).toBe(true)
    expect(state.tileReadiness.misses).toEqual([])
  })

  it('rebuilds visible remote tiles when deleted local text makes remote text stale', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    if (tileId === undefined) {
      throw new Error('Expected a visible render tile key for the test viewport')
    }
    const staleRemoteTile: GridRenderTile = {
      ...createRenderTile(tileId),
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 120,
          clipX: 100,
          clipY: 80,
          color: '#000000',
          col: 1,
          font: '12px sans-serif',
          fontSize: 12,
          height: 20,
          row: 4,
          strike: false,
          text: 'deleted value',
          underline: false,
          width: 120,
          wrap: false,
          x: 100,
          y: 80,
        },
      ],
    }

    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([staleRemoteTile]),
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 10, rowStart: 0 },
      }),
    )

    expect(state.residentBodyPane?.tile).not.toBe(staleRemoteTile)
    expect(state.residentBodyPane?.tile.textRuns.some((run) => run.text === 'deleted value')).toBe(false)
  })

  it('requires coherent sheet id and ordinal for remote tiles when both are known', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]

    const sameOrdinalWrongSheetId = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileId, 99, 7)]),
        sheetId: 7,
        sheetOrdinal: 7,
      }),
    )
    const sameSheetIdWrongOrdinal = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileId, 7, 2)]),
        sheetId: 7,
        sheetOrdinal: 7,
      }),
    )

    expect(sameOrdinalWrongSheetId.residentDataPanes).toHaveLength(1)
    expect(sameOrdinalWrongSheetId.residentDataPanes[0]?.tile.coord).toMatchObject({ sheetId: 7, sheetOrdinal: 7 })
    expect(sameOrdinalWrongSheetId.residentDataPanes[0]?.tile.rectCount).toBeGreaterThan(0)
    expect(sameSheetIdWrongOrdinal.residentDataPanes).toHaveLength(1)
    expect(sameSheetIdWrongOrdinal.residentDataPanes[0]?.tile.coord).toMatchObject({ sheetId: 7, sheetOrdinal: 7 })
    expect(sameSheetIdWrongOrdinal.residentDataPanes[0]?.tile.rectCount).toBeGreaterThan(0)
  })

  it('does not retain local fallback panes as authoritative remote panes', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const first = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: true,
        renderTileSource: createRenderTileSource([]),
        sceneRevision: 1,
      }),
    )
    const second = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: true,
        renderTileSource: createRenderTileSource([]),
        sceneRevision: 2,
      }),
    )

    expect(first.residentDataPanes).toHaveLength(1)
    expect(second.residentDataPanes).toHaveLength(1)
    expect(second.residentDataPanes).not.toBe(first.residentDataPanes)
  })

  it('requests local cell invalidation only when local tiles are the active source', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        renderTileSource: undefined,
      }),
    )

    expect(state.residentBodyPane?.tile.coord.sheetId).toBe(7)
    expect(state.needsLocalCellInvalidation).toBe(true)
    expect(state.residentDataPanes).toHaveLength(1)
  })

  it('owns local cell invalidation and clears retained remote panes', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    let invalidationListener: (() => void) | null = null
    let mergeInvalidationListener: (() => void) | null = null
    let subscribedSheetName = ''
    let subscribedAddresses: readonly string[] = []
    let unsubscribed = false
    let mergeUnsubscribed = false
    const engine: GridEngineLike = {
      ...LOCAL_EMPTY_ENGINE,
      subscribeCells: (sheetName, addresses, listener) => {
        subscribedSheetName = sheetName
        subscribedAddresses = addresses
        invalidationListener = listener
        return () => {
          unsubscribed = true
        }
      },
      subscribeSheetChannel: (_sheetName, channel, listener) => {
        expect(channel).toBe('merges')
        mergeInvalidationListener = listener
        return () => {
          mergeUnsubscribed = true
        }
      },
    }
    const ready = runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileId)]),
      }),
    )
    const invalidations: string[] = []
    const unsubscribe = runtime.connectLocalCellInvalidation(
      {
        engine,
        needsLocalCellInvalidation: true,
        sheetName: 'Sheet1',
        visibleAddresses: ['A1', 'B2'],
      },
      () => invalidations.push('invalidated'),
    )

    expect(subscribedSheetName).toBe('Sheet1')
    expect(subscribedAddresses).toEqual(['A1', 'B2'])
    expect(
      runtime.resolve(
        createInput({
          gridRuntimeHost: host,
          renderTileSource: createRenderTileSource([]),
        }),
      ).residentBodyPane?.tile,
    ).toBe(ready.residentBodyPane?.tile)

    invalidationListener?.()
    mergeInvalidationListener?.()

    expect(invalidations).toEqual(['invalidated', 'invalidated'])
    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: true,
      localFallbackRevision: 2,
      renderTileRevision: 0,
    })
    expect(
      runtime.resolve(
        createInput({
          engine: LOCAL_EMPTY_ENGINE,
          gridRuntimeHost: host,
          renderTileSource: createRenderTileSource([]),
        }),
      ).residentDataPanes,
    ).not.toBe(ready.residentDataPanes)
    unsubscribe?.()
    expect(unsubscribed).toBe(true)
    expect(mergeUnsubscribed).toBe(true)
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
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
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
    expect(renderTileSource.captured()?.tileInterest).toMatchObject({
      axisSeqX: gridRuntimeHost.snapshot().axisSeqX,
      axisSeqY: gridRuntimeHost.snapshot().axisSeqY,
      freezeSeq: gridRuntimeHost.snapshot().freezeSeq,
      reason: 'scroll',
      sheetOrdinal: 7,
    })
    expect(renderTileSource.captured()?.tileInterest?.visibleTileKeys).toEqual([
      packTileKey53({
        colTile: 0,
        dprBucket: 2,
        rowTile: 0,
        sheetOrdinal: 7,
      }),
    ])
    expect(renderTileSource.captured()?.warmTileKeys).toContain(
      packTileKey53({
        colTile: 1,
        dprBucket: 2,
        rowTile: 0,
        sheetOrdinal: 7,
      }),
    )
    expect(renderTileSource.captured()?.tileInterest?.warmTileKeys).toContain(
      packTileKey53({
        colTile: 1,
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

  it('builds frozen-pane tile interest from disjoint body and frozen strips instead of the origin-to-body rectangle', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const gridRuntimeHost = createHost()
    const renderTileSource = createCapturingRenderTileSource()
    const sheetOrdinal = 7
    const dprBucket = 1

    runtime.connectRenderTileDeltas(
      {
        dprBucket,
        freezeCols: 1,
        freezeRows: 1,
        gridRuntimeHost,
        renderTileSource: renderTileSource.source,
        renderTileViewport: { colEnd: 767, colStart: 0, rowEnd: 50079, rowStart: 0 },
        residentViewport: { colEnd: 767, colStart: 512, rowEnd: 50079, rowStart: 49984 },
        sheetId: 7,
        sheetName: 'Sheet1',
      },
      () => {},
    )

    const visibleTileKeys = renderTileSource.captured()?.tileInterest?.visibleTileKeys ?? []

    expect(visibleTileKeys).toHaveLength(12)
    expect(visibleTileKeys).toContain(packTileKey53({ colTile: 4, dprBucket, rowTile: 1562, sheetOrdinal }))
    expect(visibleTileKeys).toContain(packTileKey53({ colTile: 4, dprBucket, rowTile: 0, sheetOrdinal }))
    expect(visibleTileKeys).toContain(packTileKey53({ colTile: 0, dprBucket, rowTile: 1562, sheetOrdinal }))
    expect(visibleTileKeys).toContain(packTileKey53({ colTile: 0, dprBucket, rowTile: 0, sheetOrdinal }))
    expect(visibleTileKeys).not.toContain(packTileKey53({ colTile: 4, dprBucket, rowTile: 100, sheetOrdinal }))
    expect(visibleTileKeys).not.toContain(packTileKey53({ colTile: 0, dprBucket, rowTile: 100, sheetOrdinal }))
  })

  it('materializes frozen deep-scroll panes from disjoint resident strips instead of the origin rectangle', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const gridRuntimeHost = createHost()
    const sheetOrdinal = 7
    const dprBucket = 1
    let getCellCount = 0
    const countingEngine: GridEngineLike = {
      ...LOCAL_EMPTY_ENGINE,
      getCell: (_sheetName, address) => {
        getCellCount += 1
        return createEmptyCellSnapshot(address)
      },
    }

    const state = runtime.resolve(
      createInput({
        dprBucket,
        engine: countingEngine,
        freezeCols: 1,
        freezeRows: 1,
        frozenColumnWidth: 100,
        frozenRowHeight: 20,
        gridRuntimeHost,
        renderTileSource: createRenderTileSource([]),
        renderTileViewport: { colEnd: 255, colStart: 0, rowEnd: 959, rowStart: 0 },
        residentViewport: { colEnd: 255, colStart: 128, rowEnd: 959, rowStart: 928 },
        visibleViewport: { colEnd: 255, colStart: 128, rowEnd: 959, rowStart: 928 },
      }),
    )

    const resolvedTileIds = new Set(state.renderTilePanes.map((pane) => pane.tile.tileId))

    expect(resolvedTileIds).toEqual(
      new Set([
        packTileKey53({ colTile: 1, dprBucket, rowTile: 29, sheetOrdinal }),
        packTileKey53({ colTile: 1, dprBucket, rowTile: 0, sheetOrdinal }),
        packTileKey53({ colTile: 0, dprBucket, rowTile: 29, sheetOrdinal }),
        packTileKey53({ colTile: 0, dprBucket, rowTile: 0, sheetOrdinal }),
      ]),
    )
    expect(state.renderTilePanes).toHaveLength(4)
    expect(getCellCount).toBeLessThanOrEqual(50_000)
    expect(resolvedTileIds).not.toContain(packTileKey53({ colTile: 0, dprBucket, rowTile: 10, sheetOrdinal }))
    expect(resolvedTileIds).not.toContain(packTileKey53({ colTile: 1, dprBucket, rowTile: 10, sheetOrdinal }))
  })

  it('applies render tile delta changes to the host-owned coordinator before React recomputes panes', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    const renderTileSource = createMutableRenderTileSource([
      {
        ...createRenderTile(tileId),
        version: {
          axisX: 1,
          axisY: 1,
          freeze: 0,
          styles: 1,
          text: 1,
          values: 1,
        },
      },
    ])

    runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
      }),
    )
    renderTileSource.setTile({
      ...createRenderTile(tileId),
      version: {
        axisX: 2,
        axisY: 3,
        freeze: 4,
        styles: 5,
        text: 6,
        values: 7,
      },
    })
    const listenerChanges: unknown[] = []
    const unsubscribe = runtime.connectRenderTileDeltas(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
        sheetId: 7,
        sheetName: 'Sheet1',
      },
      (change) => listenerChanges.push(change),
    )

    renderTileSource.emit({ changedTileIds: [tileId] })

    expect(host.tiles.residency.getExact(tileId)).toMatchObject({
      axisSeqX: 2,
      axisSeqY: 3,
      freezeSeq: 4,
      styleSeq: 5,
      textSeq: 6,
      valueSeq: 7,
    })
    expect(listenerChanges).toHaveLength(1)
    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 0,
      renderTileRevision: 1,
    })

    renderTileSource.emit({ invalidatedTileIds: [tileId] })
    expect(host.tiles.residency.getExact(tileId)).toBeNull()
    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 0,
      renderTileRevision: 2,
    })

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

  it('applies workbook delta damage to the host-owned tile coordinator', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const renderTileSource = createWorkbookDeltaSource()
    const listenerBatches: WorkbookDeltaBatchLikeV3[] = []
    const tileId = packTileKey53({
      colTile: 0,
      dprBucket: 1,
      rowTile: 0,
      sheetOrdinal: 7,
    })

    const unsubscribe = runtime.connectWorkbookDeltaDamage(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        sheetId: 7,
      },
      (batch) => listenerBatches.push(batch),
    )

    renderTileSource.emit(createWorkbookDeltaBatch())
    renderTileSource.emit(createWorkbookDeltaBatch({ seq: 1 }))
    renderTileSource.emit(createWorkbookDeltaBatch({ seq: 2, sheetId: 8, sheetOrdinal: 8 }))

    expect(listenerBatches.map((batch) => batch.seq)).toEqual([1])
    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 0,
      renderTileRevision: 1,
    })
    expect(
      host.tiles.reconcileInterest({
        axisSeqX: 1,
        axisSeqY: 1,
        cameraSeq: 1,
        freezeSeq: 1,
        pinnedTileKeys: [],
        reason: 'mutation',
        seq: 1,
        sheetId: 7,
        sheetOrdinal: 7,
        visibleTileKeys: [tileId],
        warmTileKeys: [],
      }).visibleDirtyTileKeys,
    ).toEqual([tileId])

    unsubscribe?.()
    expect(renderTileSource.unsubscribed()).toBe(true)
  })

  it('reconciles render tile connection lifecycles without React-owned resubscribe churn', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const renderTileSource = createCapturingRenderTileSource()
    let subscribedSheetName = ''
    let subscribedAddresses: readonly string[] = []
    let localUnsubscribeCount = 0
    const engine: GridEngineLike = {
      ...LOCAL_EMPTY_ENGINE,
      subscribeCells: (sheetName, addresses) => {
        subscribedSheetName = sheetName
        subscribedAddresses = addresses
        return () => {
          localUnsubscribeCount += 1
        }
      },
    }
    const visibleAddresses = ['A1', 'B2']

    runtime.syncConnections({
      dprBucket: 1,
      engine,
      gridRuntimeHost: host,
      needsLocalCellInvalidation: true,
      renderTileSource: renderTileSource.source,
      renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
      sheetId: 7,
      sheetName: 'Sheet1',
      visibleAddresses,
    })
    const firstSubscription = renderTileSource.captured()

    runtime.syncConnections({
      dprBucket: 1,
      engine,
      gridRuntimeHost: host,
      needsLocalCellInvalidation: true,
      renderTileSource: renderTileSource.source,
      renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
      sheetId: 7,
      sheetName: 'Sheet1',
      visibleAddresses: [...visibleAddresses],
    })

    expect(renderTileSource.captured()).toBe(firstSubscription)
    expect(renderTileSource.unsubscribed()).toBe(false)
    expect(renderTileSource.subscribeCount()).toBe(1)
    expect(subscribedSheetName).toBe('Sheet1')
    expect(subscribedAddresses).toBe(visibleAddresses)
    expect(localUnsubscribeCount).toBe(0)

    runtime.syncConnections({
      dprBucket: 1,
      engine,
      gridRuntimeHost: host,
      needsLocalCellInvalidation: true,
      renderTileSource: renderTileSource.source,
      renderTileViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
      sheetId: 7,
      sheetName: 'Sheet1',
      visibleAddresses,
    })

    expect(renderTileSource.unsubscribed()).toBe(true)
    expect(renderTileSource.subscribeCount()).toBe(2)
    expect(renderTileSource.unsubscribeCount()).toBe(1)
    expect(renderTileSource.captured()).not.toBe(firstSubscription)
    expect(localUnsubscribeCount).toBe(0)

    runtime.disconnectConnections()

    expect(renderTileSource.unsubscribeCount()).toBe(2)
    expect(localUnsubscribeCount).toBe(1)
  })

  it('matches workbook delta damage by sheet ordinal when sheet id differs from order', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const renderTileSource = createWorkbookDeltaSource()
    const tileId = packTileKey53({
      colTile: 0,
      dprBucket: 1,
      rowTile: 0,
      sheetOrdinal: 2,
    })
    const listenerBatches: WorkbookDeltaBatchLikeV3[] = []

    const unsubscribe = runtime.connectWorkbookDeltaDamage(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        sheetId: 99,
        sheetOrdinal: 2,
      },
      (batch) => listenerBatches.push(batch),
    )

    renderTileSource.emit(createWorkbookDeltaBatch({ seq: 1, sheetId: 99, sheetOrdinal: 2 }))
    renderTileSource.emit(createWorkbookDeltaBatch({ seq: 2, sheetId: 7, sheetOrdinal: 7 }))

    expect(listenerBatches.map((batch) => batch.seq)).toEqual([1])
    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 0,
      renderTileRevision: 1,
    })
    expect(
      host.tiles.reconcileInterest({
        axisSeqX: 1,
        axisSeqY: 1,
        cameraSeq: 1,
        freezeSeq: 1,
        pinnedTileKeys: [],
        reason: 'mutation',
        seq: 1,
        sheetId: 99,
        sheetOrdinal: 2,
        visibleTileKeys: [tileId],
        warmTileKeys: [],
      }).visibleDirtyTileKeys,
    ).toEqual([tileId])

    unsubscribe?.()
  })

  it('uses local tiles for worker-authoritative damage until fresh render tiles arrive', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    const remoteTile: GridRenderTile = {
      ...createRenderTile(tileId),
      textCount: 1,
      textRuns: [
        {
          col: 0,
          row: 0,
          text: 'stale remote text',
          x: 0,
          y: 0,
          width: 80,
          height: 20,
          clipX: 0,
          clipY: 0,
          clipWidth: 80,
          clipHeight: 20,
          font: '12px sans-serif',
          fontSize: 12,
          color: '#000000',
          underline: false,
          strike: false,
        },
      ],
    }
    const renderTileSource = createMutableWorkbookDeltaRenderTileSource([remoteTile])

    const initial = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
      }),
    )
    expect(initial.residentBodyPane?.tile.textRuns.some((run) => run.text === 'stale remote text')).toBe(false)

    runtime.connectWorkbookDeltaDamage(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        sheetId: 7,
      },
      () => undefined,
    )
    runtime.connectRenderTileDeltas(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
        sheetId: 7,
        sheetName: 'Sheet1',
      },
      () => undefined,
    )
    renderTileSource.emitWorkbookDelta({
      ...createWorkbookDeltaBatch(),
      source: 'workerAuthoritative',
    })

    const fallback = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: runtime.snapshotBridgeState().forceLocalTiles,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
      }),
    )
    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: true,
      localFallbackRevision: 1,
      renderTileRevision: 1,
    })
    expect(fallback.residentBodyPane?.tile.textRuns.some((run) => run.text === 'stale remote text')).toBe(false)
    expect(fallback.residentBodyPane?.tile.textRuns).toEqual([])

    const freshRemoteTile: GridRenderTile = {
      ...remoteTile,
      lastBatchId: 2,
      lastCameraSeq: 2,
      textRuns: [],
      textCount: 0,
      version: {
        ...remoteTile.version,
        text: remoteTile.version.text + 1,
        values: remoteTile.version.values + 1,
      },
    }
    renderTileSource.setTile(freshRemoteTile)
    renderTileSource.emitRenderTileDelta({ changedTileIds: [tileId] })

    const refreshed = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: runtime.snapshotBridgeState().forceLocalTiles,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
      }),
    )
    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 1,
      renderTileRevision: 2,
    })
    expect(refreshed.residentBodyPane?.tile.textRuns).toEqual([])
  })

  it('rebuilds resident local fallback tiles when the projected workbook revision advances', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    let projectedRevision = 0
    let greenFillVisible = false
    const engineWithChangingStyle: GridEngineLike = {
      getCell: (_sheetName, address) => ({
        ...(address === 'E6' && greenFillVisible ? { styleId: 'style-green' } : {}),
        ...createEmptyCellSnapshot(address),
      }),
      getCellStyle: (styleId) => (styleId === 'style-green' ? { id: 'style-green', fill: { backgroundColor: '#00ff00' } } : undefined),
      getRenderRevisionSnapshot: () => ({
        authoritativeRevision: projectedRevision,
        projectedRevision,
        tileSceneCameraSeq: null,
        tileSceneRevision: null,
      }),
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }
    const renderTileSource = createRenderTileSource([])

    const initial = runtime.resolve(
      createInput({
        engine: engineWithChangingStyle,
        gridRuntimeHost: host,
        renderTileSource,
        sceneRevision: 0,
      }),
    )
    expect(hasOpaqueGreenFillRect(initial.residentBodyPane?.tile)).toBe(false)

    projectedRevision = 1
    greenFillVisible = true

    const refreshed = runtime.resolve(
      createInput({
        engine: engineWithChangingStyle,
        gridRuntimeHost: host,
        renderTileSource,
        sceneRevision: 0,
      }),
    )

    expect(refreshed.residentBodyPane?.tile).not.toBe(initial.residentBodyPane?.tile)
    expect(hasOpaqueGreenFillRect(refreshed.residentBodyPane?.tile)).toBe(true)
  })

  it('keeps local tiles through stale render tile deltas until the renderer batch catches up', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    if (tileId === undefined) {
      throw new Error('Expected a visible render tile key')
    }
    const staleRemoteTile: GridRenderTile = {
      ...createRenderTile(tileId),
      lastBatchId: 9,
      lastCameraSeq: 9,
      textCount: 1,
      textRuns: [
        {
          col: 0,
          row: 0,
          text: 'deleted text from stale tile',
          x: 0,
          y: 0,
          width: 140,
          height: 20,
          clipX: 0,
          clipY: 0,
          clipWidth: 140,
          clipHeight: 20,
          font: '12px sans-serif',
          fontSize: 12,
          color: '#000000',
          underline: false,
          strike: false,
        },
      ],
      version: {
        axisX: 9,
        axisY: 9,
        freeze: 0,
        styles: 9,
        text: 9,
        values: 9,
      },
    }
    const renderTileSource = createMutableWorkbookDeltaRenderTileSource([staleRemoteTile])

    runtime.connectWorkbookDeltaDamage(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        sheetId: 7,
      },
      () => undefined,
    )
    runtime.connectRenderTileDeltas(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
        sheetId: 7,
        sheetName: 'Sheet1',
      },
      () => undefined,
    )

    renderTileSource.emitWorkbookDelta({
      ...createWorkbookDeltaBatch({ seq: 10 }),
      source: 'localOptimistic',
    })
    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: true,
      localFallbackRevision: 1,
      renderTileRevision: 1,
    })

    renderTileSource.setTile(staleRemoteTile)
    renderTileSource.emitRenderTileDelta({ batchId: 9, cameraSeq: 9, changedTileIds: [tileId] })

    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: true,
      localFallbackRevision: 1,
      renderTileRevision: 2,
    })
    const staleDeltaState = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: runtime.snapshotBridgeState().forceLocalTiles,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
      }),
    )
    expect(staleDeltaState.residentBodyPane?.tile.textRuns.some((run) => run.text === 'deleted text from stale tile')).toBe(false)
    expect(staleDeltaState.residentBodyPane?.tile.textRuns).toEqual([])

    const freshRemoteTile: GridRenderTile = {
      ...staleRemoteTile,
      lastBatchId: 10,
      lastCameraSeq: 10,
      textCount: 0,
      textRuns: [],
      version: {
        ...staleRemoteTile.version,
        text: 10,
        values: 10,
      },
    }
    renderTileSource.setTile(freshRemoteTile)
    renderTileSource.emitRenderTileDelta({ batchId: 10, cameraSeq: 10, changedTileIds: [tileId] })

    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 1,
      renderTileRevision: 3,
    })
    const caughtUpState = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: runtime.snapshotBridgeState().forceLocalTiles,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
      }),
    )
    expect(caughtUpState.residentBodyPane?.tile.textRuns).toEqual([])
  })

  it('keeps clean remote tiles resident when local fallback only needs dirty tiles', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const [dirtyTileId, cleanTileId] = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
    })
    if (dirtyTileId === undefined || cleanTileId === undefined) {
      throw new Error('Expected two render tile keys for the test viewport')
    }
    const cleanRemoteTile: GridRenderTile = {
      ...createRenderTile(cleanTileId),
      bounds: { colEnd: 255, colStart: 128, rowEnd: 31, rowStart: 0 },
      coord: {
        colTile: 1,
        dprBucket: 1,
        paneKind: 'body',
        rowTile: 0,
        sheetId: 7,
        sheetOrdinal: 7,
      },
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 80,
          clipX: 0,
          clipY: 0,
          color: '#000000',
          col: 128,
          font: '12px sans-serif',
          fontSize: 12,
          height: 20,
          row: 0,
          strike: false,
          text: 'clean remote text',
          underline: false,
          width: 80,
          wrap: false,
          x: 0,
          y: 0,
        },
      ],
    }
    const renderTileSource = createRenderTileSource([createRenderTile(dirtyTileId), cleanRemoteTile])
    const cleanRemoteTextAddress = formatAddress(0, 128)
    const cleanRemoteEngine: GridEngineLike = {
      ...LOCAL_EMPTY_ENGINE,
      getCell: (_sheetName, address) =>
        address === cleanRemoteTextAddress
          ? createStringCellSnapshot(cleanRemoteTextAddress, 'clean remote text')
          : createEmptyCellSnapshot(address),
    }

    host.tiles.applyWorkbookDelta(createWorkbookDeltaBatch(), { dprBucket: 1 })
    const fallback = runtime.resolve(
      createInput({
        engine: cleanRemoteEngine,
        forceLocalTiles: true,
        gridRuntimeHost: host,
        renderTileSource,
        renderTileViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
        residentViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
        visibleViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
      }),
    )

    expect(fallback.renderTilePanes.map((pane) => pane.tile.tileId)).toEqual([dirtyTileId, cleanTileId])
    expect(fallback.renderTilePanes[0]?.tile.textRuns).toEqual([])
    expect(fallback.renderTilePanes[0]?.tile.dirtyLocalRows).toEqual(new Uint32Array([0, 0]))
    expect(fallback.renderTilePanes[0]?.tile.dirtyLocalCols).toEqual(new Uint32Array([0, 0]))
    expect(fallback.renderTilePanes[1]?.tile.textRuns[0]?.text).toBe('clean remote text')
  })

  it('promotes an offscreen dirty warm tile to local geometry before it first becomes visible', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const [visibleTileId, warmTileId] = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
    })
    if (visibleTileId === undefined || warmTileId === undefined) {
      throw new Error('Expected visible and warm render tile keys for the test viewport')
    }
    const staleWarmTile: GridRenderTile = {
      ...createRenderTile(warmTileId),
      bounds: { colEnd: 255, colStart: 128, rowEnd: 31, rowStart: 0 },
      coord: {
        colTile: 1,
        dprBucket: 1,
        paneKind: 'body',
        rowTile: 0,
        sheetId: 7,
        sheetOrdinal: 7,
      },
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 80,
          clipX: 0,
          clipY: 0,
          color: '#000000',
          col: 128,
          font: '12px sans-serif',
          fontSize: 12,
          height: 20,
          row: 0,
          strike: false,
          text: 'stale warm remote text',
          underline: false,
          width: 80,
          wrap: false,
          x: 0,
          y: 0,
        },
      ],
    }
    const renderTileSource = createRenderTileSource([createRenderTile(visibleTileId), staleWarmTile])
    host.tiles.applyWorkbookDelta(
      createWorkbookDeltaBatch({
        dirty: {
          axisX: new Uint32Array([128, 128, DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array(),
        },
      }),
      { dprBucket: 1 },
    )

    const offscreen = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: true,
        gridRuntimeHost: host,
        renderTileSource,
        renderTileViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
        residentViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
      }),
    )

    const offscreenWarmPane = offscreen.renderTilePanes.find((pane) => pane.tile.tileId === warmTileId)
    expect(offscreenWarmPane?.tile.textRuns[0]?.text).toBe('stale warm remote text')
    expect(host.tiles.dirtyTiles.getUnconsumedMask(warmTileId)).not.toBe(0)

    const visible = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: true,
        gridRuntimeHost: host,
        renderTileSource,
        renderTileViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
        residentViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
        visibleViewport: { colEnd: 255, colStart: 128, rowEnd: 31, rowStart: 0 },
      }),
    )

    const visibleWarmPane = visible.renderTilePanes.find((pane) => pane.tile.tileId === warmTileId)
    expect(visibleWarmPane?.tile.textRuns).toEqual([])
    expect(visibleWarmPane?.tile.dirtyLocalCols).toEqual(new Uint32Array([0, 127]))
    expect(visible.tileReadiness.visibleDirtyTileKeys).toContain(warmTileId)
    expect(host.tiles.dirtyTiles.getUnconsumedMask(warmTileId)).toBe(0)
  })

  it('localizes dirty warm preload tiles before stale remote text can be staged', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const [visibleTileId, warmTileId] = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
    })
    if (visibleTileId === undefined || warmTileId === undefined) {
      throw new Error('Expected visible and warm render tile keys for the test viewport')
    }
    const staleWarmTile: GridRenderTile = {
      ...createRenderTile(warmTileId),
      bounds: { colEnd: 255, colStart: 128, rowEnd: 31, rowStart: 0 },
      coord: {
        colTile: 1,
        dprBucket: 1,
        paneKind: 'body',
        rowTile: 0,
        sheetId: 7,
        sheetOrdinal: 7,
      },
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 80,
          clipX: 0,
          clipY: 0,
          color: '#000000',
          col: 128,
          font: '12px sans-serif',
          fontSize: 12,
          height: 20,
          row: 0,
          strike: false,
          text: 'stale warm remote text',
          underline: false,
          width: 80,
          wrap: false,
          x: 0,
          y: 0,
        },
      ],
    }
    const renderTileSource = createRenderTileSource([createRenderTile(visibleTileId), staleWarmTile])
    host.tiles.applyWorkbookDelta(
      createWorkbookDeltaBatch({
        dirty: {
          axisX: new Uint32Array([128, 128, DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array(),
        },
      }),
      { dprBucket: 1 },
    )

    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource,
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
        residentViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
      }),
    )

    const warmPreloadPane = state.preloadDataPanes.find((pane) => pane.tile.tileId === warmTileId)
    expect(warmPreloadPane).toBeDefined()
    expect(warmPreloadPane?.tile).not.toBe(staleWarmTile)
    expect(warmPreloadPane?.tile.textRuns.some((run) => run.text === 'stale warm remote text')).toBe(false)
    expect(warmPreloadPane?.tile.dirtyLocalCols).toEqual(new Uint32Array([0, 127]))
    expect(host.tiles.dirtyTiles.getUnconsumedMask(warmTileId)).not.toBe(0)
  })

  it('preserves dirty spans when local fallback has no remote render tile source', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    if (tileId === undefined) {
      throw new Error('Expected a visible render tile key for the test viewport')
    }

    host.tiles.applyWorkbookDelta(
      createWorkbookDeltaBatch({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([6, 6, 5, 5, DirtyMaskV3.Value | DirtyMaskV3.Text]),
        },
      }),
      { dprBucket: 1 },
    )
    const fallback = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: true,
        gridRuntimeHost: host,
        renderTileSource: undefined,
      }),
    )

    expect(fallback.renderTilePanes[0]?.tile.tileId).toBe(tileId)
    expect(fallback.renderTilePanes[0]?.tile.dirtyLocalRows).toEqual(new Uint32Array([6, 6]))
    expect(fallback.renderTilePanes[0]?.tile.dirtyLocalCols).toEqual(new Uint32Array([5, 5]))
    expect(fallback.renderTilePanes[0]?.tile.dirtyMasks).toEqual(new Uint32Array([DirtyMaskV3.Value | DirtyMaskV3.Text]))
  })

  it('materializes every missing resident tile when the remote source has not caught up', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const renderTileViewport = { colEnd: 127, colStart: 0, rowEnd: 95, rowStart: 0 }
    const expectedTileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: renderTileViewport,
    })

    const fallback = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([]),
        renderTileViewport,
        residentViewport: renderTileViewport,
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
      }),
    )

    expect(fallback.renderTilePanes.map((pane) => pane.tile.tileId)).toEqual(expectedTileIds)
  })

  it('materializes a visible remote tile locally when its text payload is incomplete', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    if (tileId === undefined) {
      throw new Error('Expected a visible render tile key for the test viewport')
    }
    const remoteTileWithoutText = createRenderTile(tileId)
    const engineWithVisibleText: GridEngineLike = {
      getCell: (_sheetName, address) =>
        address === 'A15' ? createStringCellSnapshot('A15', 'Amortization Schedule Examples') : createEmptyCellSnapshot(address),
      getCellStyle: () => undefined,
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }

    const fallback = runtime.resolve(
      createInput({
        engine: engineWithVisibleText,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([remoteTileWithoutText]),
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 34, rowStart: 3 },
      }),
    )

    expect(fallback.renderTilePanes[0]?.tile.tileId).toBe(tileId)
    expect(fallback.renderTilePanes[0]?.tile).not.toBe(remoteTileWithoutText)
    expect(
      fallback.renderTilePanes[0]?.tile.textRuns.some((run) => run.row === 14 && run.col === 0 && run.text.includes('Amortization')),
    ).toBe(true)
  })

  it('caches visible text freshness for clean remote tiles across selection-only resolves', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    if (tileId === undefined) {
      throw new Error('Expected a visible render tile key for the test viewport')
    }
    const remoteTileWithMatchingText: GridRenderTile = {
      ...createRenderTile(tileId),
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 120,
          clipX: 0,
          clipY: 280,
          color: '#000000',
          col: 0,
          font: '12px sans-serif',
          fontSize: 12,
          height: 20,
          row: 14,
          strike: false,
          text: 'Amortization Schedule Examples',
          underline: false,
          width: 120,
          wrap: false,
          x: 0,
          y: 280,
        },
      ],
    }
    let getCellCallCount = 0
    const engineWithMatchingVisibleText: GridEngineLike = {
      getCell: (_sheetName, address) => {
        getCellCallCount += 1
        return address === 'A15' ? createStringCellSnapshot('A15', 'Amortization Schedule Examples') : createEmptyCellSnapshot(address)
      },
      getCellStyle: () => undefined,
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }
    const renderTileSource = createRenderTileSource([remoteTileWithMatchingText])
    const visibleViewport = { colEnd: 10, colStart: 0, rowEnd: 34, rowStart: 3 }

    const initial = runtime.resolve(
      createInput({
        engine: engineWithMatchingVisibleText,
        gridRuntimeHost: host,
        renderTileSource,
        selectedCell: [1, 1],
        selectedCellSnapshot: createEmptyCellSnapshot('B2'),
        visibleViewport,
      }),
    )
    const callsAfterInitialResolve = getCellCallCount

    expect(initial.residentBodyPane?.tile).toBe(remoteTileWithMatchingText)
    expect(callsAfterInitialResolve).toBeGreaterThan(0)

    const selectionOnly = runtime.resolve(
      createInput({
        engine: engineWithMatchingVisibleText,
        gridRuntimeHost: host,
        renderTileSource,
        selectedCell: [2, 2],
        selectedCellSnapshot: createEmptyCellSnapshot('C3'),
        visibleViewport,
      }),
    )

    expect(selectionOnly.residentBodyPane?.tile).toBe(remoteTileWithMatchingText)
    expect(getCellCallCount).toBe(callsAfterInitialResolve)

    runtime.resolve(
      createInput({
        engine: engineWithMatchingVisibleText,
        gridRuntimeHost: host,
        renderTileSource,
        sceneRevision: 2,
        selectedCell: [3, 3],
        selectedCellSnapshot: createEmptyCellSnapshot('D4'),
        visibleViewport,
      }),
    )

    expect(getCellCallCount).toBeGreaterThan(callsAfterInitialResolve)
  })

  it('rechecks visible remote tile text when the authoritative workbook revision changes', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    if (tileId === undefined) {
      throw new Error('Expected a visible render tile key for the test viewport')
    }
    const staleRemoteTile = createRenderTile(tileId)
    let authoritativeRevision = 1
    let visibleText = ''
    const engineWithChangingVisibleText: GridEngineLike = {
      getCell: (_sheetName, address) =>
        address === 'A1' && visibleText ? createStringCellSnapshot('A1', visibleText) : createEmptyCellSnapshot(address),
      getCellStyle: () => undefined,
      getRenderRevisionSnapshot: () => ({
        authoritativeRevision,
        projectedRevision: authoritativeRevision,
        tileSceneCameraSeq: null,
        tileSceneRevision: null,
      }),
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }
    const renderTileSource = createRenderTileSource([staleRemoteTile])

    const initial = runtime.resolve(
      createInput({
        engine: engineWithChangingVisibleText,
        gridRuntimeHost: host,
        renderTileSource,
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 10, rowStart: 0 },
      }),
    )
    expect(initial.residentBodyPane?.tile).toBe(staleRemoteTile)

    authoritativeRevision = 2
    visibleText = 'Prepaid Expense Template'

    const refreshed = runtime.resolve(
      createInput({
        engine: engineWithChangingVisibleText,
        gridRuntimeHost: host,
        renderTileSource,
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 10, rowStart: 0 },
      }),
    )

    expect(refreshed.residentBodyPane?.tile).not.toBe(staleRemoteTile)
    expect(refreshed.residentBodyPane?.tile.textRuns.some((run) => run.row === 0 && run.col === 0 && run.text === visibleText)).toBe(true)
  })

  it('rechecks visible remote tile text when the local workbook revision changes', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    if (tileId === undefined) {
      throw new Error('Expected a visible render tile key for the test viewport')
    }
    const remoteTile = createRenderTile(tileId)
    let localRevision = 0
    let visibleText = ''
    let getCellCallCount = 0
    const engineWithLocalVisibleText: GridEngineLike = {
      getCell: (_sheetName, address) => {
        getCellCallCount += 1
        return address === 'B25' && visibleText ? createStringCellSnapshot('B25', visibleText) : createEmptyCellSnapshot(address)
      },
      getCellStyle: () => undefined,
      getRenderRevisionSnapshot: () => ({
        authoritativeRevision: 1,
        localRevision,
        projectedRevision: 1,
        tileSceneCameraSeq: null,
        tileSceneRevision: null,
      }),
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }
    const renderTileSource = createRenderTileSource([remoteTile])

    const initial = runtime.resolve(
      createInput({
        engine: engineWithLocalVisibleText,
        gridRuntimeHost: host,
        renderTileSource,
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 30, rowStart: 0 },
      }),
    )
    const callsAfterInitialResolve = getCellCallCount
    expect(initial.residentBodyPane?.tile).toBe(remoteTile)

    runtime.resolve(
      createInput({
        engine: engineWithLocalVisibleText,
        gridRuntimeHost: host,
        renderTileSource,
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 30, rowStart: 0 },
      }),
    )
    expect(getCellCallCount).toBe(callsAfterInitialResolve)

    localRevision = 1
    visibleText = 'abcdef'

    const refreshed = runtime.resolve(
      createInput({
        engine: engineWithLocalVisibleText,
        gridRuntimeHost: host,
        renderTileSource,
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 30, rowStart: 0 },
      }),
    )

    expect(refreshed.residentBodyPane?.tile).not.toBe(remoteTile)
    expect(refreshed.residentBodyPane?.tile.textRuns.some((run) => run.row === 24 && run.col === 1 && run.text === visibleText)).toBe(true)
  })

  it('checks resident rendered rows for local text even when the logical visible window is shorter', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    if (tileId === undefined) {
      throw new Error('Expected a visible render tile key for the test viewport')
    }
    const remoteTile = createRenderTile(tileId)
    const engineWithResidentText: GridEngineLike = {
      getCell: (_sheetName, address) =>
        address === 'B25' ? createStringCellSnapshot('B25', 'click-away text') : createEmptyCellSnapshot(address),
      getCellStyle: () => undefined,
      getRenderRevisionSnapshot: () => ({
        authoritativeRevision: 1,
        localRevision: 1,
        projectedRevision: 1,
        tileSceneCameraSeq: 1,
        tileSceneRevision: 1,
      }),
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }
    const renderTileSource = createRenderTileSource([remoteTile])

    const refreshed = runtime.resolve(
      createInput({
        engine: engineWithResidentText,
        gridRuntimeHost: host,
        renderTileSource,
        residentViewport: { colEnd: 10, colStart: 0, rowEnd: 30, rowStart: 0 },
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 10, rowStart: 0 },
      }),
    )

    expect(refreshed.residentBodyPane?.tile).not.toBe(remoteTile)
    expect(refreshed.residentBodyPane?.tile.textRuns.some((run) => run.row === 24 && run.col === 1 && run.text === 'click-away text')).toBe(
      true,
    )
  })

  it('reuses remote static rect buffers for text-only local dirty tiles', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    if (tileId === undefined) {
      throw new Error('Expected a visible render tile key for the test viewport')
    }
    const baseTile = createRenderTile(tileId)
    const rectCount = expectedGridBorderRectCount(baseTile.bounds)
    const rectInstances = createGridBorderRectInstances(rectCount)
    rectInstances[0] = 42
    const remoteTile: GridRenderTile = {
      ...baseTile,
      rectCount,
      rectInstances,
    }

    host.tiles.applyWorkbookDelta(
      createWorkbookDeltaBatch({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([6, 6, 5, 5, DirtyMaskV3.Value | DirtyMaskV3.Text]),
        },
      }),
      { dprBucket: 1 },
    )
    const fallback = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: true,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([remoteTile]),
      }),
    )

    expect(fallback.renderTilePanes[0]?.tile.tileId).toBe(tileId)
    expect(fallback.renderTilePanes[0]?.tile.rectCount).toBe(rectCount)
    expect(fallback.renderTilePanes[0]?.tile.rectInstances).toBe(rectInstances)
    expect(fallback.renderTilePanes[0]?.tile.dirty?.rectSpans).toEqual([])
    expect(fallback.renderTilePanes[0]?.tile.dirty?.textSpans).toEqual([])
  })

  it('reuses resident static rect buffers for text-only local dirty tiles when source misses', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    if (tileId === undefined) {
      throw new Error('Expected a visible render tile key for the test viewport')
    }
    const baseTile = createRenderTile(tileId)
    const rectCount = expectedGridBorderRectCount(baseTile.bounds)
    const rectInstances = createGridBorderRectInstances(rectCount)
    rectInstances[0] = 77
    const residentTile: GridRenderTile = {
      ...baseTile,
      rectCount,
      rectInstances,
    }

    runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([residentTile]),
      }),
    )
    host.tiles.applyWorkbookDelta(
      createWorkbookDeltaBatch({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([6, 6, 5, 5, DirtyMaskV3.Value | DirtyMaskV3.Text]),
        },
      }),
      { dprBucket: 1 },
    )
    const fallback = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: true,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([]),
      }),
    )

    expect(fallback.renderTilePanes[0]?.tile.tileId).toBe(tileId)
    expect(fallback.renderTilePanes[0]?.tile.rectCount).toBe(rectCount)
    expect(fallback.renderTilePanes[0]?.tile.rectInstances).toBe(rectInstances)
    expect(fallback.renderTilePanes[0]?.tile.dirty?.rectSpans).toEqual([])
  })

  it('applies local optimistic workbook deltas after higher authoritative seqs on the same sheet', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const renderTileSource = createWorkbookDeltaSource()

    runtime.connectWorkbookDeltaDamage(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        sheetId: 7,
      },
      () => undefined,
    )

    renderTileSource.emit({
      ...createWorkbookDeltaBatch({ seq: 10 }),
      source: 'workerAuthoritative',
    })

    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: true,
      localFallbackRevision: 1,
      renderTileRevision: 1,
    })

    renderTileSource.emit({
      ...createWorkbookDeltaBatch({ seq: 1 }),
      source: 'localOptimistic',
    })

    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: true,
      localFallbackRevision: 2,
      renderTileRevision: 2,
    })
  })

  it('uses local fallback for local optimistic axis damage while waiting for worker render tiles', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const renderTileSource = createWorkbookDeltaSource()

    runtime.connectWorkbookDeltaDamage(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        sheetId: 7,
      },
      () => undefined,
    )

    renderTileSource.emit({
      dirty: {
        axisX: new Uint32Array([1, 1, DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
        axisY: new Uint32Array(),
        cellRanges: new Uint32Array(),
      },
      seq: 1,
      sheetId: 7,
      sheetOrdinal: 7,
      source: 'localOptimistic',
    })

    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: true,
      localFallbackRevision: 1,
      renderTileRevision: 1,
    })
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
    expect(switched.residentDataPanes).toHaveLength(1)
    expect(switched.residentDataPanes[0]?.tile.coord).toMatchObject({ sheetId: 8, sheetOrdinal: 8 })
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
